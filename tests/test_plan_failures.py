"""
計画の組み立てが失敗したときの扱い
==================================
plan_row() は既存課題の照合のため API を呼ぶ。ここで失敗しても、
行の処理を続けられること・集計が壊れないことを固定する。
"""

import excel_to_backlog as etb
from backlog_client import BacklogAPIError
from conftest import FakeBacklog
from run_log import RunLog, load_completed
from summary_index import SummaryIndex


class BrokenIndex(FakeBacklog):
    """件名索引の構築に失敗するクライアント。"""

    def get_issues(self, project_id, params=None):
        raise BacklogAPIError("一覧の取得に失敗しました", status=500)


class BrokenLookup(FakeBacklog):
    """issueKey での照合に失敗するクライアント。"""

    def get_issue(self, issue_id_or_key):
        raise BacklogAPIError("課題の取得に失敗しました", status=500)


def upsert_cfg(source_cfg, **upsert):
    return source_cfg(
        ["件名", "Backlog番号"], [["課題A", "DEMO-1"], ["課題B", "DEMO-2"]],
        upsert={"enabled": True, **upsert},
    )


class TestLookupFailure:
    """
    照合の失敗で UnboundLocalError にならないこと。
    map_row を try の中へ移した際、例外処理が params を参照したまま残っており、
    トレースバックで実行全体が停止しサマリーも出なくなっていた。
    """

    def test_索引構築の失敗はエラーとして計上される(self, source_cfg, master):
        cfg = upsert_cfg(source_cfg, match_summary=True)
        client = BrokenIndex()

        counts = etb.process_source(
            cfg, client, master, dry_run=False,
            summary_index=SummaryIndex(client, master.project_id),
        )

        assert counts["error"] == 2          # 例外は外へ出ない
        assert counts["created"] == 0

    def test_課題取得の失敗もエラーとして計上される(self, source_cfg, master):
        cfg = upsert_cfg(source_cfg, key_col="Backlog番号")

        counts = etb.process_source(cfg, BrokenLookup(), master, dry_run=False)

        assert counts["error"] == 2

    def test_失敗した行がログに記録される(self, source_cfg, master, tmp_path):
        import csv
        cfg = upsert_cfg(source_cfg, key_col="Backlog番号")
        log_path = tmp_path / "run.csv"

        with RunLog(log_path) as log:
            etb.process_source(cfg, BrokenLookup(), master, dry_run=False, run_log=log)

        with open(log_path, encoding="utf-8-sig", newline="") as f:
            records = list(csv.DictReader(f))
        assert [r["action"] for r in records] == ["error", "error"]
        # 件名は確定していないため空。前の行の件名が漏れないこと
        assert all(r["summary"] == "" for r in records)


class TestNoSleepWithoutRequest:
    """通信していない行で待機しないこと。"""

    def _timed(self, monkeypatch):
        slept = []
        monkeypatch.setattr(etb.time, "sleep", lambda s: slept.append(s))
        return slept

    def test_再開スキップでは待機しない(self, source_cfg, master, tmp_path, monkeypatch):
        """
        5,000 行を再開すると、待たなくてよい行で 25 分待つことになっていた。
        """
        cfg = source_cfg(["件名"], [[f"課題{i}"] for i in range(1, 11)],
                         upsert={"enabled": True, "match_summary": True})
        log_path = tmp_path / "run.csv"
        first = FakeBacklog()
        with RunLog(log_path) as log:
            etb.process_source(
                cfg, first, master, dry_run=False, run_log=log,
                summary_index=SummaryIndex(first, master.project_id),
            )

        slept = self._timed(monkeypatch)
        second = FakeBacklog()
        counts = etb.process_source(
            cfg, second, master, dry_run=False, completed=load_completed(log_path),
            summary_index=SummaryIndex(second, master.project_id),
        )

        assert counts["resumed"] == 10
        assert slept == []

    def test_スキップした行でも待機しない(self, source_cfg, master, monkeypatch):
        slept = self._timed(monkeypatch)
        cfg = source_cfg(["件名", "期限"], [["", "2025/01/05"], ["", "2025/01/06"]])

        etb.process_source(cfg, FakeBacklog(), master, dry_run=False)

        assert slept == []

    def test_作成した行では待機する(self, source_cfg, master, monkeypatch):
        slept = self._timed(monkeypatch)

        etb.process_source(
            source_cfg(["件名"], [["課題A"], ["課題B"]]),
            FakeBacklog(), master, dry_run=False,
        )

        assert len(slept) == 2
