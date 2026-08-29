"""
再開時の行識別のテスト
======================
再開の識別は (ソース名, 行番号, 件名)。

件名だけで識別すると同じ件名の行を区別できず、未処理の行を「処理済み」と
誤判定して永久に飛ばしてしまう。また再開スキップをログに残さないと、
--resume を繰り返すたびにログが痩せ、次の再開で作成済みの行が再作成される。
"""

import excel_to_backlog as etb
from conftest import FakeBacklog
from run_log import COMPLETED_ACTIONS, RunLog, completion_key, load_completed


def actions(path):
    import csv
    with open(path, encoding="utf-8-sig", newline="") as f:
        return [r["action"] for r in csv.DictReader(f)]


class TestIdentity:
    def test_行番号が違えば別の行(self):
        assert completion_key("S", "2", "同じ件名") != completion_key("S", "5", "同じ件名")

    def test_ソースが違えば別の行(self):
        assert completion_key("S1", "2", "件名") != completion_key("S2", "2", "件名")

    def test_同じ件名の未処理行を飛ばさない(self, source_cfg, master, tmp_path):
        """1回目で到達する前に落ちた同名の行が、永久に作成されない問題。"""
        cfg = source_cfg(["件名"], [["定例会議"], ["別件"], ["定例会議"]])
        log_path = tmp_path / "run.csv"

        # Excel 4行目（3件目）の作成で失敗させる
        with RunLog(log_path) as log:
            etb.process_source(
                cfg, FakeBacklog(fail_create_at=3), master, dry_run=False, run_log=log
            )

        second = FakeBacklog()
        counts = etb.process_source(
            cfg, second, master, dry_run=False, completed=load_completed(log_path)
        )

        assert counts["resumed"] == 2
        assert second.create_calls == 1          # 未処理の行だけ作られる


class TestChaining:
    """--resume を繰り返してもログが痩せないこと。"""

    def test_resumed_は完了扱い(self):
        assert "resumed" in COMPLETED_ACTIONS

    def test_再開スキップもログに記録される(self, source_cfg, master, tmp_path):
        cfg = source_cfg(["件名"], [["A"], ["B"], ["C"]])

        first = tmp_path / "run1.csv"
        with RunLog(first) as log:
            etb.process_source(
                cfg, FakeBacklog(fail_create_at=2), master, dry_run=False, run_log=log
            )

        second = tmp_path / "run2.csv"
        with RunLog(second) as log:
            etb.process_source(
                cfg, FakeBacklog(), master, dry_run=False,
                run_log=log, completed=load_completed(first),
            )

        assert actions(first) == ["created", "error", "created"]
        assert actions(second) == ["resumed", "created", "resumed"]

    def test_2回目のログだけで再開しても重複しない(self, source_cfg, master, tmp_path):
        cfg = source_cfg(["件名"], [["A"], ["B"], ["C"]])

        first = tmp_path / "run1.csv"
        with RunLog(first) as log:
            etb.process_source(
                cfg, FakeBacklog(fail_create_at=2), master, dry_run=False, run_log=log
            )
        second = tmp_path / "run2.csv"
        with RunLog(second) as log:
            etb.process_source(
                cfg, FakeBacklog(), master, dry_run=False,
                run_log=log, completed=load_completed(first),
            )

        third = FakeBacklog()
        counts = etb.process_source(
            cfg, third, master, dry_run=False, completed=load_completed(second)
        )

        assert counts["resumed"] == 3
        assert third.create_calls == 0           # 以前は A・C が再作成された
