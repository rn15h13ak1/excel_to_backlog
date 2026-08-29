"""
ドライランが実行を予測することのテスト
======================================
以前のドライランは行ループの手前で return しており、upsert の照合を
一切行わなかった。このツールで最も影響が大きい「作成するのか更新するのか」
を事前に確認できず、スキップ件数も集計されなかった。

いまは実行と同じ plan_row() を通す。照合は読み取りのみで Backlog を変更しない。
"""

import excel_to_backlog as etb
from conftest import FakeBacklog
from summary_index import SummaryIndex


def run(cfg, backlog, master, dry_run):
    return etb.process_source(
        cfg, backlog, master, dry_run=dry_run,
        summary_index=SummaryIndex(backlog, master.project_id),
    )


class TestDryRunMatchesExecute:
    def _cfg(self, source_cfg):
        return source_cfg(
            ["件名", "期限"],
            [["既にある課題", "2025/01/05"], ["新しい課題", "R7/1/20"], ["", "2025/02/01"]],
            issue_mapping={"due_date_col": "期限"},
            upsert={"enabled": True, "match_summary": True},
        )

    def test_集計がまったく同じになる(self, source_cfg, master):
        cfg = self._cfg(source_cfg)

        dry = run(cfg, FakeBacklog({"DEMO-1": "既にある課題"}), master, dry_run=True)
        real = run(cfg, FakeBacklog({"DEMO-1": "既にある課題"}), master, dry_run=False)

        assert dry == real

    def test_ドライランで作成と更新を区別する(self, source_cfg, master, capsys):
        cfg = self._cfg(source_cfg)

        counts = run(cfg, FakeBacklog({"DEMO-1": "既にある課題"}), master, dry_run=True)

        assert counts["updated"] == 1
        assert counts["created"] == 1
        out = capsys.readouterr().out
        assert "更新 → DEMO-1" in out
        assert "新規作成" in out

    def test_ドライランでスキップを集計する(self, source_cfg, master):
        """以前は 500 行の出力を目で追うしかなかった。"""
        counts = run(self._cfg(source_cfg), FakeBacklog(), master, dry_run=True)
        assert counts["skipped"] == 1

    def test_ドライランで一部未設定を集計する(self, source_cfg, master):
        counts = run(self._cfg(source_cfg), FakeBacklog(), master, dry_run=True)
        assert counts["partial"] == 1


class TestDryRunDoesNotWrite:
    def test_課題を作成しない(self, source_cfg, master):
        backlog = FakeBacklog()
        run(source_cfg(["件名"], [["課題A"]]), backlog, master, dry_run=True)
        assert backlog.create_calls == 0

    def test_課題を更新しない(self, source_cfg, master):
        backlog = FakeBacklog({"DEMO-1": "既存"})
        cfg = source_cfg(["件名"], [["既存"]],
                         upsert={"enabled": True, "match_summary": True})

        run(cfg, backlog, master, dry_run=True)

        assert backlog.updates == []

    def test_照合のための読み取りは行う(self, source_cfg, master):
        """作成か更新かを判定するには既存課題の取得が必要。"""
        backlog = FakeBacklog({"DEMO-1": "既存"})
        cfg = source_cfg(["件名"], [["既存"]],
                         upsert={"enabled": True, "match_summary": True})

        run(cfg, backlog, master, dry_run=True)

        assert backlog.get_issues_calls == 1


class TestDryRunHonorsResume:
    def test_再開スキップもドライランに反映される(self, source_cfg, master, tmp_path):
        from run_log import RunLog, load_completed

        cfg = source_cfg(["件名"], [["課題A"], ["課題B"]])
        log_path = tmp_path / "run.csv"
        with RunLog(log_path) as log:
            etb.process_source(cfg, FakeBacklog(), master, dry_run=False, run_log=log)

        counts = etb.process_source(
            cfg, FakeBacklog(), master, dry_run=True, completed=load_completed(log_path)
        )

        assert counts["resumed"] == 2
        assert counts["created"] == 0


class TestPlanRow:
    def test_件名が空なら_skip(self, source_cfg, master, make_excel):
        from excel_reader import ExcelReader
        from mapper import IssueMapper

        cfg = source_cfg(["件名", "期限"], [["", "2025/01/05"]])
        _, rows = ExcelReader(cfg["excel"]).read()
        mapper = IssueMapper(cfg["issue_mapping"], master, headers=["件名", "期限"])

        plan = etb.plan_row(rows[0], cfg, mapper)

        assert plan.action == "skip"
        assert "件名" in plan.reason

    def test_upsert_無効なら常に_create(self, source_cfg, master):
        from excel_reader import ExcelReader
        from mapper import IssueMapper

        cfg = source_cfg(["件名"], [["既存"]])
        _, rows = ExcelReader(cfg["excel"]).read()
        mapper = IssueMapper(cfg["issue_mapping"], master, headers=["件名"])
        backlog = FakeBacklog({"DEMO-1": "既存"})

        plan = etb.plan_row(rows[0], cfg, mapper, client=backlog, master=master)

        assert plan.action == "create"
        assert plan.existing_key is None


class TestDuplicateSummaryInSheet:
    """
    実行時は作成した課題を索引に加えるため、同じ件名の後続行は「更新」になる。
    ドライランは何も作成しないため索引が更新されず、そのままでは「2件作成」と
    予測してしまう。作成予定の件名を索引に入れて判断を揃える。

    この経路は「集計が一致する」テストのデータに同一件名が無かったため
    見逃されていた。
    """

    def _cfg(self, source_cfg):
        return source_cfg(
            ["件名"], [["同じ件名"], ["同じ件名"]],
            upsert={"enabled": True, "match_summary": True},
        )

    def test_同じ件名が2行あっても予測が一致する(self, source_cfg, master):
        cfg = self._cfg(source_cfg)

        dry = run(cfg, FakeBacklog(), master, dry_run=True)
        real = run(cfg, FakeBacklog(), master, dry_run=False)

        assert dry == real
        assert dry["created"] == 1
        assert dry["updated"] == 1

    def test_ドライランは索引を更新しても書き込まない(self, source_cfg, master):
        backlog = FakeBacklog()

        run(self._cfg(source_cfg), backlog, master, dry_run=True)

        assert backlog.create_calls == 0
        assert backlog.updates == []

    def test_3行以上でも2行目以降は更新になる(self, source_cfg, master):
        cfg = source_cfg(
            ["件名"], [["同じ件名"]] * 3,
            upsert={"enabled": True, "match_summary": True},
        )

        dry = run(cfg, FakeBacklog(), master, dry_run=True)

        assert dry["created"] == 1
        assert dry["updated"] == 2
