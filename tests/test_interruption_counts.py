"""
中断時の集計と終了コードのテスト
================================
「実行結果の報告が事実と違う」問題を固定する。
"""

import pytest

import excel_to_backlog as etb
from backlog_client import BacklogAPIError
from conftest import FakeBacklog


class TestCountsSurviveInterruption:
    """
    process_source は集計辞書をその場で更新する。
    途中で例外が送出されても、それまでの集計が呼び出し元に残る。
    """

    def test_中断してもそれまでの作成件数が残る(self, source_cfg, master):
        backlog = FakeBacklog()
        original = backlog.create_issue

        def fail_after_three(params):
            if backlog.create_calls >= 3:
                raise BacklogAPIError("認証失敗", status=401, fatal=True)
            return original(params)

        backlog.create_issue = fail_after_three
        cfg = source_cfg(["件名"], [[f"課題{i}"] for i in range(1, 6)])

        total = etb.new_counts()
        with pytest.raises(BacklogAPIError):
            etb.process_source(cfg, backlog, master, dry_run=False, counts=total)

        assert total["created"] == 3        # 以前は 0 だった
        assert total["error"] == 1

    def test_Ctrl_C_でもそれまでの集計が残る(self, source_cfg, master):
        backlog = FakeBacklog()
        original = backlog.create_issue

        def interrupt_after_two(params):
            if backlog.create_calls >= 2:
                raise KeyboardInterrupt
            return original(params)

        backlog.create_issue = interrupt_after_two
        cfg = source_cfg(["件名"], [[f"課題{i}"] for i in range(1, 6)])

        total = etb.new_counts()
        with pytest.raises(KeyboardInterrupt):
            etb.process_source(cfg, backlog, master, dry_run=False, counts=total)

        assert total["created"] == 2

    def test_複数ソースの集計が積み上がる(self, source_cfg, master):
        backlog = FakeBacklog()
        total = etb.new_counts()

        for i in (1, 2):
            etb.process_source(
                source_cfg(["件名"], [[f"S{i}課題A"], [f"S{i}課題B"]], name=f"S{i}"),
                backlog, master, dry_run=False, counts=total,
            )

        assert total["created"] == 4

    def test_counts_を省略しても動く(self, source_cfg, master):
        counts = etb.process_source(
            source_cfg(["件名"], [["課題A"]]), FakeBacklog(), master, dry_run=False
        )
        assert counts["created"] == 1


class TestConfigErrorIsNotARowSkip:
    """設定ミスは全行に等しく影響する。データ不備として報告しない。"""

    def test_種別名の誤りはソースを中止する(self, source_cfg, master, capsys):
        cfg = source_cfg(
            ["件名"], [[f"課題{i}"] for i in range(1, 6)],
            issue_mapping={"issue_type": "存在しない種別"},
        )

        counts = etb.process_source(cfg, FakeBacklog(), master, dry_run=False)

        assert counts["error"] == 1         # 以前は skipped=5 / error=0
        assert counts["skipped"] == 0
        assert "issue_mapping の設定を確認" in capsys.readouterr().err

    def test_優先度名の誤りもソースを中止する(self, source_cfg, master):
        cfg = source_cfg(["件名"], [["課題A"]],
                         issue_mapping={"priority": "存在しない優先度"})

        counts = etb.process_source(cfg, FakeBacklog(), master, dry_run=False)

        assert counts["error"] == 1
        assert counts["skipped"] == 0

    def test_API_を呼ばずに中止する(self, source_cfg, master):
        cfg = source_cfg(["件名"], [["課題A"]],
                         issue_mapping={"issue_type": "存在しない種別"})
        backlog = FakeBacklog()

        etb.process_source(cfg, backlog, master, dry_run=False)

        assert backlog.create_calls == 0
