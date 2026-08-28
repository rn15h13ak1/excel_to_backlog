"""
upsert の振る舞いのテスト
=========================
「Excel を正とする」運用の根幹となる契約を固定する。

key_col と match_summary は、件名が変わったときの振る舞いが正反対になる。
この違いは運用方法の選択に直結するため、意図した挙動であることを明示する。

    | 件名の変更元 | key_col            | match_summary |
    |--------------|--------------------|---------------|
    | Excel 側     | Backlog を更新     | 重複作成      |
    | Backlog 側   | Excel の値へ戻す   | 重複作成      |
"""

import excel_to_backlog as etb
from conftest import FakeBacklog
from summary_index import SummaryIndex

OLD = "ログイン不具合"
NEW = "ログイン画面が表示されない"


def run(cfg, backlog, master, **kwargs):
    return etb.process_source(
        cfg, backlog, master, dry_run=False,
        summary_index=SummaryIndex(backlog, master.project_id), **kwargs
    )


# ------------------------------------------------------------------
# key_col 方式
# ------------------------------------------------------------------

class TestKeyCol:
    """issueKey で課題を特定するため、件名が変わっても追跡できる。"""

    def test_Excel側で件名を変更すると_Backlog_の件名も更新される(self, source_cfg, master):
        """Excel を正とする運用で期待される動作。"""
        backlog = FakeBacklog({"DEMO-1": OLD})
        cfg = source_cfg(
            ["件名", "Backlog番号"], [[NEW, "DEMO-1"]],
            upsert={"enabled": True, "key_col": "Backlog番号"},
        )

        counts = run(cfg, backlog, master)

        assert counts["updated"] == 1
        assert counts["created"] == 0
        assert backlog.issues == {"DEMO-1": NEW}      # 課題は増えない

    def test_Backlog側で件名が変更されていても_Excel_の値で上書きする(
        self, source_cfg, master
    ):
        """
        Excel を正とするため、Backlog 側の編集は失われる。
        これは意図した動作。Backlog 側でも件名を編集する運用には使えない。
        """
        backlog = FakeBacklog({"DEMO-1": "【修正済】ログイン不具合(担当:山田)"})
        cfg = source_cfg(
            ["件名", "Backlog番号"], [[OLD, "DEMO-1"]],
            upsert={"enabled": True, "key_col": "Backlog番号"},
        )

        run(cfg, backlog, master)

        assert backlog.issues == {"DEMO-1": OLD}

    def test_更新時に件名が送信される(self, source_cfg, master):
        """
        update_params から summary を除外すると Excel の件名変更が
        反映されなくなるため、送信されていることを明示的に固定する。
        """
        backlog = FakeBacklog({"DEMO-1": OLD})
        cfg = source_cfg(
            ["件名", "Backlog番号"], [[NEW, "DEMO-1"]],
            upsert={"enabled": True, "key_col": "Backlog番号"},
        )

        run(cfg, backlog, master)

        _, params = backlog.updates[0]
        assert params["summary"] == NEW
        assert "projectId" not in params          # 更新時は不要

    def test_key_col_が空の行は新規作成される(self, source_cfg, master):
        backlog = FakeBacklog()
        cfg = source_cfg(
            ["件名", "Backlog番号"], [[OLD, ""]],
            upsert={"enabled": True, "key_col": "Backlog番号"},
        )

        counts = run(cfg, backlog, master)

        assert counts["created"] == 1

    def test_存在しない_issueKey_は新規作成される(self, source_cfg, master, capsys):
        """課題が削除された場合など。エラーにせず新規作成する。"""
        backlog = FakeBacklog()
        cfg = source_cfg(
            ["件名", "Backlog番号"], [[OLD, "DEMO-999"]],
            upsert={"enabled": True, "key_col": "Backlog番号"},
        )

        counts = run(cfg, backlog, master)

        assert counts["created"] == 1
        assert "DEMO-999" in capsys.readouterr().err


# ------------------------------------------------------------------
# match_summary 方式
# ------------------------------------------------------------------

class TestMatchSummary:
    """件名そのものが照合キーのため、件名が変わると追跡できない。"""

    def test_件名が同じなら更新される(self, source_cfg, master):
        backlog = FakeBacklog({"DEMO-1": OLD})
        cfg = source_cfg(["件名"], [[OLD]],
                         upsert={"enabled": True, "match_summary": True})

        counts = run(cfg, backlog, master)

        assert counts["updated"] == 1
        assert len(backlog.issues) == 1

    def test_Excel側で件名を変更すると重複が作られる(self, source_cfg, master):
        """
        既知の制約。件名を照合キーにしている以上避けられない。
        Excel 側で件名を変更する運用では key_col を使うこと。
        """
        backlog = FakeBacklog({"DEMO-1": OLD})
        cfg = source_cfg(["件名"], [[NEW]],
                         upsert={"enabled": True, "match_summary": True})

        counts = run(cfg, backlog, master)

        assert counts["created"] == 1                 # 更新ではなく作成
        assert len(backlog.issues) == 2               # 旧課題が取り残される
        assert OLD in backlog.issues.values()
        assert NEW in backlog.issues.values()

    def test_件名の表記揺れは吸収される(self, source_cfg, master):
        """改行・タブ・連続スペースは正規化して照合する。"""
        backlog = FakeBacklog({"DEMO-1": "ログイン\t不具合  です"})
        cfg = source_cfg(["件名"], [["ログイン不具合 です"]],
                         upsert={"enabled": True, "match_summary": True})

        counts = run(cfg, backlog, master)

        assert counts["updated"] == 1


# ------------------------------------------------------------------
# upsert 無効
# ------------------------------------------------------------------

class TestUpsertDisabled:
    def test_既定では常に新規作成する(self, source_cfg, master):
        """upsert.enabled の既定は False。再実行すると重複する。"""
        backlog = FakeBacklog({"DEMO-1": OLD})
        cfg = source_cfg(["件名"], [[OLD]])

        counts = run(cfg, backlog, master)

        assert counts["created"] == 1
        assert len(backlog.issues) == 2

    def test_既存課題を探すための_API_を呼ばない(self, source_cfg, master):
        backlog = FakeBacklog()
        run(source_cfg(["件名"], [[OLD]]), backlog, master)
        assert backlog.get_issues_calls == 0


# ------------------------------------------------------------------
# key_col と match_summary の併用
# ------------------------------------------------------------------

class TestBothConfigured:
    """key_col に値がある行は issueKey、無い行は件名で照合する。"""

    def test_key_col_に値があればそちらを優先する(self, source_cfg, master):
        backlog = FakeBacklog({"DEMO-1": OLD, "DEMO-2": NEW})
        cfg = source_cfg(
            ["件名", "Backlog番号"], [[NEW, "DEMO-1"]],
            upsert={"enabled": True, "key_col": "Backlog番号", "match_summary": True},
        )

        run(cfg, backlog, master)

        # 件名で照合すれば DEMO-2 だが、key_col が優先されて DEMO-1 が更新される
        assert backlog.updates[0][0] == "DEMO-1"

    def test_key_col_が空なら件名で照合する(self, source_cfg, master):
        backlog = FakeBacklog({"DEMO-1": OLD})
        cfg = source_cfg(
            ["件名", "Backlog番号"], [[OLD, ""]],
            upsert={"enabled": True, "key_col": "Backlog番号", "match_summary": True},
        )

        counts = run(cfg, backlog, master)

        assert counts["updated"] == 1
        assert backlog.updates[0][0] == "DEMO-1"
