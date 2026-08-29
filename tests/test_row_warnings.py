"""
行ごとの警告のテスト
====================
担当者・日付・ステータス・カスタム属性が解決できないと、その行は
「作成されたが一部フィールドが未設定」という状態になる。

以前は stderr に直接出していたため行番号も件名も付かず、500 行の実行では
どの警告がどの行のものか対応づけられなかった。実行ログの detail 列に
記録し、サマリーにも件数を出す。
"""

import csv

import excel_to_backlog as etb
from conftest import FakeBacklog
from mapper import IssueMapper
from run_log import RunLog


def details(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        return {r["row"]: r["detail"] for r in csv.DictReader(f)}


class TestMapperCollectsWarnings:
    def base_cfg(self, **extra):
        cfg = {"issue_type": "タスク", "priority": "中", "summary_col": "件名"}
        cfg.update(extra)
        return cfg

    def test_解決できた行に警告は付かない(self, master):
        m = IssueMapper(self.base_cfg(due_date_col="期限"), master)
        m.map_row({"件名": "課題A", "期限": "2025/01/05"})
        assert m.warnings == []

    def test_日付が解釈できないと警告が付く(self, master):
        m = IssueMapper(self.base_cfg(due_date_col="期限"), master)
        m.map_row({"件名": "課題A", "期限": "R7/1/5"})
        assert any("R7/1/5" in w for w in m.warnings)

    def test_担当者が見つからないと警告が付く(self, master):
        m = IssueMapper(self.base_cfg(assignee_col="担当"), master)
        m.map_row({"件名": "課題A", "担当": "存在しない人"})
        assert any("存在しない人" in w for w in m.warnings)

    def test_複数の警告が積まれる(self, master):
        m = IssueMapper(self.base_cfg(due_date_col="期限", assignee_col="担当"), master)
        m.map_row({"件名": "課題A", "期限": "R7/1/5", "担当": "存在しない人"})
        assert len(m.warnings) == 2

    def test_行ごとにリセットされる(self, master):
        m = IssueMapper(self.base_cfg(due_date_col="期限"), master)
        m.map_row({"件名": "課題A", "期限": "R7/1/5"})
        m.map_row({"件名": "課題B", "期限": "2025/01/05"})
        assert m.warnings == []

    def test_警告は1行に畳まれる(self, master):
        """改行を含む警告でも CSV の1セルに収まること。"""
        m = IssueMapper(self.base_cfg(assignee_col="担当"), master)
        m.map_row({"件名": "課題A", "担当": "存在しない人"})
        assert all("\n" not in w for w in m.warnings)


class TestWarningsInRunLog:
    def test_作成した行の警告が_detail_に入る(self, source_cfg, master, tmp_path):
        cfg = source_cfg(
            ["件名", "期限"], [["課題A", "2025/01/05"], ["課題B", "R7/1/20"]],
            issue_mapping={"due_date_col": "期限"},
        )
        log_path = tmp_path / "run.csv"

        with RunLog(log_path) as log:
            etb.process_source(cfg, FakeBacklog(), master, dry_run=False, run_log=log)

        d = details(log_path)
        assert d["2"] == ""                      # 正常な行
        assert "R7/1/20" in d["3"]               # 落ちたフィールドが分かる

    def test_更新した行にも記録される(self, source_cfg, master, tmp_path):
        cfg = source_cfg(
            ["件名", "期限"], [["既存課題", "R7/1/20"]],
            issue_mapping={"due_date_col": "期限"},
            upsert={"enabled": True, "match_summary": True},
        )
        log_path = tmp_path / "run.csv"
        backlog = FakeBacklog({"DEMO-1": "既存課題"})

        from summary_index import SummaryIndex
        with RunLog(log_path) as log:
            etb.process_source(
                cfg, backlog, master, dry_run=False, run_log=log,
                summary_index=SummaryIndex(backlog, master.project_id),
            )

        assert "R7/1/20" in details(log_path)["2"]


class TestPartialCount:
    def test_一部未設定の行が集計される(self, source_cfg, master):
        cfg = source_cfg(
            ["件名", "期限"],
            [["課題A", "2025/01/05"], ["課題B", "R7/1/20"], ["課題C", "未定"]],
            issue_mapping={"due_date_col": "期限"},
        )

        counts = etb.process_source(cfg, FakeBacklog(), master, dry_run=False)

        assert counts["created"] == 3            # すべて作成はされている
        assert counts["partial"] == 2

    def test_サマリーに件数と対処が出る(self, capsys):
        total = etb.new_counts()
        total["created"], total["partial"] = 3, 2

        etb.print_summary(total, dry_run=False)

        out = capsys.readouterr().out
        assert "うち一部フィールド未設定: 2 件" in out
        assert "detail 列" in out

    def test_警告がなければ表示しない(self, capsys):
        total = etb.new_counts()
        total["created"] = 3

        etb.print_summary(total, dry_run=False)

        assert "一部フィールド未設定" not in capsys.readouterr().out


class TestWarningsOnAllOutcomes:
    """
    警告は created / updated だけでなく、変更なし・ステータス変更失敗の行にも
    付く。特定の結果でだけ落ちると「detail 列に記録しています」という案内が
    嘘になり、ドライランと実行の partial 件数もずれる。
    """

    def _cfg(self, source_cfg, **extra):
        return source_cfg(
            ["件名", "期限"], [["既存課題", "R7/1/20"]],
            issue_mapping={"due_date_col": "期限"}, **extra,
        )

    def test_変更なしの行にも警告が残る(self, source_cfg, master, tmp_path):
        from summary_index import SummaryIndex

        cfg = self._cfg(source_cfg, upsert={"enabled": True, "match_summary": True})
        backlog = FakeBacklog({"DEMO-1": "既存課題"}, no_change_on_update=True)
        log_path = tmp_path / "run.csv"

        with RunLog(log_path) as log:
            counts = etb.process_source(
                cfg, backlog, master, dry_run=False, run_log=log,
                summary_index=SummaryIndex(backlog, master.project_id),
            )

        assert counts["unchanged"] == 1
        assert counts["partial"] == 1
        assert "R7/1/20" in details(log_path)["2"]

    def test_ステータス変更に失敗した行にも警告が残る(self, source_cfg, master, tmp_path):
        cfg = source_cfg(
            ["件名", "期限", "状態"], [["課題A", "R7/1/20", "完了"]],
            issue_mapping={"due_date_col": "期限", "status_col": "状態",
                           "status_map": {"完了": "完了"}},
        )
        backlog = FakeBacklog(fail_status_update=True)
        log_path = tmp_path / "run.csv"

        with RunLog(log_path) as log:
            counts = etb.process_source(
                cfg, backlog, master, dry_run=False, run_log=log
            )

        assert counts["created"] == 1
        assert counts["status_failed"] == 1
        assert counts["partial"] == 1
        detail = details(log_path)["2"]
        assert "ステータス" in detail          # 失敗の理由
        assert "R7/1/20" in detail             # 落ちたフィールド
