"""
分岐カバレッジで見つかった未検証経路のテスト
============================================
行カバレッジでは「実行済み」と数えられるが、条件が偽になる側を一度も
通っていなかった分岐。行が実行されたかではなく、両方の枝を通ったかを見る
と、こうした穴が見える。
"""

import pytest

import excel_to_backlog as etb
from conftest import FakeBacklog
from mapper import BacklogMaster, IssueMapper
from run_log import RunLog
from summary_index import SummaryIndex


class TestFindExistingIssue:
    def test_件名が空なら照合しない(self, master):
        """summary が空のときに索引を引かないこと（153->156 の分岐）。"""
        backlog = FakeBacklog({"DEMO-1": "既存"})
        index = SummaryIndex(backlog, master.project_id)

        got = etb.find_existing_issue(
            backlog, {"match_summary": True}, {}, {"summary": ""}, master,
            summary_index=index,
        )

        assert got is None
        assert backlog.get_issues_calls == 0      # 索引の構築すらしない


class TestPreviewErrorDetail:
    def test_詳細が無いエラーでも出力できる(self, master, tmp_path):
        """
        SourceLoadError の detail は読み込み失敗のときだけ付く。
        列名不一致など detail が無い場合の分岐（573->576）。
        """
        cfg = {
            "name": "S",
            "excel": {"path": str(tmp_path / "x.xlsx")},
            "issue_mapping": {"issue_type": "タスク", "priority": "中",
                              "summary_col": "無い列"},
        }
        # 列名不一致で失敗させるため、実在するファイルを用意する
        from openpyxl import Workbook
        wb = Workbook(); wb.active["A1"] = "件名"; wb.active["A2"] = "課題A"
        wb.save(tmp_path / "x.xlsx")
        out = tmp_path / "preview.md"

        count = etb.generate_preview_for_source(
            cfg, master, etb.build_master_labels(master), out, "2026-01-01"
        )

        text = out.read_text(encoding="utf-8")
        assert count == 0
        assert "ヘッダーに存在しません" in text
        assert "詳細:" not in text                # detail が無いので出さない


class TestWarningsAbsentBranches:
    """警告が無い行では partial を増やさないこと。"""

    def test_変更なしで警告が無ければ_partial_を増やさない(self, source_cfg, master, tmp_path):
        cfg = source_cfg(["件名"], [["既存課題"]],
                         upsert={"enabled": True, "match_summary": True})
        backlog = FakeBacklog({"DEMO-1": "既存課題"}, no_change_on_update=True)

        with RunLog(tmp_path / "run.csv") as log:
            counts = etb.process_source(
                cfg, backlog, master, dry_run=False, run_log=log,
                summary_index=SummaryIndex(backlog, master.project_id),
            )

        assert counts["unchanged"] == 1
        assert counts["partial"] == 0

    def test_ステータス変更失敗で警告が無ければ_partial_を増やさない(
        self, source_cfg, master
    ):
        cfg = source_cfg(["件名", "状態"], [["課題A", "完了"]],
                         issue_mapping={"status_col": "状態",
                                        "status_map": {"完了": "完了"}})

        counts = etb.process_source(
            cfg, FakeBacklog(fail_status_update=True), master, dry_run=False
        )

        assert counts["status_failed"] == 1
        assert counts["partial"] == 0


class TestValueMapRegex:
    @pytest.fixture
    def cf_master(self, master):
        master.custom_field_map = {"メモ": {"id": 7, "typeId": 1, "items": {}}}
        return master

    def _mapper(self, master, value_map):
        return IssueMapper({
            "issue_type": "タスク", "priority": "中", "summary_col": "件名",
            "custom_fields": [{"field_name": "メモ", "col_name": "C",
                               "value_map": value_map}],
        }, master)

    def test_どのパターンにも一致しない場合(self, cf_master):
        """for を最後まで回して break しない経路（548->555）。"""
        m = self._mapper(cf_master, {"設計.*": "設計", ".*テスト": "QA"})
        params = m.map_row({"件名": "t", "C": "該当なし"})

        assert "customField_7" not in params
        assert any("該当なし" in w for w in m.warnings)

    def test_先頭のパターンが外れて次で一致する(self, cf_master):
        """1つ目が不一致でループを続ける経路（550->548）。"""
        m = self._mapper(cf_master, {"設計.*": "設計", ".*テスト": "QA"})
        assert m.map_row({"件名": "t", "C": "単体テスト"})["customField_7"] == "QA"


class TestSummaryIndexAdd:
    def test_正規化して空になる件名は索引に入れない(self, master):
        """タブや改行だけの件名（95->exit の分岐）。"""
        backlog = FakeBacklog()
        index = SummaryIndex(backlog, master.project_id)
        index.find("何か")                      # 索引を構築させる

        index.add("\t\n ", "DEMO-9")

        assert index.find("") is None
        assert index.find("\t\n ") is None


class TestUserMapWithoutLoginId:
    def test_ログインIDが無いユーザーも登録される(self, monkeypatch):
        """u.get(\"userId\") が偽の経路（mapper.py 65->63）。"""
        backlog = FakeBacklog()
        monkeypatch.setattr(
            backlog, "get_project_users",
            lambda k: [{"name": "表示名のみ", "id": 20},
                       {"name": "山田", "id": 21, "userId": "yamada"}],
        )

        built = BacklogMaster.build(backlog, "DEMO")

        assert built.user_map["表示名のみ"] == 20
        assert built.user_map["yamada"] == 21


class TestRunLogClosedTwice:
    def test_二重に閉じても落ちない(self, tmp_path):
        """__exit__ が self._file を None にしたあとの分岐（55->58）。"""
        log = RunLog(tmp_path / "run.csv")
        with log:
            log.record(source="S", row=1, action="created")
        log.__exit__(None, None, None)          # 2 回目
        assert (tmp_path / "run.csv").exists()


class TestHttpErrorWithoutStructuredBody:
    def test_JSON_でない応答でも本文を出す(self):
        """errors 配列が無く raw_body だけある経路（160->162）。"""
        import io
        import urllib.error

        from backlog_client import BacklogAPIError, BacklogClient

        err = urllib.error.HTTPError(
            "u", 500, "Server Error", {}, io.BytesIO(b"<html>Internal Error</html>")
        )
        with pytest.raises(BacklogAPIError) as exc:
            BacklogClient("example.com", "k")._handle_http_error(err, "/issues")

        assert "Internal Error" in str(exc.value)
