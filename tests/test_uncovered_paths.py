"""
カバレッジの穴を埋めるテスト
============================
カバレッジ計測で残っていた未到達経路のうち、実際の運用で通りうるものを対象にする。
デバッグ出力や import フォールバックなど、環境依存で意味の薄いものは対象外。
"""

import csv

import pytest

import excel_to_backlog as etb
from conftest import FakeBacklog
from mapper import BacklogMaster, IssueMapper
from run_log import RunLog


def mapper_for(master, **cfg):
    base = {"issue_type": "タスク", "priority": "中", "summary_col": "件名"}
    base.update(cfg)
    return IssueMapper(base, master, headers=list(base.get("_headers", [])) or None)


class TestStartDate:
    """start_date_col は設定できるのに一度も検証されていなかった。"""

    def test_開始日が設定される(self, master):
        m = mapper_for(master, start_date_col="開始")
        params = m.map_row({"件名": "t", "開始": "2025/01/05"})
        assert params["startDate"] == "2025-01-05"

    def test_ゼロ埋めなしでも解釈する(self, master):
        m = mapper_for(master, start_date_col="開始")
        assert m.map_row({"件名": "t", "開始": "2025/1/5"})["startDate"] == "2025-01-05"

    def test_解釈できない値は警告して未設定(self, master):
        m = mapper_for(master, start_date_col="開始")
        params = m.map_row({"件名": "t", "開始": "未定"})
        assert "startDate" not in params
        assert any("開始日" in w for w in m.warnings)

    def test_テンプレートでも指定できる(self, master):
        m = mapper_for(master, start_date_col="{{年}}/{{月}}/1")
        params = m.map_row({"件名": "t", "年": "2025", "月": "4"})
        assert params["startDate"] == "2025-04-01"

    def test_期限日と併用できる(self, master):
        m = mapper_for(master, start_date_col="開始", due_date_col="期限")
        params = m.map_row({"件名": "t", "開始": "2025/01/05", "期限": "2025/01/31"})
        assert params["startDate"] == "2025-01-05"
        assert params["dueDate"] == "2025-01-31"


class TestRequiredSettingsMissing:
    """種別・優先度そのものが未設定のケース。"""

    def test_種別が未設定ならエラー(self, master):
        m = IssueMapper({"priority": "中", "summary_col": "件名"}, master)
        with pytest.raises(ValueError, match="issue_type"):
            m.resolve_fixed_fields()

    def test_優先度が未設定ならエラー(self, master):
        m = IssueMapper({"issue_type": "タスク", "summary_col": "件名"}, master)
        with pytest.raises(ValueError, match="priority"):
            m.resolve_fixed_fields()


class TestStatusEdgeCases:
    def test_status_col_が未設定なら何もしない(self, master):
        m = mapper_for(master)
        assert "statusId" not in m.map_row({"件名": "t"})

    def test_セルが空なら何もしない(self, master):
        m = mapper_for(master, status_col="S", status_map={"完了": "完了"})
        params = m.map_row({"件名": "t", "S": ""})
        assert "statusId" not in params
        assert m.warnings == []

    def test_Backlog_に無いステータス名は警告して未設定(self, master):
        """status_map の変換先が Backlog に存在しない場合。"""
        m = mapper_for(master, status_col="S", status_map={"完了": "存在しない状態"})
        params = m.map_row({"件名": "t", "S": "完了"})
        assert "statusId" not in params
        assert any("存在しない状態" in w for w in m.warnings)


class TestCustomFieldEdgeCases:
    @pytest.fixture
    def cf_master(self, master):
        master.custom_field_map = {
            "カテゴリ": {"id": 5, "typeId": 5, "items": {"設計": 51}},
            "メモ": {"id": 7, "typeId": 1, "items": {}},
        }
        return master

    def test_未定義のカスタム属性は警告して未設定(self, cf_master):
        m = mapper_for(cf_master,
                       custom_fields=[{"field_name": "存在しない属性", "col_name": "C"}])
        params = m.map_row({"件名": "t", "C": "値"})
        assert not any(k.startswith("customField_") for k in params)
        assert any("存在しない属性" in w for w in m.warnings)

    def test_不正な正規表現を含む_value_map_でも落ちない(self, cf_master):
        """value_map のキーは正規表現として評価される。"""
        m = mapper_for(cf_master, custom_fields=[
            {"field_name": "メモ", "col_name": "C", "value_map": {"[": "壊れた", "Z": "z"}}
        ])
        assert m.map_row({"件名": "t", "C": "Z"})["customField_7"] == "z"

    def test_選択肢が見つからない場合は警告して未設定(self, cf_master):
        m = mapper_for(cf_master,
                       custom_fields=[{"field_name": "カテゴリ", "col_name": "C"}])
        params = m.map_row({"件名": "t", "C": "無い選択肢"})
        assert "customField_5" not in params
        assert any("無い選択肢" in w for w in m.warnings)


class TestPreviewFormatting:
    """format_preview の未到達だった表示項目。"""

    def test_開始日とカスタム属性が表示される(self, master):
        master.custom_field_map = {"メモ": {"id": 7, "typeId": 1, "items": {}}}
        m = mapper_for(master, start_date_col="開始",
                       custom_fields=[{"field_name": "メモ", "col_name": "C"}])
        text = m.format_preview({"件名": "t", "開始": "2025/01/05", "C": "自由記述"}, 1)
        assert "**開始日:** 2025-01-05" in text
        assert "customField_7" in text

    def test_スキップされる行は理由を返す(self, master):
        m = mapper_for(master)
        assert "スキップ" in m.format_preview({"件名": ""}, 1)


class TestSourceLevelPaths:
    def test_読み込みに失敗したソースはエラーとして返る(self, master, tmp_path):
        cfg = {
            "name": "壊れたソース",
            "excel": {"path": str(tmp_path / "存在しない.xlsx")},
            "issue_mapping": {"issue_type": "タスク", "priority": "中", "summary_col": "件名"},
        }
        counts = etb.process_source(cfg, FakeBacklog(), master, dry_run=False)
        assert counts["error"] == 1

    def test_対象行が無いソースは何もしない(self, source_cfg, master, capsys):
        cfg = source_cfg(["件名", "状態"], [["課題A", "完了"]],
                         filters=[{"col_name": "状態", "value": "対応要"}])
        backlog = FakeBacklog()

        counts = etb.process_source(cfg, backlog, master, dry_run=False)

        assert counts == etb.new_counts()
        assert backlog.create_calls == 0
        assert "対象行がないため" in capsys.readouterr().out

    def test_作成時点で既に目的のステータスなら成功扱い(self, source_cfg, master, tmp_path):
        """ステータス変更が「変更なし」で返るケース。"""
        cfg = source_cfg(["件名", "状態"], [["課題A", "未対応"]],
                         issue_mapping={"status_col": "状態",
                                        "status_map": {"未対応": "未対応"}})
        backlog = FakeBacklog(no_change_on_update=True)
        log_path = tmp_path / "run.csv"

        with RunLog(log_path) as log:
            counts = etb.process_source(cfg, backlog, master, dry_run=False, run_log=log)

        assert counts["created"] == 1
        assert counts["status_failed"] == 0
        with open(log_path, encoding="utf-8-sig", newline="") as f:
            assert next(csv.DictReader(f))["action"] == "created"


class TestPatchListValues:
    """PATCH のリスト値展開（POST 側だけ検証されていた）。"""

    def test_複数選択のカスタム属性を更新できる(self, monkeypatch):
        import json
        import urllib.request
        from backlog_client import BacklogClient

        sent = {}

        def fake_urlopen(req, timeout=None, context=None):
            sent["body"] = req.data.decode("utf-8")

            class _Res:
                def __enter__(self_inner): return self_inner
                def __exit__(self_inner, *a): return False
                def read(self_inner): return json.dumps({"id": 1}).encode()

            return _Res()

        monkeypatch.setattr(urllib.request, "urlopen", fake_urlopen)
        BacklogClient("example.com", "k")._patch("/issues/D-1", {"customField_6": [61, 62]})

        assert "customField_6[]=61" in sent["body"]
        assert "customField_6[]=62" in sent["body"]
        assert "%5B%5D" not in sent["body"]
