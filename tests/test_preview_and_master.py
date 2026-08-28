"""
プレビュー生成と BacklogMaster のテスト
=======================================
--preview の出力と、マスターデータ取得が部分的に失敗したときの挙動。

BacklogMaster.build は取得失敗を握り潰して空のまま続行する分岐を持つ。
この分岐が「カスタム属性の選択肢一覧が空」状態の発生源になる。
"""

import pytest

import excel_to_backlog as etb
from backlog_client import BacklogAPIError
from conftest import FakeBacklog
from mapper import BacklogMaster


# ------------------------------------------------------------------
# プレビュー生成
# ------------------------------------------------------------------

class TestGeneratePreview:
    def test_課題の内容が出力される(self, source_cfg, master, tmp_path):
        cfg = source_cfg(
            ["件名", "期限", "概要"],
            [["ログイン不具合", "2025/01/05", "落ちる"]],
            issue_mapping={"due_date_col": "期限", "description_format": "auto"},
        )
        out = tmp_path / "preview.md"

        count = etb.generate_preview_for_source(
            cfg, master, etb.build_master_labels(master), out, "2026-01-01"
        )

        text = out.read_text(encoding="utf-8")
        assert count == 1
        assert "ログイン不具合" in text
        assert "2025-01-05" in text
        assert "落ちる" in text

    def test_ID_ではなく名前で表示される(self, source_cfg, master, tmp_path):
        """種別・優先度・担当者・ステータスは ID のままだと読めない。"""
        cfg = source_cfg(
            ["件名", "担当", "状態"], [["課題A", "山田太郎", "完了"]],
            issue_mapping={
                "assignee_col": "担当",
                "status_col": "状態", "status_map": {"完了": "完了"},
            },
        )
        out = tmp_path / "preview.md"

        etb.generate_preview_for_source(
            cfg, master, etb.build_master_labels(master), out, "2026-01-01"
        )

        text = out.read_text(encoding="utf-8")
        assert "**種別:** タスク" in text
        assert "**優先度:** 中" in text
        assert "**担当者:** 山田太郎" in text
        assert "**ステータス:** 完了" in text

    def test_スキップされる行は理由が出る(self, source_cfg, master, tmp_path):
        cfg = source_cfg(["件名"], [[""]])
        out = tmp_path / "preview.md"

        etb.generate_preview_for_source(
            cfg, master, etb.build_master_labels(master), out, "2026-01-01"
        )

        assert "スキップ" in out.read_text(encoding="utf-8")

    def test_対象行がなければその旨を出す(self, source_cfg, master, tmp_path):
        cfg = source_cfg(["件名", "状態"], [["課題A", "完了"]],
                         filters=[{"col_name": "状態", "value": "対応要"}])
        out = tmp_path / "preview.md"

        count = etb.generate_preview_for_source(
            cfg, master, etb.build_master_labels(master), out, "2026-01-01"
        )

        assert count == 0
        assert "対象行がありません" in out.read_text(encoding="utf-8")

    def test_読み込み失敗時は原因を出力して0件を返す(self, master, tmp_path):
        cfg = {
            "name": "壊れたソース",
            "excel": {"path": str(tmp_path / "存在しない.xlsx")},
            "issue_mapping": {"issue_type": "タスク", "priority": "中", "summary_col": "件名"},
        }
        out = tmp_path / "preview.md"

        count = etb.generate_preview_for_source(
            cfg, master, etb.build_master_labels(master), out, "2026-01-01"
        )

        assert count == 0
        assert "Excel 読み込みエラー" in out.read_text(encoding="utf-8")

    def test_ソースごとにファイルが分かれる(self, source_cfg, master, tmp_path):
        sources = [
            source_cfg(["件名"], [["課題A"]], name="ソース1"),
            source_cfg(["件名"], [["課題B"]], name="ソース2"),
        ]

        results = etb.generate_preview_file(sources, master, tmp_path, "20260101_000000")

        assert [count for _, count in results] == [1, 1]
        assert {p.name for p, _ in results} == {
            "preview_20260101_000000_ソース1.md",
            "preview_20260101_000000_ソース2.md",
        }


class TestSafeFilename:
    @pytest.mark.parametrize("name,expected", [
        ("タスク管理表", "タスク管理表"),
        ("A/B", "A_B"),
        ("問い合わせ 一覧", "問い合わせ_一覧"),
        ('a:b*c?d"e<f>g|h', "a_b_c_d_e_f_g_h"),
        ("...", "source"),
        ("", "source"),
    ])
    def test_ファイル名として安全な文字列にする(self, name, expected):
        assert etb._safe_filename(name) == expected


class TestBuildMasterLabels:
    def test_ID_から名前を引ける(self, master):
        labels = etb.build_master_labels(master)

        assert labels["issue_type"][1] == "タスク"
        assert labels["priority"][3] == "中"
        assert labels["status"][4] == "完了"

    def test_ユーザーは表示名が優先される(self, master):
        """user_map には表示名とログインIDの両方が入っている。"""
        assert etb.build_master_labels(master)["user"][10] == "山田太郎"

    def test_種別と優先度は別々の空間として扱う(self, master):
        """どちらも ID 1〜3 を持ちうるためフラットにマージしてはいけない。"""
        labels = etb.build_master_labels(master)
        assert labels["issue_type"] is not labels["priority"]
        assert labels["issue_type"][2] == "バグ"
        assert labels["priority"][2] == "高"


# ------------------------------------------------------------------
# BacklogMaster.build
# ------------------------------------------------------------------

class TestBacklogMasterBuild:
    def test_すべて取得できる(self):
        master = BacklogMaster.build(FakeBacklog(), "DEMO")

        assert master.project_id == 42
        assert master.issue_type_map == {"タスク": 1, "バグ": 2}
        assert master.priority_map == {"高": 2, "中": 3}
        assert master.status_map == {"未対応": 1, "完了": 4}

    def test_ユーザーは表示名とログインIDの両方で引ける(self):
        master = BacklogMaster.build(FakeBacklog(), "DEMO")

        assert master.user_map["山田太郎"] == 10
        assert master.user_map["yamada"] == 10

    @pytest.mark.parametrize("method,attr,label", [
        ("get_project_users", "user_map", "メンバー"),
        ("get_custom_fields", "custom_field_map", "カスタム属性"),
        ("get_statuses", "status_map", "ステータス"),
    ])
    def test_任意項目の取得失敗は空のまま続行する(
        self, monkeypatch, capsys, method, attr, label
    ):
        """
        権限不足などで一部が取れなくても実行は続く。
        ただし空のまま進むと後段で問題が起きるため、警告は必ず出す。
        """
        client = FakeBacklog()
        monkeypatch.setattr(
            client, method,
            lambda *a, **kw: (_ for _ in ()).throw(BacklogAPIError("権限がありません", status=403)),
        )

        master = BacklogMaster.build(client, "DEMO")

        assert getattr(master, attr) == {}
        assert "⚠" in capsys.readouterr().err

    def test_必須項目の取得失敗は送出される(self, monkeypatch):
        """プロジェクト・種別・優先度が無いと課題を作れないため中止する。"""
        client = FakeBacklog()
        monkeypatch.setattr(
            client, "get_project",
            lambda *a: (_ for _ in ()).throw(BacklogAPIError("見つかりません", status=404)),
        )

        with pytest.raises(BacklogAPIError):
            BacklogMaster.build(client, "DEMO")

    def test_カスタム属性の選択肢が取れないと空の_items_になる(self, monkeypatch):
        """
        この状態が「選択肢型なのに選択肢一覧が空」の発生源。
        mapper 側で検出してスキップする（test_mapper.py 参照）。
        """
        client = FakeBacklog()
        monkeypatch.setattr(
            client, "get_custom_fields",
            lambda *a: [{"name": "カテゴリ", "id": 5, "typeId": 5}],   # items なし
        )

        master = BacklogMaster.build(client, "DEMO")

        assert master.custom_field_map["カテゴリ"]["items"] == {}
