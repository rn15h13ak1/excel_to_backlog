"""
IssueMapper のテスト
====================
Excel 行 → Backlog API パラメータの変換。外部 API を呼ばない純粋な処理のため、
BacklogMaster をその場で組み立てて検証する。
"""

import pytest

from mapper import BacklogMaster, IssueMapper


@pytest.fixture
def master():
    return BacklogMaster(
        project_id=100,
        issue_type_map={"タスク": 1, "バグ": 2},
        priority_map={"高": 2, "中": 3},
        user_map={"山田太郎": 10, "yamada": 10, "佐藤": 11},
        status_map={"未対応": 1, "処理中": 2, "完了": 4},
    )


def mapper_for(cfg, master, headers=None):
    base = {"issue_type": "タスク", "priority": "中", "summary_col": "件名"}
    return IssueMapper({**base, **cfg}, master, headers=headers)


# ------------------------------------------------------------------
# 日付の正規化
# ------------------------------------------------------------------

class TestNormalizeDate:
    @pytest.mark.parametrize("value,expected", [
        ("2025-01-05", "2025-01-05"),
        ("2025/01/05", "2025-01-05"),
        ("2025/1/5", "2025-01-05"),        # ゼロ埋めなし（手入力で頻出）
        ("2025-1-5", "2025-01-05"),
        ("2025年1月5日", "2025-01-05"),
        ("2025.1.5", "2025-01-05"),
        ("2025/01/05 10:00", "2025-01-05"),  # 時刻付き
        ("  2025/1/5  ", "2025-01-05"),
    ])
    def test_受理する形式(self, value, expected):
        assert IssueMapper._normalize_date(value) == expected

    @pytest.mark.parametrize("value", [
        "",
        "   ",
        "R7/1/5",      # 和暦
        "9/1",         # 年がない
        "2025/13/1",   # 存在しない月
        "2025/2/30",   # 存在しない日
        "未定",
    ])
    def test_解釈できない値は_None(self, value):
        assert IssueMapper._normalize_date(value) is None


class TestResolveDate:
    def test_ゼロ埋めなしの日付が期限日に設定される(self, master):
        """以前は警告なく捨てられ、期限日が未設定のまま登録されていた。"""
        m = mapper_for({"due_date_col": "期限"}, master)
        params = m.map_row({"件名": "テスト", "期限": "2025/1/5"})
        assert params["dueDate"] == "2025-01-05"

    def test_解釈できない値は警告を出して未設定にする(self, master, capsys):
        m = mapper_for({"due_date_col": "期限"}, master)
        params = m.map_row({"件名": "テスト", "期限": "R7/1/5"})
        assert "dueDate" not in params
        assert "R7/1/5" in capsys.readouterr().err

    def test_空セルでは警告を出さない(self, master, capsys):
        m = mapper_for({"due_date_col": "期限"}, master)
        params = m.map_row({"件名": "テスト", "期限": ""})
        assert "dueDate" not in params
        assert capsys.readouterr().err == ""

    def test_テンプレート形式でも解決できる(self, master):
        m = mapper_for({"due_date_col": "{{年}}/{{月}}/1"}, master)
        params = m.map_row({"件名": "テスト", "年": "2025", "月": "3"})
        assert params["dueDate"] == "2025-03-01"


# ------------------------------------------------------------------
# テンプレート
# ------------------------------------------------------------------

class TestExtractTemplateColumns:
    def test_通常のプレースホルダー(self):
        assert IssueMapper.extract_template_columns("{{件名}}と{{担当}}") == {"件名", "担当"}

    def test_条件ブロックの開始と終了(self):
        assert IssueMapper.extract_template_columns(
            "項番{{項番}}{{#枝番}}-{{枝番}}{{/枝番}}"
        ) == {"項番", "枝番"}

    def test_auto_は列名ではない(self):
        assert IssueMapper.extract_template_columns("{{auto}}") == set()

    def test_前後の空白は除去される(self):
        assert IssueMapper.extract_template_columns("{{ 件名 }}") == {"件名"}

    def test_プレースホルダーがなければ空(self):
        assert IssueMapper.extract_template_columns("固定文字列") == set()


class TestRenderTemplate:
    def test_値を埋め込む(self, master):
        m = mapper_for({}, master)
        assert m._render_template("【{{分類}}】{{件名}}",
                                  {"分類": "バグ", "件名": "落ちる"}) == "【バグ】落ちる"

    def test_条件ブロックは値があれば出力される(self, master):
        m = mapper_for({}, master)
        got = m._render_template("項番{{項番}}{{#枝番}}-{{枝番}}{{/枝番}}",
                                 {"項番": "1", "枝番": "A"})
        assert got == "項番1-A"

    def test_条件ブロックは値が空なら丸ごと消える(self, master):
        m = mapper_for({}, master)
        got = m._render_template("項番{{項番}}{{#枝番}}-{{枝番}}{{/枝番}}",
                                 {"項番": "1", "枝番": ""})
        assert got == "項番1"

    def test_未知の列はプレースホルダーのまま残る(self, master):
        """
        この挙動のままだと壊れた件名が登録されうるが、実行前に
        validate_column_references() が列名を検証して停止する。
        """
        m = mapper_for({}, master)
        assert m._render_template("{{存在しない}}", {}) == "{{存在しない}}"


class TestNormalizeSummary:
    def test_改行とタブを除去する(self):
        assert IssueMapper.normalize_summary("件名\n続き\tおわり") == "件名続きおわり"

    def test_連続スペースを1つに圧縮する(self):
        assert IssueMapper.normalize_summary("件名    続き") == "件名 続き"

    def test_前後の空白を除去する(self):
        assert IssueMapper.normalize_summary("  件名  ") == "件名"


# ------------------------------------------------------------------
# 行の変換
# ------------------------------------------------------------------

class TestMapRow:
    def test_必須項目が設定される(self, master):
        params = mapper_for({}, master).map_row({"件名": "ログイン不具合"})
        assert params["projectId"] == 100
        assert params["summary"] == "ログイン不具合"
        assert params["issueTypeId"] == 1
        assert params["priorityId"] == 3

    def test_件名が空ならスキップ(self, master):
        with pytest.raises(ValueError, match="件名"):
            mapper_for({}, master).map_row({"件名": ""})

    def test_required_cols_が空ならスキップ(self, master):
        m = mapper_for({"required_cols": ["対応内容"]}, master)
        with pytest.raises(ValueError, match="必須列"):
            m.map_row({"件名": "テスト", "対応内容": ""})

    def test_未知の種別はエラー(self, master):
        m = mapper_for({"issue_type": "存在しない種別"}, master)
        with pytest.raises(ValueError, match="種別"):
            m.map_row({"件名": "テスト"})

    def test_担当者は表示名でもログインIDでも解決できる(self, master):
        m = mapper_for({"assignee_col": "担当"}, master)
        assert m.map_row({"件名": "t", "担当": "山田太郎"})["assigneeId"] == 10
        assert m.map_row({"件名": "t", "担当": "yamada"})["assigneeId"] == 10

    def test_セルが空なら_default_assignee_が使われる(self, master):
        m = mapper_for({"assignee_col": "担当", "default_assignee": "佐藤"}, master)
        assert m.map_row({"件名": "t", "担当": ""})["assigneeId"] == 11

    def test_セルに値があれば_default_assignee_より優先される(self, master):
        m = mapper_for({"assignee_col": "担当", "default_assignee": "佐藤"}, master)
        assert m.map_row({"件名": "t", "担当": "山田太郎"})["assigneeId"] == 10

    def test_ステータスは_status_map_を経て_ID_に解決される(self, master):
        m = mapper_for({"status_col": "状態", "status_map": {"完了": "完了"}}, master)
        assert m.map_row({"件名": "t", "状態": "完了"})["statusId"] == 4

    def test_status_map_にない値は未設定のまま(self, master, capsys):
        m = mapper_for({"status_col": "状態", "status_map": {"完了": "完了"}}, master)
        params = m.map_row({"件名": "t", "状態": "保留"})
        assert "statusId" not in params
        assert "保留" in capsys.readouterr().err


class TestRenderAuto:
    def test_列名が見出しになる(self, master):
        m = mapper_for({"description_format": "auto"}, master,
                       headers=["件名", "概要"])
        params = m.map_row({"件名": "t", "概要": "本文"})
        assert "# 概要\n本文" in params["description"]

    def test_複数行ヘッダーは階層見出しになる(self, master):
        m = mapper_for({"description_format": "auto", "description_cols": ["大分類 / 小分類"]},
                       master, headers=["件名", "大分類 / 小分類"])
        params = m.map_row({"件名": "t", "大分類 / 小分類": "値"})
        assert "# 大分類\n## 小分類\n値" in params["description"]

    def test_空セルは値なしと表示される(self, master):
        m = mapper_for({"description_format": "auto", "description_cols": ["概要"]},
                       master, headers=["件名", "概要"])
        params = m.map_row({"件名": "t", "概要": ""})
        assert "（値なし）" in params["description"]

    def test_セル内改行は_br_に変換される(self, master):
        m = mapper_for({"description_format": "auto", "description_cols": ["概要"]},
                       master, headers=["件名", "概要"])
        params = m.map_row({"件名": "t", "概要": "1行目\n2行目"})
        assert "1行目<br>2行目" in params["description"]

    def test_書式付き行が渡されると本文に使われる(self, master):
        """rich_text: true のとき取り消し線が本文に反映されること。"""
        m = mapper_for({"description_format": "auto", "description_cols": ["概要"]},
                       master, headers=["件名", "概要"])
        params = m.map_row({"件名": "t", "概要": "削除済み"},
                           formatted_row={"件名": "t", "概要": "~~削除済み~~"})
        assert "~~削除済み~~" in params["description"]


class TestCustomFields:
    @pytest.fixture
    def cf_master(self, master):
        master.custom_field_map = {
            "カテゴリ": {"id": 5, "typeId": 5, "items": {"設計": 51, "開発": 52}},
            "タグ":     {"id": 6, "typeId": 6, "items": {"設計": 61, "QA": 62}},
            "メモ":     {"id": 7, "typeId": 1, "items": {}},
        }
        return master

    def test_単一選択型は_int_で渡す(self, cf_master):
        m = mapper_for({"custom_fields": [{"field_name": "カテゴリ", "col_name": "分類"}]},
                       cf_master)
        assert m.map_row({"件名": "t", "分類": "設計"})["customField_5"] == 51

    def test_複数選択型はリストで渡す(self, cf_master):
        m = mapper_for({"custom_fields": [
            {"field_name": "タグ", "col_name": "タグ", "value_separator": ","}
        ]}, cf_master)
        assert m.map_row({"件名": "t", "タグ": "設計,QA"})["customField_6"] == [61, 62]

    def test_value_map_で値を変換する(self, cf_master):
        m = mapper_for({"custom_fields": [
            {"field_name": "カテゴリ", "col_name": "分類", "value_map": {"A": "設計"}}
        ]}, cf_master)
        assert m.map_row({"件名": "t", "分類": "A"})["customField_5"] == 51

    def test_value_map_は正規表現も使える(self, cf_master):
        m = mapper_for({"custom_fields": [
            {"field_name": "カテゴリ", "col_name": "分類", "value_map": {"設計.*": "設計"}}
        ]}, cf_master)
        assert m.map_row({"件名": "t", "分類": "設計A"})["customField_5"] == 51

    def test_非選択型はそのまま文字列で渡す(self, cf_master):
        m = mapper_for({"custom_fields": [{"field_name": "メモ", "col_name": "備考"}]},
                       cf_master)
        assert m.map_row({"件名": "t", "備考": "自由記述"})["customField_7"] == "自由記述"

    def test_未知の選択肢はスキップして警告(self, cf_master, capsys):
        m = mapper_for({"custom_fields": [{"field_name": "カテゴリ", "col_name": "分類"}]},
                       cf_master)
        params = m.map_row({"件名": "t", "分類": "存在しない"})
        assert "customField_5" not in params
        assert "存在しない" in capsys.readouterr().err

    def test_空セルは送信しない(self, cf_master):
        m = mapper_for({"custom_fields": [{"field_name": "カテゴリ", "col_name": "分類"}]},
                       cf_master)
        assert "customField_5" not in m.map_row({"件名": "t", "分類": ""})
