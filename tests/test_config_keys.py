"""
設定キーの検証テスト
====================
列名は検証していたが、キー名は誰も見ていなかった。設定全体が dict.get() の
既定値に依存しているため、綴りを間違えても静かに無視される。
"""

import pytest

from config_validation import validate_config_keys, validate_source_keys


def base_source(**overrides):
    cfg = {
        "name": "S",
        "excel": {"path": "a.xlsx"},
        "issue_mapping": {"issue_type": "タスク", "summary_col": "件名"},
    }
    cfg.update(overrides)
    return cfg


class TestValidConfig:
    def test_正しい設定は通る(self):
        assert validate_source_keys(base_source()) == []

    def test_全項目を使った設定も通る(self):
        cfg = base_source(
            filters=[{"col_name": "状態", "value": "x", "match": "contains"}],
            upsert={"enabled": True, "key_col": "番号", "match_summary": True},
            issue_mapping={
                "issue_type": "タスク", "priority": "中",
                "summary_template": "{{件名}}", "description_format": "auto",
                "description_cols": ["概要"], "due_date_col": "期限",
                "start_date_col": "開始", "assignee_col": "担当",
                "default_assignee": "yamada", "required_cols": ["件名"],
                "status_col": "状態", "status_map": {"完了": "完了"},
                "rich_text": True,
                "custom_fields": [{"field_name": "分類", "col_name": "区分",
                                   "value_map": {"A": "B"}, "value_separator": ","}],
            },
        )
        assert validate_source_keys(cfg) == []


class TestUnknownKeys:
    @pytest.mark.parametrize("mapping_key", ["sumary_col", "asignee_col", "statu_col"])
    def test_issue_mapping_の綴り間違いを検出する(self, mapping_key):
        cfg = base_source(issue_mapping={"issue_type": "タスク", mapping_key: "x"})
        problems = validate_source_keys(cfg)
        assert any(mapping_key in p for p in problems)

    def test_近い名前を提案する(self):
        cfg = base_source(issue_mapping={"sumary_col": "件名"})
        assert "summary_col" in validate_source_keys(cfg)[0]

    def test_大文字小文字の違いも提案する(self):
        cfg = base_source(issue_mapping={"Summary_Col": "件名"})
        assert "summary_col" in validate_source_keys(cfg)[0]

    def test_excel_の未知キーを検出する(self):
        cfg = base_source(excel={"path": "a.xlsx", "sheet_name": "Sheet1"})
        assert any("sheet_name" in p for p in validate_source_keys(cfg))

    def test_upsert_の未知キーを検出する(self):
        cfg = base_source(upsert={"enabled": True, "key_column": "番号"})
        assert any("key_column" in p for p in validate_source_keys(cfg))

    def test_ソース直下の未知キーを検出する(self):
        cfg = base_source(fillters=[])
        assert any("fillters" in p for p in validate_source_keys(cfg))

    def test_カスタム属性の未知キーを検出する(self):
        cfg = base_source(issue_mapping={
            "custom_fields": [{"field_name": "分類", "column_name": "区分"}]
        })
        assert any("column_name" in p for p in validate_source_keys(cfg))

    def test_backlog_の未知キーを検出する(self):
        config = {"backlog": {"space_host": "x", "apikey": "y"}, "sources": []}
        assert any("apikey" in p for p in validate_config_keys(config))


class TestEnumValues:
    def test_match_の誤った値を検出する(self):
        """contain は黙って exact になり、対象行が変わってしまう。"""
        cfg = base_source(filters=[{"col_name": "状態", "value": "x", "match": "contain"}])
        problems = validate_source_keys(cfg)
        assert any("contain" in p and "contains" in p for p in problems)

    def test_description_format_の誤った値を検出する(self):
        cfg = base_source(issue_mapping={"description_format": "Auto"})
        assert any("Auto" in p for p in validate_source_keys(cfg))

    def test_filter_groups_内の_match_も検証する(self):
        cfg = base_source(filter_groups=[
            {"filters": [{"col_name": "状態", "value": "x", "match": "startwith"}]}
        ])
        assert any("startwith" in p for p in validate_source_keys(cfg))


class TestUpsertWithoutMethod:
    def test_判定方法が無い_upsert_を警告する(self):
        """常に新規作成になるが、以前は何も言わなかった。"""
        cfg = base_source(upsert={"enabled": True})
        assert any("常に新規作成" in p for p in validate_source_keys(cfg))

    def test_key_col_があれば警告しない(self):
        cfg = base_source(upsert={"enabled": True, "key_col": "番号"})
        assert validate_source_keys(cfg) == []

    def test_upsert_無効なら警告しない(self):
        cfg = base_source(upsert={"enabled": False})
        assert validate_source_keys(cfg) == []


class TestMultipleSources:
    def test_ソース番号が示される(self):
        config = {"sources": [base_source(), base_source(issue_mapping={"sumary_col": "x"})]}
        problems = validate_config_keys(config)
        assert any("sources[1]" in p for p in problems)


class TestMalformedConfig:
    """
    YAML の書き間違いでも、トレースバックではなく設定の問題として説明する。
    設定を説明するための機構が設定のせいで落ちては意味がない。
    """

    def test_sources_の項目が空なら報告する(self):
        """
        キーだけ書いて中身が空だと .get(key, {}) の既定値が効かない。
        素通りさせると実行時に AttributeError になる。
        """
        problems = validate_config_keys({"sources": [None]})
        assert any("中身が空" in p for p in problems)

    def test_issue_mapping_の中身が空なら報告する(self):
        config = {"sources": [{"name": "S", "excel": {"path": "a.xlsx"},
                               "issue_mapping": None}]}
        assert any("issue_mapping" in p and "中身が空" in p
                   for p in validate_config_keys(config))

    def test_excel_の中身が空なら報告する(self):
        config = {"sources": [{"name": "S", "excel": None,
                               "issue_mapping": {"summary_col": "件名"}}]}
        assert any("excel" in p and "中身が空" in p
                   for p in validate_config_keys(config))

    def test_issue_mapping_をリストで書いた場合(self):
        config = {"sources": [{"name": "S", "issue_mapping": [{"summary_col": "件名"}]}]}
        problems = validate_config_keys(config)
        assert any("issue_mapping" in p and "キー: 値" in p for p in problems)

    def test_filter_groups_をスカラーで書いた場合(self):
        config = {"sources": [{"name": "S", "filter_groups": "abc"}]}
        problems = validate_config_keys(config)
        assert any("filter_groups" in p and "一覧" in p for p in problems)

    def test_filters_をスカラーで書いた場合(self):
        config = {"sources": [{"name": "S", "filters": "abc"}]}
        assert any("filters" in p for p in validate_config_keys(config))

    def test_custom_fields_の項目がスカラー(self):
        config = {"sources": [{"name": "S",
                               "issue_mapping": {"custom_fields": ["カテゴリ"]}}]}
        assert validate_config_keys(config)          # 落ちずに説明する

    def test_match_にリストを書いた場合(self):
        config = {"sources": [{"name": "S",
                               "filters": [{"col_name": "x", "match": ["exact"]}]}]}
        assert any("match" in p for p in validate_config_keys(config))

    def test_backlog_をリストで書いた場合(self):
        assert validate_config_keys({"backlog": ["x"], "sources": []})

    def test_設定全体が_dict_でない場合(self):
        assert validate_config_keys([]) or validate_config_keys([]) == []


class TestFilterGroupKeys:
    """
    グループ自体のキーを検証しないと、filters の綴り間違いで条件が空になり
    ExcelReader.filter_rows(rows, None) が全行を返す。シート全体が無警告で
    登録対象になる。
    """

    def test_グループのキー名の間違いを検出する(self):
        cfg = base_source(filter_groups=[
            {"filtres": [{"col_name": "状態", "value": "対応要"}]}
        ])
        problems = validate_source_keys(cfg)
        assert any("filtres" in p and "filters" in p for p in problems)

    def test_正しいグループは通る(self):
        cfg = base_source(filter_groups=[
            {"filters": [{"col_name": "状態", "value": "対応要"}]}
        ])
        assert validate_source_keys(cfg) == []


class TestSuggestionIsStable:
    """
    set を走査すると候補が複数該当したとき選ばれる名前が実行ごとに変わる。
    同じ設定に対して毎回違う修正案が出ると混乱する。
    """

    def test_同じ入力には同じ提案を返す(self):
        from config_validation import ISSUE_MAPPING_KEYS, _closest
        assert len({_closest("status_ma", ISSUE_MAPPING_KEYS) for _ in range(20)}) == 1

    def test_より近い候補を選ぶ(self):
        from config_validation import ISSUE_MAPPING_KEYS, _closest
        assert _closest("status_ma", ISSUE_MAPPING_KEYS) == "status_map"
        assert _closest("status_co", ISSUE_MAPPING_KEYS) == "status_col"
