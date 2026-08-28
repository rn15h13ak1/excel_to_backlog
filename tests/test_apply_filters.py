"""
apply_filters のテスト
======================
filters は AND、filter_groups はグループ内 AND・グループ間 OR。
重複除去を伴うため間違えやすいが、これまで未検証だった。
"""

import excel_to_backlog as etb

HEADERS = ["項番", "枝番", "ステータス"]
ROWS = [
    {"項番": "1", "枝番": "A", "ステータス": "対応要"},
    {"項番": "2", "枝番": "B", "ステータス": "完了"},
    {"項番": "3", "枝番": "B", "ステータス": "対応要"},
    {"項番": "4", "枝番": "",  "ステータス": "完了"},
]


def keys(rows):
    """比較しやすいよう項番のリストにする。"""
    return [r["項番"] for r in rows]


class TestNoFilter:
    def test_条件がなければ全行(self):
        assert keys(etb.apply_filters(ROWS, {}, HEADERS)) == ["1", "2", "3", "4"]

    def test_空のリストでも全行(self):
        cfg = {"filters": [], "filter_groups": []}
        assert keys(etb.apply_filters(ROWS, cfg, HEADERS)) == ["1", "2", "3", "4"]


class TestFilters:
    def test_単一条件(self):
        cfg = {"filters": [{"col_name": "ステータス", "value": "対応要"}]}
        assert keys(etb.apply_filters(ROWS, cfg, HEADERS)) == ["1", "3"]

    def test_複数条件は_AND(self):
        cfg = {"filters": [
            {"col_name": "ステータス", "value": "対応要"},
            {"col_name": "枝番", "value": "B"},
        ]}
        assert keys(etb.apply_filters(ROWS, cfg, HEADERS)) == ["3"]

    def test_元の順序を保つ(self):
        cfg = {"filters": [{"col_name": "ステータス", "value": "完了"}]}
        assert keys(etb.apply_filters(ROWS, cfg, HEADERS)) == ["2", "4"]


class TestFilterGroups:
    """グループ内は AND、グループ間は OR。"""

    def test_複合キーで特定行を指定できる(self):
        cfg = {"filter_groups": [
            {"filters": [{"col_name": "項番", "value": "1"},
                         {"col_name": "枝番", "value": "A"}]},
            {"filters": [{"col_name": "項番", "value": "3"},
                         {"col_name": "枝番", "value": "B"}]},
        ]}
        assert keys(etb.apply_filters(ROWS, cfg, HEADERS)) == ["1", "3"]

    def test_グループ内は_AND_なので片方でも外れると除外(self):
        cfg = {"filter_groups": [
            {"filters": [{"col_name": "項番", "value": "1"},
                         {"col_name": "枝番", "value": "B"}]},   # 項番1 は枝番 A
        ]}
        assert keys(etb.apply_filters(ROWS, cfg, HEADERS)) == []

    def test_複数グループにマッチしても重複しない(self):
        """同じ行が2つのグループの条件を満たすケース。"""
        cfg = {"filter_groups": [
            {"filters": [{"col_name": "ステータス", "value": "対応要"}]},
            {"filters": [{"col_name": "枝番", "value": "A"}]},   # 項番1 は両方に該当
        ]}
        assert keys(etb.apply_filters(ROWS, cfg, HEADERS)) == ["1", "3"]

    def test_グループの順に結果が並ぶ(self):
        """
        グループ1 の結果が先、続いてグループ2 の結果。
        元の行順ではなくグループ順になる点に注意。
        """
        cfg = {"filter_groups": [
            {"filters": [{"col_name": "項番", "value": "3"}]},
            {"filters": [{"col_name": "項番", "value": "1"}]},
        ]}
        assert keys(etb.apply_filters(ROWS, cfg, HEADERS)) == ["3", "1"]

    def test_条件が空のグループは全行にマッチする(self):
        cfg = {"filter_groups": [{"filters": []}]}
        assert keys(etb.apply_filters(ROWS, cfg, HEADERS)) == ["1", "2", "3", "4"]

    def test_filters_より_filter_groups_が優先される(self):
        cfg = {
            "filters": [{"col_name": "ステータス", "value": "完了"}],
            "filter_groups": [{"filters": [{"col_name": "項番", "value": "1"}]}],
        }
        assert keys(etb.apply_filters(ROWS, cfg, HEADERS)) == ["1"]


class TestIdentityHandling:
    """
    重複除去は id(row) で行っている。内容が同一でも別オブジェクトなら
    別の行として扱われることを明示する（行の重複を勝手に潰さない）。
    """

    def test_内容が同一の行はどちらも残る(self):
        rows = [
            {"項番": "1", "枝番": "A", "ステータス": "対応要"},
            {"項番": "1", "枝番": "A", "ステータス": "対応要"},   # 別オブジェクト
        ]
        cfg = {"filter_groups": [{"filters": [{"col_name": "項番", "value": "1"}]}]}
        assert len(etb.apply_filters(rows, cfg, HEADERS)) == 2

    def test_返る行は元のオブジェクトそのもの(self):
        """
        書式付き行との対応付けが id() に依存しているため、
        フィルタがコピーを返すと対応が取れなくなる。
        """
        cfg = {"filters": [{"col_name": "項番", "value": "1"}]}
        result = etb.apply_filters(ROWS, cfg, HEADERS)
        assert result[0] is ROWS[0]


class TestMissingColumnWarning:
    def test_存在しない列は警告を出す(self, capsys):
        cfg = {"filters": [{"col_name": "無い列", "value": "x"}]}
        etb.apply_filters(ROWS, cfg, HEADERS)
        assert "無い列" in capsys.readouterr().err

    def test_filter_groups_でも警告を出す(self, capsys):
        cfg = {"filter_groups": [{"filters": [{"col_name": "無い列", "value": "x"}]}]}
        etb.apply_filters(ROWS, cfg, HEADERS)
        err = capsys.readouterr().err
        assert "無い列" in err
        assert "filter_groups[0]" in err
