"""
同名ヘッダーの本文出力
======================
ヘッダー名は Excel の表記のまま（連番なし）。
本文（description）には同名の列も**すべて**出力する。
列名で参照したとき（summary_col / filters など）は左端の列を使う。

行データは {ヘッダー名: 値} の dict のため同名の列を保持できない。
ExcelReader が列順の値リスト（_excel_cells）を別に持たせることで、
本文だけは全列を出力できるようにしている。
"""

import excel_to_backlog as etb
from conftest import FakeBacklog
from excel_reader import ExcelReader
from mapper import IssueMapper


def mapper_for(master, headers, **cfg):
    base = {"issue_type": "タスク", "priority": "中", "summary_col": "件名",
            "description_format": "auto"}
    base.update(cfg)
    return IssueMapper(base, master, headers=headers)


class TestBothColumnsInDescription:
    def test_同名の列が両方出力される(self, make_excel, master):
        path = make_excel(
            ["件名", "備考", "対応内容", "備考"],
            [["課題A", "B列の備考", "調査中", "D列の備考"]],
        )
        headers, rows = ExcelReader({"path": str(path)}).read()

        body = mapper_for(master, headers).map_row(rows[0])["description"]

        assert "B列の備考" in body
        assert "D列の備考" in body
        assert body.count("# 備考") == 2       # 見出しも2回

    def test_列の並び順が保たれる(self, make_excel, master):
        path = make_excel(["件名", "備考", "対応内容", "備考"],
                          [["課題A", "左", "中央", "右"]])
        headers, rows = ExcelReader({"path": str(path)}).read()

        body = mapper_for(master, headers).map_row(rows[0])["description"]

        assert body.index("左") < body.index("中央") < body.index("右")

    def test_3つ以上でもすべて出力される(self, make_excel, master):
        path = make_excel(["件名", "備考", "備考", "備考"],
                          [["課題A", "1つ目", "2つ目", "3つ目"]])
        headers, rows = ExcelReader({"path": str(path)}).read()

        body = mapper_for(master, headers).map_row(rows[0])["description"]

        for v in ["1つ目", "2つ目", "3つ目"]:
            assert v in body

    def test_description_cols_で絞っても同名列は両方出る(self, make_excel, master):
        path = make_excel(["件名", "備考", "対応内容", "備考"],
                          [["課題A", "左", "出さない", "右"]])
        headers, rows = ExcelReader({"path": str(path)}).read()

        body = mapper_for(master, headers,
                          description_cols=["備考"]).map_row(rows[0])["description"]

        assert "左" in body and "右" in body
        assert "出さない" not in body

    def test_取り消し線も列ごとに反映される(self, make_rich_excel, master):
        path = make_rich_excel(
            ["件名", "備考", "備考"],
            [["課題A", [("左", True)], [("右", False)]]],
        )
        headers, plain, formatted = ExcelReader(
            {"path": str(path)}
        ).read_with_format()

        body = mapper_for(master, headers).map_row(
            plain[0], formatted_row=formatted[0]
        )["description"]

        assert "~~左~~" in body
        assert "右" in body and "~~右~~" not in body


class TestLookupUsesLeftmost:
    """列名での参照は左端の列。件名・フィルタ・カスタム属性などが対象。"""

    def test_件名は左端の列(self, make_excel, master):
        path = make_excel(["件名", "件名"], [["左", "右"]])
        _, rows = ExcelReader({"path": str(path)}).read()

        params = IssueMapper(
            {"issue_type": "タスク", "priority": "中", "summary_col": "件名"},
            master, headers=["件名", "件名"],
        ).map_row(rows[0])

        assert params["summary"] == "左"

    def test_フィルタも左端の列で判定する(self, make_excel, master):
        path = make_excel(["件名", "状態", "状態"],
                          [["課題A", "対応要", "完了"], ["課題B", "完了", "対応要"]])
        headers, rows = ExcelReader({"path": str(path)}).read()

        got = etb.apply_filters(
            rows, {"filters": [{"col_name": "状態", "value": "対応要"}]}, headers
        )

        assert [r["件名"] for r in got] == ["課題A"]


class TestHeaderNamesUnchanged:
    def test_連番を付けない(self, make_excel):
        path = make_excel(["備考", "件名", "備考"], [["A", "課題", "B"]])
        headers, _ = ExcelReader({"path": str(path)}).read()

        assert headers == ["備考", "件名", "備考"]
        assert not any("(2)" in h for h in headers)

    def test_メタキーは本文に出ない(self, make_excel, master):
        path = make_excel(["件名", "備考"], [["課題A", "メモ"]])
        headers, rows = ExcelReader({"path": str(path)}).read()

        body = mapper_for(master, headers).map_row(rows[0])["description"]

        assert "_excel_row" not in body
        assert "_excel_cells" not in body
