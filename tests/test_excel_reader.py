"""
ExcelReader のテスト
====================
取り消し線（cell_to_markdown）は継承ロジックの追加・削除・再追加を繰り返して
きた箇所のため、往復した4パターンをすべて固定する。
"""

import pytest
from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

from excel_reader import ExcelReader, cell_to_markdown, cell_to_str


# ------------------------------------------------------------------
# ヘッダーの一意化
# ------------------------------------------------------------------

class TestUniquifyHeaders:
    """重複ヘッダーは列の内容を失わせるため、連番で区別する。"""

    def test_重複がなければそのまま返す(self):
        assert ExcelReader._uniquify_headers(["項番", "件名"]) == ["項番", "件名"]

    def test_重複した2つ目に連番が付く(self):
        assert ExcelReader._uniquify_headers(["備考", "件名", "備考"]) == [
            "備考", "件名", "備考 (2)"
        ]

    def test_3つ以上の重複でも連番が続く(self):
        assert ExcelReader._uniquify_headers(["備考"] * 3) == [
            "備考", "備考 (2)", "備考 (3)"
        ]

    def test_連番付きの名前が既にあっても衝突しない(self):
        assert ExcelReader._uniquify_headers(["備考", "備考 (2)", "備考"]) == [
            "備考", "備考 (2)", "備考 (3)"
        ]

    def test_1つ目は元の名前のまま残る(self):
        """既存の設定が左端の列を参照し続けられること。"""
        result = ExcelReader._uniquify_headers(["備考", "備考"])
        assert result[0] == "備考"


def test_重複ヘッダーでも両方の列の値を読める(tmp_path):
    """dict の後勝ちで C 列の内容が失われていたバグの回帰テスト。"""
    wb = Workbook()
    ws = wb.active
    for col, header in zip("ABCD", ["項番", "件名", "備考", "備考"]):
        ws[f"{col}1"] = header
    ws["A2"], ws["B2"] = 1, "ログイン不具合"
    ws["C2"], ws["D2"] = "C列の内容", "D列の内容"
    path = tmp_path / "dup.xlsx"
    wb.save(path)

    headers, rows = ExcelReader({"path": str(path)}).read()

    assert headers == ["項番", "件名", "備考", "備考 (2)"]
    assert rows[0]["備考"] == "C列の内容"
    assert rows[0]["備考 (2)"] == "D列の内容"


# ------------------------------------------------------------------
# セル値の文字列化
# ------------------------------------------------------------------

class TestCellToStr:
    def test_None_は空文字列(self):
        assert cell_to_str(None) == ""

    def test_整数はそのまま(self):
        """openpyxl は整数セルを int で返すため "1.0" にはならない。"""
        assert cell_to_str(1) == "1"

    def test_小数は小数のまま(self):
        assert cell_to_str(0.5) == "0.5"

    def test_前後の空白は除去される(self):
        assert cell_to_str("  値  ") == "値"


# ------------------------------------------------------------------
# 取り消し線 → Markdown
# ------------------------------------------------------------------

class _FakeFont:
    def __init__(self, strike):
        self.strike = strike


class _FakeCell:
    """
    cell_to_markdown() が参照するのは .value と .font.strike だけのため、
    その2つだけを持つ最小のセルを使う。

    openpyxl 3.1 は CellRichText を含むブックの保存に失敗するため、
    実ファイルを往復させる方式は使えない。値そのものは openpyxl の
    CellRichText / TextBlock を使って組み立てるので、読み込み時に得られる
    構造と同じものを検証できる。
    """

    def __init__(self, value, *, strike=False):
        self.value = value
        self.font = _FakeFont(strike)


def _cell(value, *, strike=False):
    """セルレベルの取り消し線を持つ通常セルを作る。"""
    return _FakeCell(value, strike=strike)


def _rich_cell(runs, *, cell_strike=False):
    """
    CellRichText セルを作る。

    runs: [(テキスト, 取り消し線 or None), ...]
          None を渡すと <rPr> を持たないプレーン文字列ランになる
          （openpyxl は書式指定のないランを素の str で返す）。
    """
    blocks = [
        text if struck is None else TextBlock(InlineFont(strike=struck), text)
        for text, struck in runs
    ]
    return _FakeCell(CellRichText(*blocks), strike=cell_strike)


class TestCellToMarkdown:
    """
    継承ロジックが追加・削除・再追加された経緯があるため、
    判断が分かれた4パターンをすべて固定する。
    """

    def test_空セルは空文字列(self):
        assert cell_to_markdown(_cell(None)) == ""

    def test_書式なしのセルはそのまま(self):
        assert cell_to_markdown(_cell("通常テキスト")) == "通常テキスト"

    def test_セル全体の取り消し線は全体を囲む(self):
        """CellRichText でないセル（プレーンテキスト・日付）が対象。"""
        assert cell_to_markdown(_cell("削除済み", strike=True)) == "~~削除済み~~"

    def test_リッチテキストの一部だけ取り消し線(self):
        cell = _rich_cell([("残る", False), ("消える", True)])
        assert cell_to_markdown(cell) == "残る ~~消える~~"

    def test_リッチテキストではランの情報がセルスタイルより優先される(self):
        """
        セルの一部に取り消し線を付けると Excel はセルレベルにも strike=True を
        記録することがある。ランに明示的な strike=True があるときは、
        プレーン文字列ランへ継承してはいけない。
        """
        cell = _rich_cell([("残る", None), ("消える", True)], cell_strike=True)
        assert cell_to_markdown(cell) == "残る ~~消える~~"

    def test_ランに取り消し線がなければセルスタイルを継承する(self):
        """
        「セル全体に取り消し線 → 一部を解除」の操作では、解除したランだけが
        strike=False の TextBlock になり、残りは <rPr> を省略してセルレベルの
        スタイルを継承する。この場合はプレーン文字列ランも取り消し線扱いにする。
        """
        cell = _rich_cell([("消える", None), ("残る", False)], cell_strike=True)
        assert cell_to_markdown(cell) == "~~消える~~ 残る"

    def test_複数行セルは行ごとに取り消し線を適用する(self):
        """Markdown の ~~ は改行をまたげないため行単位で囲む。"""
        cell = _cell("1行目\n2行目", strike=True)
        assert cell_to_markdown(cell) == "~~1行目~~\n~~2行目~~"


# ------------------------------------------------------------------
# 行のフィルタリング
# ------------------------------------------------------------------

class TestFilterRows:
    ROWS = [
        {"ステータス": "対応要", "種別": "バグ"},
        {"ステータス": "完了", "種別": "バグ"},
        {"ステータス": "完了", "種別": "タスク"},
    ]

    def test_条件なしなら全行(self):
        assert ExcelReader.filter_rows(self.ROWS, []) == self.ROWS

    def test_完全一致(self):
        got = ExcelReader.filter_rows(self.ROWS, [{"col_name": "ステータス", "value": "完了"}])
        assert len(got) == 2

    def test_values_は_OR_条件(self):
        got = ExcelReader.filter_rows(
            self.ROWS, [{"col_name": "種別", "values": ["バグ", "タスク"]}]
        )
        assert len(got) == 3

    def test_複数条件は_AND(self):
        got = ExcelReader.filter_rows(self.ROWS, [
            {"col_name": "ステータス", "value": "完了"},
            {"col_name": "種別", "value": "タスク"},
        ])
        assert len(got) == 1

    def test_部分一致(self):
        got = ExcelReader.filter_rows(
            self.ROWS, [{"col_name": "ステータス", "value": "対応", "match": "contains"}]
        )
        assert len(got) == 1

    def test_存在しない列の条件は無視され全行が通る(self):
        """
        この挙動自体は危険だが、実行前に validate_column_references() が
        列名を検証して停止するため、ここへは到達しない。
        挙動を変える場合は検証側とあわせて見直すこと。
        """
        got = ExcelReader.filter_rows(
            self.ROWS, [{"col_name": "存在しない列", "value": "x"}]
        )
        assert len(got) == 3


# ------------------------------------------------------------------
# 設定のバリデーション
# ------------------------------------------------------------------

class TestExcelConfigValidation:
    def test_header_start_row_は1以上(self):
        with pytest.raises(ValueError, match="header_start_row"):
            ExcelReader({"path": "x.xlsx", "header_start_row": 0})

    def test_header_end_row_は_start_以上(self):
        with pytest.raises(ValueError, match="header_end_row"):
            ExcelReader({"path": "x.xlsx", "header_start_row": 3, "header_end_row": 2})

    def test_data_start_row_は_header_end_row_より大きい(self):
        with pytest.raises(ValueError, match="data_start_row"):
            ExcelReader({"path": "x.xlsx", "header_end_row": 3, "data_start_row": 3})

    def test_存在しないファイルは_FileNotFoundError(self, tmp_path):
        reader = ExcelReader({"path": str(tmp_path / "ない.xlsx")})
        with pytest.raises(FileNotFoundError):
            reader.read()
