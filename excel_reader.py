"""
Excel 読み込み・フィルタリングモジュール
=========================================
excel_md_tool (useExcel.ts) と同等のロジックを Python / openpyxl で実装。

対応仕様:
  - ヘッダー行の範囲指定（複数行ヘッダーは " / " 結合）
  - データ開始行の指定
  - 列範囲の指定（A, B, C ... の列ラベル形式）
  - フィルター条件（列名＋値の AND フィルター、value: 単一値 / values: 複数値 OR）
  - 日付セルは "YYYY/MM/DD" 文字列に変換
  - 空行はスキップ
"""

from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

try:
    import openpyxl
    from openpyxl.utils import column_index_from_string
except ImportError:
    raise ImportError(
        "openpyxl が必要です。`pip install openpyxl` を実行してください。"
    )


# ------------------------------------------------------------------
# ユーティリティ
# ------------------------------------------------------------------

def col_letter_to_index(letter: str) -> int:
    """列ラベル（"A", "B", ..., "AA"）を 0 始まりの整数インデックスに変換"""
    return column_index_from_string(letter.upper()) - 1


def cell_to_str(value: Any) -> str:
    """セル値を文字列に変換（None/空は空文字列、日付は YYYY/MM/DD）"""
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        if isinstance(value, datetime):
            return value.strftime("%Y/%m/%d")
        return value.strftime("%Y/%m/%d")
    return str(value).strip()


# ------------------------------------------------------------------
# ExcelReader
# ------------------------------------------------------------------

class ExcelReader:
    """
    config (sources[i].excel) に従って Excel を読み込み、
    ヘッダー名をキーにした行データのリストを返す。

    Parameters
    ----------
    excel_config : dict
        path            : str   Excel ファイルパス
        sheet           : str   シート名（省略時: 最初のシート）
        header_start_row: int   ヘッダー開始行（1始まり、デフォルト 1）
        header_end_row  : int   ヘッダー終了行（デフォルト = header_start_row）
        data_start_row  : int   データ開始行（デフォルト = header_end_row + 1）
        col_start       : str   読み込み開始列ラベル（デフォルト "A"）
        col_end         : str   読み込み終了列ラベル（デフォルト: シート最終列）
    """

    MULTI_HEADER_SEP = " / "  # 複数行ヘッダーの結合区切り文字

    def __init__(self, excel_config: dict):
        self.path = Path(excel_config["path"]).expanduser()
        self.sheet_name: str | None = excel_config.get("sheet")
        self.header_start_row: int = int(excel_config.get("header_start_row", 1))
        self.header_end_row: int = int(
            excel_config.get("header_end_row", self.header_start_row)
        )
        self.data_start_row: int = int(
            excel_config.get("data_start_row", self.header_end_row + 1)
        )
        self.col_start_str: str = str(excel_config.get("col_start", "A")).upper()
        self.col_end_str: str | None = (
            str(excel_config.get("col_end")).upper()
            if excel_config.get("col_end")
            else None
        )

        # バリデーション
        if self.header_start_row < 1:
            raise ValueError("header_start_row は 1 以上を指定してください。")
        if self.header_end_row < self.header_start_row:
            raise ValueError("header_end_row は header_start_row 以上を指定してください。")
        if self.data_start_row <= self.header_end_row:
            raise ValueError("data_start_row は header_end_row より大きい値を指定してください。")

    # ------------------------------------------------------------------

    def _load_sheet(self):
        """ワークブックを開いてシートを返す"""
        if not self.path.exists():
            raise FileNotFoundError(f"Excel ファイルが見つかりません: {self.path}")

        wb = openpyxl.load_workbook(str(self.path), data_only=True)

        if self.sheet_name:
            if self.sheet_name not in wb.sheetnames:
                available = ", ".join(wb.sheetnames)
                raise ValueError(
                    f"シート「{self.sheet_name}」が見つかりません。"
                    f"（利用可能: {available}）"
                )
            ws = wb[self.sheet_name]
        else:
            ws = wb.active

        return ws

    def _resolve_col_range(self, ws) -> tuple[int, int]:
        """
        col_start / col_end を 0 始まりの列インデックスに解決する。
        col_end が未指定の場合はシートの最終列を使う。
        """
        col_start_idx = col_letter_to_index(self.col_start_str)

        if self.col_end_str:
            col_end_idx = col_letter_to_index(self.col_end_str)
        else:
            # シートの最終列（max_column は 1 始まり）
            col_end_idx = (ws.max_column or 1) - 1

        if col_start_idx > col_end_idx:
            raise ValueError(
                f"col_start ({self.col_start_str}) は col_end ({self.col_end_str}) より前の列を指定してください。"
            )
        return col_start_idx, col_end_idx

    def _build_headers(self, ws, col_start_idx: int, col_end_idx: int) -> list[str]:
        """
        ヘッダー行（header_start_row ～ header_end_row）を読み取り、
        複数行の場合は MULTI_HEADER_SEP で結合したヘッダー名リストを返す。
        """
        headers = []
        for col_idx in range(col_start_idx, col_end_idx + 1):
            parts = []
            for row_idx in range(self.header_start_row, self.header_end_row + 1):
                # openpyxl は 1 始まり
                cell = ws.cell(row=row_idx, column=col_idx + 1)
                val = cell_to_str(cell.value)
                if val:
                    parts.append(val)
            headers.append(self.MULTI_HEADER_SEP.join(parts) if parts else f"Col{col_idx + 1}")
        return headers

    def _build_rows(
        self, ws, headers: list[str], col_start_idx: int, col_end_idx: int
    ) -> list[dict[str, str]]:
        """
        データ行（data_start_row ～ シート最終行）を読み取り、
        {ヘッダー名: セル値} の dict リストを返す。
        空行（全セルが空）はスキップする。
        """
        rows = []
        max_row = ws.max_row or 0

        for row_idx in range(self.data_start_row, max_row + 1):
            row_data = {}
            is_empty = True
            for i, col_idx in enumerate(range(col_start_idx, col_end_idx + 1)):
                cell = ws.cell(row=row_idx, column=col_idx + 1)
                val = cell_to_str(cell.value)
                if val:
                    is_empty = False
                row_data[headers[i]] = val

            if not is_empty:
                rows.append(row_data)

        return rows

    # ------------------------------------------------------------------

    def read(self) -> tuple[list[str], list[dict[str, str]]]:
        """
        Excel ファイルを読み込んでヘッダーと行データを返す。

        Returns
        -------
        headers : list[str]
            ヘッダー名リスト（複数行ヘッダーは " / " 結合）
        rows    : list[dict[str, str]]
            データ行リスト。各要素は {ヘッダー名: 値} の dict
        """
        ws = self._load_sheet()
        col_start_idx, col_end_idx = self._resolve_col_range(ws)
        headers = self._build_headers(ws, col_start_idx, col_end_idx)
        rows = self._build_rows(ws, headers, col_start_idx, col_end_idx)
        return headers, rows

    @staticmethod
    def filter_rows(
        rows: list[dict[str, str]],
        filters: list[dict] | None,
    ) -> list[dict[str, str]]:
        """
        AND 条件でフィルタリングして合致する行だけを返す。

        filters の各要素:
            col_name : str        対象のヘッダー名
            value    : str        一致すべき単一値（values と排他）
            values   : list[str]  一致すべき値のリスト（いずれかに一致すれば OK、OR 条件）
            match    : str        "exact"（デフォルト）/ "contains" / "startswith"
                                  value / values いずれにも適用される

        複数の filter 条件は AND で評価される。
        filters が None または空の場合は全行を返す。
        存在しない col_name はその条件をスキップする（警告は呼び出し元で出す）。
        """
        if not filters:
            return rows

        result = []
        for row in rows:
            match_all = True
            for cond in filters:
                col = cond.get("col_name", "")
                match_type = cond.get("match", "exact")

                if col not in row:
                    # 条件列がヘッダーに存在しない場合はスキップ
                    continue

                actual = row[col]

                # values（リスト）が指定されていれば OR マッチ、なければ単一 value でマッチ
                raw_values = cond.get("values")
                if raw_values is not None:
                    candidates = [str(v) for v in raw_values]
                else:
                    candidates = [str(cond.get("value", ""))]

                if match_type == "contains":
                    matched = any(c in actual for c in candidates)
                elif match_type == "startswith":
                    matched = any(actual.startswith(c) for c in candidates)
                else:  # exact
                    matched = actual in candidates

                if not matched:
                    match_all = False
                    break

            if match_all:
                result.append(row)

        return result
