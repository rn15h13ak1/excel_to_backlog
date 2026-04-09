"""
取り消し線デバッグスクリプト
==============================
使い方:
    python debug_strike.py <Excelファイルパス> <シート名> <行番号> <列名>

例:
    python debug_strike.py "C:/path/to/book.xlsx" "Sheet1" 2 "詳細"

引数省略時はインタラクティブに入力を求めます。
"""

import sys
from pathlib import Path


def main():
    # 引数またはインタラクティブ入力
    args = sys.argv[1:]
    if len(args) >= 1:
        excel_path = args[0]
    else:
        excel_path = input("Excelファイルのパス: ").strip().strip('"')

    if len(args) >= 2:
        sheet_name = args[1]
    else:
        sheet_name = input("シート名: ").strip()

    if len(args) >= 3:
        row_num = int(args[2])
    else:
        row_num = int(input("確認したい行番号 (データ行、例: 2): ").strip())

    if len(args) >= 4:
        col_name = args[3]
    else:
        col_name = input("確認したい列名 (例: 詳細): ").strip()

    print()
    print("=" * 60)
    print("【チェック1】openpyxl バージョン")
    print("=" * 60)
    import openpyxl
    print(f"openpyxl: {openpyxl.__version__}")

    try:
        from openpyxl.cell.rich_text import CellRichText, TextBlock
        print("CellRichText: 利用可能 ✓")
    except ImportError:
        print("CellRichText: 利用不可 ✗ → 取り消し線は検出できません")
        print("  → openpyxl >= 3.1.0 が必要です")
        return

    print()
    print("=" * 60)
    print("【チェック2】ファイルを rich_text=True で読み込み")
    print("=" * 60)

    try:
        wb = openpyxl.load_workbook(excel_path, rich_text=True, data_only=True)
    except TypeError:
        print("rich_text 引数非対応 ✗ → openpyxl のバージョンを上げてください")
        return
    except FileNotFoundError:
        print(f"ファイルが見つかりません: {excel_path}")
        return

    print(f"ファイル読み込み成功 ✓")
    print(f"シート一覧: {wb.sheetnames}")

    if sheet_name not in wb.sheetnames:
        print(f"シート '{sheet_name}' が見つかりません")
        sheet_name = wb.sheetnames[0]
        print(f"  → 最初のシート '{sheet_name}' を使います")

    ws = wb[sheet_name]

    print()
    print("=" * 60)
    print(f"【チェック3】行 {row_num} の全セルの型を確認")
    print("=" * 60)

    row_cells = list(ws.iter_rows(min_row=row_num, max_row=row_num))[0]
    for cell in row_cells:
        col_letter = cell.column_letter
        val = cell.value
        val_type = type(val).__name__
        print(f"  {col_letter}{row_num}: type={val_type}, value={repr(val)[:80]}")

    print()
    print("=" * 60)
    print(f"【チェック4】列 '{col_name}' のセル詳細（行 {row_num}）")
    print("=" * 60)

    # ヘッダー行から列インデックスを探す（1行目を想定）
    target_cell = None
    col_letter_found = None
    for cell in list(ws.iter_rows(min_row=1, max_row=1))[0]:
        if str(cell.value).strip() == col_name:
            col_letter_found = cell.column_letter
            target_cell = ws.cell(row=row_num, column=cell.column)
            break

    if target_cell is None:
        print(f"ヘッダー行に '{col_name}' が見つかりません")
        print("ヘッダー一覧:")
        for c in list(ws.iter_rows(min_row=1, max_row=1))[0]:
            print(f"  {c.column_letter}: {repr(c.value)}")
    else:
        print(f"列: {col_letter_found}")
        val = target_cell.value
        print(f"値の型: {type(val).__name__}")
        print(f"値: {repr(val)}")

        if isinstance(val, CellRichText):
            print()
            print("  ▼ CellRichText のラン構造:")
            for i, run in enumerate(val):
                if isinstance(run, TextBlock):
                    strike = bool(run.font and run.font.strike)
                    bold = bool(run.font and run.font.bold)
                    print(f"    [{i}] TextBlock: text={repr(run.text)}, strike={strike}, bold={bold}")
                else:
                    print(f"    [{i}] 文字列ラン: text={repr(run)}")
        else:
            # CellRichText でない場合はセルレベルのフォント確認
            font = target_cell.font
            if font:
                print(f"セルフォント strike: {font.strike}")
            else:
                print("セルフォント: None")

    print()
    print("=" * 60)
    print("【チェック5】cell_to_markdown() の出力")
    print("=" * 60)

    from excel_reader import cell_to_markdown

    if target_cell is not None:
        md = cell_to_markdown(target_cell)
        print(f"cell_to_markdown() = {repr(md)}")
        print()
        print("実際の内容（\\n を改行で表示）:")
        print(md)

    print()
    print("=" * 60)
    print("【チェック6】config.yaml の rich_text 設定")
    print("=" * 60)

    config_path = Path("config.yaml")
    if config_path.exists():
        content = config_path.read_text(encoding="utf-8")
        if "rich_text" in content:
            for i, line in enumerate(content.splitlines(), 1):
                if "rich_text" in line:
                    print(f"  行{i}: {line}")
        else:
            print("  rich_text の設定が見つかりません → デフォルト=false（取り消し線なし）")
            print("  → config.yaml に 'rich_text: true' を追加してください")
    else:
        print(f"  config.yaml が見つかりません（{Path.cwd()}）")

    print()
    print("=" * 60)
    print("デバッグ完了")
    print("=" * 60)


if __name__ == "__main__":
    main()
