"""
継続行の結合
============
1 件の内容が Excel の複数行に分かれている表を、1 件にまとめる。

セル内に収まらない内容を次の行へ書き足す運用があり、その行は
required_cols が空になっている。そのままでは必須列が空の行として
スキップされ、書き足した内容が失われる。

    | 項番 | 件名           | 対応内容              |
    |------|----------------|-----------------------|
    | 1    | ログイン不具合 | 手順1: 再現条件を確認 |  ← 1 件目
    | 1    |                | 手順2: ログを採取     |  ← 続き（項番が同じ）
    |      |                | 手順3: 設定を修正     |  ← 続き（空欄）
    | 2    | 表示崩れ       | CSS を修正            |  ← 2 件目（項番が変わる）

    判定に使う列（required_cols）が「空」または「直前の行と同じ値」なら
    続きとみなす。項番などを継続行にも振っている表に対応するため。
    書式で文字色を変えて見えなくしている場合もあり、目視では空に見えても
    値が入っていることがある。

結合は絞り込み（filters）より前に行う。続きの行は絞り込み条件の列も
空になっているため、先に絞ると結合前に失われてしまう。
"""

from __future__ import annotations

import sys

from excel_reader import ExcelReader

# 結合した値をつなぐ区切り。Markdown で段落として分かれるよう空行を挟む。
JOIN_SEPARATOR = "\n\n"


def is_continuation(row: dict, previous: dict | None, required_cols: list[str]) -> bool:
    """
    その行が直前の行の続きか判定する。

    required_cols に指定した列が、すべて
      ・空である、または
      ・直前の行と同じ値である
    とき、続きの行とみなす。

    「空であること」だけを条件にすると、項番などを継続行にも振っている表
    （書式で見えなくしている場合もある）を扱えない。同じ値なら同じ 1 件を
    指しているとみなす。

    いずれかの列で値が変われば新しい 1 件の開始となる。
    """
    if not required_cols or previous is None:
        return False

    for col in required_cols:
        value = str(row.get(col, "")).strip()
        if not value:
            continue                      # 空欄は継続とみなす
        if value != str(previous.get(col, "")).strip():
            return False                  # 値が変わった＝新しい 1 件
    return True


def merge_continuation_rows(
    rows: list[dict],
    headers: list[str],
    required_cols: list[str],
    single_value_cols: set[str],
) -> list[dict]:
    """
    続きの行を直前の行へ結合した新しい行リストを返す。

    Parameters
    ----------
    single_value_cols : set[str]
        1 件につき 1 つの値しか持てない列（件名・期限日・担当者・
        ステータス・カスタム属性など）。ここに続きの行が値を持っていても
        連結できないため、警告して無視する。

    元の行は変更しない。結合後の行の _excel_row は先頭行の番号を保つ
    （実行ログや画面表示から Excel を辿れるようにするため）。
    """
    if not required_cols:
        return rows

    merged: list[dict] = []
    joined: dict[str, list[str]] = {}      # {先頭行: [続きの行, ...]}

    for row in rows:
        if merged and is_continuation(row, merged[-1], required_cols):
            _append_into(merged[-1], row, headers, single_value_cols, required_cols)
            head = merged[-1].get(ExcelReader.ROW_NUMBER_KEY, "?")
            joined.setdefault(head, []).append(
                row.get(ExcelReader.ROW_NUMBER_KEY, "?")
            )
        else:
            merged.append(dict(row))

    _report(rows, merged, joined, required_cols)
    return merged


def _report(
    rows: list[dict],
    merged: list[dict],
    joined: dict[str, list[str]],
    required_cols: list[str],
) -> None:
    """
    どの行がどこへ結合されたかを表示する。

    件数だけでは「なぜ結合されなかったのか」が分からないため、
    対応関係と判定条件を示す。
    """
    if joined:
        print(f"  継続行を結合: {len(rows)} 行 → {len(merged)} 件")
        for head, tail in joined.items():
            print(f"      {head}行目 ← {'、'.join(t + '行目' for t in tail)}")
        return

    print(
        f"  ℹ merge_continuation_rows が有効ですが、結合対象の行はありません"
        f"（{len(rows)} 行）。\n"
        f"    継続行と判定するのは、次の列がすべて「空」または"
        f"「直前の行と同じ値」の行です: "
        f"{'、'.join(required_cols)}",
        file=sys.stderr,
    )


def _append_into(
    target: dict,
    extra: dict,
    headers: list[str],
    single_value_cols: set[str],
    key_cols: list[str],
) -> None:
    """
    extra の内容を target へ連結する（target を直接更新）。

    判定に使った列（key_cols）は先頭行の値を保つ。同じ値が入っている前提の
    ため連結すると「1\n\n1」のようになり、次の行との比較が壊れる。
    """
    target_row = target.get(ExcelReader.ROW_NUMBER_KEY, "?")
    extra_row = extra.get(ExcelReader.ROW_NUMBER_KEY, "?")

    ignored: list[str] = []
    for column in dict.fromkeys(headers):
        value = str(extra.get(column, "")).strip()
        if not value:
            continue
        if column in key_cols:
            continue                      # 同じ値の前提。先頭行を保つ
        if column in single_value_cols:
            # 1 件につき 1 つしか持てない列は連結できない
            ignored.append(column)
            continue
        current = target.get(column, "")
        target[column] = f"{current}{JOIN_SEPARATOR}{value}" if current else value

    if ignored:
        print(
            f"  ⚠ {extra_row}行目は{target_row}行目の続きとして扱いますが、"
            f"次の列の値は連結できないため無視します: {'/ '.join(ignored)}",
            file=sys.stderr,
        )

    _merge_cell_values(target, extra, headers, single_value_cols, key_cols)


def _merge_cell_values(
    target: dict,
    extra: dict,
    headers: list[str],
    single_value_cols: set[str],
    key_cols: list[str],
) -> None:
    """
    本文生成に使う列順の値リストも同じ規則で連結する。

    同名の列を本文に両方出力するため、行データの dict とは別に
    列順の値リストを持っている（ExcelReader.CELL_VALUES_KEY）。
    """
    key = ExcelReader.CELL_VALUES_KEY
    base = target.get(key)
    added = extra.get(key)
    if base is None or added is None or len(base) != len(added):
        return

    values = list(base)
    for i, value in enumerate(added):
        value = str(value).strip()
        if not value or headers[i] in single_value_cols or headers[i] in key_cols:
            continue
        values[i] = f"{values[i]}{JOIN_SEPARATOR}{value}" if values[i] else value
    target[key] = values


def single_value_columns(source_cfg: dict) -> set[str]:
    """
    1 件につき 1 つの値しか持てない列名を集める。

    件名・日付・担当者・ステータス・カスタム属性・issueKey の列が対象。
    これらは続きの行に値があっても連結できない。
    """
    mapping = source_cfg.get("issue_mapping") or {}
    upsert = source_cfg.get("upsert") or {}

    cols = {
        mapping.get("summary_col"),
        mapping.get("assignee_col"),
        mapping.get("status_col"),
        upsert.get("key_col"),
    }
    for key in ("due_date_col", "start_date_col"):
        value = mapping.get(key)
        # テンプレート指定（{{列名}}）は個別の列を指さないため対象外
        if value and "{{" not in str(value):
            cols.add(value)
    for cf in mapping.get("custom_fields") or []:
        cols.add(cf.get("col_name"))

    return {c for c in cols if isinstance(c, str) and c}
