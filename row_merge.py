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
    |      |                | 手順2: ログを採取     |  ← 続き
    |      |                | 手順3: 設定を修正     |  ← 続き
    | 2    | 表示崩れ       | CSS を修正            |  ← 2 件目

結合は絞り込み（filters）より前に行う。続きの行は絞り込み条件の列も
空になっているため、先に絞ると結合前に失われてしまう。
"""

from __future__ import annotations

import sys

from excel_reader import ExcelReader

# 結合した値をつなぐ区切り。Markdown で段落として分かれるよう空行を挟む。
JOIN_SEPARATOR = "\n\n"


def is_continuation(row: dict, required_cols: list[str]) -> bool:
    """
    その行が直前の行の続きか判定する。

    required_cols に指定された列がすべて空なら続きとみなす。
    （全セルが空の行は ExcelReader が読み込み時点で除外している）
    """
    if not required_cols:
        return False
    return all(not str(row.get(col, "")).strip() for col in required_cols)


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
    for row in rows:
        if merged and is_continuation(row, required_cols):
            _append_into(merged[-1], row, headers, single_value_cols)
        else:
            merged.append(dict(row))

    return merged


def _append_into(
    target: dict,
    extra: dict,
    headers: list[str],
    single_value_cols: set[str],
) -> None:
    """extra の内容を target へ連結する（target を直接更新）。"""
    target_row = target.get(ExcelReader.ROW_NUMBER_KEY, "?")
    extra_row = extra.get(ExcelReader.ROW_NUMBER_KEY, "?")

    ignored: list[str] = []
    for column in dict.fromkeys(headers):
        value = str(extra.get(column, "")).strip()
        if not value:
            continue
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

    _merge_cell_values(target, extra, headers, single_value_cols)


def _merge_cell_values(
    target: dict,
    extra: dict,
    headers: list[str],
    single_value_cols: set[str],
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
        if not value or headers[i] in single_value_cols:
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
