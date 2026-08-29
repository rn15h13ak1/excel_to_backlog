"""
設定キーの検証
==============
列名は validate_column_references() が検証するが、キー名は誰も見ていない。
設定全体が dict.get() の既定値に依存しているため、綴りを間違えても静かに
無視され、意図と違う動作になる。

  summary_col → sumary_col          件名が空になり全行スキップ
  match: contains → contain         黙って完全一致になり対象行が変わる
  description_format: Auto          黙って template 扱い
"""

from __future__ import annotations

# 各階層で許可するキー名
SOURCE_KEYS = {"name", "excel", "filters", "filter_groups", "issue_mapping", "upsert"}

EXCEL_KEYS = {
    "path", "sheet", "header_start_row", "header_end_row", "data_start_row",
    "col_start", "col_end",
}

ISSUE_MAPPING_KEYS = {
    "issue_type", "priority", "summary_col", "summary_template",
    "description_format", "description_template", "description_cols",
    "due_date_col", "start_date_col", "assignee_col", "default_assignee",
    "required_cols", "custom_fields", "status_col", "status_map", "rich_text",
}

CUSTOM_FIELD_KEYS = {"field_name", "col_name", "value_map", "value_separator"}
FILTER_KEYS = {"col_name", "value", "values", "match"}
UPSERT_KEYS = {"enabled", "key_col", "match_summary"}
BACKLOG_KEYS = {"space_host", "api_key", "project_key", "ssl_verify", "base_path"}

# 値が列挙で決まっている項目
ALLOWED_VALUES = {
    "match": {"exact", "contains", "startswith"},
    "description_format": {"template", "auto"},
}


def _check_keys(cfg: dict, allowed: set[str], path: str) -> list[str]:
    """未知のキーを見つけて説明行のリストを返す。"""
    if not isinstance(cfg, dict):
        return []
    problems = []
    for key in cfg:
        if key not in allowed:
            near = _closest(key, allowed)
            hint = f"（{near} の書き間違いでは？）" if near else ""
            problems.append(f"    {path}.{key} は認識できないキーです{hint}")
    return problems


def _closest(key: str, candidates: set[str]) -> str | None:
    """よくある綴り間違いを見つける（1文字違い・大文字小文字違い）。"""
    lowered = key.lower()
    for candidate in candidates:
        if candidate.lower() == lowered:
            return candidate
    # 1 文字の欠落・重複
    for candidate in candidates:
        if abs(len(candidate) - len(key)) <= 1 and (
            candidate.startswith(key[:3]) or key.startswith(candidate[:3])
        ):
            return candidate
    return None


def _check_value(cfg: dict, key: str, path: str) -> list[str]:
    """列挙で決まっている値を検証する。"""
    if not isinstance(cfg, dict) or key not in cfg:
        return []
    value = cfg[key]
    allowed = ALLOWED_VALUES[key]
    if value in allowed:
        return []
    return [f"    {path}.{key}: 「{value}」は指定できません（{' / '.join(sorted(allowed))}）"]


def validate_source_keys(source_cfg: dict, index: int = 0) -> list[str]:
    """
    1 ソース分の設定キーと列挙値を検証する。

    問題があれば説明行のリストを返す。無ければ空リスト。
    """
    path = f"sources[{index}]"
    problems = _check_keys(source_cfg, SOURCE_KEYS, path)
    problems += _check_keys(source_cfg.get("excel") or {}, EXCEL_KEYS, f"{path}.excel")
    problems += _check_keys(source_cfg.get("upsert") or {}, UPSERT_KEYS, f"{path}.upsert")

    mapping = source_cfg.get("issue_mapping") or {}
    problems += _check_keys(mapping, ISSUE_MAPPING_KEYS, f"{path}.issue_mapping")
    problems += _check_value(mapping, "description_format", f"{path}.issue_mapping")

    for i, cf in enumerate(mapping.get("custom_fields") or []):
        problems += _check_keys(
            cf, CUSTOM_FIELD_KEYS, f"{path}.issue_mapping.custom_fields[{i}]"
        )

    for i, cond in enumerate(source_cfg.get("filters") or []):
        problems += _check_keys(cond, FILTER_KEYS, f"{path}.filters[{i}]")
        problems += _check_value(cond, "match", f"{path}.filters[{i}]")

    for gi, group in enumerate(source_cfg.get("filter_groups") or []):
        for i, cond in enumerate(group.get("filters") or []):
            group_path = f"{path}.filter_groups[{gi}].filters[{i}]"
            problems += _check_keys(cond, FILTER_KEYS, group_path)
            problems += _check_value(cond, "match", group_path)

    # upsert を有効にしたのに判定方法が無いと、毎行が新規作成になる
    upsert = source_cfg.get("upsert") or {}
    if upsert.get("enabled") and not (upsert.get("key_col") or upsert.get("match_summary")):
        problems.append(
            f"    {path}.upsert: enabled: true ですが key_col と match_summary の"
            "どちらも設定されていません（既存課題を探せないため常に新規作成になります）"
        )

    return problems


def validate_config_keys(config: dict) -> list[str]:
    """設定ファイル全体のキーと列挙値を検証する。"""
    problems = _check_keys(config.get("backlog") or {}, BACKLOG_KEYS, "backlog")
    for i, source_cfg in enumerate(config.get("sources") or []):
        problems += validate_source_keys(source_cfg, i)
    return problems
