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


def _as_dict(value, path: str, problems: list[str], *, required: bool = False) -> dict:
    """
    dict であることを確かめて返す。違えば説明を積んで空 dict を返す。

    YAML の書き間違い（項目が空・リストとして書いた等）でも、
    トレースバックではなく設定の問題として説明する。

    required=True の項目は、キーだけ書いて中身が空（None）の場合も報告する。
    「キーはあるが値が None」は .get(key, {}) の既定値が効かないため、
    そのまま進むと実行時に AttributeError になる。
    """
    if value is None:
        if required:
            problems.append(
                f"    {path} の中身が空です（キーだけ書かれています）"
            )
        return {}
    if not isinstance(value, dict):
        problems.append(
            f"    {path} は設定のまとまり（キー: 値）で書いてください"
            f"（現在: {type(value).__name__}）"
        )
        return {}
    return value


def _as_list(value, path: str, problems: list[str]) -> list:
    """list であることを確かめて返す。違えば説明を積んで空リストを返す。"""
    if value is None:
        return []
    if not isinstance(value, list):
        problems.append(f"    {path} は一覧（- で並べる形）で書いてください"
                        f"（現在: {type(value).__name__}）")
        return []
    return value


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
    """
    よくある綴り間違いに対して、近い候補を1つ返す。

    set をそのまま走査すると、複数が該当したときに選ばれる候補が
    実行ごとに変わってしまう（文字列ハッシュのランダム化）。
    「同じ設定に対して毎回違う修正案が出る」ことを避けるため、
    共通接頭辞の長さと名前順で並べて決定的に選ぶ。
    """
    lowered = key.lower()
    for candidate in sorted(candidates):
        if candidate.lower() == lowered:
            return candidate

    def common_prefix(a: str, b: str) -> int:
        n = 0
        for ca, cb in zip(a, b):
            if ca != cb:
                break
            n += 1
        return n

    # 長さが 1 文字以内の違いで、先頭 3 文字が共通するものを候補にする
    near = [
        c for c in candidates
        if abs(len(c) - len(key)) <= 1
        and (c.startswith(key[:3]) or key.startswith(c[:3]))
    ]
    if not near:
        return None
    # 共通接頭辞が長い順、同点なら名前順（実行ごとに変わらない）
    near.sort(key=lambda c: (-common_prefix(c, key), c))
    return near[0]


def _check_value(cfg: dict, key: str, path: str) -> list[str]:
    """列挙で決まっている値を検証する。"""
    if not isinstance(cfg, dict) or key not in cfg:
        return []
    value = cfg[key]
    allowed = ALLOWED_VALUES[key]
    try:
        if value in allowed:
            return []
    except TypeError:
        # リストなど、集合の要素にできない型が書かれている
        pass
    return [f"    {path}.{key}: 「{value}」は指定できません（{' / '.join(sorted(allowed))}）"]


def validate_source_keys(source_cfg: dict, index: int = 0) -> list[str]:
    """
    1 ソース分の設定キーと列挙値を検証する。

    問題があれば説明行のリストを返す。無ければ空リスト。
    """
    path = f"sources[{index}]"
    problems: list[str] = []

    source_cfg = _as_dict(source_cfg, path, problems, required=True)
    if not source_cfg:
        return problems

    problems += _check_keys(source_cfg, SOURCE_KEYS, path)
    problems += _check_keys(
        _as_dict(source_cfg.get("excel"), f"{path}.excel", problems, required=True),
        EXCEL_KEYS, f"{path}.excel",
    )
    problems += _check_keys(
        _as_dict(source_cfg.get("upsert"), f"{path}.upsert", problems),
        UPSERT_KEYS, f"{path}.upsert",
    )

    mapping = _as_dict(
        source_cfg.get("issue_mapping"), f"{path}.issue_mapping", problems,
        required=True,
    )
    problems += _check_keys(mapping, ISSUE_MAPPING_KEYS, f"{path}.issue_mapping")
    problems += _check_value(mapping, "description_format", f"{path}.issue_mapping")

    cf_path = f"{path}.issue_mapping.custom_fields"
    for i, cf in enumerate(_as_list(mapping.get("custom_fields"), cf_path, problems)):
        problems += _check_keys(
            _as_dict(cf, f"{cf_path}[{i}]", problems),
            CUSTOM_FIELD_KEYS, f"{cf_path}[{i}]",
        )

    for i, cond in enumerate(_as_list(
        source_cfg.get("filters"), f"{path}.filters", problems
    )):
        cond = _as_dict(cond, f"{path}.filters[{i}]", problems)
        problems += _check_keys(cond, FILTER_KEYS, f"{path}.filters[{i}]")
        problems += _check_value(cond, "match", f"{path}.filters[{i}]")

    for gi, group in enumerate(_as_list(
        source_cfg.get("filter_groups"), f"{path}.filter_groups", problems
    )):
        # グループ自体のキーも検証する。filters を綴り間違えると条件が空になり、
        # ExcelReader.filter_rows(rows, None) が全行を返すため、シート全体が
        # 無警告で登録対象になる。
        group = _as_dict(group, f"{path}.filter_groups[{gi}]", problems)
        problems += _check_keys(group, {"filters"}, f"{path}.filter_groups[{gi}]")
        for i, cond in enumerate(_as_list(
            group.get("filters"), f"{path}.filter_groups[{gi}].filters", problems
        )):
            group_path = f"{path}.filter_groups[{gi}].filters[{i}]"
            problems += _check_keys(cond, FILTER_KEYS, group_path)
            problems += _check_value(cond, "match", group_path)

    # upsert を有効にしたのに判定方法が無いと、毎行が新規作成になる
    upsert = _as_dict(source_cfg.get("upsert"), f"{path}.upsert", [])
    if upsert.get("enabled") and not (upsert.get("key_col") or upsert.get("match_summary")):
        problems.append(
            f"    {path}.upsert: enabled: true ですが key_col と match_summary の"
            "どちらも設定されていません（既存課題を探せないため常に新規作成になります）"
        )

    return problems


def validate_config_keys(config: dict) -> list[str]:
    """設定ファイル全体のキーと列挙値を検証する。"""
    problems: list[str] = []
    config = _as_dict(config, "設定ファイル", problems)
    problems += _check_keys(
        _as_dict(config.get("backlog"), "backlog", problems), BACKLOG_KEYS, "backlog"
    )
    for i, source_cfg in enumerate(_as_list(config.get("sources"), "sources", problems)):
        problems += validate_source_keys(source_cfg, i)
    return problems
