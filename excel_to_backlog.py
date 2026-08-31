#!/usr/bin/env python3
"""
Excel → Backlog 課題登録ツール
================================
複数の Excel ファイルから特定行を抽出し、Backlog 課題として登録・更新する。

【デフォルト動作はドライランです】
引数なしで実行すると変換結果の確認のみ行い、Backlog への登録は行いません。
実際に登録・更新するには --execute を付けて実行してください。

使い方:
  # ドライラン（デフォルト: 実際には作成/更新せず変換結果を確認）
  python excel_to_backlog.py

  # 実際に課題を作成/更新
  python excel_to_backlog.py --execute

  # 特定のソースのみ処理（ドライラン）
  python excel_to_backlog.py --source "タスク管理表"

  # 特定のソースのみ実際に登録
  python excel_to_backlog.py --source "タスク管理表" --execute

  # 設定ファイルを指定
  python excel_to_backlog.py --config path/to/config.yaml

  # 登録内容をMarkdownファイルに出力して確認（本文の全内容を含む）
  python excel_to_backlog.py --preview

  # API リクエスト詳細を表示（デバッグ）
  python excel_to_backlog.py --execute --debug
"""

import argparse
import contextlib
import io
import re
import sys
import time
from contextlib import ExitStack
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import yaml

from backlog_client import BacklogAPIError, BacklogClient, BacklogNoChangeError
from config_validation import validate_config_keys
from excel_reader import ExcelReader, col_letter_to_index
from mapper import BacklogMaster, IssueMapper
from row_merge import merge_continuation_rows, single_value_columns
from run_log import RunLog, completion_key, default_log_path, load_completed
from summary_index import SummaryIndex


# ------------------------------------------------------------------
# 設定ファイル読み込み
# ------------------------------------------------------------------

def load_config(config_path: str) -> dict:
    path = Path(config_path)
    if not path.exists():
        print(f"エラー: 設定ファイルが見つかりません: {config_path}", file=sys.stderr)
        sys.exit(1)
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def validate_backlog_config(backlog_cfg: dict) -> None:
    for key, placeholder in [
        ("space_host", "yourcompany.backlog.com"),
        ("api_key",    "YOUR_API_KEY_HERE"),
        ("project_key", "YOUR_PROJECT_KEY"),
    ]:
        val = backlog_cfg.get(key, "")
        if not val or val == placeholder:
            print(f"エラー: config.yaml の backlog.{key} を設定してください。", file=sys.stderr)
            sys.exit(1)


# ------------------------------------------------------------------
# 集計
# ------------------------------------------------------------------

def new_counts() -> dict:
    """
    処理結果の集計辞書を返す。

    created       : 新規作成に成功した件数
    updated       : 既存課題の更新に成功した件数
    unchanged     : 既存課題と内容が同一で変更が発生しなかった件数
                    （以前は skipped に混ぜていたが、「処理しなかった」行と
                     「処理したが変わらなかった」行は意味が異なるため分離した）
    partial       : 作成・更新はできたが一部フィールドを設定できなかった件数
                    （created / updated にも計上される）
    skipped       : 必須列が空・確認でキャンセル等で処理しなかった件数
    resumed       : --resume により前回処理済みとして飛ばした件数
    status_failed : 作成には成功したがステータス変更に失敗した件数
                    （created にも計上される。課題は Backlog に存在する）
    error         : 作成・更新に失敗した件数
    """
    return {
        "created": 0,
        "updated": 0,
        "unchanged": 0,
        "partial": 0,
        "skipped": 0,
        "resumed": 0,
        "status_failed": 0,
        "error": 0,
    }


# ------------------------------------------------------------------
# upsert ロジック
# ------------------------------------------------------------------

def find_existing_issue(
    client: BacklogClient,
    upsert_cfg: dict,
    row: dict,
    params: dict,
    master: BacklogMaster,
    summary_index: SummaryIndex | None = None,
) -> dict | None:
    """
    upsert 設定に従い既存課題を返す（見つからない場合は None）。

    issueKey だけでなく課題そのものを返す。更新前に「本当に変わるのか」を
    比較するために、現在の値が必要なため。

    upsert_cfg キー:
        key_col       : str  Excel の列名（issueKey が記入されている列）
        match_summary : bool 件名で一致する課題を探す
    """
    # ① Excel の key_col に issueKey が記入されている場合
    key_col = upsert_cfg.get("key_col")
    if key_col:
        issue_key = row.get(key_col, "").strip()
        if issue_key:
            existing = client.get_issue(issue_key)
            if existing:
                return existing
            # key_col に値はあるが Backlog に存在しない → 新規作成
            print(
                f"    ℹ issueKey「{issue_key}」は Backlog に存在しません → 新規作成",
                file=sys.stderr,
            )
            return None

    # ② 件名で照合
    # params["summary"] は map_row() で normalize_summary() 済み。
    # 索引側にも同じ正規化を適用しているため表記の揺れを吸収できる。
    if upsert_cfg.get("match_summary") and summary_index is not None:
        summary = params.get("summary", "")
        if summary:
            return summary_index.find(summary)

    return None


# ------------------------------------------------------------------
# 更新内容の差分判定
# ------------------------------------------------------------------

def _existing_value(issue: dict, key: str):
    """
    Backlog の課題オブジェクトから、送信パラメータ key に対応する現在値を返す。

    比較できない項目は None ではなく _UNKNOWN を返し、呼び出し元が
    「変更あり」に倒せるようにする。
    """
    if key == "summary":
        return IssueMapper.normalize_summary(issue.get("summary") or "")
    if key == "description":
        return issue.get("description") or ""
    if key in ("dueDate", "startDate"):
        # "2025-01-05T00:00:00Z" のような形式で返ることがあるため日付部分だけ見る
        value = issue.get(key)
        return str(value)[:10] if value else ""
    if key == "assigneeId":
        return (issue.get("assignee") or {}).get("id")
    if key == "statusId":
        return (issue.get("status") or {}).get("id")
    if key == "issueTypeId":
        return (issue.get("issueType") or {}).get("id")
    if key == "priorityId":
        return (issue.get("priority") or {}).get("id")
    if key.startswith("customField_"):
        field_id = int(key.removeprefix("customField_"))
        for cf in issue.get("customFields") or []:
            if cf.get("id") == field_id:
                value = cf.get("value")
                if isinstance(value, dict):
                    return value.get("id")
                if isinstance(value, list):
                    return sorted(
                        v.get("id") if isinstance(v, dict) else v for v in value
                    )
                return value
        return None
    return _UNKNOWN


class _Unknown:
    """比較できないことを表す番兵。None（値なし）と区別する。"""

    def __repr__(self) -> str:
        return "<比較不可>"


_UNKNOWN = _Unknown()


def has_changes(params: dict, issue: dict) -> bool:
    """
    送信予定の params が、既存課題 issue に対して変更をもたらすか判定する。

    Backlog は「変更が無く、コメントも無い」更新に対して
    "No comment content." (code=7) を返す。実際に PATCH を投げるまで
    分からなかったが、既存課題は照合の時点で取得済みのため、
    事前に比較できる。

    判定できない項目が 1 つでもあれば True（変更あり）を返す。
    誤って「変更なし」と判断して更新を飛ばすより、余分に PATCH を
    投げるほうが安全なため。
    """
    for key, new_value in params.items():
        if key == "projectId":
            continue
        current = _existing_value(issue, key)
        if current is _UNKNOWN:
            return True
        if key == "summary":
            if IssueMapper.normalize_summary(str(new_value)) != current:
                return True
            continue
        if isinstance(new_value, list):
            if sorted(new_value) != (sorted(current) if isinstance(current, list) else current):
                return True
            continue
        if str(new_value) != ("" if current is None else str(current)):
            return True
    return False


# ------------------------------------------------------------------
# メタキー注入
# ------------------------------------------------------------------

def inject_meta(row: dict, source_cfg: dict) -> dict:
    """
    行データに Excel ソース由来のメタ情報を注入して返す。

    注入キー（アンダースコア始まりで Excel 列名とは区別できる）:
        _source_name  : sources[i].name の値
        _excel_path   : sources[i].excel.path の値
        _excel_sheet  : sources[i].excel.sheet の値

    これらは summary_template / description_template などの
    {{キー名}} プレースホルダーで参照できる。
    """
    excel_cfg = source_cfg.get("excel", {})
    meta = {
        "_source_name": source_cfg.get("name", ""),
        "_excel_path":  excel_cfg.get("path", ""),
        "_excel_sheet": excel_cfg.get("sheet", ""),
    }
    # 元の row は変更しない（コピーして返す）
    return {**meta, **row}


# ------------------------------------------------------------------
# フィルタリング（filters / filter_groups 共通処理）
# ------------------------------------------------------------------

def apply_filters(
    rows: list,
    source_cfg: dict,
    headers: list,
) -> list:
    """
    source_cfg の filters または filter_groups に従い行を絞り込む。

    filters       : 複数条件を AND 評価（従来通り）
    filter_groups : 各グループを AND 評価し、グループ間を OR 評価
                    同じ行が複数グループにマッチしても重複しない
    両方省略時は全行を返す。filters と filter_groups が両方指定された場合は
    filter_groups を優先する。
    """
    filter_groups_cfg = source_cfg.get("filter_groups") or []
    filters_cfg = source_cfg.get("filters") or []

    if filter_groups_cfg:
        # 列名チェック（全グループ対象）
        for gi, group in enumerate(filter_groups_cfg):
            for cond in group.get("filters") or []:
                col = cond.get("col_name", "")
                if col and col not in headers:
                    print(
                        f"  ⚠ filter_groups[{gi}] の列「{col}」がヘッダーに存在しません。"
                        f"（ヘッダー: {headers}）",
                        file=sys.stderr,
                    )

        # 各グループを AND 評価 → グループ間を OR（重複除去しつつ順序保持）
        seen_ids: set = set()
        result = []
        for group in filter_groups_cfg:
            group_filters = group.get("filters") or []
            for row in ExcelReader.filter_rows(rows, group_filters):
                rid = id(row)
                if rid not in seen_ids:
                    seen_ids.add(rid)
                    result.append(row)
        return result

    else:
        # 従来の filters（AND 評価）
        for cond in filters_cfg:
            col = cond.get("col_name", "")
            if col and col not in headers:
                print(
                    f"  ⚠ フィルター列「{col}」がヘッダーに存在しません。"
                    f"（ヘッダー: {headers}）",
                    file=sys.stderr,
                )
        return ExcelReader.filter_rows(rows, filters_cfg)


# ------------------------------------------------------------------
# 列名参照の事前検証
# ------------------------------------------------------------------

# inject_meta() が行データに注入するキー。Excel の列ではないがテンプレートから
# 参照できるため、列名検証では既知の名前として扱う。
# _excel_row は ExcelReader が行データに直接持たせる（inject_meta 経由ではない）
META_KEYS = {"_source_name", "_excel_path", "_excel_sheet", ExcelReader.ROW_NUMBER_KEY}


def collect_referenced_columns(source_cfg: dict) -> list[tuple[str, str]]:
    """
    source_cfg が参照している Excel 列名を (設定パス, 列名) のリストで返す。

    テンプレート項目（summary_template / description_template）は
    IssueMapper.extract_template_columns() でプレースホルダーを展開する。
    due_date_col / start_date_col は map_row() と同じ判定（"{{" を含むか）で
    テンプレートと列名を区別する。
    """
    refs: list[tuple[str, str]] = []
    mapping_cfg = source_cfg.get("issue_mapping") or {}
    upsert_cfg = source_cfg.get("upsert") or {}

    def add(path: str, value) -> None:
        # 列名は strip せずそのまま登録する。
        # filter_rows() や _resolve_*() は設定値を strip せずに row の
        # キーと突き合わせるため、前後の空白の有無まで含めて一致させる必要がある。
        # （テンプレートの {{列名}} だけは _render_template() 側が strip するため、
        #   add_template() 経由で strip 済みの名前が渡る）
        if isinstance(value, str) and value.strip():
            refs.append((path, value))

    def add_template(path: str, template: str) -> None:
        for col in sorted(IssueMapper.extract_template_columns(template)):
            refs.append((path, col))

    # ---- フィルター ----
    # apply_filters() は filter_groups があれば filters を無視するため、
    # 検証も同じ優先順位に合わせる。移行途中で古い filters が残っていても
    # 実行時に読まれない設定でソースを止めない。
    filter_groups_cfg = source_cfg.get("filter_groups") or []
    if filter_groups_cfg:
        for gi, group in enumerate(filter_groups_cfg):
            for i, cond in enumerate(group.get("filters") or []):
                add(f"filter_groups[{gi}].filters[{i}].col_name", cond.get("col_name"))
    else:
        for i, cond in enumerate(source_cfg.get("filters") or []):
            add(f"filters[{i}].col_name", cond.get("col_name"))

    # ---- 件名 ----
    summary_template = mapping_cfg.get("summary_template", "")
    if summary_template:
        add_template("issue_mapping.summary_template", summary_template)
    else:
        add("issue_mapping.summary_col", mapping_cfg.get("summary_col"))

    # ---- 本文 ----
    # description_template は template モードでのみ使われる
    is_auto = mapping_cfg.get("description_format", "template") == "auto"
    description_template = mapping_cfg.get("description_template", "")
    if not is_auto:
        add_template("issue_mapping.description_template", description_template)

    # description_cols を読むのは _render_auto() だけで、それが動くのは
    # auto モードか、テンプレートに {{auto}} がある場合に限られる。
    # どちらでもないときに検証すると、使われない設定でソースを止めてしまう。
    # _render_template() はプレースホルダー名を strip して判定するため、
    # {{ auto }} のような書き方も同じように扱う。
    placeholders = {
        name.strip()
        for name in IssueMapper.TEMPLATE_PLACEHOLDER_RE.findall(description_template)
    }
    uses_auto_body = is_auto or "auto" in placeholders
    if uses_auto_body:
        for i, col in enumerate(mapping_cfg.get("description_cols") or []):
            add(f"issue_mapping.description_cols[{i}]", col)

    # ---- 日付（列名またはテンプレート）----
    for key in ("due_date_col", "start_date_col"):
        value = mapping_cfg.get(key)
        if not value:
            continue
        if "{{" in str(value):
            add_template(f"issue_mapping.{key}", str(value))
        else:
            add(f"issue_mapping.{key}", value)

    # ---- その他の単一列 ----
    add("issue_mapping.assignee_col", mapping_cfg.get("assignee_col"))
    add("issue_mapping.status_col", mapping_cfg.get("status_col"))

    # key_col を読む find_existing_issue() は upsert 有効時しか呼ばれない。
    # 書き戻し列をこれから用意する段階で設定だけ先に書いておけるよう、
    # 無効時は検証しない。
    if upsert_cfg.get("enabled", False):
        add("upsert.key_col", upsert_cfg.get("key_col"))

    # ---- 必須列・カスタム属性 ----
    for i, col in enumerate(mapping_cfg.get("required_cols") or []):
        add(f"issue_mapping.required_cols[{i}]", col)
    for i, cf in enumerate(mapping_cfg.get("custom_fields") or []):
        add(f"issue_mapping.custom_fields[{i}].col_name", cf.get("col_name"))

    return refs


def validate_column_references(source_cfg: dict, headers: list[str]) -> list[str]:
    """
    設定が参照する列名がすべてヘッダーに存在するか検証する。

    存在しない参照が1件でもあれば、その内容を説明する行のリストを返す。
    すべて解決できる場合は空リストを返す。

    列名の不一致は「その条件が無視される」「プレースホルダーが未展開のまま
    件名になる」といった無言の誤動作につながるため、警告ではなく実行前の
    エラーとして扱う（呼び出し元がソースの処理を中止する）。
    """
    known = set(headers) | META_KEYS
    unknown = [(path, col) for path, col in collect_referenced_columns(source_cfg)
               if col not in known]
    if not unknown:
        return []

    lines = [
        f"設定が参照する列名が Excel のヘッダーに存在しません（{len(unknown)} 件）:",
    ]
    for path, col in unknown:
        # 前後の空白は「」で囲んでも見えないため、strip すると一致する場合は明示する
        hint = ""
        if col.strip() != col and col.strip() in known:
            hint = f"  ← 前後の空白を除けば一致します（{col!r}）"
        lines.append(f"    {path}: 「{col}」{hint}")
    lines.append(f"  ヘッダー: {headers}")
    lines.append(
        "  → 列名の綴り・前後の空白・複数行ヘッダーの結合結果（\" / \" 区切り）を確認してください。"
    )
    return lines


# ------------------------------------------------------------------
# ソースの読み込み
# ------------------------------------------------------------------

class SourceLoadError(Exception):
    """ソースの読み込み・検証に失敗した。message は利用者向けの説明。"""

    def __init__(self, message: str, detail: str = ""):
        super().__init__(message)
        self.message = message
        self.detail = detail


@dataclass
class LoadedSource:
    """
    1 ソース分の読み込み結果。

    実処理（process_source）とプレビュー生成（generate_preview_for_source）は
    同じ手順で Excel を読み、列名を検証し、フィルターを適用し、マッパーを
    用意する。以前は両者が同じ流れを別々に実装しており、実際に乖離が
    始まっていた（プレビュー側だけ書式付き行の取得方法が異なる等）。
    """

    headers: list[str]
    rows: list[dict]                    # フィルター適用後
    mapper: IssueMapper
    formatted_by_row: dict[str, dict]   # {_excel_row: 書式付き行}

    def formatted_for(self, row: dict) -> dict | None:
        """
        平文行に対応する書式付き行を返す。rich_text 無効時は None。

        対応付けには Excel の行番号を使う。以前は id(dict) を使っており、
        フィルターがコピーを返すようになった瞬間に壊れる作りだった。
        壊れても例外は出ず、本文だけが静かに元テキストへ戻る。
        """
        if not self.formatted_by_row:
            return None
        return self.formatted_by_row.get(row.get(ExcelReader.ROW_NUMBER_KEY))


def load_source(source_cfg: dict, master: BacklogMaster, *, limit: int | None = None) -> LoadedSource:
    """
    ソースを読み込み、列名を検証し、フィルターを適用して返す。

    Raises
    ------
    SourceLoadError : 読み込み失敗・列名不一致・設定不備のいずれか
    """
    import traceback

    excel_cfg = source_cfg.get("excel") or {}
    mapping_cfg = source_cfg.get("issue_mapping") or {}

    # ---- Excel 読み込み ----
    try:
        reader = ExcelReader(excel_cfg)
        if mapping_cfg.get("rich_text"):
            headers, rows, formatted_rows = reader.read_with_format()
        else:
            headers, rows = reader.read()
            formatted_rows = None
    except Exception as e:
        raise SourceLoadError(
            f"Excel の読み込みに失敗しました: {e}\n"
            "  ヒント: Excel ファイルが破損または非標準形式の可能性があります。",
            detail=traceback.format_exc(),
        ) from e

    print(f"  読込行数: {len(rows)} 行（フィルター前）")

    # ---- 継続行の結合 ----
    # 1 件の内容が複数行に分かれている表を 1 件にまとめる。
    # 絞り込みより前に行う。続きの行は絞り込み条件の列も空になっているため、
    # 先に絞ると結合前に失われてしまう。
    if mapping_cfg.get("merge_continuation_rows"):
        before = len(rows)
        rows = merge_continuation_rows(
            rows, headers,
            mapping_cfg.get("required_cols") or [],
            single_value_columns(source_cfg),
        )
        if len(rows) != before:
            print(f"  継続行を結合: {before} 行 → {len(rows)} 件")

    # ---- 列名参照の検証 ----
    # 列名が1つでも一致しないと、フィルター条件が無視されて全行が対象になる、
    # プレースホルダーが未展開のまま件名になる等の無言の誤動作が起きるため、
    # 行を処理する前に中止する。
    errors = validate_column_references(source_cfg, headers)
    if errors:
        raise SourceLoadError("\n".join(errors))

    # ---- フィルタリング ----
    filtered = apply_filters(rows, source_cfg, headers)
    print(f"  対象行数: {len(filtered)} 行（フィルター後）")

    if limit is not None and len(filtered) > limit:
        print(f"  → --limit {limit} のため先頭 {limit} 行のみ処理します"
              f"（残り {len(filtered) - limit} 行は対象外）")
        filtered = filtered[:limit]

    # ---- マッパー ----
    mapper = IssueMapper(mapping_cfg, master, headers=headers)
    # 種別・優先度は全行に共通の設定のため、行ごとではなくここで一度だけ解決する。
    # 行ごとに判定すると、設定のタイプミスが「全行スキップ」というデータ不備の
    # ような報告になり、原因が設定だと分からない。
    try:
        mapper.resolve_fixed_fields()
    except ValueError as e:
        raise SourceLoadError(f"{e}\n  → issue_mapping の設定を確認してください。") from e

    formatted_by_row = (
        {r.get(ExcelReader.ROW_NUMBER_KEY): r for r in formatted_rows}
        if formatted_rows is not None else {}
    )
    return LoadedSource(headers, filtered, mapper, formatted_by_row)


# ------------------------------------------------------------------
# プレビューファイル生成
# ------------------------------------------------------------------

def build_master_labels(master: BacklogMaster) -> dict:
    """
    ID → 表示名 の逆引き辞書を生成する（プレビュー表示用）。

    種別・優先度・ユーザーは ID 空間が独立しているため、
    フラットにマージせずカテゴリ別のネスト構造で返す。

    Returns
    -------
    dict
        {
          "issue_type": {id: 種別名, ...},
          "priority":   {id: 優先度名, ...},
          "user":       {id: ユーザー名, ...},
          "status":     {id: ステータス名, ...},
        }
    """
    # user_map には同一ユーザーの表示名・ログインIDが両方登録されているため、
    # ID が初出のエントリ（登録順で先に来る表示名）のみを逆引きに採用する
    user_labels: dict[int, str] = {}
    for name, id_ in master.user_map.items():
        if id_ not in user_labels:
            user_labels[id_] = name

    return {
        "issue_type": {id_: name for name, id_ in master.issue_type_map.items()},
        "priority":   {id_: name for name, id_ in master.priority_map.items()},
        "user":       user_labels,
        "status":     {id_: name for name, id_ in master.status_map.items()},
    }


def _safe_filename(name: str) -> str:
    """
    sources[].name をファイル名として安全な文字列に変換する。
    ファイル名に使えない文字（/ \\ : * ? " < > | など）はアンダースコアに置換し、
    前後の空白・ドットを除去する。
    """
    safe = re.sub(r'[\\/:*?"<>|\s]+', "_", name)
    safe = safe.strip("._")
    return safe or "source"


def generate_preview_for_source(
    source_cfg: dict,
    master: BacklogMaster,
    master_labels: dict,
    output_path: Path,
    now: str,
) -> int:
    """
    1ソース分の登録予定内容を Markdown ファイルに書き出す。

    Returns
    -------
    int : プレビュー生成した課題件数
    """
    name = source_cfg.get("name", "（名前なし）")
    excel_cfg = source_cfg.get("excel", {})
    mapping_cfg = source_cfg.get("issue_mapping", {})

    lines = [
        f"# Backlog 課題登録 プレビュー — {name}",
        "",
        f"> 生成日時: {now}  ",
        f"> ※ このファイルは登録前の確認用です。実際の登録は `--execute` で行います。",
        "",
        f"- ファイル: `{excel_cfg.get('path', '（未設定）')}`",
        f"- シート: `{excel_cfg.get('sheet', '（最初のシート）')}`",
        "",
        "---",
        "",
    ]

    issue_count = 0
    # ---- 読み込み・検証・フィルタ（実処理と共通）----
    try:
        loaded = load_source(source_cfg, master)
    except SourceLoadError as e:
        lines.append(f"> ⚠ {e.message}")
        lines.append("")
        if e.detail:
            lines.append(f"> 詳細: `{e.detail}`")
            lines.append("")
        output_path.write_text("\n".join(lines), encoding="utf-8")
        return 0

    lines.append(f"対象行数: **{len(loaded.rows)} 件**（フィルター後）")
    lines.append("")

    if not loaded.rows:
        lines.append("_対象行がありません。_")
        lines.append("")
        output_path.write_text("\n".join(lines), encoding="utf-8")
        return 0

    issue_count = 0
    for row in loaded.rows:
        i = row.get(ExcelReader.ROW_NUMBER_KEY, "?")
        enriched = inject_meta(row, source_cfg)
        fmt_row = loaded.formatted_for(row)
        fmt_enriched = inject_meta(fmt_row, source_cfg) if fmt_row is not None else None
        lines.append(loaded.mapper.format_preview(
            enriched, i, master_labels=master_labels, formatted_row=fmt_enriched
        ))
        lines.append("")
        lines.append("---")
        lines.append("")
        issue_count += 1

    output_path.write_text("\n".join(lines), encoding="utf-8")
    return issue_count


def generate_preview_file(
    sources_cfg: list,
    master: BacklogMaster,
    output_dir: Path,
    timestamp: str,
) -> list[tuple[Path, int]]:
    """
    sources[].name の単位で Markdown プレビューファイルを生成する。
    Backlog API のデータ書き込みは行わない。

    Returns
    -------
    list[tuple[Path, int]]
        各ソースの (出力ファイルパス, 課題件数) のリスト
    """
    master_labels = build_master_labels(master)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    results = []

    for source_cfg in sources_cfg:
        name = source_cfg.get("name", "source")
        safe = _safe_filename(name)
        output_path = output_dir / f"preview_{timestamp}_{safe}.md"
        count = generate_preview_for_source(
            source_cfg, master, master_labels, output_path, now
        )
        results.append((output_path, count))

    return results


# ------------------------------------------------------------------
# 課題新規作成（2段階: 未対応で作成 → statusId を更新）
# ------------------------------------------------------------------

class StatusUpdateFailed(Exception):
    """
    課題の作成には成功したが、その後のステータス変更に失敗した。

    課題は Backlog に存在するため、呼び出し元は「作成成功」として issueKey を
    表示しなければならない。これをエラーとして扱うと、実際には存在する課題が
    「作成 0 件 / エラー 1 件」と報告され、issueKey も表示されないまま
    Backlog 上に取り残される。
    """

    def __init__(self, issue: dict, cause: Exception):
        super().__init__(str(cause))
        self.issue = issue
        self.cause = cause


def create_issue_with_status(client: BacklogClient, params: dict) -> dict:
    """
    課題を新規作成する。

    Backlog API は新規作成時に「完了」などの終了ステータスを直接設定できないため、
    statusId をいったん除いた状態（デフォルト「未対応」）で作成し、
    statusId が指定されていた場合は作成後に update_issue で変更する。

    Raises
    ------
    StatusUpdateFailed
        作成には成功したがステータス変更に失敗した場合。
        作成済みの課題を .issue に保持する。
    """
    status_id = params.pop("statusId", None)
    try:
        issue = client.create_issue(params)
    finally:
        # pop したのでロールバックしておく（呼び出し元の params を汚さない）
        if status_id is not None:
            params["statusId"] = status_id

    if status_id is not None:
        try:
            client.update_issue(issue["issueKey"], {"statusId": status_id})
        except BacklogNoChangeError:
            # 作成時点で既にそのステータスだった場合はそのまま続行
            pass
        except BacklogAPIError as e:
            # 課題は作成済み。呼び出し元が issueKey を表示できるよう課題ごと渡す。
            raise StatusUpdateFailed(issue, e) from e

    return issue


# ------------------------------------------------------------------
# 新規作成の確認
# ------------------------------------------------------------------

class ConfirmationDeclined(Exception):
    """ユーザーが確認で実行を取りやめた。"""


class RowConfirmer:
    """
    1 行ずつ書き込みの可否を確認する。

    まとめて一気に書き込むと、件数が多いときに何が起きたのか把握しきれない。
    作成・更新のどちらも、実行前に内容を見て判断できるようにする。

    選択肢:
        y  この行を実行する
        n  この行を飛ばす（スキップとして計上）
        a  以降はすべて確認せず実行する
        q  ここで実行を中止する

    assume_yes=True（--yes）のときは一切確認しない。cron やパイプ経由で
    使う場合に必要。
    """

    def __init__(self, assume_yes: bool = False, master_labels: dict | None = None):
        self.assume_all = assume_yes
        self.master_labels = master_labels or {}

    def confirm(self, plan, mapper: IssueMapper) -> bool:
        """
        実行してよければ True、飛ばすなら False を返す。

        Raises
        ------
        ConfirmationDeclined : ユーザーが中止（q）を選んだ場合
        """
        if self.assume_all:
            return True

        action = (
            f"更新 → {plan.existing_key}" if plan.action == "update" else "新規作成"
        )
        print(f"\n  [{plan.row_number}] {action}")
        print(mapper.format_plan(plan, master_labels=self.master_labels))

        while True:
            try:
                answer = input(
                    "    実行しますか？ [y=実行 / n=スキップ / a=以降すべて / q=中止]: "
                ).strip().lower()
            except (EOFError, OSError):
                # 非対話環境。本来は main() が事前に弾いているが、
                # 途中で標準入力が閉じられた場合に無限ループしないよう中止する。
                raise ConfirmationDeclined(
                    "確認の入力を受け取れませんでした。--yes を付けて実行してください。"
                ) from None

            if answer in ("y", "yes"):
                return True
            if answer in ("n", "no", ""):
                return False
            if answer in ("a", "all"):
                self.assume_all = True
                print("    → 以降は確認せずに実行します。")
                return True
            if answer in ("q", "quit"):
                raise ConfirmationDeclined("ユーザーが実行を中止しました。")
            print("    y / n / a / q のいずれかを入力してください。")


def confirm_run(
    sources_cfg: list,
    master: BacklogMaster,
    assume_yes: bool,
    planned: dict | None = None,
) -> None:
    """
    Backlog への書き込みを始める前に、全体の見込みを表示する。

    件数の内訳を示したうえで、実際の可否は 1 行ずつ RowConfirmer が確認する
    （ここで重ねて聞くと二重の確認になるため、この関数は表示のみ）。

    以前は行ごとに新規作成の確認を出していたが、以下の問題があった:
      - 200 行あれば 200 回の入力が必要で、実質「全部 y」しか選べない
      - 表示が件名と期限日だけで、本文・担当者・ステータス・カスタム属性は
        隠れており、行ごとに判断する材料が無かった
      - 既存課題の「更新」は無確認だった。上書きという破壊的な操作の方が
        確認なしで、追加でしかない作成の方に確認を求めていた
      - 非対話環境では input() が毎回 EOFError になり全行スキップ。
        1 件も作らずに「作成: 0 件」と表示して正常終了していた

    Raises
    ------
    ConfirmationDeclined
        ユーザーが「いいえ」を選んだ場合、または非対話環境で --yes が
        指定されていない場合。
    """
    names = [s.get("name", "（名前なし）") for s in sources_cfg]

    print()
    print("─" * 55)
    print("Backlog への書き込みを開始します")
    print("─" * 55)
    print(f"  対象ソース: {', '.join(names)}")

    # 何件を作成し何件を更新するのかを、書き込む前に示す。
    # 「更新のつもりが全部作成になっていた」といった取り違えを
    # 実行前に気づけるようにする。
    if planned is not None:
        print(f"  作成予定: {planned['created']} 件 / 更新予定: {planned['updated']} 件")
        for key, label in (("unchanged", "変更なし"), ("resumed", "再開スキップ"),
                           ("skipped", "スキップ"), ("partial", "一部フィールド未設定")):
            if planned.get(key):
                print(f"  {label}: {planned[key]} 件")
    print("  内容は --preview で詳しく確認できます。")
    print()

    if assume_yes:
        print("  --yes が指定されているため、確認せずに実行します。")
        return

    if not sys.stdin.isatty():
        # 非対話環境で黙って進めると、確認を求められないまま
        # 全件スキップして「成功したように見える無処理」になる。
        # cron やパイプ経由で使う場合は --yes を明示させる。
        raise ConfirmationDeclined(
            "非対話環境では確認を求められません。実行するには --yes を付けてください。"
        )

    print("  この後、1 件ずつ内容を確認します。")


# ------------------------------------------------------------------
# マスターデータの一覧表示
# ------------------------------------------------------------------

# Backlog のカスタム属性の型（typeId → 表示名）
CUSTOM_FIELD_TYPES = {
    1: "文字列", 2: "文章", 3: "数値", 4: "日付",
    5: "単一リスト", 6: "複数リスト", 7: "チェックボックス", 8: "ラジオ",
}


def print_master_data(master: BacklogMaster) -> None:
    """
    設定に書ける名前の一覧を表示する。

    config.yaml を書くには Backlog 側の種別名・優先度名・ステータス名・
    担当者名・カスタム属性名が必要だが、これらは設定を書き終えて実行して
    初めて分かるという順序になっていた。特にカスタム属性名はどこにも
    表示されず、「見つかりません」と言われても正しい名前を知る手段がなかった。
    """
    def section(title: str, values) -> None:
        print(f"\n{title}")
        if not values:
            print("  （取得できませんでした）")
            return
        for v in values:
            print(f"  {v}")

    print("=" * 55)
    print("設定に使える名前の一覧")
    print("=" * 55)

    section("種別（issue_mapping.issue_type）", list(master.issue_type_map))
    section("優先度（issue_mapping.priority）", list(master.priority_map))
    section("ステータス（issue_mapping.status_map の変換先）", list(master.status_map))

    # user_map は表示名とログインIDの両方を持つため、ID ごとにまとめ直す
    by_id: dict[int, list[str]] = {}
    for name, uid in master.user_map.items():
        by_id.setdefault(uid, []).append(name)
    section(
        "担当者（issue_mapping.assignee_col の値 / default_assignee）",
        [" / ".join(names) for names in by_id.values()],
    )

    print("\nカスタム属性（issue_mapping.custom_fields.field_name）")
    if not master.custom_field_map:
        print("  （このプロジェクトには定義されていません）")
    for name, info in master.custom_field_map.items():
        type_name = CUSTOM_FIELD_TYPES.get(info.get("typeId"), f"typeId={info.get('typeId')}")
        print(f"  {name}  [{type_name}]")
        items = list(info.get("items") or {})
        if items:
            print(f"      選択肢: {' / '.join(items)}")

    print()
    print("─" * 55)
    print("  これらの名前を config.yaml にそのまま記述してください。")
    print("  Excel の列名は --show-columns で確認できます。")


# ------------------------------------------------------------------
# Excel の列名一覧
# ------------------------------------------------------------------

def print_source_columns(sources_cfg: list) -> int:
    """
    各ソースの Excel から読み取れる列名を一覧表示する。

    設定に書く列名は、複数行ヘッダーの " / " 結合結果や、重複時に付く
    " (2)" の連番まで含めて正確に一致させる必要がある。これらは実際に
    読み込ませないと分からないため、設定を書く前に確認できるようにする。

    Returns
    -------
    int : 読み込みに失敗したソースの数
    """
    from openpyxl.utils import get_column_letter

    failures = 0
    for source_cfg in sources_cfg:
        name = source_cfg.get("name", "（名前なし）")
        excel_cfg = source_cfg.get("excel") or {}
        print(f"\n{'=' * 55}")
        print(f"[{name}] {excel_cfg.get('path', '（path 未設定）')}")
        sheet = excel_cfg.get("sheet")
        print(f"  シート: {sheet or '（最初のシート）'}")
        print("=" * 55)

        try:
            reader = ExcelReader(excel_cfg)
            headers, rows = reader.read()
        except Exception as e:
            print(f"  ⚠ 読み込みに失敗しました: {e}", file=sys.stderr)
            failures += 1
            continue

        start = col_letter_to_index(reader.col_start_str)
        seen: set[str] = set()
        for i, header in enumerate(headers):
            letter = get_column_letter(start + i + 1)
            if header in seen:
                # 同名の列。本文には出力されるが、列名で参照すると左端になる
                print(f"  {letter:>3}: {header}  ← 同名（本文には出力／列名指定は左端）")
                continue
            seen.add(header)
            # その列に値が入っている最初の行を例として添える
            sample = next(
                (r[header] for r in rows if r.get(header)), ""
            )
            sample = f"  例: {sample[:28]}" if sample else ""
            print(f"  {letter:>3}: {header}{sample}")

        print(f"\n  読込 {len(rows)} 行（フィルター前）")

    print()
    print("─" * 55)
    print("  この列名を config.yaml にそのまま記述してください。")
    print("  複数行ヘッダーは \" / \" で結合されます。")
    return failures


# ------------------------------------------------------------------
# 行ごとの処理計画
# ------------------------------------------------------------------

@dataclass
class RowPlan:
    """
    1 行をどう処理するかの決定。

    ドライラン・実行の両方がこれを組み立てるため、ドライランの表示が
    実行結果と食い違わない。以前はドライランが行ループの手前で return して
    おり、upsert の照合を一切行わなかった。このツールで最も影響が大きい
    「作成するのか更新するのか」を事前に確認できなかった。

    action:
        "create"  新規作成する
        "update"  existing_key の課題を更新する
        "skip"      処理しない（必須列が空・件名が空など）
        "resume"    前回の実行で完了済みのため飛ばす
        "unchanged" 既存課題と内容が同じで、更新しても変わらない
    """

    row_number: str                     # Excel シート上の行番号
    action: str
    params: dict = field(default_factory=dict)
    existing: dict | None = None        # 既存課題（更新・変更なしの場合）
    warnings: list[str] = field(default_factory=list)
    reason: str = ""                    # skip の理由

    @property
    def summary(self) -> str:
        return self.params.get("summary", "")

    @property
    def existing_key(self) -> str | None:
        return (self.existing or {}).get("issueKey")


def plan_row(
    row: dict,
    source_cfg: dict,
    mapper: IssueMapper,
    *,
    formatted_row: dict | None = None,
    client: BacklogClient | None = None,
    master: BacklogMaster | None = None,
    summary_index: SummaryIndex | None = None,
    completed: set[tuple[str, str, str]] | None = None,
) -> RowPlan:
    """
    1 行分の処理計画を組み立てる。Backlog への書き込みは行わない。

    upsert が有効な場合は既存課題の照合（読み取りのみ）を行うため、
    ドライランでも「作成か更新か」を確定できる。
    """
    name = source_cfg.get("name", "（名前なし）")
    upsert_cfg = source_cfg.get("upsert") or {}
    row_number = row.get(ExcelReader.ROW_NUMBER_KEY, "?")

    enriched = inject_meta(row, source_cfg)
    fmt_enriched = (
        inject_meta(formatted_row, source_cfg) if formatted_row is not None else None
    )

    try:
        params = mapper.map_row(enriched, formatted_row=fmt_enriched)
    except ValueError as e:
        return RowPlan(row_number=row_number, action="skip", reason=str(e))

    # map_row は次の呼び出しでリセットするため、この行の分を控える
    warnings = list(mapper.warnings)
    summary = params.get("summary", "")

    # --resume: 前回の実行で作成・更新まで完了した行は飛ばす
    if completed is not None and completion_key(name, row_number, summary) in completed:
        return RowPlan(row_number=row_number, action="resume", params=params,
                       warnings=warnings)

    existing = None
    if upsert_cfg.get("enabled", False) and client is not None and master is not None:
        existing = find_existing_issue(
            client, upsert_cfg, enriched, params, master, summary_index=summary_index
        )

    if existing is None:
        action = "create"
    elif has_changes(params, existing):
        action = "update"
    else:
        # 既存課題と内容が同じ。PATCH を投げても Backlog が
        # "No comment content." を返すだけなので、送らずに済ませる。
        action = "unchanged"

    return RowPlan(
        row_number=row_number,
        action=action,
        params=params,
        existing=existing,
        warnings=warnings,
    )


# ------------------------------------------------------------------
# 1ソースの処理
# ------------------------------------------------------------------

def process_source(
    source_cfg: dict,
    client: BacklogClient,
    master: BacklogMaster,
    dry_run: bool,
    run_log: RunLog | None = None,
    completed: set[tuple[str, str, str]] | None = None,
    summary_index: SummaryIndex | None = None,
    counts: dict | None = None,
    limit: int | None = None,
    confirmer: RowConfirmer | None = None,
) -> dict:
    """
    1つのソース（Excel ファイル）を処理して作成・更新件数を返す。

    Parameters
    ----------
    limit : int | None
        処理する行数の上限。初回に少数だけ試すために使う。
        フィルター適用後の先頭から数える。
    confirmer : RowConfirmer | None
        1 行ずつ書き込みの可否を確認する。省略時は確認しない。
    counts : dict | None
        集計を積む辞書。呼び出し元が渡した辞書をその場で更新するため、
        途中で例外が送出されてもそこまでの集計が呼び出し元に残る。
        以前は戻り値でのみ返しており、認証エラー等で中断すると
        「3 件作成したのにサマリーは 0 件」と報告していた。
        省略時は新しい辞書を作る。

    Returns
    -------
    dict: new_counts() と同じキーを持つ集計結果（counts と同一オブジェクト）
    """
    name = source_cfg.get("name", "（名前なし）")
    excel_cfg = source_cfg.get("excel", {})
    mapping_cfg = source_cfg.get("issue_mapping", {})
    upsert_cfg = source_cfg.get("upsert") or {}
    upsert_enabled = upsert_cfg.get("enabled", False)

    if counts is None:
        counts = new_counts()

    print(f"\n{'='*55}")
    print(f"ソース: {name}")
    print(f"{'='*55}")
    print(f"  ファイル: {excel_cfg.get('path', '（未設定）')}")
    print(f"  シート : {excel_cfg.get('sheet', '（最初のシート）')}")
    print(f"  upsert : {'有効' if upsert_enabled else '無効（常に新規作成）'}")

    # ---- 読み込み・検証・フィルタ（プレビュー生成と共通）----
    try:
        loaded = load_source(source_cfg, master, limit=limit)
    except SourceLoadError as e:
        print(f"\n  エラー: {e.message}", file=sys.stderr)
        if e.detail:
            print(e.detail, file=sys.stderr)
        counts["error"] += 1
        return counts

    filtered_rows = loaded.rows
    mapper = loaded.mapper

    if not filtered_rows:
        print("  → 対象行がないためスキップします。")
        return counts

    def get_formatted_row(plain_row: dict) -> dict | None:
        """plain_row に対応する書式付き行を返す。rich_text 無効時は None。"""
        return loaded.formatted_for(plain_row)

    # ---- ドライラン ----
    # 実行と同じ plan_row() を通すため、作成か更新かまで確認できる。
    # upsert の照合は読み取りのみで Backlog を変更しない。
    if dry_run:
        print(f"\n  [DRY RUN] 以下の課題を作成/更新します:\n")
        for row in filtered_rows:
            plan = plan_row(
                row, source_cfg, mapper,
                formatted_row=get_formatted_row(row),
                client=client, master=master,
                summary_index=summary_index, completed=completed,
            )
            i = plan.row_number

            if plan.action == "skip":
                print(f"  [{i}] ⚠ スキップ: {plan.reason}", file=sys.stderr)
                counts["skipped"] += 1
                continue
            if plan.action == "resume":
                print(f"  [{i}] — 再開スキップ（前回処理済み）: {plan.summary}")
                counts["resumed"] += 1
                continue

            if plan.action == "unchanged":
                print(f"  [{i}] 変更なし（{plan.existing_key}）")
                counts["unchanged"] += 1
                continue

            if plan.action == "update":
                print(f"  [{i}] 更新 → {plan.existing_key}")
                counts["updated"] += 1
            else:
                print(f"  [{i}] 新規作成")
                counts["created"] += 1
                # 実行時は作成した課題を索引に加えるため、同じ件名の後続行は
                # 「更新」になる。ドライランでも同じ判断になるよう、作成予定の
                # 件名を索引に入れておく（Backlog へは書き込まない）。
                if summary_index is not None:
                    # 実行時と同じく索引へ加える。実在しないことが分かる
                    # issueKey にして、万一表示されても取り違えないようにする。
                    summary_index.add(plan.summary, {
                        "issueKey": "（この実行で作成予定）",
                        "summary": plan.summary,
                    })
            if plan.warnings:
                counts["partial"] += 1

            print(mapper.format_plan(plan))
        return counts

    # ---- 実処理 ----
    def log(*, row: int, action: str, issue_key: str = "", summary: str = "", detail: str = "") -> None:
        """実行ログに1件記録する（--log-file 未指定時は何もしない）。"""
        if run_log is not None:
            run_log.record(
                source=name, row=row, action=action,
                issue_key=issue_key, summary=summary, detail=detail,
            )

    for row in filtered_rows:
        # 表示・ログ・再開判定には Excel シート上の行番号を使う。
        # フィルタ後の連番ではシートの何行目か辿れず、失敗した行を特定できない。
        i = row.get(ExcelReader.ROW_NUMBER_KEY, "?")

        # 例外処理で参照するため、計画の組み立て前に初期化しておく。
        # plan_row() は照合のために API を呼ぶため、ここで失敗しうる。
        params: dict = {}
        row_warnings: list[str] = []

        # 実際に通信した行だけレート制限用の待機を入れる
        # （スキップ・再開スキップした行は通信していないため待つ意味がない）
        api_called = False
        try:
            # ドライランと同じ関数で計画を組み立てる（照合は読み取りのみ）
            api_called = upsert_enabled
            plan = plan_row(
                row, source_cfg, mapper,
                formatted_row=get_formatted_row(row),
                client=client, master=master,
                summary_index=summary_index, completed=completed,
            )
            params = plan.params
            row_warnings = plan.warnings
            existing_key = plan.existing_key

            if plan.action == "skip":
                print(f"  [{i}] ⚠ スキップ: {plan.reason}", file=sys.stderr)
                counts["skipped"] += 1
                log(row=i, action="skipped", detail=plan.reason)
                api_called = False          # 照合前に確定するため通信していない
                continue

            if plan.action == "resume":
                print(f"  [{i}] — 再開スキップ（前回処理済み）: {plan.summary}")
                counts["resumed"] += 1
                # 再開スキップもログに残す。残さないと --resume を繰り返した
                # ときにログが痩せ、次の再開で作成済みの行が再作成される。
                log(row=i, action="resumed", summary=plan.summary)
                api_called = False          # 照合より前に判定するため通信していない
                continue

            if plan.action == "unchanged":
                # 内容が同じ。PATCH を投げても Backlog が
                # "No comment content." を返すだけなので送らない。
                # 確認も求めない（判断することが無いため）。
                print(f"  [{i}] — 変更なし: {plan.existing_key} — {plan.summary}")
                counts["unchanged"] += 1
                if row_warnings:
                    counts["partial"] += 1
                log(row=i, action="unchanged", issue_key=plan.existing_key,
                    summary=plan.summary, detail=" / ".join(row_warnings))
                api_called = False
                continue

            # 書き込む直前に 1 行ずつ確認する。
            # まとめて一気に書き込むと、件数が多いときに何が起きたのか
            # 把握しきれないため、作成・更新のどちらも都度確認する。
            if confirmer is not None and not confirmer.confirm(plan, mapper):
                label = "更新" if existing_key else "新規作成"
                print(f"  [{i}] — スキップ（{label}を見送り）: {plan.summary}")
                counts["skipped"] += 1
                log(row=i, action="skipped", summary=plan.summary,
                    detail=f"{label}を見送り")
                api_called = False
                continue

            if existing_key:
                # projectId は更新時不要なので除去
                update_params = {k: v for k, v in params.items() if k != "projectId"}
                try:
                    client.update_issue(existing_key, update_params)
                    print(f"  [{i}] ✅ 更新: {existing_key} — {params.get('summary', '')}")
                    counts["updated"] += 1
                    if row_warnings:
                        counts["partial"] += 1
                    log(row=i, action="updated", issue_key=existing_key,
                        summary=params.get("summary", ""),
                        detail=" / ".join(row_warnings))
                except BacklogNoChangeError as nce:
                    # 実際の Backlog エラーメッセージを表示して誤検出を確認できるようにする
                    print(f"  [{i}] — 変更なし: {existing_key} — {params.get('summary', '')}")
                    print(f"    Backlog message: {nce}", file=sys.stderr)
                    counts["unchanged"] += 1
                    if row_warnings:
                        counts["partial"] += 1
                    log(row=i, action="unchanged", issue_key=existing_key,
                        summary=params.get("summary", ""),
                        detail=" / ".join([str(nce), *row_warnings]))
            else:
                try:
                    api_called = True
                    issue = create_issue_with_status(client, params)
                    print(f"  [{i}] ✅ 作成: {issue['issueKey']} — {issue['summary']}")
                    counts["created"] += 1
                    if row_warnings:
                        counts["partial"] += 1
                    log(row=i, action="created", issue_key=issue["issueKey"],
                        summary=issue["summary"], detail=" / ".join(row_warnings))
                    if summary_index is not None:
                        summary_index.add(issue["summary"], issue)
                except StatusUpdateFailed as e:
                    # 課題は作成済み。issueKey を必ず表示する（Backlog 上に
                    # 取り残されたことに気づけないと重複作成につながる）
                    issue = e.issue
                    print(
                        f"  [{i}] ⚠ 作成（ステータス変更は失敗）: "
                        f"{issue['issueKey']} — {issue['summary']}"
                    )
                    print(f"    {e.cause}", file=sys.stderr)
                    counts["created"] += 1
                    counts["status_failed"] += 1
                    if summary_index is not None:
                        summary_index.add(issue["summary"], issue)
                    if row_warnings:
                        counts["partial"] += 1
                    log(row=i, action="created_status_failed",
                        issue_key=issue["issueKey"], summary=issue["summary"],
                        detail=" / ".join([str(e.cause), *row_warnings]))

        except BacklogAPIError as e:
            print(f"  [{i}] ❌ 失敗: {params.get('summary', '')}", file=sys.stderr)
            print(f"    {e}", file=sys.stderr)
            counts["error"] += 1
            log(row=i, action="error", summary=params.get("summary", ""), detail=str(e))
            # 認証・権限エラーは以降の行でも必ず失敗するため実行全体を中止する
            if e.fatal:
                raise
        finally:
            # API レート制限対策。以前は成功パスにしか無かったため、
            # エラーが続くと逆に速くリクエストが飛んでいた。
            if api_called:
                time.sleep(0.3)

    return counts


# ------------------------------------------------------------------
# メイン
# ------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Excel から Backlog 課題を登録・更新するツール",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
例:
  python excel_to_backlog.py                        # ドライラン（デフォルト）
  python excel_to_backlog.py --preview              # プレビューファイルを生成
  python excel_to_backlog.py --execute              # 実際に登録/更新
  python excel_to_backlog.py --source "タスク管理表"          # ソース指定（ドライラン）
  python excel_to_backlog.py --source "タスク管理表" --execute # ソース指定して実行
  python excel_to_backlog.py --config ./config.yaml --execute
""",
    )
    # 設定ファイルの既定値。カレントディレクトリを優先し、無ければ
    # スクリプトと同じ場所を見る。excel-to-backlog コマンドとして
    # どこからでも実行できるようにするため。
    default_config = str(
        Path("config.yaml") if Path("config.yaml").exists()
        else Path(__file__).parent / "config.yaml"
    )
    parser.add_argument(
        "--config",
        default=default_config,
        help="設定ファイルのパス（デフォルト: カレントの config.yaml、無ければスクリプトと同じ場所）",
    )
    parser.add_argument(
        "--source",
        metavar="NAME",
        help="処理するソース名（省略時: 全ソースを処理）",
    )
    parser.add_argument(
        "--preview",
        action="store_true",
        help="登録予定の課題内容（本文全文含む）を Markdown ファイルに出力して確認する",
    )
    parser.add_argument(
        "--execute",
        action="store_true",
        help="実際に Backlog へ課題を作成/更新する（省略時はドライラン）",
    )
    parser.add_argument(
        "--limit",
        type=int,
        metavar="N",
        help="各ソースで処理する行数を先頭 N 行に制限する（初回に少数だけ試す用）",
    )
    parser.add_argument(
        "--show-columns",
        action="store_true",
        help="Excel から読み取れる列名を一覧表示する（Backlog へは接続しない）",
    )
    parser.add_argument(
        "--list-master",
        action="store_true",
        help="設定に使える名前（種別・優先度・ステータス・担当者・カスタム属性）を一覧表示する",
    )
    parser.add_argument(
        "-y", "--yes",
        action="store_true",
        help="実行前の確認を省略する（非対話環境ではこの指定が必要）",
    )
    parser.add_argument(
        "--resume",
        metavar="CSV",
        help="過去の実行ログ（run_*.csv）を読み、作成・更新済みの行を飛ばして再開する",
    )
    parser.add_argument(
        "--no-log",
        action="store_true",
        help="実行ログ（run_*.csv）を出力しない",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="API リクエストの詳細を表示する",
    )
    args = parser.parse_args()
    # デフォルトはドライラン。--execute が指定された場合のみ実処理を行う。
    dry_run = not args.execute

    if args.limit is not None and args.limit < 1:
        parser.error("--limit は 1 以上を指定してください。")

    if args.preview and args.execute:
        parser.error("--preview と --execute は同時に指定できません。")

    # 設定読み込み
    config = load_config(args.config)
    backlog_cfg = config.get("backlog", {})
    sources_cfg = config.get("sources") or []

    validate_backlog_config(backlog_cfg)

    # 設定キーの綴り間違いは dict.get() の既定値で静かに無視され、
    # 意図と違う動作になる。読み込み直後に検出する。
    key_problems = validate_config_keys(config)
    if key_problems:
        print("エラー: 設定ファイルに問題があります:", file=sys.stderr)
        for line in key_problems:
            print(line, file=sys.stderr)
        sys.exit(1)

    # --list-master は設定を書く前に使うため、sources がまだ無くても通す
    if not sources_cfg and not args.list_master:
        print("エラー: config.yaml に sources が設定されていません。", file=sys.stderr)
        sys.exit(1)

    # ソースの絞り込み
    if args.source:
        sources_cfg = [s for s in sources_cfg if s.get("name") == args.source]
        if not sources_cfg:
            names = [s.get("name", "（名前なし）") for s in config.get("sources", [])]
            print(
                f"エラー: ソース「{args.source}」が見つかりません。"
                f"（定義済み: {names}）",
                file=sys.stderr,
            )
            sys.exit(1)

    # ヘッダー
    print("=" * 55)
    print("Excel → Backlog 課題登録ツール")
    print("=" * 55)
    print(f"スペース    : {backlog_cfg['space_host']}")
    print(f"プロジェクト : {backlog_cfg['project_key']}")
    print(f"ソース数    : {len(sources_cfg)}")
    if args.preview:
        print("モード      : PREVIEW（登録内容をMarkdownファイルに出力します）")
    elif dry_run:
        print("モード      : DRY RUN（実際の作成/更新は行いません）")
    else:
        print("モード      : EXECUTE（Backlog に登録/更新します）")
    print()

    # --show-columns: Excel の列名だけ表示して終了。
    # 設定を書いている途中で使うため、Backlog への接続は行わない。
    if args.show_columns:
        sys.exit(1 if print_source_columns(sources_cfg) else 0)

    # BacklogClient 初期化
    client = BacklogClient(
        space_host=backlog_cfg["space_host"],
        api_key=backlog_cfg["api_key"],
        ssl_verify=backlog_cfg.get("ssl_verify", True),
        base_path=backlog_cfg.get("base_path", ""),
        debug=args.debug,
    )

    # マスターデータ取得（ドライランでも接続確認のため取得）
    print("マスターデータを取得中...")
    master = BacklogMaster.build(client, backlog_cfg["project_key"])

    # --list-master: 設定に書ける名前を一覧表示して終了
    if args.list_master:
        print_master_data(master)
        return

    print(
        f"  種別: {list(master.issue_type_map.keys())}\n"
        f"  優先度: {list(master.priority_map.keys())}\n"
        f"  ステータス: {list(master.status_map.keys())}\n"
        f"  メンバー数: {len(master.user_map)} 名"
    )

    # --preview モード: Markdown ファイルを生成して終了
    if args.preview:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = Path(args.config).parent
        print(f"プレビューファイルを生成中...")
        results = generate_preview_file(sources_cfg, master, output_dir, timestamp)
        total_issues = sum(count for _, count in results)
        print(f"\n{'='*55}")
        print("プレビュー生成完了")
        print(f"{'='*55}")
        for path, count in results:
            print(f"  {path.name}  （{count} 件）")
        print(f"{'─'*55}")
        print(f"  合計: {total_issues} 件")
        print()
        print("  内容を確認後、実際に登録するには --execute を付けて再実行してください。")
        return

    # 各ソースを処理
    # 途中で中断・失敗しても、それまでの処理結果を必ず表示する。
    # 以前は通信エラーがトレースバックのまま抜けてサマリーが出ず、
    # 何件作成済みかを知る手段がターミナルのログしか無かった。
    total = new_counts()
    interrupted = ""

    # 実行ログ。ドライランでは書き込みが発生しないため出力しない。
    log_path = None
    if not dry_run and not args.no_log:
        log_path = default_log_path(
            Path(args.config).parent, datetime.now().strftime("%Y%m%d_%H%M%S")
        )

    completed = load_completed(args.resume) if args.resume else None

    # 件名索引は遅延構築のため、match_summary を使うソースが無ければ API を呼ばない
    summary_index = SummaryIndex(client, master.project_id)

    # 確認は Backlog へ書き込む場合のみ。ドライランは何も変更しないため、
    # 確認を求めると無意味な入力待ちになり、非対話環境では実行すらできなくなる。
    if not dry_run:
        # 確認の前に、書き込まないモードで計画だけ算出する。
        # ドライランと実行は同じ plan_row() を通るため、ここで出る内訳は
        # 実際の結果と一致する。照合の読み取りだけ行い Backlog は変更しない。
        planned = new_counts()
        try:
            # 索引は本処理と共有しない。ドライランは「作成予定」を表す
            # ダミーの issueKey を索引へ入れるため、共有すると本処理が
            # それを既存課題と誤認し、存在しないキーに対して更新を実行する。
            planning_index = SummaryIndex(client, master.project_id)
            with contextlib.redirect_stdout(io.StringIO()):
                for source_cfg in sources_cfg:
                    process_source(
                        source_cfg, client, master, dry_run=True,
                        completed=completed, summary_index=planning_index,
                        counts=planned, limit=args.limit,
                    )
        except (BacklogAPIError, KeyboardInterrupt):
            # 事前算出に失敗しても本処理は試みる（そこで改めて報告される）
            planned = None

        try:
            confirm_run(sources_cfg, master, assume_yes=args.yes, planned=planned)
        except ConfirmationDeclined as e:
            print(f"\n  {e}", file=sys.stderr)
            sys.exit(1)

    with ExitStack() as stack:
        run_log = stack.enter_context(RunLog(log_path)) if log_path else None
        try:
            # 実行時は 1 行ずつ確認する（ドライランでは確認しない）
            confirmer = (
                None if dry_run
                else RowConfirmer(assume_yes=args.yes,
                                  master_labels=build_master_labels(master))
            )
            for source_cfg in sources_cfg:
                # total を直接渡す。ソースの途中で中断しても、それまでの
                # 集計が total に残りサマリーに反映される。
                process_source(
                    source_cfg, client, master, dry_run=dry_run,
                    run_log=run_log, completed=completed,
                    summary_index=summary_index, counts=total,
                    limit=args.limit, confirmer=confirmer,
                )
        except ConfirmationDeclined as e:
            interrupted = str(e)
        except KeyboardInterrupt:
            interrupted = "ユーザーによる中断（Ctrl-C）"
        except BacklogAPIError as e:
            interrupted = f"API エラーのため中止しました\n  {e}"
        finally:
            print_summary(
                total, dry_run=dry_run, interrupted=interrupted, log_path=log_path
            )

    # 中断だけでなく、1 件でも失敗があれば異常終了とする。
    # 以前は全行が設定ミスで失敗しても終了コード 0 で終わっていた。
    if interrupted or total["error"]:
        sys.exit(1)


def print_summary(total: dict, *, dry_run: bool, interrupted: str = "",
                  log_path=None) -> None:
    """処理結果のサマリーを表示する。"""
    print(f"\n{'='*55}")
    print("処理完了" if not interrupted else "処理中断")
    print(f"{'='*55}")

    if interrupted:
        print(f"  ⚠ {interrupted}")
        print()

    if dry_run:
        # ドライランは実行と同じ計画を組み立てているため、作成/更新の内訳まで
        # 予測できる。算出しているのに表示しないと RowPlan を入れた意味がない。
        print("（DRY RUN のため実際の登録は行っていません）")
        print(f"  作成予定: {total['created']} 件")
        print(f"  更新予定: {total['updated']} 件")
        if total["resumed"]:
            print(f"  再開スキップ: {total['resumed']} 件（前回処理済み）")
        if total["skipped"]:
            print(f"  スキップ: {total['skipped']} 件")
        if total["partial"]:
            print(f"  うち一部フィールド未設定: {total['partial']} 件")
        if total["error"]:
            print(f"  エラー: {total['error']} 件  ← 読み込み・設定に問題があります")
        print("  実際に登録するには --execute を付けて再実行してください。")
        return

    print(f"  作成: {total['created']} 件")
    print(f"  更新: {total['updated']} 件")
    print(f"  変更なし: {total['unchanged']} 件")
    print(f"  スキップ: {total['skipped']} 件")
    if total["partial"]:
        print(f"  うち一部フィールド未設定: {total['partial']} 件")
        print("    実行ログの detail 列に、どのフィールドが落ちたか記録しています。")
    if total["resumed"]:
        print(f"  再開スキップ: {total['resumed']} 件（前回処理済み）")
    print(f"  エラー: {total['error']} 件")
    if total["status_failed"]:
        print()
        print(
            f"  ⚠ うち {total['status_failed']} 件は作成できましたが"
            f"ステータス変更に失敗しています。"
        )
        print("    課題は Backlog に存在します。上のログの issueKey を確認してください。")

    if log_path is not None:
        print()
        print(f"  実行ログ: {log_path}")
        if interrupted or total["error"]:
            print(f"  続きから再開するには: --resume {log_path.name}")


if __name__ == "__main__":
    main()
