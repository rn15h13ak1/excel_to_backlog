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
import re
import sys
import time
from contextlib import ExitStack
from datetime import datetime
from pathlib import Path

import yaml

from backlog_client import BacklogAPIError, BacklogClient, BacklogNoChangeError
from excel_reader import ExcelReader
from mapper import BacklogMaster, IssueMapper
from run_log import RunLog, default_log_path, load_completed
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
) -> str | None:
    """
    upsert 設定に従い既存課題の issueKey を返す。
    見つからない場合は None を返す。

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
                return existing["issueKey"]
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
META_KEYS = {"_source_name", "_excel_path", "_excel_sheet"}


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
    for i, cond in enumerate(source_cfg.get("filters") or []):
        add(f"filters[{i}].col_name", cond.get("col_name"))
    for gi, group in enumerate(source_cfg.get("filter_groups") or []):
        for i, cond in enumerate(group.get("filters") or []):
            add(f"filter_groups[{gi}].filters[{i}].col_name", cond.get("col_name"))

    # ---- 件名 ----
    summary_template = mapping_cfg.get("summary_template", "")
    if summary_template:
        add_template("issue_mapping.summary_template", summary_template)
    else:
        add("issue_mapping.summary_col", mapping_cfg.get("summary_col"))

    # ---- 本文 ----
    # description_template は template モードでのみ使われる
    if mapping_cfg.get("description_format", "template") != "auto":
        add_template(
            "issue_mapping.description_template",
            mapping_cfg.get("description_template", ""),
        )
    # description_cols は auto モードと {{auto}} の両方で使われるため常に検証する
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
    use_rich_text = bool(mapping_cfg.get("rich_text"))

    # Excel 読み込み
    try:
        reader = ExcelReader(excel_cfg)
        if use_rich_text:
            headers, rows, formatted_rows_all = reader.read_with_format()
        else:
            headers, rows = reader.read()
            formatted_rows_all = None
    except Exception as e:
        import traceback
        lines.append(f"> ⚠ Excel 読み込みエラー: {e}")
        lines.append("")
        lines.append("> ヒント: Excelファイルが破損または非標準形式の可能性があります。")
        lines.append(f"> 詳細: `{traceback.format_exc()}`")
        lines.append("")
        output_path.write_text("\n".join(lines), encoding="utf-8")
        return 0

    # 列名参照の検証（process_source と同じ基準で中止する）
    errors = validate_column_references(source_cfg, headers)
    if errors:
        lines.append(f"> ⚠ {errors[0]}")
        lines.append("")
        for line in errors[1:]:
            lines.append(f"> `{line.strip()}`")
        lines.append("")
        output_path.write_text("\n".join(lines), encoding="utf-8")
        return 0

    filtered_rows = apply_filters(rows, source_cfg, headers)
    lines.append(f"対象行数: **{len(filtered_rows)} 件**（フィルター後）")
    lines.append("")

    if not filtered_rows:
        lines.append("_対象行がありません。_")
        lines.append("")
        output_path.write_text("\n".join(lines), encoding="utf-8")
        return 0

    mapper = IssueMapper(mapping_cfg, master, headers=headers)

    # フィルタ後の行インデックスを plain_rows 全体の中から特定する
    # （formatted_rows_all は plain rows と同じ順序・同じ件数）
    plain_row_ids = {id(r): idx for idx, r in enumerate(rows)} if formatted_rows_all else {}

    for i, row in enumerate(filtered_rows, 1):
        enriched = inject_meta(row, source_cfg)
        if formatted_rows_all is not None:
            orig_idx = plain_row_ids.get(id(row))
            fmt_row = formatted_rows_all[orig_idx] if orig_idx is not None else None
            fmt_enriched = inject_meta(fmt_row, source_cfg) if fmt_row is not None else None
        else:
            fmt_enriched = None
        lines.append(mapper.format_preview(enriched, i, master_labels=master_labels, formatted_row=fmt_enriched))
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

def confirm_create(params: dict, index: int) -> bool:
    """
    新規作成前にユーザーへ確認を求める。
    y / yes を入力した場合のみ True を返す。デフォルトは No（スキップ）。

    表示する情報:
        件名・種別ID・優先度ID・期限日（設定されている場合）
    """
    summary   = params.get("summary", "（件名なし）")
    due_date  = params.get("dueDate", "")

    print(f"\n  [{index}] 新規作成の確認:")
    print(f"    件名  : {summary}")
    if due_date:
        print(f"    期限日: {due_date}")

    try:
        answer = input("    Backlog に新規作成しますか？ [y/N]: ").strip().lower()
    except EOFError:
        # 非対話環境（パイプ等）ではデフォルト No
        answer = ""

    return answer in ("y", "yes")


# ------------------------------------------------------------------
# 1ソースの処理
# ------------------------------------------------------------------

def process_source(
    source_cfg: dict,
    client: BacklogClient,
    master: BacklogMaster,
    dry_run: bool,
    run_log: RunLog | None = None,
    completed: set[tuple[str, str]] | None = None,
    summary_index: SummaryIndex | None = None,
) -> dict:
    """
    1つのソース（Excel ファイル）を処理して作成・更新件数を返す。

    Returns
    -------
    dict: new_counts() と同じキーを持つ集計結果
    """
    name = source_cfg.get("name", "（名前なし）")
    excel_cfg = source_cfg.get("excel", {})
    mapping_cfg = source_cfg.get("issue_mapping", {})
    upsert_cfg = source_cfg.get("upsert") or {}
    upsert_enabled = upsert_cfg.get("enabled", False)

    counts = new_counts()

    print(f"\n{'='*55}")
    print(f"ソース: {name}")
    print(f"{'='*55}")
    print(f"  ファイル: {excel_cfg.get('path', '（未設定）')}")
    print(f"  シート : {excel_cfg.get('sheet', '（最初のシート）')}")
    print(f"  upsert : {'有効' if upsert_enabled else '無効（常に新規作成）'}")

    use_rich_text = bool(mapping_cfg.get("rich_text"))

    # ---- Excel 読み込み ----
    try:
        reader = ExcelReader(excel_cfg)
        if use_rich_text:
            headers, rows, formatted_rows_all = reader.read_with_format()
        else:
            headers, rows = reader.read()
            formatted_rows_all = None
    except Exception as e:
        import traceback
        print(f"\n  エラー: Excel の読み込みに失敗しました: {e}", file=sys.stderr)
        print("  ヒント: Excelファイルが破損または非標準形式の可能性があります。", file=sys.stderr)
        print("         詳細:", file=sys.stderr)
        traceback.print_exc(file=sys.stderr)
        counts["error"] += 1
        return counts

    print(f"  読込行数: {len(rows)} 行（フィルター前）")

    # ---- 列名参照の検証 ----
    # 列名が1つでも一致しないと、フィルター条件が無視されて全行が対象になる、
    # プレースホルダーが未展開のまま件名になる等の無言の誤動作が起きるため、
    # 行を処理する前に中止する。
    errors = validate_column_references(source_cfg, headers)
    if errors:
        print(f"\n  エラー: {errors[0]}", file=sys.stderr)
        for line in errors[1:]:
            print(line, file=sys.stderr)
        counts["error"] += 1
        return counts

    # フィルタリング（filters / filter_groups 共通処理）
    filtered_rows = apply_filters(rows, source_cfg, headers)
    print(f"  対象行数: {len(filtered_rows)} 行（フィルター後）")

    if not filtered_rows:
        print("  → 対象行がないためスキップします。")
        return counts

    # ---- マッパー初期化 ----
    mapper = IssueMapper(mapping_cfg, master, headers=headers)

    # フィルタ後の行を plain_rows 全体のインデックスに対応付ける
    plain_row_ids = {id(r): idx for idx, r in enumerate(rows)} if formatted_rows_all else {}

    def get_formatted_row(plain_row: dict) -> dict | None:
        """plain_row に対応する書式付き行を返す。rich_text 無効時は None。"""
        if formatted_rows_all is None:
            return None
        orig_idx = plain_row_ids.get(id(plain_row))
        return formatted_rows_all[orig_idx] if orig_idx is not None else None

    # ---- ドライラン ----
    if dry_run:
        print(f"\n  [DRY RUN] 以下の課題を作成/更新します:\n")
        for i, row in enumerate(filtered_rows, 1):
            fmt_row = get_formatted_row(row)
            enriched = inject_meta(row, source_cfg)
            fmt_enriched = inject_meta(fmt_row, source_cfg) if fmt_row is not None else None
            print(mapper.format_dry_run(enriched, i, formatted_row=fmt_enriched))
        return counts

    # ---- 実処理 ----
    def log(*, row: int, action: str, issue_key: str = "", summary: str = "", detail: str = "") -> None:
        """実行ログに1件記録する（--log-file 未指定時は何もしない）。"""
        if run_log is not None:
            run_log.record(
                source=name, row=row, action=action,
                issue_key=issue_key, summary=summary, detail=detail,
            )

    for i, row in enumerate(filtered_rows, 1):
        fmt_row = get_formatted_row(row)
        enriched = inject_meta(row, source_cfg)
        fmt_enriched = inject_meta(fmt_row, source_cfg) if fmt_row is not None else None
        try:
            params = mapper.map_row(enriched, formatted_row=fmt_enriched)
        except ValueError as e:
            print(f"  [{i}] ⚠ スキップ: {e}", file=sys.stderr)
            counts["skipped"] += 1
            log(row=i, action="skipped", detail=str(e))
            continue

        # --resume: 前回の実行で作成・更新まで完了した行は飛ばす
        if completed is not None and (name, params.get("summary", "")) in completed:
            print(f"  [{i}] — 再開スキップ（前回処理済み）: {params.get('summary', '')}")
            counts["resumed"] += 1
            continue

        # API を1度でも呼んだ行だけレート制限用の待機を入れる
        # （確認プロンプトでキャンセルした行は通信していないため待つ意味がない）
        api_called = upsert_enabled
        try:
            existing_key = (
                find_existing_issue(client, upsert_cfg, enriched, params, master,
                                    summary_index=summary_index)
                if upsert_enabled
                else None
            )

            if existing_key:
                # projectId は更新時不要なので除去
                update_params = {k: v for k, v in params.items() if k != "projectId"}
                try:
                    client.update_issue(existing_key, update_params)
                    print(f"  [{i}] ✅ 更新: {existing_key} — {params.get('summary', '')}")
                    counts["updated"] += 1
                    log(row=i, action="updated", issue_key=existing_key,
                        summary=params.get("summary", ""))
                except BacklogNoChangeError as nce:
                    # 実際の Backlog エラーメッセージを表示して誤検出を確認できるようにする
                    print(f"  [{i}] — 変更なし: {existing_key} — {params.get('summary', '')}")
                    print(f"    Backlog message: {nce}", file=sys.stderr)
                    counts["unchanged"] += 1
                    log(row=i, action="unchanged", issue_key=existing_key,
                        summary=params.get("summary", ""), detail=str(nce))
            else:
                if not confirm_create(params, i):
                    print(f"  [{i}] — スキップ（新規作成をキャンセル）: {params.get('summary', '')}")
                    counts["skipped"] += 1
                    log(row=i, action="skipped", summary=params.get("summary", ""),
                        detail="新規作成をキャンセル")
                    continue
                try:
                    api_called = True
                    issue = create_issue_with_status(client, params)
                    print(f"  [{i}] ✅ 作成: {issue['issueKey']} — {issue['summary']}")
                    counts["created"] += 1
                    log(row=i, action="created", issue_key=issue["issueKey"],
                        summary=issue["summary"])
                    if summary_index is not None:
                        summary_index.add(issue["summary"], issue["issueKey"])
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
                        summary_index.add(issue["summary"], issue["issueKey"])
                    log(row=i, action="created_status_failed",
                        issue_key=issue["issueKey"], summary=issue["summary"],
                        detail=str(e.cause))

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
    default_config = str(Path(__file__).parent / "config.yaml")
    parser.add_argument(
        "--config",
        default=default_config,
        help="設定ファイルのパス（デフォルト: スクリプトと同じディレクトリの config.yaml）",
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

    if args.preview and args.execute:
        parser.error("--preview と --execute は同時に指定できません。")

    # 設定読み込み
    config = load_config(args.config)
    backlog_cfg = config.get("backlog", {})
    sources_cfg = config.get("sources") or []

    validate_backlog_config(backlog_cfg)

    if not sources_cfg:
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

    with ExitStack() as stack:
        run_log = stack.enter_context(RunLog(log_path)) if log_path else None
        try:
            for source_cfg in sources_cfg:
                counts = process_source(
                    source_cfg, client, master, dry_run=dry_run,
                    run_log=run_log, completed=completed,
                    summary_index=summary_index,
                )
                for k in total:
                    total[k] += counts[k]
        except KeyboardInterrupt:
            interrupted = "ユーザーによる中断（Ctrl-C）"
        except BacklogAPIError as e:
            interrupted = f"API エラーのため中止しました\n  {e}"
        finally:
            print_summary(
                total, dry_run=dry_run, interrupted=interrupted, log_path=log_path
            )

    if interrupted:
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
        print("（DRY RUN のため実際の登録は行っていません）")
        if total["skipped"]:
            print(f"  スキップ: {total['skipped']} 件")
        if total["error"]:
            print(f"  エラー: {total['error']} 件  ← 読み込み・設定に問題があります")
        print("  実際に登録するには --execute を付けて再実行してください。")
        return

    print(f"  作成: {total['created']} 件")
    print(f"  更新: {total['updated']} 件")
    print(f"  変更なし: {total['unchanged']} 件")
    print(f"  スキップ: {total['skipped']} 件")
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
