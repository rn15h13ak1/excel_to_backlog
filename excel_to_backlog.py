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
import sys
import time
from datetime import datetime
from pathlib import Path

import yaml

from backlog_client import BacklogClient, BacklogNoChangeError
from excel_reader import ExcelReader
from mapper import BacklogMaster, IssueMapper


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
# upsert ロジック
# ------------------------------------------------------------------

def find_existing_issue(
    client: BacklogClient,
    upsert_cfg: dict,
    row: dict,
    params: dict,
    master: BacklogMaster,
) -> str | None:
    """
    upsert 設定に従い既存課題の issueKey を返す。
    見つからない場合は None を返す。

    upsert_cfg キー:
        key_col       : str  Excel の列名（issueKey が記入されている列）
        match_summary : bool 件名で検索して一致する課題を探す
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

    # ② 件名で検索
    if upsert_cfg.get("match_summary"):
        summary = params.get("summary", "")
        if summary:
            candidates = client.search_issues_by_summary(master.project_id, summary)
            exact = [i for i in candidates if i.get("summary") == summary]
            if exact:
                return exact[0]["issueKey"]

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
        }
    """
    return {
        "issue_type": {id_: name for name, id_ in master.issue_type_map.items()},
        "priority":   {id_: name for name, id_ in master.priority_map.items()},
        "user":       {id_: name for name, id_ in master.user_map.items()},
    }


def generate_preview_file(
    sources_cfg: list,
    client: BacklogClient,
    master: BacklogMaster,
    output_path: Path,
) -> int:
    """
    全ソースの登録予定内容を Markdown ファイルに書き出す。
    Backlog API には接続するがデータの書き込みは行わない。

    Returns
    -------
    int : プレビュー生成した課題の総件数
    """
    master_labels = build_master_labels(master)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    total_issues = 0

    lines = [
        "# Backlog 課題登録 プレビュー",
        "",
        f"> 生成日時: {now}  ",
        f"> ※ このファイルは登録前の確認用です。実際の登録は `--execute` で行います。",
        "",
        "---",
        "",
    ]

    for source_cfg in sources_cfg:
        name = source_cfg.get("name", "（名前なし）")
        excel_cfg = source_cfg.get("excel", {})
        mapping_cfg = source_cfg.get("issue_mapping", {})

        lines.append(f"# ソース: {name}")
        lines.append("")
        lines.append(f"- ファイル: `{excel_cfg.get('path', '（未設定）')}`")
        lines.append(f"- シート: `{excel_cfg.get('sheet', '（最初のシート）')}`")
        lines.append("")

        # Excel 読み込み
        try:
            reader = ExcelReader(excel_cfg)
            headers, rows = reader.read()
        except Exception as e:
            lines.append(f"> ⚠ Excel 読み込みエラー: {e}")
            lines.append("")
            lines.append("---")
            lines.append("")
            continue

        filtered_rows = apply_filters(rows, source_cfg, headers)
        lines.append(f"対象行数: **{len(filtered_rows)} 件**（フィルター後）")
        lines.append("")

        if not filtered_rows:
            lines.append("_対象行がありません。_")
            lines.append("")
            lines.append("---")
            lines.append("")
            continue

        mapper = IssueMapper(mapping_cfg, master, headers=headers)

        for i, row in enumerate(filtered_rows, 1):
            enriched = inject_meta(row, source_cfg)
            lines.append(mapper.format_preview(enriched, i, master_labels=master_labels))
            lines.append("")
            lines.append("---")
            lines.append("")
            total_issues += 1

    output_path.write_text("\n".join(lines), encoding="utf-8")
    return total_issues


# ------------------------------------------------------------------
# 1ソースの処理
# ------------------------------------------------------------------

def process_source(
    source_cfg: dict,
    client: BacklogClient,
    master: BacklogMaster,
    dry_run: bool,
) -> dict:
    """
    1つのソース（Excel ファイル）を処理して作成・更新件数を返す。

    Returns
    -------
    dict: {"created": int, "updated": int, "skipped": int, "error": int}
    """
    name = source_cfg.get("name", "（名前なし）")
    excel_cfg = source_cfg.get("excel", {})
    mapping_cfg = source_cfg.get("issue_mapping", {})
    upsert_cfg = source_cfg.get("upsert") or {}
    upsert_enabled = upsert_cfg.get("enabled", False)

    counts = {"created": 0, "updated": 0, "skipped": 0, "error": 0}

    print(f"\n{'='*55}")
    print(f"ソース: {name}")
    print(f"{'='*55}")
    print(f"  ファイル: {excel_cfg.get('path', '（未設定）')}")
    print(f"  シート : {excel_cfg.get('sheet', '（最初のシート）')}")
    print(f"  upsert : {'有効' if upsert_enabled else '無効（常に新規作成）'}")

    # ---- Excel 読み込み ----
    try:
        reader = ExcelReader(excel_cfg)
        headers, rows = reader.read()
    except (FileNotFoundError, ValueError, Exception) as e:
        print(f"\n  エラー: Excel の読み込みに失敗しました: {e}", file=sys.stderr)
        counts["error"] += 1
        return counts

    print(f"  読込行数: {len(rows)} 行（フィルター前）")

    # フィルタリング（filters / filter_groups 共通処理）
    filtered_rows = apply_filters(rows, source_cfg, headers)
    print(f"  対象行数: {len(filtered_rows)} 行（フィルター後）")

    if not filtered_rows:
        print("  → 対象行がないためスキップします。")
        return counts

    # ---- マッパー初期化 ----
    mapper = IssueMapper(mapping_cfg, master, headers=headers)

    # ---- ドライラン ----
    if dry_run:
        print(f"\n  [DRY RUN] 以下の課題を作成/更新します:\n")
        for i, row in enumerate(filtered_rows, 1):
            print(mapper.format_dry_run(inject_meta(row, source_cfg), i))
        return counts

    # ---- 実処理 ----
    for i, row in enumerate(filtered_rows, 1):
        enriched = inject_meta(row, source_cfg)
        try:
            params = mapper.map_row(enriched)
        except ValueError as e:
            print(f"  [{i}] ⚠ スキップ: {e}", file=sys.stderr)
            counts["skipped"] += 1
            continue

        try:
            if upsert_enabled:
                existing_key = find_existing_issue(client, upsert_cfg, enriched, params, master)
                if existing_key:
                    # projectId は更新時不要なので除去
                    update_params = {k: v for k, v in params.items() if k != "projectId"}
                    try:
                        client.update_issue(existing_key, update_params)
                        print(f"  [{i}] ✅ 更新: {existing_key} — {params.get('summary', '')}")
                        counts["updated"] += 1
                    except BacklogNoChangeError:
                        print(f"  [{i}] — スキップ（変更なし）: {existing_key} — {params.get('summary', '')}")
                        counts["skipped"] += 1
                        continue
                else:
                    issue = client.create_issue(params)
                    print(f"  [{i}] ✅ 作成: {issue['issueKey']} — {issue['summary']}")
                    counts["created"] += 1
            else:
                issue = client.create_issue(params)
                print(f"  [{i}] ✅ 作成: {issue['issueKey']} — {issue['summary']}")
                counts["created"] += 1

        except SystemExit:
            # BacklogClient のエラーは sys.exit(1) を呼ぶが、
            # 1件の失敗で全体を止めないように続行する
            counts["error"] += 1
            continue

        # API レート制限対策
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
        f"  メンバー数: {len(master.user_map)} 名"
    )

    # --preview モード: Markdown ファイルを生成して終了
    if args.preview:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        preview_path = Path(args.config).parent / f"preview_{timestamp}.md"
        print(f"プレビューファイルを生成中: {preview_path}")
        total_issues = generate_preview_file(sources_cfg, client, master, preview_path)
        print(f"\n{'='*55}")
        print("プレビュー生成完了")
        print(f"{'='*55}")
        print(f"  出力ファイル : {preview_path}")
        print(f"  課題数      : {total_issues} 件")
        print()
        print("  内容を確認後、実際に登録するには --execute を付けて再実行してください。")
        return

    # 各ソースを処理
    total = {"created": 0, "updated": 0, "skipped": 0, "error": 0}
    for source_cfg in sources_cfg:
        counts = process_source(source_cfg, client, master, dry_run=dry_run)
        for k in total:
            total[k] += counts[k]

    # サマリー
    print(f"\n{'='*55}")
    print("処理完了")
    print(f"{'='*55}")
    if dry_run:
        print("（DRY RUN のため実際の登録は行っていません）")
        print("  実際に登録するには --execute を付けて再実行してください。")
    else:
        print(f"  作成: {total['created']} 件")
        print(f"  更新: {total['updated']} 件")
        print(f"  スキップ: {total['skipped']} 件")
        print(f"  エラー: {total['error']} 件")


if __name__ == "__main__":
    main()
