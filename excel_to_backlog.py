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

  # API リクエスト詳細を表示（デバッグ）
  python excel_to_backlog.py --execute --debug
"""

import argparse
import sys
import time
from pathlib import Path

import yaml

from backlog_client import BacklogClient
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
    filters_cfg = source_cfg.get("filters") or []
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

    # フィルター条件の列名チェック（警告のみ）
    for cond in filters_cfg:
        col = cond.get("col_name", "")
        if col and col not in headers:
            print(
                f"  ⚠ フィルター列「{col}」がヘッダーに存在しません。"
                f"（ヘッダー: {headers}）",
                file=sys.stderr,
            )

    # フィルタリング
    filtered_rows = ExcelReader.filter_rows(rows, filters_cfg)
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
            print(mapper.format_dry_run(row, i))
        return counts

    # ---- 実処理 ----
    for i, row in enumerate(filtered_rows, 1):
        try:
            params = mapper.map_row(row)
        except ValueError as e:
            print(f"  [{i}] ⚠ スキップ: {e}", file=sys.stderr)
            counts["skipped"] += 1
            continue

        try:
            if upsert_enabled:
                existing_key = find_existing_issue(client, upsert_cfg, row, params, master)
                if existing_key:
                    # projectId は更新時不要なので除去
                    update_params = {k: v for k, v in params.items() if k != "projectId"}
                    client.update_issue(existing_key, update_params)
                    print(f"  [{i}] ✅ 更新: {existing_key} — {params.get('summary', '')}")
                    counts["updated"] += 1
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
    if dry_run:
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
