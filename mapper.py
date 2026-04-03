"""
Excel 行 → Backlog 課題パラメータ 変換モジュール
=================================================
config の issue_mapping 設定に従い、Excel の1行データを
Backlog API の課題作成/更新パラメータ（dict）に変換する。
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass, field


# ------------------------------------------------------------------
# BacklogMaster: 名前 → ID マッピングを保持するコンテナ
# ------------------------------------------------------------------

@dataclass
class BacklogMaster:
    """BacklogClient から取得したマスターデータを格納する"""

    project_id: int = 0
    issue_type_map: dict[str, int] = field(default_factory=dict)   # {種別名: ID}
    priority_map: dict[str, int] = field(default_factory=dict)     # {優先度名: ID}
    user_map: dict[str, int] = field(default_factory=dict)         # {ユーザー名: ID}
    status_map: dict[str, int] = field(default_factory=dict)       # {ステータス名: ID}
    # {属性名: {id, typeId, items: {選択肢名: ID}}}
    custom_field_map: dict[str, dict] = field(default_factory=dict)

    @classmethod
    def build(cls, client, project_key: str) -> "BacklogMaster":
        """
        BacklogClient を使ってマスターデータを一括取得して BacklogMaster を生成する。
        """
        master = cls()

        # プロジェクト
        print("  プロジェクト情報を取得中...")
        project = client.get_project(project_key)
        master.project_id = project["id"]

        # 種別
        print("  種別一覧を取得中...")
        issue_types = client.get_issue_types(project_key)
        master.issue_type_map = {it["name"]: it["id"] for it in issue_types}

        # 優先度
        print("  優先度一覧を取得中...")
        priorities = client.get_priorities()
        master.priority_map = {p["name"]: p["id"] for p in priorities}

        # プロジェクトメンバー
        print("  プロジェクトメンバーを取得中...")
        try:
            users = client.get_project_users(project_key)
            # 表示名（name）とログインID（userId）の両方でルックアップできるようにする
            # 同じ numeric id に複数キーが紐づく場合があるが問題なし
            user_map: dict[str, int] = {}
            for u in users:
                user_map[u["name"]] = u["id"]
                if u.get("userId"):
                    user_map[u["userId"]] = u["id"]
            master.user_map = user_map
        except SystemExit:
            # 権限不足で取得できない場合は空のまま続行
            print("  ⚠ プロジェクトメンバーの取得に失敗しました（担当者の解決はスキップされます）",
                  file=sys.stderr)

        # カスタム属性
        print("  カスタム属性一覧を取得中...")
        try:
            custom_fields = client.get_custom_fields(project_key)
            master.custom_field_map = {
                cf["name"]: {
                    "id": cf["id"],
                    "typeId": cf.get("typeId"),
                    "items": {
                        item["name"]: item["id"]
                        for item in cf.get("items", [])
                    },
                }
                for cf in custom_fields
            }
        except SystemExit:
            print("  ⚠ カスタム属性の取得に失敗しました", file=sys.stderr)

        # ステータス
        print("  ステータス一覧を取得中...")
        try:
            statuses = client.get_statuses(project_key)
            master.status_map = {s["name"]: s["id"] for s in statuses}
        except SystemExit:
            print("  ⚠ ステータスの取得に失敗しました", file=sys.stderr)

        return master


# ------------------------------------------------------------------
# IssueMapper: Excel 行 → Backlog API パラメータ
# ------------------------------------------------------------------

class IssueMapper:
    """
    mapping_config (sources[i].issue_mapping) に従い、
    Excel の行データを Backlog API パラメータに変換する。

    mapping_config キー:
        issue_type          : str        種別名（固定値）
        priority            : str        優先度名（固定値、デフォルト: "中"）
        summary_col         : str        件名として使う列名（summary_template と排他）
        summary_template    : str        件名テンプレート（{{列名}} でセル値を埋め込み、summary_col より優先）
        description_template: str        詳細欄テンプレート（{{列名}} でセル値を埋め込み）
        due_date_col        : str        期限日列名、または {{列名}} テンプレート（任意）
        start_date_col      : str        開始日列名、または {{列名}} テンプレート（任意）
        assignee_col        : str        担当者列名（任意）
        default_assignee    : str        担当者のデフォルト値（任意）
                                         assignee_col が未設定、またはセルが空の場合に使用する
                                         担当者名（Backlog 表示名 or ログインID）。
                                         セルに値がある場合はセル値が優先される。
        required_cols       : list[str]  値が空の場合にスキップする列名リスト（任意）
                                         リスト内のいずれか1列でも空であればその行を処理しない。
                                         項番だけ記入されその他が未記入の行を除外したい場合などに使用。
        custom_fields       : list       カスタム属性マッピングリスト
            - field_name    : str        Backlog カスタム属性名
              col_name      : str        Excel 列名
              value_map     : dict       Excel 値 → Backlog 値 の変換テーブル（任意）
                                         定義した場合は Excel のセル値をテーブルで変換してから Backlog に渡す。
                                         テーブルに存在しない値はスキップ（警告を出力）。
                                         省略時は Excel のセル値をそのまま使用する。
                                         例: {"A": "カテゴリA", "B": "カテゴリB"}
        status_col          : str        Excel のステータス列名（任意）
        status_map          : dict       Excel ステータス値 → Backlog ステータス名 のマッピング（任意）
                                         例: {"未着手": "未対応", "対応中": "処理中", "完了": "完了"}
                                         status_col が設定されている場合に使用。
                                         マッピングに存在しない値はスキップ（警告を出力）。
    description_format : str  "template"（デフォルト）または "auto"
        "auto" の場合は description_template を無視し、excel_md_tool と同じ形式で
        列名を見出し・セル値を本文として自動生成する。
    description_cols   : list  "auto" 時に出力する列名リスト（省略時: 全列）
    """

    def __init__(self, mapping_config: dict, master: BacklogMaster, headers: list[str] = None):
        self.cfg = mapping_config
        self.master = master
        self.headers = headers or []  # auto モードでの列順序に使用

    # ------------------------------------------------------------------
    # テンプレート処理
    # ------------------------------------------------------------------

    def _render_template(self, template: str, row: dict[str, str]) -> str:
        """
        {{列名}} を行のセル値に置換する。
        存在しない列名はそのまま残す（警告なし）。

        特殊プレースホルダー:
          {{auto}}          : _render_auto() の出力に展開される。
                              description_format が "template" でも auto 方式の出力を
                              任意の位置に埋め込めるため、ヘッダー・フッターの付与に使える。

          {{#列名}}...{{/列名}} : 条件ブロック。
                              指定列の値が空でなければブロック内を出力し、
                              空であればブロック全体を出力しない（セパレーター等の
                              「値がある場合のみ表示したい文字列」に使用）。
                              例: "項番{{項番}}{{#枝番}}-{{枝番}}{{/枝番}}"
                                → 枝番="1" → "項番1-1"
                                → 枝番=""  → "項番1"
        """
        # Step 1: 条件ブロック {{#列名}}...{{/列名}} を処理
        # 値が非空 → ブロック内テキストをそのまま残す（Step 2 でさらに展開）
        # 値が空   → ブロック全体を除去
        def cond_replacer(m: re.Match) -> str:
            col = m.group(1).strip()
            inner = m.group(2)
            return inner if row.get(col, "") else ""

        result = re.sub(
            r"\{\{#(.+?)\}\}(.*?)\{\{/\1\}\}",
            cond_replacer,
            template,
            flags=re.DOTALL,
        )

        # Step 2: 通常プレースホルダー {{列名}} を展開
        def replacer(m: re.Match) -> str:
            col = m.group(1).strip()
            if col == "auto":
                return self._render_auto(row)
            return row.get(col, m.group(0))  # 未マッチはそのまま

        return re.sub(r"\{\{(.+?)\}\}", replacer, result)

    def _render_auto(self, row: dict[str, str]) -> str:
        """
        excel_md_tool (MarkdownEditor.tsx) と同じ形式で Markdown を生成する。

        仕様:
          - description_cols が指定されていればその列のみ、省略時は全列を出力
          - 列名を # 見出し（複数行ヘッダーは " / " で階層化: 1段目=#, 2段目=##）
          - セル値を本文として見出しの直後に出力
          - セル内の改行（\\n / \\r\\n）は <br> に変換
          - 空セルは「（値なし）」を出力
        """
        # 出力する列を決定（description_cols 指定 > headers の順序 > row のキー順）
        specified = self.cfg.get("description_cols")
        if specified:
            cols = specified
        elif self.headers:
            cols = self.headers
        else:
            cols = list(row.keys())

        parts = []
        for header in cols:
            if header not in row:
                continue

            # 複数行ヘッダーを " / " で分割して階層見出しを生成
            # 例: "大分類 / 小分類" → "# 大分類\n## 小分類\n"
            levels = [lv.strip() for lv in header.split(" / ")]
            heading_lines = [
                f"{'#' * (i + 1)} {lv}"
                for i, lv in enumerate(levels)
                if lv
            ]
            heading = "\n".join(heading_lines)

            value = row.get(header, "")
            if value:
                body = value.replace("\r\n", "<br>").replace("\n", "<br>").replace("\r", "<br>")
            else:
                body = "（値なし）"

            parts.append(f"{heading}\n{body}")

        return "\n\n".join(parts)

    # ------------------------------------------------------------------
    # 件名の正規化
    # ------------------------------------------------------------------

    @staticmethod
    def normalize_summary(text: str) -> str:
        """
        件名文字列から特殊文字を除去・正規化して返す。

        処理内容:
          - 改行（\\r\\n / \\n / \\r）を除去
          - タブ（\\t）を除去
          - 連続スペースを1つに圧縮
          - 先頭・末尾のスペースを除去

        match_summary: true の比較にも同じメソッドを使うことで
        検索キーと Backlog 保存済み件名の表記を統一する。
        """
        normalized = text.replace("\r\n", "").replace("\r", "").replace("\n", "")
        normalized = normalized.replace("\t", "")
        normalized = re.sub(r" {2,}", " ", normalized)
        return normalized.strip()

    # ------------------------------------------------------------------
    # 各フィールドの解決
    # ------------------------------------------------------------------

    def _resolve_issue_type_id(self) -> int:
        name = self.cfg.get("issue_type", "")
        if not name:
            raise ValueError("issue_mapping.issue_type が設定されていません。")
        iid = self.master.issue_type_map.get(name)
        if iid is None:
            available = list(self.master.issue_type_map.keys())
            raise ValueError(
                f"種別「{name}」が見つかりません。利用可能: {available}"
            )
        return iid

    def _resolve_priority_id(self) -> int:
        name = self.cfg.get("priority", "")
        if not name:
            raise ValueError("issue_mapping.priority が設定されていません。")
        pid = self.master.priority_map.get(name)
        if pid is None:
            available = list(self.master.priority_map.keys())
            raise ValueError(
                f"優先度「{name}」が見つかりません。利用可能: {available}"
            )
        return pid

    def _resolve_assignee_id(self, row: dict[str, str]) -> int | None:
        col = self.cfg.get("assignee_col")
        default = self.cfg.get("default_assignee", "").strip()

        # Excel 列からユーザー名を取得。列が未設定または空の場合は default_assignee にフォールバック
        if col:
            name = row.get(col, "").strip()
        else:
            name = ""

        if not name:
            # default_assignee が設定されていなければ担当者なし
            if not default:
                return None
            name = default

        uid = self.master.user_map.get(name)
        if uid is None:
            # 重複のない表示名リストを作成（name と userId で同じ id が入るため）
            seen_ids: set[int] = set()
            unique_names: list[str] = []
            for k, v in self.master.user_map.items():
                if v not in seen_ids:
                    seen_ids.add(v)
                    unique_names.append(k)
            print(
                f"  ⚠ 担当者「{name}」がプロジェクトメンバーに見つかりません（スキップ）\n"
                f"    利用可能（表示名 or ログインID）: {unique_names}",
                file=sys.stderr,
            )
        return uid

    def _resolve_status_id(self, row: dict[str, str]) -> int | None:
        """
        status_col と status_map の設定に従い、Backlog ステータス ID を解決する。

        Config キー:
            status_col : str   Excel のステータス列名
            status_map : dict  Excel ステータス値 → Backlog ステータス名 のマッピング
                               例: {"未着手": "未対応", "対応中": "処理中"}
        """
        status_col = self.cfg.get("status_col")
        if not status_col:
            return None
        excel_status = row.get(status_col, "").strip()
        if not excel_status:
            return None
        status_map_cfg = self.cfg.get("status_map") or {}
        backlog_status_name = status_map_cfg.get(excel_status)
        if backlog_status_name is None:
            print(
                f"  ⚠ ステータス「{excel_status}」は status_map に定義されていません（スキップ）",
                file=sys.stderr,
            )
            return None
        sid = self.master.status_map.get(backlog_status_name)
        if sid is None:
            available = list(self.master.status_map.keys())
            print(
                f"  ⚠ Backlog ステータス「{backlog_status_name}」が見つかりません（スキップ）\n"
                f"    利用可能: {available}",
                file=sys.stderr,
            )
            return None
        return sid

    @staticmethod
    def _normalize_date(value: str) -> str | None:
        """
        "YYYY/MM/DD" → "YYYY-MM-DD" に変換。
        既に "YYYY-MM-DD" 形式ならそのまま返す。
        空文字列や変換不可は None を返す。
        """
        if not value:
            return None
        # YYYY/MM/DD → YYYY-MM-DD
        normalized = value.replace("/", "-")
        # 簡易バリデーション: YYYY-MM-DD
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", normalized):
            return normalized
        return None

    def _resolve_custom_fields(self, row: dict[str, str]) -> dict:
        """
        custom_fields 設定を解決して {customField_{id}: value} の dict を返す。
        """
        params = {}
        for cf_cfg in self.cfg.get("custom_fields") or []:
            field_name = cf_cfg.get("field_name", "")
            col_name = cf_cfg.get("col_name", "")

            if field_name not in self.master.custom_field_map:
                print(
                    f"  ⚠ カスタム属性「{field_name}」が見つかりません（スキップ）",
                    file=sys.stderr,
                )
                continue

            cf_info = self.master.custom_field_map[field_name]
            field_id = cf_info["id"]
            type_id = cf_info.get("typeId")
            items_map = cf_info.get("items", {})

            value = row.get(col_name, "").strip()
            if not value:
                continue

            # value_map が定義されている場合は Excel 値を Backlog 値に変換する
            # マッチング順序:
            #   1. 完全一致（dict.get）→ 後方互換・高速
            #   2. 定義順に re.fullmatch でパターンマッチ → 最初にマッチしたキーを採用
            value_map = cf_cfg.get("value_map") or {}
            if value_map:
                mapped = value_map.get(value)
                if mapped is None:
                    for pattern, target in value_map.items():
                        try:
                            if re.fullmatch(str(pattern), value):
                                mapped = target
                                break
                        except re.error:
                            pass  # 不正な正規表現はスキップ
                if mapped is None:
                    print(
                        f"  ⚠ カスタム属性「{field_name}」の値「{value}」は value_map に定義されていません（スキップ）",
                        file=sys.stderr,
                    )
                    continue
                value = str(mapped).strip()

            # 選択肢型（typeId 5=単一リスト, 6=複数, 7=チェックボックス, 8=ラジオ）
            # → 選択肢名を ID に変換
            list_types = {5, 6, 7, 8}
            if type_id in list_types and items_map:
                resolved = items_map.get(value)
                if resolved is None:
                    print(
                        f"  ⚠ カスタム属性「{field_name}」の選択肢「{value}」が見つかりません（スキップ）",
                        file=sys.stderr,
                    )
                    continue
                params[f"customField_{field_id}"] = [resolved]
            else:
                params[f"customField_{field_id}"] = value

        return params

    # ------------------------------------------------------------------
    # メイン変換処理
    # ------------------------------------------------------------------

    def map_row(self, row: dict[str, str]) -> dict:
        """
        Excel の1行データを Backlog API の課題パラメータに変換して返す。

        Returns
        -------
        dict
            Backlog API の create_issue / update_issue に渡せるパラメータ dict
        """
        params: dict = {}

        # required_cols チェック: 指定列のいずれかが空ならスキップ
        required_cols = self.cfg.get("required_cols") or []
        empty_cols = [col for col in required_cols if not row.get(col, "").strip()]
        if empty_cols:
            raise ValueError(
                f"必須列が空のためスキップします。空の列: {empty_cols}"
            )

        # 必須: projectId
        params["projectId"] = self.master.project_id

        # 必須: summary（件名）
        # summary_template が指定されていればテンプレート展開、なければ summary_col の値を使用
        # いずれの場合も normalize_summary() で改行・タブなどの特殊文字を除去する
        summary_template = self.cfg.get("summary_template", "")
        if summary_template:
            summary = self.normalize_summary(self._render_template(summary_template, row))
            if not summary:
                raise ValueError(
                    f"summary_template の展開結果が空です。この行はスキップします。"
                )
        else:
            summary_col = self.cfg.get("summary_col", "")
            summary = self.normalize_summary(row.get(summary_col, ""))
            if not summary:
                raise ValueError(
                    f"件名列「{summary_col}」の値が空です。この行はスキップします。"
                )
        params["summary"] = summary

        # 必須: issueTypeId（種別）
        params["issueTypeId"] = self._resolve_issue_type_id()

        # 必須: priorityId（優先度）
        params["priorityId"] = self._resolve_priority_id()

        # 任意: description（詳細）
        # description_format: "auto"  → excel_md_tool と同じ形式で自動生成
        # description_format: "template"（デフォルト）→ description_template を使用
        desc_format = self.cfg.get("description_format", "template")
        if desc_format == "auto":
            params["description"] = self._render_auto(row)
        else:
            template = self.cfg.get("description_template", "")
            if template:
                params["description"] = self._render_template(template, row)

        # 任意: dueDate（期限日）
        due_col = self.cfg.get("due_date_col")
        if due_col:
            # {{列名}} を含む場合はテンプレート展開して日付文字列を得る
            # 含まない場合は列名として row から値を取得する（後方互換）
            due_value = (
                self._render_template(due_col, row)
                if "{{" in due_col
                else row.get(due_col, "")
            )
            due = self._normalize_date(due_value)
            if due:
                params["dueDate"] = due

        # 任意: startDate（開始日）
        start_col = self.cfg.get("start_date_col")
        if start_col:
            start_value = (
                self._render_template(start_col, row)
                if "{{" in start_col
                else row.get(start_col, "")
            )
            start = self._normalize_date(start_value)
            if start:
                params["startDate"] = start

        # 任意: assigneeId（担当者）
        assignee_id = self._resolve_assignee_id(row)
        if assignee_id is not None:
            params["assigneeId"] = assignee_id

        # 任意: カスタム属性
        params.update(self._resolve_custom_fields(row))

        # 任意: statusId（ステータス）
        status_id = self._resolve_status_id(row)
        if status_id is not None:
            params["statusId"] = status_id

        return params

    def format_dry_run(self, row: dict[str, str], index: int) -> str:
        """
        ドライラン用: 変換結果を人間が読みやすい形式で返す。
        変換に失敗した場合はエラーメッセージを返す。
        """
        try:
            params = self.map_row(row)
        except ValueError as e:
            return f"  [{index}] ⚠ スキップ: {e}"

        lines = [f"  [{index}] 件名: {params.get('summary', '（なし）')}"]
        if "description" in params:
            # 最初の3行だけ表示
            desc_lines = params["description"].splitlines()[:3]
            for dl in desc_lines:
                lines.append(f"         {dl}")
            if len(params["description"].splitlines()) > 3:
                lines.append("         ...")
        if "dueDate" in params:
            lines.append(f"         期限日: {params['dueDate']}")
        if "assigneeId" in params:
            lines.append(f"         担当者ID: {params['assigneeId']}")
        # カスタム属性
        for k, v in params.items():
            if k.startswith("customField_"):
                lines.append(f"         {k}: {v}")
        return "\n".join(lines)

    def format_preview(self, row: dict[str, str], index: int, master_labels: dict = None) -> str:
        """
        プレビューファイル用: 課題の全内容を Markdown ブロックとして返す。

        master_labels : build_master_labels() が返すネスト構造
                        {
                          "issue_type": {id: 種別名},
                          "priority":   {id: 優先度名},
                          "user":       {id: ユーザー名},
                        }
                        省略時は ID をそのまま表示。
        """
        try:
            params = self.map_row(row)
        except ValueError as e:
            return f"## 課題 {index}\n\n> ⚠ スキップ: {e}\n"

        labels = master_labels or {}
        issue_type_labels = labels.get("issue_type", {})
        priority_labels   = labels.get("priority", {})
        user_labels       = labels.get("user", {})
        status_labels     = labels.get("status", {})

        lines = [f"## 課題 {index}"]
        lines.append("")

        # 基本フィールド
        lines.append(f"**件名:** {params.get('summary', '（なし）')}  ")
        issue_type_label = issue_type_labels.get(params.get("issueTypeId"), str(params.get("issueTypeId", "")))
        lines.append(f"**種別:** {issue_type_label}  ")
        priority_label = priority_labels.get(params.get("priorityId"), str(params.get("priorityId", "")))
        lines.append(f"**優先度:** {priority_label}  ")
        if "dueDate" in params:
            lines.append(f"**期限日:** {params['dueDate']}  ")
        if "startDate" in params:
            lines.append(f"**開始日:** {params['startDate']}  ")
        if "assigneeId" in params:
            assignee_label = user_labels.get(params["assigneeId"], str(params["assigneeId"]))
            lines.append(f"**担当者:** {assignee_label}  ")
        if "statusId" in params:
            status_label = status_labels.get(params["statusId"], str(params["statusId"]))
            lines.append(f"**ステータス:** {status_label}  ")
        for k, v in params.items():
            if k.startswith("customField_"):
                lines.append(f"**{k}:** {v}  ")

        # 本文（description）を全文表示
        lines.append("")
        lines.append("### 本文")
        lines.append("")
        if "description" in params and params["description"]:
            lines.append(params["description"])
        else:
            lines.append("_（本文なし）_")

        return "\n".join(lines)
