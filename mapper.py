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
            master.user_map = {u["name"]: u["id"] for u in users}
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

        return master


# ------------------------------------------------------------------
# IssueMapper: Excel 行 → Backlog API パラメータ
# ------------------------------------------------------------------

class IssueMapper:
    """
    mapping_config (sources[i].issue_mapping) に従い、
    Excel の行データを Backlog API パラメータに変換する。

    mapping_config キー:
        issue_type          : str   種別名（固定値）
        priority            : str   優先度名（固定値、デフォルト: "中"）
        summary_col         : str   件名として使う列名（summary_template と排他）
        summary_template    : str   件名テンプレート（{{列名}} でセル値を埋め込み、summary_col より優先）
        description_template: str   詳細欄テンプレート（{{列名}} でセル値を埋め込み）
        due_date_col        : str   期限日列名、または {{列名}} テンプレート（任意）
        start_date_col      : str   開始日列名、または {{列名}} テンプレート（任意）
        assignee_col        : str   担当者列名（任意）
        custom_fields       : list  カスタム属性マッピングリスト
            - field_name    : str   Backlog カスタム属性名
              col_name      : str   Excel 列名
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
          {{auto}} : _render_auto() の出力に展開される。
                     description_format が "template" でも auto 方式の出力を
                     任意の位置に埋め込めるため、ヘッダー・フッターの付与に使える。
        """
        def replacer(m: re.Match) -> str:
            col = m.group(1).strip()
            if col == "auto":
                return self._render_auto(row)
            return row.get(col, m.group(0))  # 未マッチはそのまま

        return re.sub(r"\{\{(.+?)\}\}", replacer, template)

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
        if not col:
            return None
        name = row.get(col, "").strip()
        if not name:
            return None
        uid = self.master.user_map.get(name)
        if uid is None:
            print(
                f"  ⚠ 担当者「{name}」がプロジェクトメンバーに見つかりません（スキップ）",
                file=sys.stderr,
            )
        return uid

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

        # 必須: projectId
        params["projectId"] = self.master.project_id

        # 必須: summary（件名）
        # summary_template が指定されていればテンプレート展開、なければ summary_col の値を使用
        summary_template = self.cfg.get("summary_template", "")
        if summary_template:
            summary = self._render_template(summary_template, row).strip()
            if not summary:
                raise ValueError(
                    f"summary_template の展開結果が空です。この行はスキップします。"
                )
        else:
            summary_col = self.cfg.get("summary_col", "")
            summary = row.get(summary_col, "").strip()
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
