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
from datetime import date

from backlog_client import BacklogAPIError


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
        except BacklogAPIError:
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
        except BacklogAPIError:
            print("  ⚠ カスタム属性の取得に失敗しました", file=sys.stderr)

        # ステータス
        print("  ステータス一覧を取得中...")
        try:
            statuses = client.get_statuses(project_key)
            master.status_map = {s["name"]: s["id"] for s in statuses}
        except BacklogAPIError:
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
        priority            : str        優先度名（固定値・必須。例: "高" / "中" / "低"）
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
              value_separator: str       セル値の区切り文字（任意）。指定すると分割した各値を
                                         個別に value_map で変換・items_map で解決し、
                                         複数IDのリストとして渡す（複数選択型 typeId 6・7 向け）。
                                         省略時は分割せず1値として処理。
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

    # テンプレート内のプレースホルダーを抽出する正規表現。
    # _render_template() の Step 2 と同じパターンを使う。
    # テンプレート構文を変更する場合は両方を更新すること。
    TEMPLATE_PLACEHOLDER_RE = re.compile(r"\{\{(.+?)\}\}")

    # 列名ではない特殊プレースホルダー（列名検証の対象外）
    SPECIAL_PLACEHOLDERS = {"auto"}

    def __init__(self, mapping_config: dict, master: BacklogMaster, headers: list[str] = None):
        self.cfg = mapping_config
        self.master = master
        self.headers = headers or []  # auto モードでの列順序に使用
        # map_row() 1 回分の警告。担当者・日付・ステータス・カスタム属性が
        # 解決できず未設定のまま登録される場合に積む。
        # 以前は stderr へ直接出していたため行番号も件名も付かず、
        # 「作成されたが一部フィールドが落ちた」ことを実行ログから追えなかった。
        self.warnings: list[str] = []

    def _warn(self, message: str) -> None:
        """この行の警告として記録し、標準エラーにも出力する。"""
        self.warnings.append(" ".join(message.split()))
        print(f"  ⚠ {message}", file=sys.stderr)

    # ------------------------------------------------------------------
    # テンプレート処理
    # ------------------------------------------------------------------

    @classmethod
    def extract_template_columns(cls, template: str) -> set[str]:
        """
        テンプレート文字列が参照している列名を抽出して返す。

        対象:
            {{列名}}       通常のプレースホルダー
            {{#列名}}      条件ブロックの開始
            {{/列名}}      条件ブロックの終了

        除外:
            {{auto}} などの特殊プレースホルダー（SPECIAL_PLACEHOLDERS）

        起動時の列名検証（validate_column_references）から使用する。
        """
        cols: set[str] = set()
        for m in cls.TEMPLATE_PLACEHOLDER_RE.finditer(template or ""):
            # 条件ブロックの {{#列名}} / {{/列名}} から記号を除去して列名を得る
            name = m.group(1).strip().lstrip("#/").strip()
            if name and name not in cls.SPECIAL_PLACEHOLDERS:
                cols.add(name)
        return cols

    def _render_template(
        self,
        template: str,
        row: dict[str, str],
        formatted_row: dict[str, str] | None = None,
    ) -> str:
        """
        {{列名}} を行のセル値に置換する。
        存在しない列名はそのまま残す（警告なし）。

        特殊プレースホルダー:
          {{auto}}          : _render_auto() の出力に展開される。
                              description_format が "template" でも auto 方式の出力を
                              任意の位置に埋め込めるため、ヘッダー・フッターの付与に使える。
                              rich_text: true のとき formatted_row を渡して取り消し線を反映する。

          {{#列名}}...{{/列名}} : 条件ブロック。
                              指定列の値が空でなければブロック内を出力し、
                              空であればブロック全体を出力しない（セパレーター等の
                              「値がある場合のみ表示したい文字列」に使用）。
                              例: "項番{{項番}}{{#枝番}}-{{枝番}}{{/枝番}}"
                                → 枝番="1" → "項番1-1"
                                → 枝番=""  → "項番1"

        Parameters
        ----------
        formatted_row : dict[str, str] | None
            書式付き Markdown 行（rich_text: true 時に渡す）。
            {{auto}} の展開にのみ使用される。{{列名}} はプレーンテキスト（row）のまま。
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
                # {{auto}} のみ formatted_row を渡して取り消し線を反映する
                return self._render_auto(row, formatted_row=formatted_row)
            return row.get(col, m.group(0))  # 未マッチはそのまま

        return re.sub(r"\{\{(.+?)\}\}", replacer, result)

    def _render_auto(self, row: dict[str, str], formatted_row: dict[str, str] | None = None) -> str:
        """
        excel_md_tool (MarkdownEditor.tsx) と同じ形式で Markdown を生成する。

        仕様:
          - description_cols が指定されていればその列のみ、省略時は全列を出力
          - 列名を # 見出し（複数行ヘッダーは " / " で階層化: 1段目=#, 2段目=##）
          - セル値を本文として見出しの直後に出力
          - セル内の改行（\\n / \\r\\n）は <br> に変換
          - 空セルは「（値なし）」を出力

        Parameters
        ----------
        row : dict[str, str]
            プレーンテキスト行（列の存在確認・空値判定に使用）
        formatted_row : dict[str, str] | None
            書式付き Markdown 行（rich_text: true 時に渡す）。
            指定された場合、本文の値はこちらから取得する。
            None の場合は row をそのまま使用する。
        """
        # 出力対象を (列名, 平文, 表示用) の並びで組み立てる。
        #
        # 行データは {ヘッダー名: 値} の dict のため、同名の列は 1 つしか
        # 保持できない。本文には両方の列を出したいので、ExcelReader が
        # 持たせた列順の値リスト（_excel_cells）を優先して使う。
        cells = self._ordered_cells(row, formatted_row)

        specified = self.cfg.get("description_cols")
        if specified:
            wanted = set(specified)
            cells = [c for c in cells if c[0] in wanted]

        parts = []
        for header, plain_value, display_value in cells:
            # 複数行ヘッダーを " / " で分割して階層見出しを生成
            # 例: "大分類 / 小分類" → "# 大分類\n## 小分類\n"
            levels = [lv.strip() for lv in header.split(" / ")]
            heading_lines = [
                f"{'#' * (i + 1)} {lv}"
                for i, lv in enumerate(levels)
                if lv
            ]
            heading = "\n".join(heading_lines)

            if plain_value:
                body = self._to_markdown_body(display_value)
            else:
                body = "（値なし）"

            parts.append(f"{heading}\n{body}")

        return "\n\n".join(parts)

    @staticmethod
    def _to_markdown_body(text: str) -> str:
        """
        セル値を本文用の Markdown に変換する。

        セル内の改行は <br> にする（Markdown では改行 1 つが行継続として
        扱われ、意図した位置で改行されないため）。
        ただし空行は段落の区切りとして残す。継続行を結合したときの
        区切りがここに当たる。
        """
        normalized = text.replace("\r\n", "\n").replace("\r", "\n")
        paragraphs = normalized.split("\n\n")
        return "\n\n".join(p.replace("\n", "<br>") for p in paragraphs)

    def _ordered_cells(
        self, row: dict[str, str], formatted_row: dict[str, str] | None
    ) -> list[tuple[str, str, str]]:
        """
        本文に出力するセルを (列名, 平文, 表示用) の列順リストで返す。

        ExcelReader が付与した列順の値リストがあればそれを使う。
        同名の列がある場合、dict からは 1 つしか取れないため、
        この経路でのみ両方の列を出力できる。

        値リストが無い場合（テストで手組みした行など）は dict から組み立てる。
        """
        from excel_reader import ExcelReader

        values = row.get(ExcelReader.CELL_VALUES_KEY)
        if values is not None and len(values) == len(self.headers):
            display = (formatted_row or {}).get(ExcelReader.CELL_VALUES_KEY) or values
            if len(display) != len(self.headers):
                display = values
            return list(zip(self.headers, values, display))

        # フォールバック: dict のキー順（同名の列は 1 つだけになる）
        cols = self.headers or [k for k in row if not k.startswith("_")]
        return [
            (h, row.get(h, ""), (formatted_row or row).get(h, ""))
            for h in dict.fromkeys(cols)
            if h in row
        ]

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

    def resolve_fixed_fields(self) -> None:
        """
        全行に共通の設定（種別・優先度）を解決できるか確認する。

        これらは行のデータではなく設定そのものに由来するため、行ごとに
        判定すると設定のタイプミスが「全行スキップ」というデータ不備の
        ような報告になる。処理を始める前に一度だけ検証する。

        Raises
        ------
        ValueError : 種別名・優先度名が Backlog に存在しない場合
        """
        self._resolve_issue_type_id()
        self._resolve_priority_id()

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
            self._warn(
                f"担当者「{name}」がプロジェクトメンバーに見つかりません（未設定）\n"
                f"    利用可能（表示名 or ログインID）: {unique_names}"
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
            self._warn(
                f"ステータス「{excel_status}」は status_map に定義されていません（未設定）"
            )
            return None
        sid = self.master.status_map.get(backlog_status_name)
        if sid is None:
            available = list(self.master.status_map.keys())
            self._warn(
                f"Backlog ステータス「{backlog_status_name}」が見つかりません（未設定）\n"
                f"    利用可能: {available}"
            )
            return None
        return sid

    # 年月日の区切りを許容してマッチする。ゼロ埋めの有無を問わない。
    # 例: 2025-01-05 / 2025/1/5 / 2025年1月5日
    # 末尾に時刻などが続く場合も先頭の日付部分だけを取り出す。
    _DATE_RE = re.compile(r"^\s*(\d{4})\s*[-/年.]\s*(\d{1,2})\s*[-/月.]\s*(\d{1,2})\s*[日]?")

    @classmethod
    def _normalize_date(cls, value: str) -> str | None:
        """
        日付文字列を Backlog API の "YYYY-MM-DD" 形式に変換する。

        Excel の日付型セルは ExcelReader が "YYYY/MM/DD" に整形済みだが、
        文字列として手入力された日付はゼロ埋めされていないことが多いため、
        1〜2 桁の月日や和文の区切りも受け付ける。

        受理例:
            "2025-01-05" / "2025/1/5" / "2025年1月5日" / "2025/01/05 10:00"
        None を返す例:
            "" / "R7/1/5"（和暦）/ "9/1"（年がない）/ "2025/13/1"（不正な日付）

        変換できなかった場合は None を返す。呼び出し元は値が非空なのに
        None が返ったときに警告を出すこと（無言で捨てないため）。
        """
        if not value or not value.strip():
            return None

        m = cls._DATE_RE.match(value)
        if not m:
            return None

        year, month, day = (int(g) for g in m.groups())
        try:
            # 実在する日付かを検証する（2025/13/1・2025/2/30 などを弾く）
            return date(year, month, day).strftime("%Y-%m-%d")
        except ValueError:
            return None

    def _resolve_date(self, row: dict[str, str], cfg_key: str, label: str) -> str | None:
        """
        due_date_col / start_date_col の設定値から日付文字列を解決する。

        設定値が "{{" を含む場合はテンプレートとして展開し、
        含まない場合は列名として row から値を取得する（後方互換）。

        セルに値があるのに日付として解釈できなかった場合は警告を出力する。
        担当者・ステータス・カスタム属性と同様、無言では捨てない。
        """
        col = self.cfg.get(cfg_key)
        if not col:
            return None

        raw = (
            self._render_template(col, row)
            if "{{" in str(col)
            else row.get(col, "")
        )
        normalized = self._normalize_date(raw)
        if normalized is None and raw and raw.strip():
            self._warn(
                f"{label}「{raw.strip()}」を日付として解釈できません（未設定）\n"
                f"    受理する形式: YYYY-MM-DD / YYYY/M/D / YYYY年M月D日"
            )
        return normalized

    def _resolve_custom_fields(self, row: dict[str, str]) -> dict:
        """
        custom_fields 設定を解決して {customField_{id}: value} の dict を返す。
        """
        params = {}
        for cf_cfg in self.cfg.get("custom_fields") or []:
            field_name = cf_cfg.get("field_name", "")
            col_name = cf_cfg.get("col_name", "")

            if field_name not in self.master.custom_field_map:
                self._warn(f"カスタム属性「{field_name}」が見つかりません（未設定）")
                continue

            cf_info = self.master.custom_field_map[field_name]
            field_id = cf_info["id"]
            type_id = cf_info.get("typeId")
            items_map = cf_info.get("items", {})

            value = row.get(col_name, "").strip()
            if not value:
                continue

            # value_separator が定義されている場合はセル値を分割して複数値として扱う
            # 省略時は分割せず1値として処理（後方互換）
            separator = cf_cfg.get("value_separator")
            raw_values = (
                [v.strip() for v in value.split(separator) if v.strip()]
                if separator
                else [value]
            )

            # value_map が定義されている場合は各値を Backlog 値に変換する
            # マッチング順序:
            #   1. 完全一致（dict.get）→ 後方互換・高速
            #   2. 定義順に re.fullmatch でパターンマッチ → 最初にマッチしたキーを採用
            value_map = cf_cfg.get("value_map") or {}
            mapped_values: list[str] = []
            skip = False
            for raw in raw_values:
                if value_map:
                    mapped = value_map.get(raw)
                    if mapped is None:
                        for pattern, target in value_map.items():
                            try:
                                if re.fullmatch(str(pattern), raw, re.DOTALL):
                                    mapped = target
                                    break
                            except re.error:
                                pass  # 不正な正規表現はスキップ
                    if mapped is None:
                        self._warn(
                            f"カスタム属性「{field_name}」の値「{raw}」は value_map に定義されていません（未設定）"
                        )
                        skip = True
                        break
                    mapped_values.append(str(mapped).strip())
                else:
                    mapped_values.append(raw)
            if skip:
                continue

            # 選択肢型（typeId 5=単一リスト, 6=複数, 7=チェックボックス, 8=ラジオ）
            # → 選択肢名を ID に変換して渡す
            #
            # Backlog API のパラメータ形式:
            #   typeId 5（単一リスト）: customField_{id}=ID     ← [] なし・単一値
            #   typeId 6（複数リスト）: customField_{id}[]=ID   ← [] あり・複数可
            #   typeId 7（チェックボックス）: customField_{id}[]=ID  ← [] あり・複数可
            #   typeId 8（ラジオ）: customField_{id}=ID         ← [] なし・単一値
            #
            # Python の list として渡すと _post/_patch で [] 付きになるため、
            # 単一選択型（5・8）は int として渡す（[] なし）。
            multi_select_types = {6, 7}   # [] 付き配列形式
            single_select_types = {5, 8}  # [] なし単一値形式
            all_select_types = multi_select_types | single_select_types

            # 選択肢型なのに選択肢一覧が空の場合、そのまま else へ落ちると
            # ID ではなく選択肢名の文字列を送ってしまい、Backlog は HTTP 400 を
            # 返す。原因が分かりにくいため、ここで検出して理由を示す。
            # （マスターデータ取得が権限不足などで部分的に失敗した場合に起きる）
            if type_id in all_select_types and not items_map:
                self._warn(
                    f"カスタム属性「{field_name}」は選択肢型（typeId={type_id}）ですが"
                    f"選択肢一覧を取得できていません（未設定）\n"
                    f"    → 起動時のカスタム属性取得に失敗した可能性があります。"
                    f"api_key の権限を確認してください。"
                )
                continue

            if type_id in all_select_types:
                resolved_ids = []
                for mv in mapped_values:
                    resolved = items_map.get(mv)
                    if resolved is None:
                        self._warn(
                            f"カスタム属性「{field_name}」の選択肢「{mv}」が見つかりません（未設定）"
                        )
                        skip = True
                        break
                    resolved_ids.append(resolved)
                if skip:
                    continue

                if type_id in multi_select_types:
                    # 複数選択型: リストで渡す → _post/_patch が customField_{id}[] として展開
                    params[f"customField_{field_id}"] = resolved_ids
                else:
                    # 単一選択型（5/8）: int で渡す → _post/_patch が customField_{id} として送信
                    # value_separator で複数値が指定された場合でも先頭の1件のみ使用する
                    if len(resolved_ids) > 1:
                        self._warn(
                            f"カスタム属性「{field_name}」は単一選択型（typeId={type_id}）のため"
                            f" 先頭の値のみ使用します: {mapped_values[0]}"
                        )
                    params[f"customField_{field_id}"] = resolved_ids[0]
            else:
                # 非選択肢型（文字列・数値・日付など）は最初の値のみ使用する。
                # value_separator を指定していても分割は活きないため、
                # 単一選択型と同様に捨てた値があることを警告する。
                if len(mapped_values) > 1:
                    self._warn(
                        f"カスタム属性「{field_name}」は選択肢型ではない（typeId={type_id}）ため"
                        f" value_separator による分割は無効です。先頭の値のみ使用します: "
                        f"{mapped_values[0]}"
                    )
                params[f"customField_{field_id}"] = mapped_values[0] if mapped_values else value

        return params

    # ------------------------------------------------------------------
    # メイン変換処理
    # ------------------------------------------------------------------

    def map_row(self, row: dict[str, str], formatted_row: dict[str, str] | None = None) -> dict:
        """
        Excel の1行データを Backlog API の課題パラメータに変換して返す。

        Parameters
        ----------
        row : dict[str, str]
            プレーンテキスト行。フィルタ・件名・担当者・日付解決などに使用。
        formatted_row : dict[str, str] | None
            書式付き Markdown 行（rich_text: true 時に渡す）。
            description_format: auto の本文生成にのみ使用される。
            None の場合は row をそのまま使用する。

        Returns
        -------
        dict
            Backlog API の create_issue / update_issue に渡せるパラメータ dict
        """
        params: dict = {}
        # この行の警告を集める（呼び出しごとにリセット）
        self.warnings = []

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
            # formatted_row が渡されている場合は書式付き値を本文に使用する
            params["description"] = self._render_auto(row, formatted_row=formatted_row)
        else:
            template = self.cfg.get("description_template", "")
            if template:
                # formatted_row を渡す: {{auto}} プレースホルダーがある場合にのみ使用される
                # {{列名}} はプレーンテキスト（row）のままで変化しない
                params["description"] = self._render_template(template, row, formatted_row=formatted_row)

        # 任意: dueDate（期限日）
        due = self._resolve_date(row, "due_date_col", "期限日")
        if due:
            params["dueDate"] = due

        # 任意: startDate（開始日）
        start = self._resolve_date(row, "start_date_col", "開始日")
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

    def format_plan(
        self,
        plan,
        master_labels: dict | None = None,
        indent: str = "         ",
        description_lines: int = 3,
    ) -> str:
        """
        RowPlan の内容を人間が読める形で返す。ドライラン表示と実行前の
        1 件ずつの確認で共用する。

        master_labels を渡すと、担当者・ステータスを ID ではなく名前で表示する
        （build_master_labels() の戻り値）。確認画面では ID を見せても判断
        できないため、名前を出す。
        """
        params = plan.params
        labels = master_labels or {}
        lines = [f"{indent}件名: {params.get('summary', '（なし）')}"]

        if description_lines and "description" in params:
            desc = params["description"].splitlines()
            for dl in desc[:description_lines]:
                lines.append(f"{indent}{dl}")
            if len(desc) > description_lines:
                lines.append(f"{indent}...")

        def named(key: str, group: str) -> str:
            value = params[key]
            return str(labels.get(group, {}).get(value, value))

        if "startDate" in params:
            lines.append(f"{indent}開始日: {params['startDate']}")
        if "dueDate" in params:
            lines.append(f"{indent}期限日: {params['dueDate']}")
        if "assigneeId" in params:
            lines.append(f"{indent}担当者: {named('assigneeId', 'user')}")
        if "statusId" in params:
            lines.append(f"{indent}ステータス: {named('statusId', 'status')}")

        for k, v in params.items():
            if k.startswith("customField_"):
                lines.append(f"{indent}{k}: {v}")

        for warning in plan.warnings:
            lines.append(f"{indent}⚠ {warning}")

        return "\n".join(lines)

    def format_preview(
        self,
        row: dict[str, str],
        index: int,
        master_labels: dict = None,
        formatted_row: dict[str, str] | None = None,
    ) -> str:
        """
        プレビューファイル用: 課題の全内容を Markdown ブロックとして返す。

        master_labels : build_master_labels() が返すネスト構造
                        {
                          "issue_type": {id: 種別名},
                          "priority":   {id: 優先度名},
                          "user":       {id: ユーザー名},
                        }
                        省略時は ID をそのまま表示。
        formatted_row : 書式付き Markdown 行（rich_text: true 時に渡す）。
        """
        try:
            params = self.map_row(row, formatted_row=formatted_row)
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
