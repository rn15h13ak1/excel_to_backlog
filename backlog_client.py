"""
Backlog API クライアント
========================
backlog_report/backlog_weekly_report.py の BacklogClient をベースに、
課題の作成（POST）・更新（PATCH）を追加した拡張版。
"""

import json
import ssl
import sys
import time
import urllib.error
import urllib.parse
import urllib.request


class BacklogNoChangeError(Exception):
    """
    更新内容が現在の課題と同一のため変更なしと判断されたエラー。
    sys.exit(1) ではなくスキップ扱いにしたい呼び出し元で使用する。
    """


class BacklogClient:
    def __init__(
        self,
        space_host: str,
        api_key: str,
        ssl_verify: bool = True,
        base_path: str = "",
        debug: bool = False,
    ):
        base_path = "/" + base_path.strip("/") if base_path.strip("/") else ""
        self.base_url = f"https://{space_host}{base_path}/api/v2"
        self.api_key = api_key
        self.debug = debug

        if ssl_verify:
            self.ssl_context = None
        else:
            self.ssl_context = ssl.create_default_context()
            self.ssl_context.check_hostname = False
            self.ssl_context.verify_mode = ssl.CERT_NONE

    # ------------------------------------------------------------------
    # 内部ユーティリティ
    # ------------------------------------------------------------------

    def _build_query(self, params: dict) -> str:
        """パラメータ dict をクエリ文字列に変換（リスト値は [] 展開）"""
        parts = []
        for key, value in params.items():
            if isinstance(value, list):
                for v in value:
                    parts.append(
                        f"{urllib.parse.quote(str(key))}%5B%5D={urllib.parse.quote(str(v))}"
                    )
            else:
                parts.append(
                    f"{urllib.parse.quote(str(key))}={urllib.parse.quote(str(value))}"
                )
        return "&".join(parts)

    def _handle_http_error(
        self,
        e: urllib.error.HTTPError,
        endpoint: str,
        *,
        raise_no_change: bool = False,
    ) -> None:
        """
        HTTPError を整形して標準エラーに出力し sys.exit(1)。

        raise_no_change=True のとき、HTTP 400 かつ Backlog エラーコード 7
        （InvalidRequestError：変更内容なし等）であれば sys.exit の代わりに
        BacklogNoChangeError を raise する。
        """
        detail = ""
        raw_body = ""
        errors: list = []
        try:
            raw_body = e.read().decode("utf-8")
            body = json.loads(raw_body)
            errors = body.get("errors", [])
            if errors:
                detail = " / ".join(
                    f"{err.get('message', '')}（code={err.get('code')}）"
                    for err in errors
                )
        except Exception:
            pass

        # 変更なしエラーの検出: 更新時に HTTP 400 + error code 7 が返る
        if raise_no_change and e.code == 400 and any(
            err.get("code") == 7 for err in errors
        ):
            raise BacklogNoChangeError(detail or "変更内容が同一のためスキップ")

        print(
            f"エラー: API呼び出しに失敗しました（HTTP {e.code}）: {endpoint}",
            file=sys.stderr,
        )
        if detail:
            print(f"  詳細: {detail}", file=sys.stderr)
        elif raw_body:
            print(f"  レスポンス: {raw_body[:500]}", file=sys.stderr)

        hints = {
            400: "リクエストパラメータを確認してください。",
            401: "api_key を確認してください。",
            403: "api_key の権限を確認してください。",
            404: "space_host または project_key を確認してください。",
        }
        if e.code in hints:
            print(f"  → {hints[e.code]}", file=sys.stderr)
        sys.exit(1)

    def _get(self, endpoint: str, params: dict = None) -> dict | list:
        """GET リクエストを送信して JSON を返す"""
        params = dict(params or {})
        params["apiKey"] = self.api_key
        query = self._build_query(params)
        url = f"{self.base_url}{endpoint}?{query}"

        if self.debug:
            debug_parts = [p for p in query.split("&") if not p.startswith("apiKey=")]
            print(f"  [DEBUG GET] {endpoint} ?" + "&".join(debug_parts), file=sys.stderr)

        req = urllib.request.Request(url)
        try:
            with urllib.request.urlopen(req, timeout=30, context=self.ssl_context) as res:
                return json.loads(res.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            self._handle_http_error(e, endpoint)

    def _post(self, endpoint: str, params: dict) -> dict:
        """POST リクエストを送信して JSON を返す"""
        url = f"{self.base_url}{endpoint}?apiKey={urllib.parse.quote(self.api_key)}"

        # リスト値を展開（例: categoryId[] → categoryId%5B%5D=1&...）
        body_parts = []
        for key, value in params.items():
            if isinstance(value, list):
                for v in value:
                    body_parts.append(
                        (f"{key}[]", str(v))
                    )
            else:
                body_parts.append((key, str(value)))

        body = urllib.parse.urlencode(body_parts).encode("utf-8")

        if self.debug:
            print(f"  [DEBUG POST] {endpoint}", file=sys.stderr)
            for k, v in body_parts:
                print(f"    {k}={v}", file=sys.stderr)

        req = urllib.request.Request(
            url,
            data=body,
            method="POST",
            headers={"Content-Type": "application/x-www-form-urlencoded"},
        )
        try:
            with urllib.request.urlopen(req, timeout=30, context=self.ssl_context) as res:
                return json.loads(res.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            self._handle_http_error(e, endpoint)

    def _patch(self, endpoint: str, params: dict, *, raise_no_change: bool = False) -> dict:
        """PATCH リクエストを送信して JSON を返す"""
        url = f"{self.base_url}{endpoint}?apiKey={urllib.parse.quote(self.api_key)}"

        body_parts = []
        for key, value in params.items():
            if isinstance(value, list):
                for v in value:
                    body_parts.append((f"{key}[]", str(v)))
            else:
                body_parts.append((key, str(value)))

        body = urllib.parse.urlencode(body_parts).encode("utf-8")

        if self.debug:
            print(f"  [DEBUG PATCH] {endpoint}", file=sys.stderr)
            for k, v in body_parts:
                print(f"    {k}={v}", file=sys.stderr)

        req = urllib.request.Request(
            url,
            data=body,
            method="PATCH",
            headers={"Content-Type": "application/x-www-form-urlencoded"},
        )
        try:
            with urllib.request.urlopen(req, timeout=30, context=self.ssl_context) as res:
                return json.loads(res.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            self._handle_http_error(e, endpoint, raise_no_change=raise_no_change)

    # ------------------------------------------------------------------
    # マスターデータ取得
    # ------------------------------------------------------------------

    def get_project(self, project_key: str) -> dict:
        """プロジェクト情報を取得"""
        return self._get(f"/projects/{project_key}")

    def get_issue_types(self, project_id_or_key) -> list:
        """種別一覧を取得"""
        return self._get(f"/projects/{project_id_or_key}/issueTypes")

    def get_custom_fields(self, project_id_or_key) -> list:
        """カスタム属性一覧を取得"""
        return self._get(f"/projects/{project_id_or_key}/customFields")

    def get_statuses(self, project_id_or_key) -> list:
        """ステータス一覧を取得"""
        return self._get(f"/projects/{project_id_or_key}/statuses")

    def get_priorities(self) -> list:
        """優先度一覧を取得"""
        return self._get("/priorities")

    def get_project_users(self, project_id_or_key) -> list:
        """プロジェクトメンバー一覧を取得"""
        return self._get(f"/projects/{project_id_or_key}/users")

    # ------------------------------------------------------------------
    # 課題の取得
    # ------------------------------------------------------------------

    def get_issues(self, project_id: int, params: dict = None) -> list:
        """課題一覧を全件取得（ページネーション対応）"""
        all_issues = []
        offset = 0
        count = 100
        base_params = dict(params or {})
        base_params["projectId"] = [project_id]
        base_params["count"] = count

        while True:
            base_params["offset"] = offset
            issues = self._get("/issues", base_params.copy())
            if not issues:
                break
            all_issues.extend(issues)
            if len(issues) < count:
                break
            offset += count
            time.sleep(0.3)

        return all_issues

    def get_issue(self, issue_id_or_key: str) -> dict | None:
        """
        issueKey（例: PROJ-123）または数値IDで課題を1件取得。
        存在しない場合（404）は None を返す。
        """
        url = (
            f"{self.base_url}/issues/{urllib.parse.quote(str(issue_id_or_key))}"
            f"?apiKey={urllib.parse.quote(self.api_key)}"
        )
        req = urllib.request.Request(url)
        try:
            with urllib.request.urlopen(req, timeout=30, context=self.ssl_context) as res:
                return json.loads(res.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code == 404:
                return None
            self._handle_http_error(e, f"/issues/{issue_id_or_key}")

    def search_issues_by_summary(self, project_id: int, summary: str) -> list:
        """
        件名の前方一致で課題を検索して返す。
        Backlog API に全文検索はないため keyword パラメータを利用する。
        """
        return self._get("/issues", {
            "projectId": [project_id],
            "keyword": summary,
            "count": 100,
        })

    # ------------------------------------------------------------------
    # 課題の作成・更新
    # ------------------------------------------------------------------

    def create_issue(self, params: dict) -> dict:
        """
        課題を新規作成する。

        必須パラメータ（呼び出し側で設定）:
            projectId    (int)
            summary      (str)
            issueTypeId  (int)
            priorityId   (int)

        任意パラメータ（例）:
            description  (str)
            startDate    (str)  "YYYY-MM-DD"
            dueDate      (str)  "YYYY-MM-DD"
            assigneeId   (int)
            categoryId   (list[int])
            milestoneId  (list[int])
            customField_{id}  (str | int | list)
        """
        return self._post("/issues", params)

    def update_issue(self, issue_id_or_key: str, params: dict) -> dict:
        """
        既存課題を更新する。
        params は create_issue と同じキー（すべて任意）。

        変更内容が同一で Backlog API がエラー（HTTP 400 / code 7）を返した場合は
        BacklogNoChangeError を raise する（sys.exit しない）。
        """
        return self._patch(
            f"/issues/{urllib.parse.quote(str(issue_id_or_key))}",
            params,
            raise_no_change=True,
        )
