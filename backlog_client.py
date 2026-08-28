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


class BacklogError(Exception):
    """このモジュールが送出する例外の基底クラス。"""


class BacklogAPIError(BacklogError):
    """
    Backlog API がエラーを返した、または通信に失敗した。

    以前は _handle_http_error() が sys.exit(1) を呼び、呼び出し元が
    SystemExit を捕捉して継続していたが、
      - 1件の失敗と「実行全体の中止」を区別できない
      - 課題の作成には成功したのに後続の更新で終了してしまう
    といった問題があったため、通常の例外として送出する。

    Attributes
    ----------
    status   : int | None   HTTP ステータスコード（通信失敗時は None）
    endpoint : str          呼び出したエンドポイント
    detail   : str          Backlog が返したエラーメッセージ
    errors   : list         Backlog のエラーオブジェクト配列
    fatal    : bool         True のとき実行全体を中止すべき（認証・権限エラー等）
    """

    def __init__(self, message, *, status=None, endpoint="", detail="", errors=None, fatal=False):
        super().__init__(message)
        self.status = status
        self.endpoint = endpoint
        self.detail = detail
        self.errors = errors or []
        self.fatal = fatal


class BacklogNoChangeError(BacklogError):
    """
    更新内容が現在の課題と同一のため変更なしと判断されたエラー。
    エラーではなくスキップ扱いにしたい呼び出し元で使用する。
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

    # 「変更内容がない」ことを示す Backlog のエラーメッセージ断片。
    # error code 7（InvalidRequestError）は不正なステータス遷移・不正なカスタム
    # 属性値・件名の長さ超過などにも使われる汎用コードのため、コードだけで
    # 「変更なし」と判断すると本当の更新失敗を握り潰してしまう。
    # メッセージが以下のいずれかを含むときのみ「変更なし」と判定する。
    NO_CHANGE_MESSAGE_HINTS = (
        "変更されていません",
        "変更がありません",
        "No changes",
        "not changed",
        "nothing to update",
    )

    @classmethod
    def _is_no_change(cls, errors: list, detail: str) -> bool:
        """HTTP 400 / code 7 のエラーが「変更なし」を意味するか判定する。"""
        if not any(err.get("code") == 7 for err in errors):
            return False
        lowered = detail.lower()
        return any(hint.lower() in lowered for hint in cls.NO_CHANGE_MESSAGE_HINTS)

    def _handle_http_error(
        self,
        e: urllib.error.HTTPError,
        endpoint: str,
        *,
        raise_no_change: bool = False,
    ) -> None:
        """
        HTTPError を BacklogAPIError（または BacklogNoChangeError）に変換して送出する。

        raise_no_change=True のとき、HTTP 400 かつ Backlog エラーコード 7 で、
        かつメッセージが「変更なし」を示す場合のみ BacklogNoChangeError を送出する。
        判定できない code 7 は安全側に倒して BacklogAPIError とする
        （本当の更新失敗を「変更なし」と表示して握り潰さないため）。
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

        if raise_no_change and e.code == 400 and self._is_no_change(errors, detail):
            raise BacklogNoChangeError(detail or "HTTP 400 / code 7（変更なし）")

        hints = {
            400: "リクエストパラメータを確認してください。",
            401: "api_key を確認してください。",
            403: "api_key の権限を確認してください。",
            404: "space_host または project_key を確認してください。",
            429: "API のレート制限に達しました。時間をおいて再実行してください。",
        }
        message = f"API 呼び出しに失敗しました（HTTP {e.code}）: {endpoint}"
        if detail:
            message += f"\n  詳細: {detail}"
        elif raw_body:
            message += f"\n  レスポンス: {raw_body[:500]}"
        if e.code in hints:
            message += f"\n  → {hints[e.code]}"

        raise BacklogAPIError(
            message,
            status=e.code,
            endpoint=endpoint,
            detail=detail,
            errors=errors,
            # 認証・権限の誤りは行ごとに再試行しても必ず失敗するため実行全体を中止する
            fatal=e.code in (401, 403),
        )

    # リトライ設定
    MAX_RETRIES = 3          # 初回を除く再送回数
    RETRY_BASE_WAIT = 2.0    # 指数バックオフの基準秒数（2 → 4 → 8 秒）
    MAX_RETRY_WAIT = 60.0    # Retry-After が極端に大きい場合の上限

    def _retry_wait(self, attempt: int, retry_after: str | None) -> float:
        """
        次の再送までの待機秒数を求める。

        Retry-After ヘッダーがあればそれを優先し、なければ指数バックオフ。
        """
        if retry_after:
            try:
                return min(float(retry_after), self.MAX_RETRY_WAIT)
            except ValueError:
                pass  # HTTP-date 形式は解釈せずバックオフにフォールバック
        return min(self.RETRY_BASE_WAIT * (2 ** attempt), self.MAX_RETRY_WAIT)

    def _send(
        self,
        req: urllib.request.Request,
        endpoint: str,
        *,
        raise_no_change: bool = False,
        idempotent: bool = True,
    ) -> dict | list:
        """
        リクエストを送信して JSON を返す。失敗時は条件付きで再送する。

        HTTPError は _handle_http_error() で BacklogAPIError に変換する。
        URLError（DNS 失敗・接続リセット・TLS エラー）とタイムアウトも
        BacklogAPIError に変換する。以前はこれが捕捉されておらず、
        通信が一度切れるだけでトレースバックとともに実行全体が停止し、
        それまでに作成した課題のサマリーも表示されなかった。

        Parameters
        ----------
        idempotent : bool
            同じリクエストを再送しても副作用が重複しないか。
            GET と PATCH は True。POST（課題作成）は False を渡すこと。

        再送する条件:
            HTTP 429            : リクエストは処理されていないため常に再送する
            HTTP 5xx / 通信断   : idempotent=True のときのみ再送する
                                  （POST は課題が作成済みかもしれず、再送すると
                                    重複作成になるためエラーとして返す）
        """
        last_error: BacklogAPIError | None = None

        for attempt in range(self.MAX_RETRIES + 1):
            retry_after = None
            try:
                with urllib.request.urlopen(
                    req, timeout=30, context=self.ssl_context
                ) as res:
                    return json.loads(res.read().decode("utf-8"))

            except urllib.error.HTTPError as e:
                retryable = e.code == 429 or (idempotent and 500 <= e.code < 600)
                if not retryable:
                    # BacklogAPIError / BacklogNoChangeError を送出する
                    self._handle_http_error(e, endpoint, raise_no_change=raise_no_change)
                retry_after = e.headers.get("Retry-After") if e.headers else None
                last_error = BacklogAPIError(
                    f"API 呼び出しに失敗しました（HTTP {e.code}）: {endpoint}",
                    status=e.code,
                    endpoint=endpoint,
                )

            except urllib.error.URLError as e:
                if not idempotent:
                    raise BacklogAPIError(
                        f"通信に失敗しました: {endpoint}\n  理由: {e.reason}\n"
                        f"  → 課題が作成済みの可能性があるため再送しません。"
                        f"Backlog 側を確認してください。",
                        endpoint=endpoint,
                        detail=str(e.reason),
                    ) from e
                last_error = BacklogAPIError(
                    f"通信に失敗しました: {endpoint}\n  理由: {e.reason}",
                    endpoint=endpoint,
                    detail=str(e.reason),
                )

            except TimeoutError as e:
                if not idempotent:
                    raise BacklogAPIError(
                        f"通信がタイムアウトしました（30秒）: {endpoint}\n"
                        f"  → 課題が作成済みの可能性があるため再送しません。"
                        f"Backlog 側を確認してください。",
                        endpoint=endpoint,
                        detail="timeout",
                    ) from e
                last_error = BacklogAPIError(
                    f"通信がタイムアウトしました（30秒）: {endpoint}",
                    endpoint=endpoint,
                    detail="timeout",
                )

            if attempt == self.MAX_RETRIES:
                break

            wait = self._retry_wait(attempt, retry_after)
            print(
                f"  ⏳ {last_error.status or '通信エラー'} のため {wait:.0f} 秒待って再試行します"
                f"（{attempt + 1}/{self.MAX_RETRIES}）: {endpoint}",
                file=sys.stderr,
            )
            time.sleep(wait)

        raise last_error

    def _get(self, endpoint: str, params: dict = None) -> dict | list:
        """GET リクエストを送信して JSON を返す"""
        params = dict(params or {})
        params["apiKey"] = self.api_key
        query = self._build_query(params)
        url = f"{self.base_url}{endpoint}?{query}"

        if self.debug:
            debug_parts = [p for p in query.split("&") if not p.startswith("apiKey=")]
            print(f"  [DEBUG GET] {endpoint} ?" + "&".join(debug_parts), file=sys.stderr)

        return self._send(urllib.request.Request(url), endpoint)

    def _post(self, endpoint: str, params: dict) -> dict:
        """POST リクエストを送信して JSON を返す"""
        url = f"{self.base_url}{endpoint}?apiKey={urllib.parse.quote(self.api_key)}"

        # リスト値を展開（例: categoryId[] → categoryId[]=1&...）
        # キー名はエンコードせず [] をそのまま送信する。
        # urllib.parse.urlencode はキーの [] を %5B%5D にエンコードするため、
        # Backlog API がリスト表記を認識できなくなる。値のみ quote_plus でエンコードする。
        body_parts = []
        for key, value in params.items():
            if isinstance(value, list):
                for v in value:
                    body_parts.append((f"{key}[]", str(v)))
            else:
                body_parts.append((key, str(value)))

        body = "&".join(
            f"{k}={urllib.parse.quote_plus(v)}"
            for k, v in body_parts
        ).encode("utf-8")

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
        # POST は冪等でない。通信断や 5xx で再送すると課題が二重に作られるため、
        # 確実に未処理と分かる 429 のみ再送させる。
        result = self._send(req, endpoint, idempotent=False)
        if self.debug:
            self._debug_custom_fields("POST", result)
        return result

    def _patch(self, endpoint: str, params: dict, *, raise_no_change: bool = False) -> dict:
        """PATCH リクエストを送信して JSON を返す"""
        url = f"{self.base_url}{endpoint}?apiKey={urllib.parse.quote(self.api_key)}"

        # _post() と同じ理由でキー名の [] をエンコードせず値のみ quote_plus でエンコードする
        body_parts = []
        for key, value in params.items():
            if isinstance(value, list):
                for v in value:
                    body_parts.append((f"{key}[]", str(v)))
            else:
                body_parts.append((key, str(value)))

        body = "&".join(
            f"{k}={urllib.parse.quote_plus(v)}"
            for k, v in body_parts
        ).encode("utf-8")

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
        result = self._send(req, endpoint, raise_no_change=raise_no_change)
        if self.debug:
            self._debug_custom_fields("PATCH", result)
        return result

    @staticmethod
    def _debug_custom_fields(method: str, result: dict) -> None:
        """カスタム属性の反映確認用に、レスポンスのカスタム属性を出力する。"""
        custom_fields = (result or {}).get("customFields", [])
        if not custom_fields:
            print(f"  [DEBUG {method} response] customFields: (なし または 空)", file=sys.stderr)
            return
        print(f"  [DEBUG {method} response] customFields:", file=sys.stderr)
        for cf in custom_fields:
            print(
                f"    id={cf.get('id')} name={cf.get('name')!r} value={cf.get('value')!r}",
                file=sys.stderr,
            )

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
        endpoint = f"/issues/{issue_id_or_key}"
        try:
            return self._send(urllib.request.Request(url), endpoint)
        except BacklogAPIError as e:
            # 存在しない課題キーは「見つからない」であってエラーではない
            if e.status == 404:
                return None
            raise

    def search_issues_by_summary(self, project_id: int, summary: str) -> list:
        """
        件名の前方一致で課題を検索して全件返す（ページネーション対応）。
        Backlog API に全文検索はないため keyword パラメータを利用する。
        1 リクエストあたりの上限（100 件）を超えるプロジェクトでも取得漏れが
        発生しないよう、get_issues() と同様にページネーションで全件取得する。
        """
        all_issues = []
        offset = 0
        count = 100
        while True:
            issues = self._get("/issues", {
                "projectId": [project_id],
                "keyword": summary,
                "count": count,
                "offset": offset,
            })
            if not issues:
                break
            all_issues.extend(issues)
            if len(issues) < count:
                break
            offset += count
            time.sleep(0.3)
        return all_issues

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
