"""
BacklogClient のエラー処理のテスト
==================================
「変更なし」の誤判定は本当の更新失敗を握り潰し、スキップとして表示していた。
判定基準を固定する。
"""

import io
import json
import urllib.error

import pytest

from backlog_client import BacklogAPIError, BacklogClient, BacklogNoChangeError


@pytest.fixture
def client():
    return BacklogClient("example.backlog.com", "dummy-key")


def http_error(status, message="エラー", code=1):
    body = json.dumps({"errors": [{"message": message, "code": code}]}).encode("utf-8")
    return urllib.error.HTTPError("url", status, "reason", {}, io.BytesIO(body))


class TestNoChangeDetection:
    """
    Backlog のエラーコード 7 は InvalidRequestError の汎用コードで、
    不正なステータス遷移・カスタム属性値不正・件名長超過でも返る。
    コードだけで「変更なし」と判断すると本当の失敗を握り潰す。
    """

    @pytest.mark.parametrize("message", [
        "変更されていません",
        "変更がありません",
        "No changes were made",
    ])
    def test_変更なしを示すメッセージはスキップ扱い(self, client, message):
        with pytest.raises(BacklogNoChangeError):
            client._handle_http_error(
                http_error(400, message, code=7), "/issues/X", raise_no_change=True
            )

    @pytest.mark.parametrize("message", [
        "ステータスの変更が許可されていません",
        "件名が長すぎます",
        "カスタム属性の値が不正です",
    ])
    def test_それ以外の_code_7_はエラーとして扱う(self, client, message):
        with pytest.raises(BacklogAPIError):
            client._handle_http_error(
                http_error(400, message, code=7), "/issues/X", raise_no_change=True
            )

    def test_code_7_以外はエラー(self, client):
        with pytest.raises(BacklogAPIError):
            client._handle_http_error(
                http_error(400, "変更されていません", code=1),
                "/issues/X", raise_no_change=True,
            )

    def test_作成時は変更なし判定をしない(self, client):
        """raise_no_change=False（POST）では常にエラー。"""
        with pytest.raises(BacklogAPIError):
            client._handle_http_error(
                http_error(400, "変更されていません", code=7), "/issues"
            )


class TestFatalClassification:
    """認証・権限エラーは行ごとに再試行しても必ず失敗するため実行全体を中止する。"""

    @pytest.mark.parametrize("status", [401, 403])
    def test_認証権限エラーは_fatal(self, client, status):
        with pytest.raises(BacklogAPIError) as exc:
            client._handle_http_error(http_error(status), "/issues")
        assert exc.value.fatal is True

    @pytest.mark.parametrize("status", [400, 404, 429, 500])
    def test_その他は_fatal_ではない(self, client, status):
        with pytest.raises(BacklogAPIError) as exc:
            client._handle_http_error(http_error(status), "/issues")
        assert exc.value.fatal is False

    def test_ステータスコードが保持される(self, client):
        with pytest.raises(BacklogAPIError) as exc:
            client._handle_http_error(http_error(429), "/issues")
        assert exc.value.status == 429

    def test_429_には対処のヒントが付く(self, client):
        with pytest.raises(BacklogAPIError) as exc:
            client._handle_http_error(http_error(429), "/issues")
        assert "レート制限" in str(exc.value)


class TestBuildQuery:
    def test_リスト値は角括弧付きで展開される(self, client):
        query = client._build_query({"projectId": [1, 2]})
        assert query == "projectId%5B%5D=1&projectId%5B%5D=2"

    def test_単一値はそのまま(self, client):
        assert client._build_query({"count": 100}) == "count=100"


class TestBasePath:
    def test_オンプレ用のパスプレフィックスが付く(self):
        c = BacklogClient("example.com", "k", base_path="/backlog")
        assert c.base_url == "https://example.com/backlog/api/v2"

    def test_前後のスラッシュは正規化される(self):
        c = BacklogClient("example.com", "k", base_path="backlog/")
        assert c.base_url == "https://example.com/backlog/api/v2"

    def test_未指定なら付かない(self):
        c = BacklogClient("example.com", "k")
        assert c.base_url == "https://example.com/api/v2"
