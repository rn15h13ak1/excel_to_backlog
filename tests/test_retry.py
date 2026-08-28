"""
リトライとバックオフのテスト
============================
429 は必ず再送する。通信断・5xx は冪等なリクエストのみ再送する
（POST を再送すると課題が二重に作られるため）。
"""

import io
import json
import urllib.error
import urllib.request

import pytest

from backlog_client import BacklogAPIError, BacklogClient


@pytest.fixture
def client():
    """BacklogClient（リトライ待機は conftest の no_sleep で無効化済み）。"""
    return BacklogClient("example.backlog.com", "dummy-key")


def http_error(status, headers=None):
    body = json.dumps({"errors": [{"message": "エラー", "code": 1}]}).encode("utf-8")
    return urllib.error.HTTPError("url", status, "reason", headers or {}, io.BytesIO(body))


class _Responder:
    """urlopen を差し替えて、あらかじめ決めた結果を順に返す。"""

    def __init__(self, *outcomes):
        self.outcomes = list(outcomes)
        self.calls = 0

    def __call__(self, req, timeout=None, context=None):
        self.calls += 1
        outcome = self.outcomes.pop(0) if self.outcomes else self.outcomes
        if isinstance(outcome, Exception):
            raise outcome

        class _Res:
            def __enter__(self_inner):
                return self_inner

            def __exit__(self_inner, *a):
                return False

            def read(self_inner):
                return json.dumps(outcome).encode("utf-8")

        return _Res()


def install(monkeypatch, responder):
    monkeypatch.setattr(urllib.request, "urlopen", responder)
    return responder


def request():
    return urllib.request.Request("https://example.com/x")


class TestRetryOn429:
    def test_429_のあと成功すれば結果を返す(self, client, monkeypatch):
        r = install(monkeypatch, _Responder(http_error(429), {"id": 1}))
        assert client._send(request(), "/x") == {"id": 1}
        assert r.calls == 2

    def test_429_は_POST_でも再送する(self, client, monkeypatch):
        """429 はリクエストが処理されていないため重複作成の心配がない。"""
        r = install(monkeypatch, _Responder(http_error(429), {"id": 1}))
        assert client._send(request(), "/x", idempotent=False) == {"id": 1}
        assert r.calls == 2

    def test_再送上限を超えたらエラー(self, client, monkeypatch):
        r = install(monkeypatch, _Responder(*[http_error(429)] * 10))
        with pytest.raises(BacklogAPIError) as exc:
            client._send(request(), "/x")
        assert exc.value.status == 429
        assert r.calls == client.MAX_RETRIES + 1


class TestRetryOnServerErrorAndNetwork:
    def test_5xx_は冪等なら再送する(self, client, monkeypatch):
        r = install(monkeypatch, _Responder(http_error(503), {"id": 1}))
        assert client._send(request(), "/x") == {"id": 1}
        assert r.calls == 2

    def test_5xx_は_POST_では再送しない(self, client, monkeypatch):
        r = install(monkeypatch, _Responder(http_error(503), {"id": 1}))
        with pytest.raises(BacklogAPIError):
            client._send(request(), "/x", idempotent=False)
        assert r.calls == 1

    def test_通信断は冪等なら再送する(self, client, monkeypatch):
        r = install(monkeypatch, _Responder(urllib.error.URLError("接続断"), {"id": 1}))
        assert client._send(request(), "/x") == {"id": 1}
        assert r.calls == 2

    def test_通信断は_POST_では再送せず注意を促す(self, client, monkeypatch):
        r = install(monkeypatch, _Responder(urllib.error.URLError("接続断")))
        with pytest.raises(BacklogAPIError, match="作成済みの可能性"):
            client._send(request(), "/x", idempotent=False)
        assert r.calls == 1

    def test_タイムアウトも_POST_では再送しない(self, client, monkeypatch):
        r = install(monkeypatch, _Responder(TimeoutError()))
        with pytest.raises(BacklogAPIError, match="作成済みの可能性"):
            client._send(request(), "/x", idempotent=False)
        assert r.calls == 1


class TestNoRetryOnClientError:
    @pytest.mark.parametrize("status", [400, 401, 403, 404])
    def test_4xx_は再送しない(self, client, monkeypatch, status):
        r = install(monkeypatch, _Responder(http_error(status)))
        with pytest.raises(BacklogAPIError):
            client._send(request(), "/x")
        assert r.calls == 1


class TestRetryWait:
    def test_指数バックオフ(self, client):
        assert client._retry_wait(0, None) == 2.0
        assert client._retry_wait(1, None) == 4.0
        assert client._retry_wait(2, None) == 8.0

    def test_Retry_After_を優先する(self, client):
        assert client._retry_wait(0, "10") == 10.0

    def test_Retry_After_にも上限を適用する(self, client):
        assert client._retry_wait(0, "99999") == client.MAX_RETRY_WAIT

    def test_解釈できない_Retry_After_はバックオフにフォールバック(self, client):
        assert client._retry_wait(0, "Wed, 21 Oct 2026 07:28:00 GMT") == 2.0

    def test_バックオフにも上限を適用する(self, client):
        assert client._retry_wait(20, None) == client.MAX_RETRY_WAIT
