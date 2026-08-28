"""
BacklogClient の HTTP 層のテスト
================================
urlopen を差し替えて、実際に送信される URL とボディを検証する。

リクエストボディの組み立ては過去に2回バグが出ている:
  1181c0b POST/PATCH ボディの [] をパーセントエンコードしていた
  ea2c28e typeId 5/8 を [] 付き配列形式で送信していた
いずれも Backlog 側がパラメータを認識できず、原因の特定に時間がかかった。
"""

import json
import urllib.parse
import urllib.request

import pytest

from backlog_client import BacklogAPIError, BacklogClient


@pytest.fixture
def client():
    return BacklogClient("example.backlog.com", "key/with+chars")


@pytest.fixture
def captured(monkeypatch):
    """
    urlopen を差し替えて送信内容を記録する。

    captured.requests に urllib.request.Request が順に積まれる。
    captured.responses に返す JSON を積んでおくと順に返す。
    """
    class Captured:
        requests = []
        responses = [{"id": 1}]

        @property
        def last(self):
            return self.requests[-1]

        def body(self):
            return self.last.data.decode("utf-8")

        def params(self):
            """ボディを {キー: [値, ...]} に分解する（[] はキー名の一部として保持）。"""
            result = {}
            for pair in self.body().split("&"):
                key, _, value = pair.partition("=")
                result.setdefault(key, []).append(urllib.parse.unquote_plus(value))
            return result

        def query(self):
            return urllib.parse.parse_qs(urllib.parse.urlparse(self.last.full_url).query)

    cap = Captured()
    cap.requests = []

    def fake_urlopen(req, timeout=None, context=None):
        cap.requests.append(req)
        payload = cap.responses[min(len(cap.requests) - 1, len(cap.responses) - 1)]

        class _Res:
            def __enter__(self_inner): return self_inner
            def __exit__(self_inner, *a): return False
            def read(self_inner): return json.dumps(payload).encode("utf-8")

        return _Res()

    monkeypatch.setattr(urllib.request, "urlopen", fake_urlopen)
    return cap


# ------------------------------------------------------------------
# URL の組み立て
# ------------------------------------------------------------------

class TestUrl:
    def test_GET_に_apiKey_が付く(self, client, captured):
        client._get("/issues", {"count": 100})
        assert captured.query()["apiKey"] == ["key/with+chars"]

    def test_apiKey_の記号がエスケープされる(self, client, captured):
        """
        + はエスケープしないとスペースとして解釈される。
        / はクエリ文字列の値として正当なためエスケープされない（RFC 3986）。
        """
        client._get("/issues")

        assert "%2B" in captured.last.full_url          # + がエスケープ済み
        assert captured.query()["apiKey"] == ["key/with+chars"]  # 復号すると元に戻る

    def test_リスト値は角括弧付きで展開される(self, client, captured):
        client._get("/issues", {"projectId": [1, 2]})
        assert captured.query()["projectId[]"] == ["1", "2"]

    def test_base_path_が_URL_に反映される(self, captured):
        c = BacklogClient("example.com", "k", base_path="/backlog")
        c._get("/issues")
        assert captured.last.full_url.startswith("https://example.com/backlog/api/v2/issues")


# ------------------------------------------------------------------
# リクエストボディ（過去にバグが出た箇所）
# ------------------------------------------------------------------

class TestRequestBody:
    def test_POST_のメソッドとヘッダー(self, client, captured):
        client._post("/issues", {"summary": "件名"})
        assert captured.last.get_method() == "POST"
        assert captured.last.headers["Content-type"] == "application/x-www-form-urlencoded"

    def test_PATCH_のメソッド(self, client, captured):
        client._patch("/issues/DEMO-1", {"summary": "件名"})
        assert captured.last.get_method() == "PATCH"

    def test_角括弧はパーセントエンコードしない(self, client, captured):
        """
        キー名の [] を %5B%5D にすると Backlog がリスト表記を認識できない。
        （1181c0b で修正した退行）
        """
        client._post("/issues", {"customField_6": [61, 62]})

        assert "customField_6[]=61" in captured.body()
        assert "%5B%5D" not in captured.body()

    def test_単一値には角括弧を付けない(self, client, captured):
        """
        typeId 5/8（単一選択）は [] なしで送る必要がある。
        （ea2c28e で修正した退行）
        """
        client._post("/issues", {"customField_5": 51})

        assert "customField_5=51" in captured.body()
        assert "customField_5[]" not in captured.body()

    def test_値はパーセントエンコードされる(self, client, captured):
        client._post("/issues", {"summary": "件名 & 記号=あり"})

        assert captured.params()["summary"] == ["件名 & 記号=あり"]
        assert "&" not in captured.body().split("summary=")[1].split("&")[0].replace("%26", "")

    def test_日本語が_UTF8_で送られる(self, client, captured):
        client._post("/issues", {"summary": "日本語"})
        assert urllib.parse.quote_plus("日本語") in captured.body()

    def test_改行を含む本文も送れる(self, client, captured):
        client._post("/issues", {"description": "1行目\n2行目"})
        assert captured.params()["description"] == ["1行目\n2行目"]

    def test_複数パラメータが_アンパサンド_で連結される(self, client, captured):
        client._post("/issues", {"projectId": 42, "summary": "件名", "issueTypeId": 1})
        params = captured.params()
        assert params["projectId"] == ["42"]
        assert params["issueTypeId"] == ["1"]


# ------------------------------------------------------------------
# 課題の取得
# ------------------------------------------------------------------

class TestGetIssue:
    def test_存在しない課題は_None(self, client, monkeypatch):
        def not_found(req, timeout=None, context=None):
            import io
            import urllib.error
            raise urllib.error.HTTPError(
                "url", 404, "Not Found", {}, io.BytesIO(b'{"errors":[]}')
            )

        monkeypatch.setattr(urllib.request, "urlopen", not_found)
        assert client.get_issue("DEMO-999") is None

    def test_404_以外はエラーとして送出される(self, client, monkeypatch):
        def server_error(req, timeout=None, context=None):
            import io
            import urllib.error
            raise urllib.error.HTTPError(
                "url", 403, "Forbidden", {}, io.BytesIO(b'{"errors":[]}')
            )

        monkeypatch.setattr(urllib.request, "urlopen", server_error)
        with pytest.raises(BacklogAPIError):
            client.get_issue("DEMO-1")

    def test_issueKey_は_URL_エンコードされる(self, client, captured):
        client.get_issue("DEMO-1")
        assert "/issues/DEMO-1" in captured.last.full_url


class TestGetIssuesPagination:
    """1リクエスト 100 件の上限を超えても全件取得すること。"""

    def test_100件未満なら1回で終わる(self, client, captured):
        captured.responses = [[{"issueKey": f"DEMO-{i}"} for i in range(30)]]
        issues = client.get_issues(42)

        assert len(issues) == 30
        assert len(captured.requests) == 1

    def test_ちょうど100件なら次ページも取りにいく(self, client, captured):
        captured.responses = [
            [{"issueKey": f"DEMO-{i}"} for i in range(100)],
            [{"issueKey": "DEMO-100"}],
        ]
        issues = client.get_issues(42)

        assert len(issues) == 101
        assert len(captured.requests) == 2

    def test_offset_が繰り上がる(self, client, captured):
        captured.responses = [
            [{"issueKey": f"DEMO-{i}"} for i in range(100)],
            [],
        ]
        client.get_issues(42)

        offsets = [
            urllib.parse.parse_qs(urllib.parse.urlparse(r.full_url).query)["offset"][0]
            for r in captured.requests
        ]
        assert offsets == ["0", "100"]

    def test_空の応答で打ち切る(self, client, captured):
        captured.responses = [[]]
        assert client.get_issues(42) == []
        assert len(captured.requests) == 1


# ------------------------------------------------------------------
# マスターデータ取得
# ------------------------------------------------------------------

class TestMasterEndpoints:
    @pytest.mark.parametrize("call,expected", [
        (lambda c: c.get_project("DEMO"),        "/projects/DEMO"),
        (lambda c: c.get_issue_types("DEMO"),    "/projects/DEMO/issueTypes"),
        (lambda c: c.get_custom_fields("DEMO"),  "/projects/DEMO/customFields"),
        (lambda c: c.get_statuses("DEMO"),       "/projects/DEMO/statuses"),
        (lambda c: c.get_project_users("DEMO"),  "/projects/DEMO/users"),
        (lambda c: c.get_priorities(),           "/priorities"),
    ])
    def test_正しいエンドポイントを呼ぶ(self, client, captured, call, expected):
        call(client)
        assert expected in captured.last.full_url
