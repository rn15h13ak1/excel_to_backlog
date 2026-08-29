"""
CLI（main）のテスト
===================
BacklogClient を差し替えて main() をそのまま実行する。
引数の解釈・確認フロー・中断時の扱いなど、これまで未検証だった経路を通す。
"""

import csv

import pytest
import yaml

import excel_to_backlog as etb
from backlog_client import BacklogAPIError
from conftest import FakeBacklog


@pytest.fixture
def workspace(tmp_path, make_excel):
    """
    config.yaml とサンプル Excel を用意し、設定を書き換えられる形で返す。
    実行ログ・プレビューは config と同じディレクトリに出力される。
    """
    excel = make_excel(["件名", "Backlog番号"], [["課題A", ""], ["課題B", ""]])

    class Workspace:
        dir = tmp_path
        config = tmp_path / "config.yaml"

        def write(self, **overrides):
            cfg = {
                "backlog": {
                    "space_host": "example.backlog.com",
                    "api_key": "dummy-key",
                    "project_key": "DEMO",
                },
                "sources": [{
                    "name": "タスク管理表",
                    "excel": {"path": str(excel)},
                    "issue_mapping": {
                        "issue_type": "タスク", "priority": "中", "summary_col": "件名",
                    },
                }],
            }
            for key, value in overrides.items():
                cfg[key] = value
            self.config.write_text(yaml.safe_dump(cfg, allow_unicode=True), encoding="utf-8")
            return self.config

        def run_logs(self):
            return sorted(self.dir.glob("run_*.csv"))

        def previews(self):
            return sorted(self.dir.glob("preview_*.md"))

    ws = Workspace()
    ws.write()
    return ws


@pytest.fixture
def backlog(monkeypatch):
    """main() が生成する BacklogClient を FakeBacklog に差し替える。"""
    instance = FakeBacklog()
    monkeypatch.setattr(etb, "BacklogClient", lambda *a, **kw: instance)
    return instance


def main_with(*args):
    """main() を実行し、SystemExit の終了コードを返す（正常終了は 0）。"""
    import sys
    sys.argv = ["excel_to_backlog.py", *args]
    try:
        etb.main()
    except SystemExit as e:
        return e.code or 0
    return 0


# ------------------------------------------------------------------
# 引数の解釈
# ------------------------------------------------------------------

class TestArguments:
    def test_preview_と_execute_は同時指定できない(self, workspace, backlog):
        # argparse の parser.error() は終了コード 2 で SystemExit を送出する
        assert main_with("--config", str(workspace.config), "--preview", "--execute") == 2

    def test_存在しない設定ファイルはエラー(self, tmp_path):
        assert main_with("--config", str(tmp_path / "ない.yaml")) == 1

    def test_source_で絞り込める(self, workspace, backlog, capsys):
        main_with("--config", str(workspace.config), "--source", "タスク管理表")
        assert "ソース数    : 1" in capsys.readouterr().out

    def test_存在しないソース名はエラー(self, workspace, backlog, capsys):
        code = main_with("--config", str(workspace.config), "--source", "無いソース")
        assert code == 1
        assert "無いソース" in capsys.readouterr().err

    def test_api_key_が未設定ならエラー(self, workspace, backlog, capsys):
        workspace.write(backlog={
            "space_host": "example.backlog.com",
            "api_key": "YOUR_API_KEY_HERE",       # プレースホルダーのまま
            "project_key": "DEMO",
        })
        assert main_with("--config", str(workspace.config)) == 1
        assert "api_key" in capsys.readouterr().err

    def test_sources_が空ならエラー(self, workspace, backlog, capsys):
        workspace.write(sources=[])
        assert main_with("--config", str(workspace.config)) == 1
        assert "sources" in capsys.readouterr().err


# ------------------------------------------------------------------
# 実行モード
# ------------------------------------------------------------------

class TestModes:
    def test_既定はドライランで書き込まない(self, workspace, backlog, capsys):
        main_with("--config", str(workspace.config))

        assert backlog.create_calls == 0
        out = capsys.readouterr().out
        assert "DRY RUN" in out
        assert workspace.run_logs() == []       # ログも作らない

    def test_execute_で課題を作成する(self, workspace, backlog):
        main_with("--config", str(workspace.config), "--execute", "--yes")
        assert backlog.create_calls == 2

    def test_preview_で_Markdown_を出力する(self, workspace, backlog):
        main_with("--config", str(workspace.config), "--preview")

        previews = workspace.previews()
        assert len(previews) == 1
        assert "課題A" in previews[0].read_text(encoding="utf-8")
        assert backlog.create_calls == 0


# ------------------------------------------------------------------
# 実行前の確認
# ------------------------------------------------------------------

class TestConfirmation:
    def test_非対話環境で_yes_なしなら停止する(self, workspace, backlog, capsys):
        """
        以前は input() の EOFError を全行スキップとして扱い、1件も作らずに
        「作成: 0 件」と表示して正常終了していた（成功に見える無処理）。
        """
        code = main_with("--config", str(workspace.config), "--execute")

        assert code == 1
        assert backlog.create_calls == 0
        assert "--yes" in capsys.readouterr().err

    def test_確認で_n_を選ぶと中止する(self, workspace, backlog, monkeypatch, capsys):
        monkeypatch.setattr("builtins.input", lambda _: "n")

        code = main_with("--config", str(workspace.config), "--execute")

        assert code == 1
        assert backlog.create_calls == 0
        assert "取り消しました" in capsys.readouterr().err

    def test_確認で_y_を選ぶと実行する(self, workspace, backlog, monkeypatch):
        monkeypatch.setattr("builtins.input", lambda _: "y")
        main_with("--config", str(workspace.config), "--execute")
        assert backlog.create_calls == 2

    def test_ドライランでは確認を求めない(self, workspace, backlog, monkeypatch):
        def no_input(_):
            raise AssertionError("ドライランで確認を求めてはいけない")

        monkeypatch.setattr("builtins.input", no_input)
        assert main_with("--config", str(workspace.config)) == 0


# ------------------------------------------------------------------
# 実行ログと再開
# ------------------------------------------------------------------

class TestRunLogAndResume:
    def test_execute_で実行ログが作られる(self, workspace, backlog, capsys):
        main_with("--config", str(workspace.config), "--execute", "--yes")

        logs = workspace.run_logs()
        assert len(logs) == 1
        with open(logs[0], encoding="utf-8-sig", newline="") as f:
            actions = [r["action"] for r in csv.DictReader(f)]
        assert actions == ["created", "created"]
        assert "実行ログ:" in capsys.readouterr().out

    def test_no_log_で実行ログを抑止できる(self, workspace, backlog):
        main_with("--config", str(workspace.config), "--execute", "--yes", "--no-log")
        assert workspace.run_logs() == []
        assert backlog.create_calls == 2

    def test_resume_で処理済みの行を飛ばす(self, workspace, backlog, capsys):
        main_with("--config", str(workspace.config), "--execute", "--yes")
        first_log = workspace.run_logs()[0]
        assert backlog.create_calls == 2

        main_with("--config", str(workspace.config), "--execute", "--yes",
                  "--resume", str(first_log))

        assert backlog.create_calls == 2                 # 増えない
        assert "再開スキップ: 2 件" in capsys.readouterr().out

    def test_存在しない再開ログはエラー(self, workspace, backlog, tmp_path):
        with pytest.raises(FileNotFoundError):
            main_with("--config", str(workspace.config), "--execute", "--yes",
                      "--resume", str(tmp_path / "ない.csv"))


# ------------------------------------------------------------------
# 中断とエラー
# ------------------------------------------------------------------

class TestInterruption:
    def test_認証エラーは実行全体を中止しサマリーを出す(
        self, workspace, backlog, monkeypatch, capsys
    ):
        def unauthorized(params):
            raise BacklogAPIError("認証に失敗しました", status=401, fatal=True)

        monkeypatch.setattr(backlog, "create_issue", unauthorized)

        code = main_with("--config", str(workspace.config), "--execute", "--yes")

        assert code == 1
        out = capsys.readouterr().out
        assert "処理中断" in out
        assert "API エラーのため中止しました" in out

    def test_Ctrl_C_でもサマリーを出す(self, workspace, backlog, monkeypatch, capsys):
        def interrupt(params):
            raise KeyboardInterrupt

        monkeypatch.setattr(backlog, "create_issue", interrupt)

        code = main_with("--config", str(workspace.config), "--execute", "--yes")

        assert code == 1
        assert "ユーザーによる中断" in capsys.readouterr().out

    def test_1件失敗しても残りを処理する(self, workspace, monkeypatch, capsys):
        instance = FakeBacklog(fail_create_at=1)
        monkeypatch.setattr(etb, "BacklogClient", lambda *a, **kw: instance)

        code = main_with("--config", str(workspace.config), "--execute", "--yes")

        # 残りの行は処理されるが、失敗があるため終了コードは 1
        assert code == 1
        out = capsys.readouterr().out
        assert "作成: 1 件" in out
        assert "エラー: 1 件" in out


# ------------------------------------------------------------------
# 列名検証との連携
# ------------------------------------------------------------------

class TestColumnValidation:
    def test_列名が一致しなければ_API_を呼ばずエラーになる(
        self, workspace, backlog, capsys
    ):
        workspace.write(sources=[{
            "name": "タスク管理表",
            "excel": {"path": yaml.safe_load(
                workspace.config.read_text(encoding="utf-8")
            )["sources"][0]["excel"]["path"]},
            "filters": [{"col_name": "ステータス ", "value": "対応要"}],  # 末尾スペース
            "issue_mapping": {
                "issue_type": "タスク", "priority": "中", "summary_col": "件名",
            },
        }])

        main_with("--config", str(workspace.config), "--execute", "--yes")

        assert backlog.create_calls == 0
        assert "ヘッダーに存在しません" in capsys.readouterr().err


class TestExitCode:
    def test_エラーがあれば終了コード_1(self, workspace, monkeypatch):
        """以前は全行失敗しても 0 で終わっていた。"""
        instance = FakeBacklog(fail_create_at=1)
        monkeypatch.setattr(etb, "BacklogClient", lambda *a, **kw: instance)

        assert main_with("--config", str(workspace.config), "--execute", "--yes") == 1

    def test_すべて成功すれば終了コード_0(self, workspace, backlog):
        assert main_with("--config", str(workspace.config), "--execute", "--yes") == 0

    def test_中断後もそれまでの作成件数がサマリーに出る(
        self, workspace, backlog, monkeypatch, capsys
    ):
        original = backlog.create_issue

        def fail_on_second(params):
            if backlog.create_calls >= 1:
                raise BacklogAPIError("認証失敗", status=401, fatal=True)
            return original(params)

        monkeypatch.setattr(backlog, "create_issue", fail_on_second)

        main_with("--config", str(workspace.config), "--execute", "--yes")

        out = capsys.readouterr().out
        assert "処理中断" in out
        assert "作成: 1 件" in out          # 以前は「作成: 0 件」だった
