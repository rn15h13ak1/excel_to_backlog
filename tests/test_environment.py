"""
依存が入っていない Python で起動したときの案内
==============================================
開発機には依存が入っているため、この経路が壊れてもほかのテストは全部通る。
実際に import を撥ねた状態を作って確かめる。

sys.meta_path に「名前を見て ModuleNotFoundError を送出する finder」を差し込む。
本ツールは `import yaml` を直接書いており、importlib.util.find_spec() で存在を
調べてはいないため、この細工で忠実に再現できる。

細工が空振りしても気づけないので、依存がある場合に通常どおり起動することも
併せて確かめる。
"""

import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
ENTRY = ROOT / "excel_to_backlog.py"

BLOCKER = """\
import sys


class _Blocker:
    \"\"\"指定した名前の import だけを撥ねる。\"\"\"

    def __init__(self, names):
        self.names = set(names)

    def find_spec(self, fullname, path=None, target=None):
        if fullname.split(".")[0] in self.names:
            raise ModuleNotFoundError(f"No module named {fullname!r}", name=fullname)
        return None


sys.meta_path.insert(0, _Blocker(__NAMES__))
"""


def run_entry(tmp_path, blocked=()):
    """excel_to_backlog.py --help を、指定モジュールを撥ねた状態で実行する。"""
    if blocked:
        # BLOCKER は f-string を含むため .format() は使えない（波括弧が衝突する）
        (tmp_path / "sitecustomize.py").write_text(
            BLOCKER.replace("__NAMES__", repr(list(blocked))), encoding="utf-8"
        )
    env = {
        "PATH": "/usr/bin:/bin",
        "PYTHONPATH": f"{tmp_path}:{ROOT}" if blocked else str(ROOT),
        "HOME": str(tmp_path),
    }
    return subprocess.run(
        [sys.executable, str(ENTRY), "--help"],
        capture_output=True, text=True, env=env, cwd=ROOT,
    )


class TestDependencyGuard:
    @pytest.mark.parametrize("module,package", [
        ("yaml", "pyyaml"),
        ("openpyxl", "openpyxl"),
    ])
    def test_不足しているライブラリを名指しで案内する(self, tmp_path, module, package):
        r = run_entry(tmp_path, blocked=[module])

        assert r.returncode == 1
        assert "Traceback" not in r.stderr
        assert f"「{package}」が入っていません" in r.stderr

    def test_実行中の_Python_のパスを出す(self, tmp_path):
        """入れた先と動かしている先が違うことに気づけるようにする。"""
        r = run_entry(tmp_path, blocked=["yaml"])
        assert sys.executable in r.stderr

    def test_導入コマンドを示す(self, tmp_path):
        r = run_entry(tmp_path, blocked=["yaml"])
        assert "pip install -e ." in r.stderr

    def test_終了コードは_1(self, tmp_path):
        """既存の対応表に合わせる。新しい番号は足さない。"""
        assert run_entry(tmp_path, blocked=["yaml"]).returncode == 1

    def test_無関係なモジュールの不足は握り潰さない(self, tmp_path):
        """
        自前モジュールの綴り間違いまで「ライブラリを入れてください」と
        案内すると、本当の原因が隠れる。
        """
        r = run_entry(tmp_path, blocked=["backlog_client"])

        assert "Traceback" in r.stderr
        assert "backlog_client" in r.stderr
        assert "入っていません" not in r.stderr

    def test_依存があれば通常どおり起動する(self, tmp_path):
        """細工が空振りしていたら、この差が出ない。"""
        r = run_entry(tmp_path)

        assert r.returncode == 0
        assert "Traceback" not in r.stderr
        assert "入っていません" not in r.stderr
