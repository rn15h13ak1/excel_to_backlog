"""
README と実装の整合テスト
=========================
README に書いた設定キー・オプションが実装と食い違っていないことを機械的に
確かめる。文書は実装に追随し忘れやすく、実際これまでに
  「スキップ（変更なし）」というもう出力されない文言
  必須設定 issue_type / priority の説明が丸ごと欠落
といったずれが発生している。
"""

import re
from pathlib import Path

import pytest
import yaml

from config_validation import (
    BACKLOG_KEYS, CUSTOM_FIELD_KEYS, EXCEL_KEYS, FILTER_KEYS,
    ISSUE_MAPPING_KEYS, UPSERT_KEYS,
)

ROOT = Path(__file__).resolve().parent.parent
README = (ROOT / "README.md").read_text(encoding="utf-8")
SAMPLE = (ROOT / "config.sample.yaml").read_text(encoding="utf-8")
CLI_SRC = (ROOT / "excel_to_backlog.py").read_text(encoding="utf-8")

ALL_KEYS = (
    BACKLOG_KEYS | EXCEL_KEYS | ISSUE_MAPPING_KEYS
    | CUSTOM_FIELD_KEYS | FILTER_KEYS | UPSERT_KEYS
)


class TestOptionsDocumented:
    def cli_options(self):
        return {
            opt
            for line in re.findall(r"parser\.add_argument\((.*?)\)", CLI_SRC, re.S)
            for opt in re.findall(r'"(--?[a-z][a-z-]*)"', line)
        }

    def test_すべてのオプションが_README_に載っている(self):
        # README は `--config path` のように引数付きで書くため、
        # オプション名に続く文字が英字でないことだけを見る
        undocumented = {
            o for o in self.cli_options()
            if not re.search(re.escape(o) + r"(?![a-z-])", README)
        }
        assert undocumented == set(), f"README に未記載: {sorted(undocumented)}"

    def test_README_のオプションがすべて実装されている(self):
        documented = set(re.findall(r"`(--[a-z][a-z-]*)`", README))
        assert documented <= self.cli_options() | {"--cov", "--cov-report"}


class TestConfigKeysDocumented:
    """許可リストにあるキーは、README か config.sample のどちらかで説明する。"""

    @pytest.mark.parametrize("key", sorted(ALL_KEYS))
    def test_設定キーが文書に載っている(self, key):
        assert key in README or key in SAMPLE, f"どこにも説明がない: {key}"

    @pytest.mark.parametrize("key", sorted(
        {"issue_type", "priority", "summary_col", "due_date_col",
         "start_date_col", "rich_text", "assignee_col", "status_col"}
    ))
    def test_主要なキーは_README_で説明されている(self, key):
        """config.sample のコメントだけでなく README にも書く。"""
        assert key in README, f"README に説明がない: {key}"


class TestSampleConfigIsValid:
    """config.sample.yaml が現在の検証を通ること。"""

    def test_サンプル設定に未知のキーが無い(self):
        from config_validation import validate_config_keys
        config = yaml.safe_load(SAMPLE)
        assert validate_config_keys(config) == []

    def test_サンプル設定が_YAML_として読める(self):
        assert isinstance(yaml.safe_load(SAMPLE), dict)


class TestQuotedMessagesExist:
    """
    README のトラブルシューティングが引用する文言が実装に存在すること。
    存在しない文言を載せると、見たメッセージで検索しても見つからない。
    """

    SOURCES = "".join(
        (ROOT / name).read_text(encoding="utf-8")
        for name in ["excel_to_backlog.py", "mapper.py", "excel_reader.py",
                     "backlog_client.py", "config_validation.py", "run_log.py",
                     "summary_index.py"]
    )

    @pytest.mark.parametrize("message", [
        "変更なし",
        "認識できないキーです",
        "Strict Open XML",
        "ヘッダー名が重複しています",
        "日付として解釈できません",
        "ヘッダーに存在しません",
    ])
    def test_引用された文言が実装にある(self, message):
        assert message in README, f"README に無い: {message}"
        assert message in self.SOURCES, f"実装に無い: {message}"


class TestExitCodesDocumented:
    def test_終了コードの表がある(self):
        assert "### 終了コード" in README

    @pytest.mark.parametrize("code", ["`0`", "`1`", "`2`"])
    def test_各コードが説明されている(self, code):
        assert code in README
