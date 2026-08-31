"""
継続行の結合
============
1 件の内容が Excel の複数行に分かれている表を 1 件にまとめる。
セル内に収まらない内容を次の行へ書き足す運用への対応。
"""

import excel_to_backlog as etb
from conftest import FakeBacklog
from excel_reader import ExcelReader
from row_merge import is_continuation, single_value_columns

KEY = ExcelReader.ROW_NUMBER_KEY


def cfg_for(source_cfg, headers, rows, **mapping):
    base = {"required_cols": ["件名"], "merge_continuation_rows": True,
            "description_format": "auto"}
    base.update(mapping)
    return source_cfg(headers, rows, issue_mapping=base)


class TestIsContinuation:
    def test_必須列が空なら継続行(self):
        assert is_continuation({"件名": "", "対応内容": "続き"}, ["件名"]) is True

    def test_必須列に値があれば継続行ではない(self):
        assert is_continuation({"件名": "課題", "対応内容": "x"}, ["件名"]) is False

    def test_空白だけなら継続行(self):
        assert is_continuation({"件名": "   "}, ["件名"]) is True

    def test_複数の必須列はすべて空である必要がある(self):
        row = {"項番": "", "件名": "課題"}
        assert is_continuation(row, ["項番", "件名"]) is False

    def test_required_cols_が無ければ結合しない(self):
        assert is_continuation({"件名": ""}, []) is False


class TestMerge:
    def test_複数行が1件にまとまる(self, source_cfg, master):
        cfg = cfg_for(
            source_cfg, ["件名", "対応内容"],
            [["ログイン不具合", "手順1"], ["", "手順2"], ["", "手順3"],
             ["表示崩れ", "CSS 修正"]],
            description_cols=["対応内容"],
        )
        loaded = etb.load_source(cfg, master)

        assert len(loaded.rows) == 2
        assert loaded.rows[0]["件名"] == "ログイン不具合"

    def test_空行で区切って連結される(self, source_cfg, master):
        cfg = cfg_for(
            source_cfg, ["件名", "対応内容"],
            [["課題A", "手順1"], ["", "手順2"]],
            description_cols=["対応内容"],
        )
        loaded = etb.load_source(cfg, master)
        body = etb.plan_row(loaded.rows[0], cfg, loaded.mapper).params["description"]

        assert "手順1\n\n手順2" in body        # 段落として分かれる
        assert "<br><br>" not in body

    def test_セル内の改行は_br_のまま(self, source_cfg, master):
        cfg = cfg_for(
            source_cfg, ["件名", "対応内容"],
            [["課題A", "1行目\n2行目"], ["", "続き"]],
            description_cols=["対応内容"],
        )
        loaded = etb.load_source(cfg, master)
        body = etb.plan_row(loaded.rows[0], cfg, loaded.mapper).params["description"]

        assert "1行目<br>2行目" in body        # セル内改行
        assert "2行目\n\n続き" in body          # 結合の区切り

    def test_行番号は先頭行のもの(self, source_cfg, master):
        cfg = cfg_for(source_cfg, ["件名", "対応内容"],
                      [["課題A", "手順1"], ["", "手順2"]])
        loaded = etb.load_source(cfg, master)
        assert loaded.rows[0][KEY] == "2"

    def test_課題は結合後の件数だけ作られる(self, source_cfg, master):
        cfg = cfg_for(source_cfg, ["件名", "対応内容"],
                      [["課題A", "手順1"], ["", "手順2"], ["課題B", "別件"]])
        backlog = FakeBacklog()

        counts = etb.process_source(cfg, backlog, master, dry_run=False)

        assert counts["created"] == 2
        assert counts["skipped"] == 0          # 継続行はスキップされない

    def test_先頭が継続行でも落ちない(self, source_cfg, master):
        """1 行目から必須列が空のケース。結合先が無いのでそのまま残す。"""
        cfg = cfg_for(source_cfg, ["件名", "対応内容"],
                      [["", "宙に浮いた行"], ["課題A", "本体"]])
        loaded = etb.load_source(cfg, master)

        assert len(loaded.rows) == 2           # スキップは map_row 側で行う


class TestSingleValueColumns:
    def test_連結できない列を集める(self):
        cfg = {
            "issue_mapping": {
                "summary_col": "件名", "assignee_col": "担当",
                "status_col": "状態", "due_date_col": "期限",
                "custom_fields": [{"field_name": "分類", "col_name": "区分"}],
            },
            "upsert": {"key_col": "Backlog番号"},
        }
        assert single_value_columns(cfg) == {
            "件名", "担当", "状態", "期限", "区分", "Backlog番号"
        }

    def test_テンプレート指定の日付は対象外(self):
        cfg = {"issue_mapping": {"summary_col": "件名",
                                 "due_date_col": "{{年}}/{{月}}/1"}}
        assert single_value_columns(cfg) == {"件名"}

    def test_連結できない列は警告して無視する(self, source_cfg, master, capsys):
        cfg = cfg_for(
            source_cfg, ["件名", "対応内容", "担当"],
            [["課題A", "手順1", "山田太郎"], ["", "手順2", "佐藤"]],
            assignee_col="担当", description_cols=["対応内容"],
        )
        loaded = etb.load_source(cfg, master)

        err = capsys.readouterr().err
        assert "3行目は2行目の続き" in err
        assert "担当" in err
        assert loaded.rows[0]["担当"] == "山田太郎"      # 先頭行の値が残る


class TestDisabledByDefault:
    def test_既定では結合しない(self, source_cfg, master):
        """設定を明示しない限り従来どおり（必須列が空ならスキップ）。"""
        cfg = source_cfg(
            ["件名", "対応内容"], [["課題A", "手順1"], ["", "手順2"]],
            issue_mapping={"required_cols": ["件名"]},
        )
        backlog = FakeBacklog()

        counts = etb.process_source(cfg, backlog, master, dry_run=False)

        assert counts["created"] == 1
        assert counts["skipped"] == 1

    def test_required_cols_が無ければ結合しない(self, source_cfg, master):
        cfg = source_cfg(
            ["件名", "対応内容"], [["課題A", "手順1"], ["", "手順2"]],
            issue_mapping={"merge_continuation_rows": True},
        )
        loaded = etb.load_source(cfg, master)
        assert len(loaded.rows) == 2


class TestMergeBeforeFilter:
    def test_絞り込みより前に結合される(self, source_cfg, master):
        """
        継続行は絞り込み条件の列も空になっている。先に絞ると
        結合前に失われるため、結合を先に行う。
        """
        cfg = source_cfg(
            ["件名", "状態", "対応内容"],
            [["課題A", "対応要", "手順1"], ["", "", "手順2"],
             ["課題B", "完了", "別件"]],
            filters=[{"col_name": "状態", "value": "対応要"}],
            issue_mapping={"required_cols": ["件名"], "merge_continuation_rows": True,
                           "description_format": "auto", "description_cols": ["対応内容"]},
        )
        loaded = etb.load_source(cfg, master)

        assert len(loaded.rows) == 1
        body = etb.plan_row(loaded.rows[0], cfg, loaded.mapper).params["description"]
        assert "手順1" in body and "手順2" in body


class TestMergeReport:
    """
    件数だけでは「なぜ結合されなかったのか」が分からないため、
    どの行がどこへ結合されたかと、結合されなかった場合の判定条件を示す。
    """

    def test_結合された対応関係を表示する(self, source_cfg, master, capsys):
        cfg = cfg_for(source_cfg, ["件名", "対応内容"],
                      [["課題A", "手順1"], ["", "手順2"], ["", "手順3"]])
        etb.load_source(cfg, master)

        out = capsys.readouterr().out
        assert "継続行を結合: 3 行 → 1 件" in out
        assert "2行目 ← 3行目、4行目" in out

    def test_結合対象が無ければ判定条件を示す(self, source_cfg, master, capsys):
        """required_cols の指定が実態と合っていない場合に気づけるように。"""
        cfg = cfg_for(
            source_cfg, ["件名", "対応内容"],
            [["課題A", "手順1"], ["", "手順2"]],
            required_cols=["件名", "対応内容"],      # 対応内容は継続行にも値がある
        )
        etb.load_source(cfg, master)

        err = capsys.readouterr().err
        assert "結合対象の行はありません" in err
        assert "件名、対応内容" in err              # 判定に使っている列

    def test_複数の課題があれば別々に表示する(self, source_cfg, master, capsys):
        cfg = cfg_for(source_cfg, ["件名", "対応内容"],
                      [["課題A", "手順1"], ["", "手順2"],
                       ["課題B", "別件1"], ["", "別件2"]])
        etb.load_source(cfg, master)

        out = capsys.readouterr().out
        assert "2行目 ← 3行目" in out
        assert "4行目 ← 5行目" in out
