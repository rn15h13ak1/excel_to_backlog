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
    """
    判定に使う列が「空」または「直前の行と同じ値」なら継続行。
    項番などを継続行にも振っている表に対応するため、空欄だけを条件に
    していない（書式で見えなくしていて目視では空に見えることもある）。
    """

    PREV = {"項番": "1", "件名": "認証改修"}

    def test_空欄なら継続行(self):
        assert is_continuation({"件名": "", "対応内容": "続き"}, self.PREV, ["件名"]) is True

    def test_直前と同じ値なら継続行(self):
        row = {"項番": "1", "件名": "", "対応内容": "続き"}
        assert is_continuation(row, self.PREV, ["項番", "件名"]) is True

    def test_同じ値と空欄が混在しても継続行(self):
        row = {"項番": "1", "件名": "認証改修"}
        assert is_continuation(row, self.PREV, ["項番", "件名"]) is True

    def test_値が変われば新しい1件(self):
        row = {"項番": "2", "件名": "表示崩れ"}
        assert is_continuation(row, self.PREV, ["項番", "件名"]) is False

    def test_一部の列だけ変わっても新しい1件(self):
        row = {"項番": "2", "件名": ""}
        assert is_continuation(row, self.PREV, ["項番", "件名"]) is False

    def test_前後の空白は無視して比較する(self):
        row = {"項番": " 1 ", "件名": ""}
        assert is_continuation(row, self.PREV, ["項番", "件名"]) is True

    def test_直前の行が無ければ継続行ではない(self):
        assert is_continuation({"件名": ""}, None, ["件名"]) is False

    def test_required_cols_が無ければ結合しない(self):
        assert is_continuation({"件名": ""}, self.PREV, []) is False


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
            [["課題A", "手順1"], ["課題B", "手順2"]],   # 値が毎行変わる
            required_cols=["件名", "対応内容"],
        )
        etb.load_source(cfg, master)

        err = capsys.readouterr().err
        assert "結合対象の行はありません" in err
        assert "直前の行と同じ値" in err
        assert "件名、対応内容" in err              # 判定に使っている列

    def test_複数の課題があれば別々に表示する(self, source_cfg, master, capsys):
        cfg = cfg_for(source_cfg, ["件名", "対応内容"],
                      [["課題A", "手順1"], ["", "手順2"],
                       ["課題B", "別件1"], ["", "別件2"]])
        etb.load_source(cfg, master)

        out = capsys.readouterr().out
        assert "2行目 ← 3行目" in out
        assert "4行目 ← 5行目" in out


class TestRepeatedKeyValue:
    """
    項番などを継続行にも振っている表。文字色を変えて見えなくしている
    ことがあり、目視では空に見えても値が入っている。
    """

    def test_項番が同じ行は1件にまとまる(self, source_cfg, master):
        cfg = cfg_for(
            source_cfg, ["項番", "件名", "対応内容"],
            [[1, "認証改修", "手順1"], [1, "", "手順2"], [1, "", "手順3"]],
            required_cols=["項番", "件名"], description_cols=["対応内容"],
        )
        loaded = etb.load_source(cfg, master)

        assert len(loaded.rows) == 1
        body = etb.plan_row(loaded.rows[0], cfg, loaded.mapper).params["description"]
        assert "手順1\n\n手順2\n\n手順3" in body

    def test_項番が変われば別の1件(self, source_cfg, master):
        cfg = cfg_for(
            source_cfg, ["項番", "件名", "対応内容"],
            [[1, "認証改修", "手順1"], [1, "", "手順2"], [2, "表示崩れ", "別件"]],
            required_cols=["項番", "件名"], description_cols=["対応内容"],
        )
        loaded = etb.load_source(cfg, master)

        assert len(loaded.rows) == 2
        assert [r["件名"] for r in loaded.rows] == ["認証改修", "表示崩れ"]

    def test_判定に使う列は連結されない(self, source_cfg, master):
        """
        同じ値の前提のため、連結すると「1\\n\\n1」となり次の行との比較が壊れる。
        先頭行の値を保つ。
        """
        cfg = cfg_for(
            source_cfg, ["項番", "件名", "対応内容"],
            [[1, "認証改修", "手順1"], [1, "", "手順2"], [1, "", "手順3"]],
            required_cols=["項番", "件名"],
        )
        loaded = etb.load_source(cfg, master)

        assert loaded.rows[0]["項番"] == "1"
        assert loaded.rows[0]["件名"] == "認証改修"

    def test_件名も繰り返されている場合(self, source_cfg, master):
        cfg = cfg_for(
            source_cfg, ["項番", "件名", "対応内容"],
            [[1, "認証改修", "手順1"], [1, "認証改修", "手順2"]],
            required_cols=["項番", "件名"], description_cols=["対応内容"],
        )
        loaded = etb.load_source(cfg, master)

        assert len(loaded.rows) == 1

    def test_本文には判定列を出しても重複しない(self, source_cfg, master):
        cfg = cfg_for(
            source_cfg, ["項番", "件名", "対応内容"],
            [[1, "認証改修", "手順1"], [1, "", "手順2"]],
            required_cols=["項番", "件名"],
        )
        loaded = etb.load_source(cfg, master)
        body = etb.plan_row(loaded.rows[0], cfg, loaded.mapper).params["description"]

        assert body.count("# 項番") == 1
        assert "1\n\n1" not in body


class TestJoinSpec:
    """
    判定に使う列（required_cols）以外の連結の仕様。

    値が空の行は区切りを入れずに飛ばす。区切りを無条件に挟むと、
    2 行目以降に値が無い列で空行だけが増えていく。
    """

    def _value(self, source_cfg, master, rows, column="備考"):
        cfg = cfg_for(
            source_cfg, ["項番", "件名", "対応内容", "備考"], rows,
            required_cols=["項番", "件名"],
            description_cols=["対応内容", "備考"],
        )
        return etb.load_source(cfg, master).rows[0][column]

    def test_継続行の値が空なら区切りを入れない(self, source_cfg, master):
        got = self._value(source_cfg, master, [
            [1, "A", "手順1", "メモ1"], [1, "", "手順2", ""], [1, "", "手順3", ""],
        ])
        assert got == "メモ1"

    def test_途中の行だけ値がある場合(self, source_cfg, master):
        got = self._value(source_cfg, master, [
            [1, "A", "手順1", ""], [1, "", "手順2", "途中のメモ"], [1, "", "手順3", ""],
        ])
        assert got == "途中のメモ"

    def test_先頭が空で継続行に値がある場合(self, source_cfg, master):
        """先頭に区切りが付かないこと。"""
        got = self._value(source_cfg, master, [
            [1, "A", "手順1", ""], [1, "", "手順2", "後から追記"],
        ])
        assert got == "後から追記"

    def test_すべて空なら空のまま(self, source_cfg, master):
        got = self._value(source_cfg, master, [
            [1, "A", "手順1", ""], [1, "", "手順2", ""],
        ])
        assert got == ""

    def test_空白だけの値も飛ばす(self, source_cfg, master):
        got = self._value(source_cfg, master, [
            [1, "A", "手順1", ""], [1, "", "   ", ""], [1, "", "手順3", ""],
        ], column="対応内容")
        assert got == "手順1\n\n手順3"

    def test_値は前後の空白を除いて連結する(self, source_cfg, master):
        """セル内の末尾改行が余分な空行にならないこと。"""
        got = self._value(source_cfg, master, [
            [1, "A", "手順1\n", ""], [1, "", "手順2", ""],
        ], column="対応内容")
        assert got == "手順1\n\n手順2"


class TestNoExcessBlankLines:
    """本文に 3 連続以上の改行が現れないこと。"""

    def _body(self, source_cfg, master, rows):
        cfg = cfg_for(
            source_cfg, ["項番", "件名", "対応内容"], rows,
            required_cols=["項番", "件名"], description_cols=["対応内容"],
        )
        loaded = etb.load_source(cfg, master)
        return etb.plan_row(loaded.rows[0], cfg, loaded.mapper).params["description"]

    def test_継続行が空でも空行が増えない(self, source_cfg, master):
        body = self._body(source_cfg, master, [
            [1, "A", "手順1"], [1, "", ""], [1, "", "手順3"],
        ])
        assert "\n\n\n" not in body
        assert "手順1\n\n手順3" in body

    def test_セル内に空行があっても増えない(self, source_cfg, master):
        body = self._body(source_cfg, master, [
            [1, "A", "手順1\n\n注記"], [1, "", "手順2"],
        ])
        assert "\n\n\n" not in body

    def test_セル内の末尾改行でも増えない(self, source_cfg, master):
        body = self._body(source_cfg, master, [
            [1, "A", "手順1\n"], [1, "", "手順2"],
        ])
        assert "\n\n\n" not in body


class TestNoDuplicateJoin:
    """
    判定に使わない列に同じ値が繰り返し入っていても重複させない。

    項番を required_cols に、枝番を入れていない表がある。枝番は 1 件を
    通して同じ値なので、そのまま連結すると「1\n\n1\n\n1」となり、
    件名テンプレートに使うと「1-111」になってしまう。
    """

    def _row(self, source_cfg, master, rows, **mapping):
        mapping.setdefault("description_cols", ["対応内容"])
        mapping.setdefault("summary_template", "{{項番}}-{{枝番}}")
        cfg = cfg_for(
            source_cfg, ["項番", "枝番", "対応内容"], rows,
            required_cols=["項番"], summary_col=None, **mapping,
        )
        loaded = etb.load_source(cfg, master)
        assert len(loaded.rows) == 1
        return loaded, cfg

    def test_同じ値は連結されない(self, source_cfg, master):
        loaded, _ = self._row(source_cfg, master, [
            [1, 1, "手順1"], [1, 1, "手順2"], [1, 1, "手順3"],
        ])
        assert loaded.rows[0]["枝番"] == "1"
        assert loaded.rows[0]["対応内容"] == "手順1\n\n手順2\n\n手順3"

    def test_値が異なれば連結される(self, source_cfg, master):
        loaded, _ = self._row(source_cfg, master, [
            [1, 1, "手順1"], [1, 2, "手順2"],
        ])
        assert loaded.rows[0]["枝番"] == "1\n\n2"

    def test_同じ値が飛び飛びでも重複しない(self, source_cfg, master):
        loaded, _ = self._row(source_cfg, master, [
            [1, "A", "手順1"], [1, "B", "手順2"], [1, "A", "手順3"],
        ])
        assert loaded.rows[0]["枝番"] == "A\n\nB"

    def test_件名テンプレートが壊れない(self, source_cfg, master):
        """1-111 ではなく 1-1 になる。既存課題との照合も保たれる。"""
        loaded, cfg = self._row(
            source_cfg, master,
            [[1, 1, "手順1"], [1, 1, "手順2"], [1, 1, "手順3"]],
            summary_template="{{項番}}-{{枝番}}",
        )
        plan = etb.plan_row(loaded.rows[0], cfg, loaded.mapper)
        assert plan.params["summary"] == "1-1"

    def test_本文にも重複して出ない(self, source_cfg, master):
        """本文用の列順の値リストにも同じ規則が効くこと。"""
        loaded, cfg = self._row(
            source_cfg, master,
            [[1, 1, "手順1"], [1, 1, "手順2"]],
            description_cols=["枝番", "対応内容"],
        )
        body = etb.plan_row(loaded.rows[0], cfg, loaded.mapper).params["description"]
        assert body == "# 枝番\n1\n\n# 対応内容\n手順1\n\n手順2"
