"""
件名インデックス
================
upsert の match_summary 判定を、行ごとの検索から一括取得＋辞書引きに置き換える。

以前は 1 行ごとに GET /issues?keyword=... を発行していた（N+1）。
keyword 検索は本文にもマッチするため 1 行で複数ページ取得することもあり、
500 行なら千回規模の往復になっていた。

課題を一度だけページネーションで全件取得して辞書を作れば、
検索回数は ⌈全課題数 / 100⌉ 回で済む。
"""

from __future__ import annotations

import sys

from mapper import IssueMapper


class SummaryIndex:
    """
    正規化した件名 → issueKey の索引。

    初回参照時にプロジェクトの全課題を取得して構築する（遅延構築）。
    match_summary を使うソースが1つも無ければ API は呼ばれない。
    """

    def __init__(self, client, project_id: int):
        self.client = client
        self.project_id = project_id
        self._index: dict[str, str] | None = None

    # ------------------------------------------------------------------

    def _build(self) -> dict[str, str]:
        """全課題を取得して {正規化件名: issueKey} を作る。"""
        print("  既存課題を取得して件名の索引を作成中...")
        issues = self.client.get_issues(self.project_id)

        index: dict[str, str] = {}
        duplicates: dict[str, int] = {}
        for issue in issues:
            key = IssueMapper.normalize_summary(issue.get("summary", ""))
            if not key:
                continue
            if key in index:
                # 同じ件名の課題が複数ある場合は最初に見つかったものを使う。
                # 以前の検索方式も候補の先頭を採用していたため挙動は変わらない。
                duplicates[key] = duplicates.get(key, 1) + 1
                continue
            index[key] = issue["issueKey"]

        print(f"    {len(issues)} 件の課題から {len(index)} 件の件名を索引化しました")
        if duplicates:
            print(
                f"  ⚠ 件名が重複している課題が {len(duplicates)} 種類あります。"
                f"更新対象は最初の 1 件になります:",
                file=sys.stderr,
            )
            for summary, count in list(duplicates.items())[:5]:
                print(f"      「{summary}」（{count} 件）→ {index[summary]}", file=sys.stderr)
            if len(duplicates) > 5:
                print(f"      ...他 {len(duplicates) - 5} 種類", file=sys.stderr)

        return index

    # ------------------------------------------------------------------

    def find(self, summary: str) -> str | None:
        """
        正規化済みの件名に一致する既存課題の issueKey を返す。

        summary は map_row() で normalize_summary() 済みの文字列を渡すこと。
        索引側も同じ正規化を適用しているため表記の揺れを吸収できる。
        """
        if self._index is None:
            self._index = self._build()
        return self._index.get(summary)

    def add(self, summary: str, issue_key: str) -> None:
        """
        実行中に作成した課題を索引へ加える。

        索引は起動時のスナップショットのため、これを行わないと Excel に
        同じ件名の行が2つある場合に2件とも新規作成してしまう。
        （行ごとに検索していた従来方式では、直前に作成した課題も検索で
          見つかっていた。同じ挙動を保つために必要。）
        """
        if self._index is None:
            # まだ一度も参照されていない場合は構築を先送りする。
            # 次の find() で全件取得すれば、この課題も含まれる。
            return
        key = IssueMapper.normalize_summary(summary)
        if key:
            self._index.setdefault(key, issue_key)
