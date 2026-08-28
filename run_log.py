"""
実行ログ（CSV）
===============
処理した行を1件ずつ CSV に追記する。

これがないと、実行が途中で落ちたときに「何件作られたか」を知る手段が
ターミナルのスクロールバックしか無く、upsert 無効のまま再実行すると
作成済みの行が重複する。

各行を書いた直後に flush するため、強制終了しても直前までの内容が残る。
"""

from __future__ import annotations

import csv
import sys
from datetime import datetime
from pathlib import Path

# 「その行の処理は完了しており、再実行時に繰り返す必要がない」と見なす結果。
# error / skipped は再実行の対象に含める（前者は失敗、後者は未処理のため）。
COMPLETED_ACTIONS = frozenset({"created", "created_status_failed", "updated", "unchanged"})

FIELDNAMES = ["time", "source", "row", "action", "issue_key", "summary", "detail"]


class RunLog:
    """
    実行結果を CSV に追記する。

    with 文で使うとファイルを確実に閉じる。
    """

    def __init__(self, path: Path):
        self.path = Path(path)
        self._file = None
        self._writer = None
        self.written = 0

    def __enter__(self) -> "RunLog":
        self._file = open(self.path, "w", encoding="utf-8-sig", newline="")
        self._writer = csv.DictWriter(self._file, fieldnames=FIELDNAMES)
        self._writer.writeheader()
        self._file.flush()
        return self

    def __exit__(self, *exc) -> bool:
        if self._file:
            self._file.close()
            self._file = None
        return False

    def record(
        self,
        *,
        source: str,
        row: int,
        action: str,
        issue_key: str = "",
        summary: str = "",
        detail: str = "",
    ) -> None:
        """1行分の結果を追記して即座に flush する。"""
        if not self._writer:
            return
        self._writer.writerow({
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "source": source,
            "row": row,
            "action": action,
            "issue_key": issue_key,
            "summary": summary,
            # 改行を含むとログが読みにくいため1行に畳む
            "detail": " ".join(str(detail).split()),
        })
        self._file.flush()
        self.written += 1


def load_completed(path: str | Path) -> set[tuple[str, str]]:
    """
    過去の実行ログを読み、処理済みの (ソース名, 件名) の集合を返す。

    --resume で「作成・更新まで終わっている行」を飛ばすために使う。
    error / skipped の行は含めないため、失敗した行は再実行される。
    """
    log_path = Path(path)
    if not log_path.exists():
        raise FileNotFoundError(f"実行ログが見つかりません: {log_path}")

    completed: set[tuple[str, str]] = set()
    with open(log_path, encoding="utf-8-sig", newline="") as f:
        for record in csv.DictReader(f):
            if record.get("action") in COMPLETED_ACTIONS:
                completed.add((record.get("source", ""), record.get("summary", "")))

    print(f"  再開: 処理済み {len(completed)} 件を読み込みました（{log_path.name}）")
    return completed


def default_log_path(output_dir: Path, timestamp: str) -> Path:
    """実行ログの既定の出力先を返す。"""
    return Path(output_dir) / f"run_{timestamp}.csv"
