"""
更新前の差分判定
================
既存課題は照合の時点で取得済みのため、PATCH を投げる前に
「本当に変わるのか」を比較できる。

以前は実際に更新してみて "No comment content." (code=7) が返るかで
判別していた。そのため確認画面でも、結果的に何も変わらない行まで
可否を尋ねることになっていた。

判定できない項目が 1 つでもあれば「変更あり」に倒す。誤って変更なしと
判断して更新を飛ばすより、余分な PATCH を投げるほうが安全なため。
"""

import excel_to_backlog as etb
from conftest import FakeBacklog
from summary_index import SummaryIndex


def issue(**overrides):
    base = {
        "issueKey": "DEMO-1", "summary": "課題A", "description": "",
        "issueType": {"id": 1}, "priority": {"id": 3},
    }
    base.update(overrides)
    return base


class TestHasChanges:
    def test_同じ内容なら変更なし(self):
        params = {"projectId": 1, "summary": "課題A", "issueTypeId": 1, "priorityId": 3}
        assert etb.has_changes(params, issue()) is False

    def test_projectId_は比較しない(self):
        """更新時には送らない項目のため、差があっても無視する。"""
        params = {"projectId": 999, "summary": "課題A", "issueTypeId": 1, "priorityId": 3}
        assert etb.has_changes(params, issue()) is False

    def test_件名が違えば変更あり(self):
        params = {"summary": "別の件名", "issueTypeId": 1, "priorityId": 3}
        assert etb.has_changes(params, issue()) is True

    def test_件名は正規化して比較する(self):
        """改行やタブの違いだけなら変更なし。"""
        params = {"summary": "課題A", "issueTypeId": 1, "priorityId": 3}
        assert etb.has_changes(params, issue(summary="課題\tA")) is False

    def test_期限日は日付部分だけ比較する(self):
        """Backlog は "2025-01-05T00:00:00Z" 形式で返すことがある。"""
        params = {"dueDate": "2025-01-05", "issueTypeId": 1, "priorityId": 3,
                  "summary": "課題A"}
        assert etb.has_changes(params, issue(dueDate="2025-01-05T00:00:00Z")) is False

    def test_期限日が違えば変更あり(self):
        params = {"dueDate": "2025-03-01", "issueTypeId": 1, "priorityId": 3,
                  "summary": "課題A"}
        assert etb.has_changes(params, issue(dueDate="2025-01-05T00:00:00Z")) is True

    def test_未設定の期限日を設定するのは変更あり(self):
        params = {"dueDate": "2025-01-05", "issueTypeId": 1, "priorityId": 3,
                  "summary": "課題A"}
        assert etb.has_changes(params, issue()) is True

    def test_担当者とステータスは入れ子から取り出して比較する(self):
        params = {"summary": "課題A", "issueTypeId": 1, "priorityId": 3,
                  "assigneeId": 10, "statusId": 4}
        same = issue(assignee={"id": 10}, status={"id": 4})
        assert etb.has_changes(params, same) is False

        params["assigneeId"] = 11
        assert etb.has_changes(params, same) is True

    def test_単一選択のカスタム属性(self):
        params = {"summary": "課題A", "issueTypeId": 1, "priorityId": 3,
                  "customField_5": 51}
        same = issue(customFields=[{"id": 5, "value": {"id": 51, "name": "設計"}}])
        assert etb.has_changes(params, same) is False

        params["customField_5"] = 52
        assert etb.has_changes(params, same) is True

    def test_複数選択のカスタム属性は順序を問わない(self):
        params = {"summary": "課題A", "issueTypeId": 1, "priorityId": 3,
                  "customField_6": [62, 61]}
        same = issue(customFields=[
            {"id": 6, "value": [{"id": 61}, {"id": 62}]},
        ])
        assert etb.has_changes(params, same) is False

    def test_比較できない項目があれば変更ありに倒す(self):
        params = {"summary": "課題A", "issueTypeId": 1, "priorityId": 3,
                  "未知のパラメータ": "x"}
        assert etb.has_changes(params, issue()) is True


class TestPlanUnchanged:
    """process_source が API を呼ばずに「変更なし」を判定すること。"""

    class Backlog(FakeBacklog):
        def get_issues(self, project_id, params=None):
            self.get_issues_calls += 1
            return [
                {"issueKey": "DEMO-1", "summary": "同じ内容",
                 "dueDate": "2025-01-05T00:00:00Z", "description": "",
                 "issueType": {"id": 1}, "priority": {"id": 3}},
                {"issueKey": "DEMO-2", "summary": "変わる課題",
                 "dueDate": "2025-01-01T00:00:00Z", "description": "",
                 "issueType": {"id": 1}, "priority": {"id": 3}},
            ]

    def _cfg(self, source_cfg):
        return source_cfg(
            ["件名", "期限"],
            [["同じ内容", "2025/01/05"], ["変わる課題", "2025/03/01"]],
            issue_mapping={"due_date_col": "期限"},
            upsert={"enabled": True, "match_summary": True},
        )

    def test_内容が同じ行は更新を送らない(self, source_cfg, master):
        backlog = self.Backlog()
        counts = etb.process_source(
            self._cfg(source_cfg), backlog, master, dry_run=False,
            summary_index=SummaryIndex(backlog, master.project_id),
        )

        assert counts["unchanged"] == 1
        assert counts["updated"] == 1
        assert len(backlog.updates) == 1        # 変わる行だけ PATCH

    def test_変更なしの行は確認を求めない(self, source_cfg, master):
        """判断することが無いため、確認画面に出さない。"""
        asked = []

        class Recorder:
            assume_all = False

            def confirm(self, plan, mapper):
                asked.append(plan.action)
                return True

        backlog = self.Backlog()
        etb.process_source(
            self._cfg(source_cfg), backlog, master, dry_run=False,
            summary_index=SummaryIndex(backlog, master.project_id),
            confirmer=Recorder(),
        )

        assert asked == ["update"]              # 変更なしの行は聞かれない

    def test_ドライランでも変更なしを予測する(self, source_cfg, master):
        backlog = self.Backlog()
        dry = etb.process_source(
            self._cfg(source_cfg), backlog, master, dry_run=True,
            summary_index=SummaryIndex(backlog, master.project_id),
        )

        assert dry["unchanged"] == 1
        assert dry["updated"] == 1

    def test_実行ログに変更なしとして残る(self, source_cfg, master, tmp_path):
        import csv

        from run_log import RunLog

        backlog = self.Backlog()
        log_path = tmp_path / "run.csv"
        with RunLog(log_path) as log:
            etb.process_source(
                self._cfg(source_cfg), backlog, master, dry_run=False, run_log=log,
                summary_index=SummaryIndex(backlog, master.project_id),
            )

        with open(log_path, encoding="utf-8-sig", newline="") as f:
            actions = {r["row"]: r["action"] for r in csv.DictReader(f)}
        assert actions == {"2": "unchanged", "3": "updated"}
