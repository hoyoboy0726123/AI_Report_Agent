"""每回合 LLM 呼叫預算追蹤。

Planner（對話）與 Reviewer（VLM）各算各的上限，
避免 reviewer 大量 docx 一次刷掉預算。
"""

from dataclasses import dataclass


@dataclass
class BudgetTracker:
    planner_limit: int = 50
    reviewer_limit: int = 100
    planner_used: int = 0
    reviewer_used: int = 0

    def can_use_planner(self) -> bool:
        return self.planner_used < self.planner_limit

    def can_use_reviewer(self) -> bool:
        return self.reviewer_used < self.reviewer_limit

    def use_planner(self):
        self.planner_used += 1

    def use_reviewer(self):
        self.reviewer_used += 1

    def reset(self):
        self.planner_used = 0
        self.reviewer_used = 0

    def status(self) -> dict:
        return {
            "planner_used": self.planner_used,
            "planner_limit": self.planner_limit,
            "reviewer_used": self.reviewer_used,
            "reviewer_limit": self.reviewer_limit,
        }

    def update_limits(self, planner_limit=None, reviewer_limit=None):
        if planner_limit is not None and planner_limit > 0:
            self.planner_limit = int(planner_limit)
        if reviewer_limit is not None and reviewer_limit > 0:
            self.reviewer_limit = int(reviewer_limit)
