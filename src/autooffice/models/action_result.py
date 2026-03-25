"""ACTION 실행 결과 모델.

각 step의 실행 결과와 전체 실행 로그를 정의한다.
"""

from __future__ import annotations

from datetime import datetime
from enum import Enum
from typing import Any

from pydantic import BaseModel, Field


class StepStatus(str, Enum):
    """step 실행 상태."""

    SUCCESS = "success"
    FAILED = "failed"
    SKIPPED = "skipped"
    WARNED = "warned"


class ActionResult(BaseModel):
    """단일 ACTION의 실행 결과."""

    success: bool
    data: Any = None
    error: str | None = None
    message: str = ""


class StepResult(BaseModel):
    """단일 step의 실행 결과 (메타 정보 포함)."""

    step: int
    action: str
    status: StepStatus
    result: ActionResult
    started_at: datetime
    finished_at: datetime
    duration_ms: float = 0.0


class ExecutionLog(BaseModel):
    """전체 plan 실행 로그."""

    task_id: str
    started_at: datetime
    finished_at: datetime | None = None
    step_results: list[StepResult] = Field(default_factory=list)
    success: bool = False
    error: str | None = None

    @property
    def total_steps(self) -> int:
        return len(self.step_results)

    @property
    def failed_steps(self) -> list[StepResult]:
        return [s for s in self.step_results if s.status == StepStatus.FAILED]

    def summary(self) -> str:
        """실행 결과 요약 문자열."""
        total = self.total_steps
        success = sum(1 for s in self.step_results if s.status == StepStatus.SUCCESS)
        failed = len(self.failed_steps)
        skipped = sum(1 for s in self.step_results if s.status == StepStatus.SKIPPED)
        return (
            f"실행 완료: {total}개 step 중 "
            f"성공 {success}, 실패 {failed}, 스킵 {skipped}"
        )
