"""execution_plan.json의 Pydantic 모델 정의.

Claude가 생성한 execution_plan.json을 파싱하고 검증하는 데 사용한다.
"""

from __future__ import annotations

from datetime import datetime
from enum import Enum
from typing import Any

from pydantic import BaseModel, Field


class ActionType(str, Enum):
    """실행 엔진이 지원하는 ACTION 타입."""

    OPEN_FILE = "OPEN_FILE"
    READ_COLUMNS = "READ_COLUMNS"
    READ_RANGE = "READ_RANGE"
    WRITE_DATA = "WRITE_DATA"
    CLEAR_RANGE = "CLEAR_RANGE"
    RECALCULATE = "RECALCULATE"
    SAVE_FILE = "SAVE_FILE"
    VALIDATE = "VALIDATE"
    FORMAT_MESSAGE = "FORMAT_MESSAGE"
    SEND_MESSENGER = "SEND_MESSENGER"
    SEND_EMAIL = "SEND_EMAIL"
    GENERATE_PPT = "GENERATE_PPT"
    LOG = "LOG"


class OnFailAction(str, Enum):
    """step 실패 시 행동."""

    STOP = "STOP"
    SKIP = "SKIP"
    RETRY = "RETRY"
    WARN_AND_CONTINUE = "WARN_AND_CONTINUE"


class InputSpec(BaseModel):
    """입력 파일/데이터 정의."""

    description: str
    expected_format: str
    file_path: str = Field(default="", description="기본 파일 경로 (사용자가 재지정 가능)")
    expected_sheets: list[str] = Field(default_factory=list)
    expected_columns: list[str] = Field(default_factory=list)


class Expect(BaseModel):
    """step 성공 기준."""

    condition: str = ""
    value: Any = None


class Step(BaseModel):
    """실행 계획의 개별 step."""

    step: int = Field(ge=1, description="step 번호")
    action: ActionType
    description: str
    params: dict[str, Any]
    expect: Expect | None = None
    on_fail: OnFailAction = OnFailAction.STOP
    store_as: str | None = None
    require_confirm: bool = False


class PlanMetadata(BaseModel):
    """plan 메타데이터."""

    task_type: str = ""
    template_hash: str = ""
    reusable: bool = True


class FinalOutput(BaseModel):
    """최종 산출물 정의."""

    type: str
    description: str


class ExecutionPlan(BaseModel):
    """Claude가 생성한 실행 계획서.

    이 모델은 execution_plan.json을 파싱하고,
    실행 엔진(runner)이 순차적으로 실행할 수 있도록 구조화한다.
    """

    task_id: str
    description: str
    created_by: str = "Claude"
    created_at: datetime
    version: str = "1.0.0"
    metadata: PlanMetadata = Field(default_factory=PlanMetadata)
    inputs: dict[str, InputSpec]
    steps: list[Step] = Field(min_length=1)
    final_output: FinalOutput

    @classmethod
    def from_json_file(cls, path: str) -> ExecutionPlan:
        """JSON 파일에서 ExecutionPlan을 로드한다."""
        import json
        from pathlib import Path

        data = json.loads(Path(path).read_text(encoding="utf-8"))
        return cls.model_validate(data)
