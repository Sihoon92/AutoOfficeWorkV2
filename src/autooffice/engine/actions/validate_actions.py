"""검증 ACTION 핸들러: VALIDATE."""

from __future__ import annotations

import logging
from typing import Any

from autooffice.engine.actions.base import ActionHandler
from autooffice.engine.context import EngineContext
from autooffice.models.action_result import ActionResult

logger = logging.getLogger(__name__)


class ValidateHandler(ActionHandler):
    """VALIDATE: 데이터 검증을 수행한다.

    params:
        check: 검증 유형
            - "row_count": 행 수 확인
            - "value_range": 값 범위 확인
            - "not_empty": 비어있지 않은지 확인
            - "sum_equals": 합계 일치 확인
            - "column_exists": 컬럼 존재 확인
        source: 검증 대상 ($변수명 참조)
        logic: 검증 로직 파라미터 (check별로 다름)
            - row_count: {"min": 1, "max": 10000}
            - value_range: {"column": "불량률", "min": 0, "max": 100}
            - sum_equals: {"column": "수량", "expected": "$total_expected"}
    """

    def execute(self, params: dict[str, Any], ctx: EngineContext) -> ActionResult:
        check = params.get("check", "")
        source = params.get("source")
        logic = params.get("logic", {})

        validators = {
            "row_count": self._check_row_count,
            "value_range": self._check_value_range,
            "not_empty": self._check_not_empty,
            "sum_equals": self._check_sum_equals,
            "column_exists": self._check_column_exists,
        }

        validator = validators.get(check)
        if validator is None:
            return ActionResult(
                success=False,
                error=f"알 수 없는 검증 유형: {check}. 사용 가능: {list(validators.keys())}",
            )

        return validator(source, logic)

    def _check_row_count(self, source: Any, logic: dict) -> ActionResult:
        if not isinstance(source, list):
            return ActionResult(success=False, error="source가 리스트가 아닙니다.")

        count = len(source)
        min_val = logic.get("min", 0)
        max_val = logic.get("max", float("inf"))

        if min_val <= count <= max_val:
            return ActionResult(
                success=True,
                data=count,
                message=f"행 수 검증 통과: {count}행 (범위: {min_val}~{max_val})",
            )
        return ActionResult(
            success=False,
            error=f"행 수 범위 초과: {count}행 (허용: {min_val}~{max_val})",
        )

    def _check_value_range(self, source: Any, logic: dict) -> ActionResult:
        if not isinstance(source, list):
            return ActionResult(success=False, error="source가 리스트가 아닙니다.")

        column = logic.get("column", "")
        min_val = logic.get("min", float("-inf"))
        max_val = logic.get("max", float("inf"))

        violations = []
        for i, row in enumerate(source):
            val = row.get(column) if isinstance(row, dict) else None
            if val is not None and isinstance(val, (int, float)):
                if not (min_val <= val <= max_val):
                    violations.append({"row": i + 1, "value": val})

        if not violations:
            return ActionResult(
                success=True,
                message=f"값 범위 검증 통과: {column} ({min_val}~{max_val})",
            )
        return ActionResult(
            success=False,
            data=violations,
            error=f"값 범위 위반 {len(violations)}건: {column}",
        )

    def _check_not_empty(self, source: Any, logic: dict) -> ActionResult:
        if source is None or (isinstance(source, list) and len(source) == 0):
            return ActionResult(success=False, error="데이터가 비어있습니다.")
        return ActionResult(success=True, message="비어있지 않음 확인")

    def _check_sum_equals(self, source: Any, logic: dict) -> ActionResult:
        if not isinstance(source, list):
            return ActionResult(success=False, error="source가 리스트가 아닙니다.")

        column = logic.get("column", "")
        expected = logic.get("expected", 0)

        total = sum(
            row.get(column, 0)
            for row in source
            if isinstance(row, dict) and isinstance(row.get(column), (int, float))
        )

        if abs(total - expected) < 0.001:
            return ActionResult(
                success=True,
                data=total,
                message=f"합계 일치: {column} = {total}",
            )
        return ActionResult(
            success=False,
            error=f"합계 불일치: {column} 실제={total}, 기대={expected}",
        )

    def _check_column_exists(self, source: Any, logic: dict) -> ActionResult:
        if not isinstance(source, list) or len(source) == 0:
            return ActionResult(success=False, error="source가 비어있습니다.")

        required = logic.get("columns", [])
        first_row = source[0]
        if not isinstance(first_row, dict):
            return ActionResult(success=False, error="source 행이 dict가 아닙니다.")

        missing = [col for col in required if col not in first_row]
        if not missing:
            return ActionResult(
                success=True,
                message=f"필수 컬럼 존재 확인: {required}",
            )
        return ActionResult(
            success=False,
            error=f"누락된 컬럼: {missing}",
        )
