"""Plan 순차 실행기.

PlanRunner는 execution_plan.json을 로드하여 각 step을 순서대로 실행한다.
각 step은 등록된 ACTION 핸들러에 위임된다.
"""

from __future__ import annotations

import logging
import re
from datetime import datetime, timezone
from typing import TYPE_CHECKING

from autooffice.models.action_result import (
    ActionResult,
    ExecutionLog,
    StepResult,
    StepStatus,
)
from autooffice.models.execution_plan import ExecutionPlan, OnFailAction, Step

if TYPE_CHECKING:
    from autooffice.engine.actions.base import ActionHandler
    from autooffice.engine.context import EngineContext

logger = logging.getLogger(__name__)

# {{dynamic:...}} 미해소 마커 감지 패턴
_UNRESOLVED_DYNAMIC = re.compile(r"\{\{dynamic:\w+(?:\.\w+)?\}\}")


class PlanRunner:
    """execution_plan.json의 step을 순차 실행하는 엔진.

    action_registry에 ACTION 타입별 핸들러를 등록하고,
    plan의 step을 하나씩 실행한다. 각 step은 원자적이며,
    실패 시 on_fail 정책에 따라 처리된다.
    """

    def __init__(self, action_registry: dict[str, ActionHandler]) -> None:
        self.action_registry = action_registry

    def run(self, plan: ExecutionPlan, ctx: EngineContext) -> ExecutionLog:
        """plan의 모든 step을 순차 실행하고 ExecutionLog를 반환한다."""
        log = ExecutionLog(
            task_id=plan.task_id,
            started_at=datetime.now(timezone.utc),
        )

        logger.info("=== 실행 시작: %s ===", plan.task_id)
        logger.info("설명: %s", plan.description)
        logger.info("총 %d개 step", len(plan.steps))

        try:
            for step in plan.steps:
                step_result = self._execute_step(step, ctx)
                log.step_results.append(step_result)

                if step_result.status == StepStatus.FAILED:
                    if step.on_fail == OnFailAction.STOP:
                        logger.error(
                            "Step %d 실패 (STOP): %s",
                            step.step,
                            step_result.result.error,
                        )
                        break
                    elif step.on_fail == OnFailAction.SKIP:
                        logger.warning("Step %d 실패 → SKIP", step.step)
                    elif step.on_fail == OnFailAction.WARN_AND_CONTINUE:
                        logger.warning(
                            "Step %d 실패 → WARN_AND_CONTINUE: %s",
                            step.step,
                            step_result.result.error,
                        )

                # store_as 처리
                if step.store_as and step_result.result.success:
                    ctx.store(step.store_as, step_result.result.data)

            # 최종 결과 판정
            log.success = len(log.failed_steps) == 0
            if not log.success:
                log.error = f"{len(log.failed_steps)}개 step 실패"

        except Exception as e:
            log.success = False
            log.error = str(e)
            logger.exception("실행 중 예외 발생")
        finally:
            log.finished_at = datetime.now(timezone.utc)
            ctx.close_all()

        logger.info("=== 실행 완료: %s ===", log.summary())
        return log

    def validate(self, plan: ExecutionPlan) -> list[str]:
        """plan을 실행하지 않고 논리적 검증만 수행한다 (dry-run).

        Returns:
            검증 오류 메시지 목록 (빈 리스트면 검증 통과)
        """
        errors: list[str] = []

        # 1) 모든 ACTION이 등록되어 있는지 확인
        for step in plan.steps:
            if step.action.value not in self.action_registry:
                errors.append(
                    f"Step {step.step}: 미등록 ACTION '{step.action.value}'"
                )

        # 2) store_as 참조 체인 검증
        defined_vars: set[str] = set()
        for step in plan.steps:
            # params에서 $variable 참조 확인
            for key, val in step.params.items():
                if isinstance(val, str) and val.startswith("$"):
                    var_name = val[1:].split(".", 1)[0]
                    if var_name not in defined_vars:
                        errors.append(
                            f"Step {step.step}: 정의되지 않은 변수 '${var_name}' 참조 "
                            f"(param: {key})"
                        )
            # store_as 등록
            if step.store_as:
                defined_vars.add(step.store_as)

        # 3) step 번호 연속성 확인
        expected = 1
        for step in plan.steps:
            if step.step != expected:
                errors.append(
                    f"Step 번호 불연속: 예상 {expected}, 실제 {step.step}"
                )
            expected += 1

        # 4) 미해소 동적 마커 경고
        for step in plan.steps:
            for key, val in step.params.items():
                if isinstance(val, str) and _UNRESOLVED_DYNAMIC.search(val):
                    errors.append(
                        f"Step {step.step}: 미해소 동적 파라미터 발견 "
                        f"(param: {key}, value: {val}). "
                        f"resolve_plan_dynamic_params()가 실행 전에 호출되어야 합니다."
                    )

        return errors

    def _execute_step(self, step: Step, ctx: EngineContext) -> StepResult:
        """단일 step을 실행하고 StepResult를 반환한다."""
        started = datetime.now(timezone.utc)
        action_name = step.action.value

        logger.info(
            "[Step %d] %s - %s", step.step, action_name, step.description
        )

        # require_confirm 처리
        if step.require_confirm and not ctx.dry_run:
            logger.info("  ⚠ 사용자 확인 필요 (require_confirm=true)")
            # 실제 환경에서는 사용자 입력 대기 로직 구현 필요
            # 현재는 로그만 남기고 진행

        # ACTION 핸들러 조회
        handler = self.action_registry.get(action_name)
        if handler is None:
            result = ActionResult(
                success=False,
                error=f"미등록 ACTION: {action_name}",
            )
        else:
            try:
                resolved_params = ctx.resolve_params(step.params)
                result = handler.execute(resolved_params, ctx)
            except Exception as e:
                result = ActionResult(
                    success=False,
                    error=f"{type(e).__name__}: {e}",
                )

        # expect 검증
        if result.success and step.expect:
            if not self._check_expect(step.expect, result, ctx):
                result = ActionResult(
                    success=False,
                    error=f"expect 조건 불만족: {step.expect.condition}",
                    data=result.data,
                )

        finished = datetime.now(timezone.utc)
        duration_ms = (finished - started).total_seconds() * 1000

        status = StepStatus.SUCCESS if result.success else StepStatus.FAILED
        if not result.success and step.on_fail == OnFailAction.SKIP:
            status = StepStatus.SKIPPED
        elif not result.success and step.on_fail == OnFailAction.WARN_AND_CONTINUE:
            status = StepStatus.WARNED

        return StepResult(
            step=step.step,
            action=action_name,
            status=status,
            result=result,
            started_at=started,
            finished_at=finished,
            duration_ms=duration_ms,
        )

    def _check_expect(self, expect, result: ActionResult, ctx: EngineContext) -> bool:
        """expect 조건을 검증한다."""
        condition = expect.condition
        expected_value = expect.value

        if condition == "row_count_gt":
            if isinstance(result.data, list):
                return len(result.data) > expected_value
            if isinstance(result.data, dict):
                # COPY_RANGE → rows_copied, WRITE_DATA → rows_written 등
                count = (
                    result.data.get("rows_copied")
                    or result.data.get("rows_written")
                    or result.data.get("row_count")
                )
                if count is not None:
                    return count > expected_value
            return False
        if condition == "not_empty":
            return result.data is not None and result.data != []
        if condition == "equals":
            return result.data == expected_value

        # 알 수 없는 조건은 통과 처리 (로그 남김)
        logger.warning("알 수 없는 expect 조건: %s", condition)
        return True
