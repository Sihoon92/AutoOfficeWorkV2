"""동적 파라미터 해소: plan의 {{dynamic:key}} 마커를 실제 값으로 치환한다.

BuiltinResolver가 type=date를 로컬에서 해소한다 (LLM 불필요, API 비용 0).
"""

from __future__ import annotations

import copy
import logging
from typing import Any

from autooffice.engine.resolvers.builtin_resolver import BuiltinDynamicResolver
from autooffice.engine.resolvers.substitution import (
    extract_dynamic_keys,
    substitute_params,
)
from autooffice.models.execution_plan import DynamicParamSpec, ExecutionPlan

logger = logging.getLogger(__name__)


def resolve_plan_dynamic_params(
    plan: ExecutionPlan,
) -> ExecutionPlan:
    """plan의 모든 동적 파라미터를 해소하여 새 plan을 반환한다.

    dynamic_params가 없는 plan은 그대로 반환한다 (하위 호환).

    Args:
        plan: 원본 실행 계획

    Returns:
        동적 마커가 치환된 새 ExecutionPlan
    """
    if not plan.dynamic_params:
        return plan

    # 실제 step에서 참조된 키만 필터링
    referenced_keys = extract_dynamic_keys(
        [step.params for step in plan.steps]
    )
    declarations_to_resolve = {
        k: v
        for k, v in plan.dynamic_params.items()
        if k in referenced_keys
    }

    if not declarations_to_resolve:
        logger.info("동적 파라미터 선언 있으나 step에서 참조 없음 — 건너뜀")
        return plan

    # BuiltinResolver로 해소
    resolver = BuiltinDynamicResolver()
    resolved = resolver.resolve(declarations_to_resolve)

    logger.info("동적 파라미터 해소 결과: %s", resolved)

    # deep copy 후 치환
    plan_copy = plan.model_copy(deep=True)
    for step in plan_copy.steps:
        step.params = substitute_params(step.params, resolved)

    return plan_copy
