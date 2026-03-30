"""체인 resolver: Builtin → LLM 순서로 동적 파라미터를 해소한다.

BuiltinResolver가 처리할 수 있는 항목(type=date 등)을 먼저 해소하고,
남은 항목만 LLMResolver에 위임하여 API 비용을 최소화한다.
"""

from __future__ import annotations

import copy
import logging
from typing import Any

from autooffice.engine.resolvers.base import DynamicResolver
from autooffice.engine.resolvers.builtin_resolver import BuiltinDynamicResolver
from autooffice.engine.resolvers.llm_resolver import LLMDynamicResolver
from autooffice.engine.resolvers.substitution import (
    extract_dynamic_keys,
    substitute_params,
)
from autooffice.models.execution_plan import DynamicParamSpec, ExecutionPlan

logger = logging.getLogger(__name__)


class ChainResolver(DynamicResolver):
    """Builtin → LLM 체인으로 동적 파라미터를 해소한다.

    1. BuiltinResolver가 로컬에서 해소 가능한 항목 처리 (API 비용 0)
    2. 남은 항목을 LLMResolver에 위임 (1회 API 호출)
    """

    def __init__(
        self,
        llm_model: str = "gpt-4o-mini",
        llm_api_key: str | None = None,
        llm_base_url: str | None = None,
    ) -> None:
        self._builtin = BuiltinDynamicResolver()
        self._llm = LLMDynamicResolver(
            model_name=llm_model,
            api_key=llm_api_key,
            base_url=llm_base_url,
        )

    def resolve(
        self, declarations: dict[str, DynamicParamSpec]
    ) -> dict[str, Any]:
        if not declarations:
            return {}

        # 1단계: Builtin 해소
        resolved = self._builtin.resolve(declarations)

        # 2단계: 미해소 항목만 LLM에 위임
        remaining = {
            k: v for k, v in declarations.items() if k not in resolved
        }
        if remaining:
            llm_resolved = self._llm.resolve(remaining)
            resolved.update(llm_resolved)

        return resolved


def resolve_plan_dynamic_params(
    plan: ExecutionPlan,
    resolver: DynamicResolver | None = None,
) -> ExecutionPlan:
    """plan의 모든 동적 파라미터를 해소하여 새 plan을 반환한다.

    dynamic_params가 없는 plan은 그대로 반환한다 (하위 호환).

    Args:
        plan: 원본 실행 계획
        resolver: 사용할 resolver (None이면 기본 ChainResolver 생성)

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

    # 해소
    if resolver is None:
        resolver = ChainResolver()
    resolved = resolver.resolve(declarations_to_resolve)

    logger.info("동적 파라미터 해소 결과: %s", resolved)

    # deep copy 후 치환
    plan_copy = plan.model_copy(deep=True)
    for step in plan_copy.steps:
        step.params = substitute_params(step.params, resolved)

    return plan_copy
