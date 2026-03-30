"""LLM 호출 없이 내장 함수로 해소 가능한 동적 파라미터 처리.

type=date 등 단순 케이스는 LLM API 없이 로컬에서 즉시 해소한다.
"""

from __future__ import annotations

import logging
from datetime import date
from typing import Any

from autooffice.engine.resolvers.base import DynamicResolver
from autooffice.models.execution_plan import DynamicParamSpec, DynamicParamType

logger = logging.getLogger(__name__)


class BuiltinDynamicResolver(DynamicResolver):
    """내장 해소기: LLM 없이 로컬에서 해소 가능한 파라미터를 처리한다.

    해소 가능한 유형:
    - type=date: date.today().isoformat() 반환

    해소하지 못한 파라미터는 결과에 포함하지 않는다 (체인의 다음 resolver에 위임).
    """

    def resolve(
        self, declarations: dict[str, DynamicParamSpec]
    ) -> dict[str, Any]:
        resolved: dict[str, Any] = {}

        for key, spec in declarations.items():
            if spec.type == DynamicParamType.DATE:
                value = date.today().isoformat()
                resolved[key] = value
                logger.info("내장 해소: %s → %s", key, value)

        return resolved
