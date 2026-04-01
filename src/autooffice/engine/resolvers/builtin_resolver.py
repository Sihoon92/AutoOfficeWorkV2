"""LLM 호출 없이 내장 함수로 해소 가능한 동적 파라미터 처리.

type=date는 prompt 키워드에 따라 다양한 날짜를 로컬에서 즉시 해소한다.
"""

from __future__ import annotations

import logging
from collections.abc import Callable
from datetime import date, timedelta
from typing import Any

from autooffice.engine.resolvers.base import DynamicResolver
from autooffice.models.execution_plan import DynamicParamSpec, DynamicParamType

logger = logging.getLogger(__name__)


def _this_week_monday() -> date:
    today = date.today()
    return today - timedelta(days=today.weekday())


def _this_month_start() -> date:
    return date.today().replace(day=1)


def _next_month_start() -> date:
    today = date.today()
    if today.month == 12:
        return date(today.year + 1, 1, 1)
    return date(today.year, today.month + 1, 1)


DATE_FUNCTIONS: dict[str, Callable[[], date]] = {
    "today": lambda: date.today(),
    "yesterday": lambda: date.today() - timedelta(days=1),
    "this_week_monday": _this_week_monday,
    "this_week_sunday": lambda: _this_week_monday() + timedelta(days=6),
    "last_week_monday": lambda: _this_week_monday() - timedelta(days=7),
    "last_week_sunday": lambda: _this_week_monday() - timedelta(days=1),
    "this_month_start": _this_month_start,
    "this_month_end": lambda: _next_month_start() - timedelta(days=1),
    "last_month_start": lambda: (_this_month_start() - timedelta(days=1)).replace(day=1),
    "last_month_end": lambda: _this_month_start() - timedelta(days=1),
}


class BuiltinDynamicResolver(DynamicResolver):
    """내장 해소기: LLM 없이 로컬에서 해소 가능한 파라미터를 처리한다.

    해소 가능한 유형:
    - type=date: prompt 키워드에 따라 날짜 계산
      (today, yesterday, this_week_monday, this_week_sunday,
       last_week_monday, last_week_sunday, this_month_start,
       this_month_end, last_month_start, last_month_end)

    해소하지 못한 파라미터는 결과에 포함하지 않는다 (체인의 다음 resolver에 위임).
    """

    def resolve(
        self, declarations: dict[str, DynamicParamSpec]
    ) -> dict[str, Any]:
        resolved: dict[str, Any] = {}

        for key, spec in declarations.items():
            if spec.type == DynamicParamType.DATE:
                func_key = spec.prompt.strip().lower()
                if func_key in DATE_FUNCTIONS:
                    value = DATE_FUNCTIONS[func_key]().isoformat()
                else:
                    value = date.today().isoformat()
                    logger.warning(
                        "알 수 없는 date prompt '%s', today로 대체", func_key
                    )
                resolved[key] = value
                logger.info(
                    "내장 해소: %s → %s (prompt: %s)", key, value, func_key
                )

        return resolved
