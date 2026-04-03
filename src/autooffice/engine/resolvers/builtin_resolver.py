"""LLM 호출 없이 내장 함수로 해소 가능한 동적 파라미터 처리.

type=date는 prompt 키워드에 따라 다양한 날짜를 로컬에서 즉시 해소한다.
base_date가 주입되면 date.today() 대신 해당 날짜를 기준으로 계산한다.
"""

from __future__ import annotations

import logging
from collections.abc import Callable
from datetime import date, timedelta
from typing import Any

from autooffice.engine.resolvers.base import DynamicResolver
from autooffice.models.execution_plan import DynamicParamSpec, DynamicParamType

logger = logging.getLogger(__name__)


def _build_date_functions(base: date) -> dict[str, Callable[[], date | str]]:
    """base 날짜를 기준으로 DATE_FUNCTIONS 매핑을 생성한다."""
    monday = base - timedelta(days=base.weekday())
    month_start = base.replace(day=1)
    if base.month == 12:
        next_month_start = date(base.year + 1, 1, 1)
    else:
        next_month_start = date(base.year, base.month + 1, 1)
    last_month_end = month_start - timedelta(days=1)
    last_month_start = last_month_end.replace(day=1)

    return {
        # ── date 객체 반환 ──
        "today": lambda: base,
        "yesterday": lambda: base - timedelta(days=1),
        "this_week_monday": lambda: monday,
        "this_week_sunday": lambda: monday + timedelta(days=6),
        "last_week_monday": lambda: monday - timedelta(days=7),
        "last_week_sunday": lambda: monday - timedelta(days=1),
        "this_month_start": lambda: month_start,
        "this_month_end": lambda: next_month_start - timedelta(days=1),
        "last_month_start": lambda: last_month_start,
        "last_month_end": lambda: last_month_end,
        # ── str 반환 (파일명 생성, 라벨 매칭용) ──
        "today_yyyymmdd": lambda: base.strftime("%Y%m%d"),
        "today_yyyy_mm_dd": lambda: base.strftime("%Y_%m_%d"),
        "week_number": lambda: str(base.isocalendar()[1]),
        "month_number": lambda: str(base.month),
        "year": lambda: str(base.year),
        "quarter": lambda: str((base.month - 1) // 3 + 1),
    }


# 하위 호환: 기존 테스트에서 DATE_FUNCTIONS를 직접 참조하는 경우
DATE_FUNCTIONS: dict[str, Callable[[], date | str]] = _build_date_functions(
    date.today()
)


class BuiltinDynamicResolver(DynamicResolver):
    """내장 해소기: LLM 없이 로컬에서 해소 가능한 파라미터를 처리한다.

    Args:
        base_date: 기준 날짜. None이면 date.today() 사용.
                   input 파일명에서 추출한 날짜를 주입하면
                   해당 날짜 기준으로 모든 날짜를 계산한다.

    해소 가능한 유형:
    - type=date: prompt 키워드에 따라 날짜 계산
      (today, yesterday, this_week_monday, this_week_sunday,
       last_week_monday, last_week_sunday, this_month_start,
       this_month_end, last_month_start, last_month_end)

    해소하지 못한 파라미터는 결과에 포함하지 않는다 (체인의 다음 resolver에 위임).
    """

    def __init__(self, base_date: date | None = None) -> None:
        self._base_date = base_date or date.today()
        self._functions = _build_date_functions(self._base_date)

    def resolve(
        self, declarations: dict[str, DynamicParamSpec]
    ) -> dict[str, Any]:
        resolved: dict[str, Any] = {}

        for key, spec in declarations.items():
            if spec.type == DynamicParamType.DATE:
                func_key = spec.prompt.strip().lower()
                if func_key in self._functions:
                    raw = self._functions[func_key]()
                else:
                    raw = self._base_date
                    logger.warning(
                        "알 수 없는 date prompt '%s', base_date(%s)로 대체",
                        func_key,
                        self._base_date,
                    )
                # date 객체 → ISO string, str은 그대로
                value = raw.isoformat() if isinstance(raw, date) else raw
                resolved[key] = value
                logger.info(
                    "내장 해소: %s → %s (prompt: %s, base: %s)",
                    key,
                    value,
                    func_key,
                    self._base_date,
                )

        return resolved
