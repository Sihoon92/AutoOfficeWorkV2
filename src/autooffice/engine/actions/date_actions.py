"""날짜 추출 ACTION 핸들러: EXTRACT_DATE.

파일명이나 임의 문자열에서 날짜를 추출하고
파생 정보(주차, 월, 분기 등)를 한번에 계산하여 반환한다.
"""

from __future__ import annotations

import logging
import re
from datetime import date, timedelta
from typing import Any

from autooffice.engine.actions.base import ActionHandler
from autooffice.engine.context import EngineContext
from autooffice.models.action_result import ActionResult

logger = logging.getLogger(__name__)

_WEEKDAY_NAMES = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

# 패턴명 → regex 매핑 (named groups가 아닌 positional: year, month, day)
_PATTERNS: dict[str, str] = {
    "YYYYMMDD": r"(\d{4})(\d{2})(\d{2})",
    "YYYY-MM-DD": r"(\d{4})-(\d{2})-(\d{2})",
    "YYYY_MM_DD": r"(\d{4})_(\d{2})_(\d{2})",
    "YYYY/MM/DD": r"(\d{4})/(\d{2})/(\d{2})",
}

# auto 모드에서 시도할 패턴 순서
_AUTO_ORDER = ["YYYYMMDD", "YYYY-MM-DD", "YYYY_MM_DD", "YYYY/MM/DD"]


def _parse_date(source: str, pattern: str) -> date | None:
    """source에서 pattern에 맞는 날짜를 추출한다."""
    if pattern == "auto":
        for pat_name in _AUTO_ORDER:
            result = _try_parse(source, _PATTERNS[pat_name])
            if result is not None:
                return result
        return None

    regex = _PATTERNS.get(pattern)
    if regex is None:
        return None
    return _try_parse(source, regex)


def _try_parse(source: str, regex: str) -> date | None:
    """regex로 source에서 날짜를 추출 시도한다."""
    m = re.search(regex, source)
    if m is None:
        return None
    try:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def derive_date_info(d: date) -> dict[str, Any]:
    """날짜에서 모든 파생 정보를 계산한다."""
    monday = d - timedelta(days=d.weekday())
    sunday = monday + timedelta(days=6)
    month_start = d.replace(day=1)
    if d.month == 12:
        next_month = date(d.year + 1, 1, 1)
    else:
        next_month = date(d.year, d.month + 1, 1)
    month_end = next_month - timedelta(days=1)

    return {
        "date": d.isoformat(),
        "yyyymmdd": d.strftime("%Y%m%d"),
        "year": d.year,
        "month": d.month,
        "day": d.day,
        "weekday": _WEEKDAY_NAMES[d.weekday()],
        "week_number": d.isocalendar()[1],
        "week_monday": monday.isoformat(),
        "week_sunday": sunday.isoformat(),
        "month_start": month_start.isoformat(),
        "month_end": month_end.isoformat(),
        "quarter": (d.month - 1) // 3 + 1,
    }


class ExtractDateHandler(ActionHandler):
    """EXTRACT_DATE: 문자열에서 날짜를 추출하고 파생 정보를 계산한다.

    params:
        source: 날짜를 추출할 문자열 (파일명, 셀 값 등)
        pattern: 날짜 패턴 (기본 "YYYYMMDD"). "auto"면 여러 패턴 자동 시도.
    """

    def execute(self, params: dict[str, Any], ctx: EngineContext) -> ActionResult:
        source = params.get("source", "")
        pattern = params.get("pattern", "YYYYMMDD")

        if not source:
            return ActionResult(success=False, error="source 파라미터가 비어있습니다.")

        parsed = _parse_date(source, pattern)
        if parsed is None:
            return ActionResult(
                success=False,
                error=f"날짜 추출 실패: '{source}' (패턴: {pattern})",
            )

        data = derive_date_info(parsed)
        logger.info("날짜 추출: '%s' → %s", source, data["date"])
        return ActionResult(
            success=True,
            data=data,
            message=f"날짜 추출 완료: {data['date']} ({data['weekday']}, W{data['week_number']})",
        )
