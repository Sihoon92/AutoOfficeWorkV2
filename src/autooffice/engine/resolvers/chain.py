"""동적 파라미터 해소: plan의 {{dynamic:key}} 마커를 실제 값으로 치환한다.

BuiltinResolver가 type=date를 로컬에서 해소한다 (LLM 불필요, API 비용 0).
input 파일명에 날짜가 포함되어 있으면 해당 날짜를 기준으로 해소한다.
"""

from __future__ import annotations

import logging
from datetime import date
from pathlib import Path
from typing import Any

from autooffice.engine.actions.date_actions import _parse_date
from autooffice.engine.resolvers.builtin_resolver import BuiltinDynamicResolver
from autooffice.engine.resolvers.substitution import (
    extract_dynamic_keys,
    substitute_params,
)
from autooffice.models.execution_plan import ExecutionPlan

logger = logging.getLogger(__name__)


def _extract_date_from_inputs(
    input_files: dict[str, str | Path],
) -> date | None:
    """input 파일명에서 날짜를 추출한다.

    모든 input 파일명에서 날짜 추출을 시도한다.
    - 단일 날짜 발견: 해당 날짜 반환
    - 복수의 서로 다른 날짜 발견: 경고 후 None 반환
    - 날짜 없음: None 반환
    """
    found_dates: dict[date, str] = {}  # date → 파일명 (로그용)

    for key, filepath in input_files.items():
        filename = Path(filepath).stem  # 확장자 제거한 파일명
        parsed = _parse_date(filename, "auto")
        if parsed is not None:
            found_dates[parsed] = str(filepath)
            logger.info(
                "input '%s' 파일명에서 날짜 추출: %s → %s",
                key,
                filename,
                parsed.isoformat(),
            )

    if not found_dates:
        return None

    unique_dates = set(found_dates.keys())
    if len(unique_dates) == 1:
        extracted = unique_dates.pop()
        logger.info("input 파일명 기준 날짜: %s", extracted.isoformat())
        return extracted

    # 복수의 서로 다른 날짜 발견
    detail = ", ".join(
        f"{d.isoformat()} ({f})" for d, f in found_dates.items()
    )
    logger.warning(
        "input 파일에서 서로 다른 날짜가 %d개 발견되어 파일명 날짜를 사용하지 않습니다: %s",
        len(unique_dates),
        detail,
    )
    return None


def resolve_plan_dynamic_params(
    plan: ExecutionPlan,
    input_files: dict[str, str | Path] | None = None,
) -> ExecutionPlan:
    """plan의 모든 동적 파라미터를 해소하여 새 plan을 반환한다.

    input_files가 주어지면 파일명에서 날짜 추출을 시도한다.
    파일명에 날짜가 있으면 해당 날짜를 기준으로, 없으면 오늘 날짜를 기준으로 해소한다.

    Args:
        plan: 원본 실행 계획
        input_files: CLI에서 전달된 input 파일 매핑 (KEY → PATH)

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

    # input 파일명에서 날짜 추출 시도
    base_date: date | None = None
    if input_files:
        base_date = _extract_date_from_inputs(input_files)
        if base_date:
            logger.info(
                "파일명 날짜(%s)를 기준으로 동적 파라미터를 해소합니다 "
                "(date.today() 대신)",
                base_date.isoformat(),
            )

    # BuiltinResolver로 해소 (base_date가 None이면 today 사용)
    resolver = BuiltinDynamicResolver(base_date=base_date)
    resolved = resolver.resolve(declarations_to_resolve)

    logger.info("동적 파라미터 해소 결과: %s", resolved)

    # deep copy 후 치환
    plan_copy = plan.model_copy(deep=True)
    for step in plan_copy.steps:
        step.params = substitute_params(step.params, resolved)

    return plan_copy
