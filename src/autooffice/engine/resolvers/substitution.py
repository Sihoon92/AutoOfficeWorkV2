"""{{dynamic:key}} 패턴 감지 및 치환 유틸리티.

step params 내의 동적 마커를 resolver가 해소한 값으로 치환한다.

마커 형식:
- {{dynamic:key}} — resolved[key] 전체 값으로 치환
- {{dynamic:key.field}} — resolved[key][field]로 치환 (JSON 응답의 개별 필드)
- 부분 문자열: "report_{{dynamic:key}}.xlsx" → "report_2026-03-30.xlsx"
"""

from __future__ import annotations

import logging
import re
from typing import Any

logger = logging.getLogger(__name__)

# {{dynamic:key}} 또는 {{dynamic:key.field}} 매칭
DYNAMIC_PATTERN = re.compile(r"\{\{dynamic:(\w+(?:\.\w+)?)\}\}")


def extract_dynamic_keys(params_list: list[dict[str, Any]]) -> set[str]:
    """여러 step의 params에서 참조된 동적 키(base key)를 수집한다.

    {{dynamic:today_target.column}} → "today_target" 추출.
    """
    keys: set[str] = set()
    for params in params_list:
        _collect_keys_recursive(params, keys)
    return keys


def _collect_keys_recursive(obj: Any, keys: set[str]) -> None:
    """재귀적으로 값을 순회하여 동적 마커의 base key를 수집한다."""
    if isinstance(obj, str):
        for match in DYNAMIC_PATTERN.finditer(obj):
            base_key = match.group(1).split(".", 1)[0]
            keys.add(base_key)
    elif isinstance(obj, dict):
        for v in obj.values():
            _collect_keys_recursive(v, keys)
    elif isinstance(obj, list):
        for item in obj:
            _collect_keys_recursive(item, keys)


def substitute_params(
    params: dict[str, Any],
    resolved: dict[str, Any],
) -> dict[str, Any]:
    """params 내 모든 {{dynamic:...}} 마커를 해소된 값으로 치환한다.

    Args:
        params: step의 params 딕셔너리
        resolved: {key: value} 매핑. value는 str 또는 dict(JSON 응답)

    Returns:
        치환된 새 params 딕셔너리 (원본 불변)
    """
    return _substitute_recursive(params, resolved)


def _substitute_recursive(obj: Any, resolved: dict[str, Any]) -> Any:
    """값을 재귀적으로 순회하며 동적 마커를 치환한다."""
    if isinstance(obj, str):
        return _substitute_string(obj, resolved)
    if isinstance(obj, dict):
        return {k: _substitute_recursive(v, resolved) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_substitute_recursive(item, resolved) for item in obj]
    return obj


def _substitute_string(s: str, resolved: dict[str, Any]) -> Any:
    """문자열 내의 동적 마커를 치환한다.

    - 문자열 전체가 마커 하나인 경우: 해소된 값의 타입을 유지 (int, dict 등)
    - 부분 문자열인 경우: str로 변환하여 삽입
    """
    # 전체 매치 (문자열 = 마커 하나)인 경우 타입 보존
    full_match = DYNAMIC_PATTERN.fullmatch(s)
    if full_match:
        return _resolve_key_path(full_match.group(1), resolved)

    # 부분 매치: 문자열 내 마커를 str로 치환
    def replacer(m: re.Match) -> str:
        val = _resolve_key_path(m.group(1), resolved)
        return str(val)

    result = DYNAMIC_PATTERN.sub(replacer, s)
    return result


def _resolve_key_path(key_path: str, resolved: dict[str, Any]) -> Any:
    """key 또는 key.field 경로를 resolved에서 조회한다.

    - "today_date" → resolved["today_date"]
    - "today_target.column" → resolved["today_target"]["column"]
    """
    parts = key_path.split(".", 1)
    base_key = parts[0]

    if base_key not in resolved:
        logger.warning("동적 파라미터 '%s' 미해소 — 마커 유지", base_key)
        return f"{{{{dynamic:{key_path}}}}}"

    value = resolved[base_key]

    if len(parts) > 1:
        field = parts[1]
        if isinstance(value, dict) and field in value:
            return value[field]
        logger.warning(
            "동적 파라미터 '%s'에서 필드 '%s' 접근 불가", base_key, field
        )
        return f"{{{{dynamic:{key_path}}}}}"

    return value
