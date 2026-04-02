"""동적 파라미터 해소 패키지.

execution_plan.json의 {{dynamic:key}} 마커를 런타임에 실제 값으로 치환한다.

BuiltinResolver가 type=date를 로컬에서 해소한다 (LLM 불필요).
"""

from autooffice.engine.resolvers.base import DynamicResolver
from autooffice.engine.resolvers.builtin_resolver import BuiltinDynamicResolver
from autooffice.engine.resolvers.chain import resolve_plan_dynamic_params
from autooffice.engine.resolvers.substitution import substitute_params

__all__ = [
    "DynamicResolver",
    "BuiltinDynamicResolver",
    "resolve_plan_dynamic_params",
    "substitute_params",
]
