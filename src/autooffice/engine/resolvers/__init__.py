"""동적 파라미터 해소 패키지.

execution_plan.json의 {{dynamic:key}} 마커를 런타임에 실제 값으로 치환한다.

해소 체인:
  BuiltinResolver (type=date → 로컬) → LLMResolver (type=lookup/text → ChatOpenAI)
"""

from autooffice.engine.resolvers.base import DynamicResolver
from autooffice.engine.resolvers.builtin_resolver import BuiltinDynamicResolver
from autooffice.engine.resolvers.chain import ChainResolver, resolve_plan_dynamic_params
from autooffice.engine.resolvers.substitution import substitute_params

__all__ = [
    "DynamicResolver",
    "BuiltinDynamicResolver",
    "ChainResolver",
    "resolve_plan_dynamic_params",
    "substitute_params",
]
