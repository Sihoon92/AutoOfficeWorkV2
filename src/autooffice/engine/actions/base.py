"""ACTION 핸들러 추상 기본 클래스."""

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any, TYPE_CHECKING

if TYPE_CHECKING:
    from autooffice.engine.context import EngineContext
    from autooffice.models.action_result import ActionResult


class ActionHandler(ABC):
    """모든 ACTION 핸들러의 기본 인터페이스.

    실행 엔진은 이 인터페이스를 통해 각 ACTION을 호출한다.
    새로운 ACTION을 추가하려면 이 클래스를 상속하고
    execute()를 구현한 뒤 action_registry에 등록한다.
    """

    @abstractmethod
    def execute(self, params: dict[str, Any], ctx: EngineContext) -> ActionResult:
        """ACTION을 실행하고 결과를 반환한다.

        Args:
            params: Claude가 지정한 ACTION 파라미터 (변수 참조 해소 완료)
            ctx: 런타임 컨텍스트

        Returns:
            ActionResult: 실행 결과 (success, data, error, message)
        """
        ...
