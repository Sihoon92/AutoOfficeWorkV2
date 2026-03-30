"""동적 파라미터 해소기 기본 인터페이스."""

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any

from autooffice.models.execution_plan import DynamicParamSpec


class DynamicResolver(ABC):
    """동적 파라미터를 해소하는 resolver의 기본 클래스.

    resolve()는 선언된 파라미터 목록을 받아
    {key: resolved_value} 매핑을 반환한다.
    """

    @abstractmethod
    def resolve(
        self, declarations: dict[str, DynamicParamSpec]
    ) -> dict[str, Any]:
        """동적 파라미터를 해소한다.

        Args:
            declarations: {key: DynamicParamSpec} 선언 목록

        Returns:
            {key: 해소된 값} 매핑. 값은 str, int, dict 등 가능.
        """
        ...
