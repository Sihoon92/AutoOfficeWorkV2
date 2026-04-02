"""BuiltinResolver 신규 DATE_FUNCTIONS 테스트."""

from __future__ import annotations

from datetime import date

from autooffice.engine.resolvers.builtin_resolver import BuiltinDynamicResolver
from autooffice.models.execution_plan import DynamicParamSpec, DynamicParamType


class TestNewDateFunctions:
    """BuiltinResolver: 신규 추가된 str 반환 함수들."""

    def _resolve_one(self, prompt: str) -> str:
        resolver = BuiltinDynamicResolver()
        declarations = {
            "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt=prompt),
        }
        return resolver.resolve(declarations)["d"]

    def test_today_yyyymmdd(self):
        result = self._resolve_one("today_yyyymmdd")
        assert result == date.today().strftime("%Y%m%d")
        assert len(result) == 8
        assert result.isdigit()

    def test_today_yyyy_mm_dd(self):
        result = self._resolve_one("today_yyyy_mm_dd")
        assert result == date.today().strftime("%Y_%m_%d")
        assert "_" in result

    def test_week_number(self):
        result = self._resolve_one("week_number")
        expected = str(date.today().isocalendar()[1])
        assert result == expected

    def test_month_number(self):
        result = self._resolve_one("month_number")
        assert result == str(date.today().month)

    def test_year(self):
        result = self._resolve_one("year")
        assert result == str(date.today().year)

    def test_quarter(self):
        result = self._resolve_one("quarter")
        expected = str((date.today().month - 1) // 3 + 1)
        assert result == expected

    def test_str_functions_return_str_not_date_iso(self):
        """str 반환 함수의 결과가 ISO date 형식이 아님을 확인."""
        result = self._resolve_one("today_yyyymmdd")
        # ISO format은 "YYYY-MM-DD" (하이픈 포함), yyyymmdd는 하이픈 없음
        assert "-" not in result

    def test_existing_functions_still_return_iso(self):
        """기존 date 반환 함수는 여전히 ISO 형식을 반환한다."""
        result = self._resolve_one("today")
        assert result == date.today().isoformat()
        assert "-" in result

    def test_case_insensitive(self):
        """신규 함수도 대소문자 무시."""
        result = self._resolve_one("Today_YYYYMMDD")
        assert result == date.today().strftime("%Y%m%d")

    def test_prompt_is_required(self):
        """prompt 필드 누락 시 Pydantic 검증 에러."""
        import pytest
        from pydantic import ValidationError

        with pytest.raises(ValidationError, match="prompt"):
            DynamicParamSpec(type=DynamicParamType.DATE)
