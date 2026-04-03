"""input 파일명 날짜 추출 및 dynamic_params 덮어쓰기 테스트."""

from __future__ import annotations

from datetime import date, timedelta

import pytest

from autooffice.engine.resolvers.builtin_resolver import BuiltinDynamicResolver
from autooffice.engine.resolvers.chain import (
    _extract_date_from_inputs,
    resolve_plan_dynamic_params,
)
from autooffice.models.execution_plan import (
    DynamicParamSpec,
    DynamicParamType,
    ExecutionPlan,
)


# ── _extract_date_from_inputs 테스트 ──


class TestExtractDateFromInputs:
    """input 파일명에서 날짜 추출 로직."""

    def test_yyyymmdd_pattern(self):
        """YYYYMMDD 패턴 파일명에서 날짜 추출."""
        inputs = {"raw_data": "20260404_생산실적.xlsx"}
        result = _extract_date_from_inputs(inputs)
        assert result == date(2026, 4, 4)

    def test_yyyy_mm_dd_pattern(self):
        """YYYY-MM-DD 패턴 파일명에서 날짜 추출."""
        inputs = {"raw_data": "생산실적_2026-04-04.xlsx"}
        result = _extract_date_from_inputs(inputs)
        assert result == date(2026, 4, 4)

    def test_yyyy_underscore_mm_dd_pattern(self):
        """YYYY_MM_DD 패턴 파일명에서 날짜 추출."""
        inputs = {"raw_data": "report_2026_04_04.xlsx"}
        result = _extract_date_from_inputs(inputs)
        assert result == date(2026, 4, 4)

    def test_no_date_in_filename(self):
        """날짜 없는 파일명 → None."""
        inputs = {"raw_data": "기준표.xlsx"}
        result = _extract_date_from_inputs(inputs)
        assert result is None

    def test_multiple_inputs_same_date(self):
        """여러 input이 같은 날짜 → 해당 날짜 반환."""
        inputs = {
            "raw_data": "20260404_생산.xlsx",
            "raw_data2": "20260404_품질.xlsx",
        }
        result = _extract_date_from_inputs(inputs)
        assert result == date(2026, 4, 4)

    def test_multiple_inputs_different_dates_returns_none(self):
        """여러 input이 다른 날짜 → 경고 후 None."""
        inputs = {
            "raw_data": "20260404_생산.xlsx",
            "raw_data2": "20260405_품질.xlsx",
        }
        result = _extract_date_from_inputs(inputs)
        assert result is None

    def test_mixed_date_and_no_date(self):
        """날짜 있는 파일 + 없는 파일 → 날짜 있는 것만 사용."""
        inputs = {
            "raw_data": "20260404_생산.xlsx",
            "ref": "기준표.xlsx",
        }
        result = _extract_date_from_inputs(inputs)
        assert result == date(2026, 4, 4)

    def test_path_with_directory(self):
        """경로 포함 파일명에서 파일명만 추출하여 날짜 추출."""
        inputs = {"raw_data": "/data/raw/20260404_생산실적.xlsx"}
        result = _extract_date_from_inputs(inputs)
        assert result == date(2026, 4, 4)

    def test_empty_inputs(self):
        """빈 input → None."""
        result = _extract_date_from_inputs({})
        assert result is None


# ── BuiltinDynamicResolver base_date 테스트 ──


class TestBuiltinResolverBaseDate:
    """BuiltinDynamicResolver에 base_date 주입 테스트."""

    def _resolve_one(self, prompt: str, base_date: date) -> str:
        resolver = BuiltinDynamicResolver(base_date=base_date)
        declarations = {
            "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt=prompt),
        }
        return resolver.resolve(declarations)["d"]

    def test_today_uses_base_date(self):
        """base_date 주입 시 today가 base_date를 반환."""
        base = date(2026, 4, 4)
        result = self._resolve_one("today", base)
        assert result == "2026-04-04"

    def test_yesterday_uses_base_date(self):
        """base_date 기준 yesterday."""
        base = date(2026, 4, 4)
        result = self._resolve_one("yesterday", base)
        assert result == "2026-04-03"

    def test_this_week_monday_uses_base_date(self):
        """base_date 기준 this_week_monday."""
        # 2026-04-04 = 토요일 → 월요일 = 2026-03-30
        base = date(2026, 4, 4)
        result = self._resolve_one("this_week_monday", base)
        assert result == "2026-03-30"

    def test_this_month_start_uses_base_date(self):
        """base_date 기준 this_month_start."""
        base = date(2026, 4, 15)
        result = self._resolve_one("this_month_start", base)
        assert result == "2026-04-01"

    def test_this_month_end_uses_base_date(self):
        """base_date 기준 this_month_end."""
        base = date(2026, 4, 15)
        result = self._resolve_one("this_month_end", base)
        assert result == "2026-04-30"

    def test_today_yyyymmdd_uses_base_date(self):
        """base_date 기준 today_yyyymmdd (str 반환)."""
        base = date(2026, 4, 4)
        result = self._resolve_one("today_yyyymmdd", base)
        assert result == "20260404"

    def test_week_number_uses_base_date(self):
        """base_date 기준 week_number."""
        base = date(2026, 4, 4)
        result = self._resolve_one("week_number", base)
        assert result == str(base.isocalendar()[1])

    def test_month_number_uses_base_date(self):
        """base_date 기준 month_number."""
        base = date(2026, 4, 4)
        result = self._resolve_one("month_number", base)
        assert result == "4"

    def test_quarter_uses_base_date(self):
        """base_date 기준 quarter."""
        base = date(2026, 4, 4)
        result = self._resolve_one("quarter", base)
        assert result == "2"

    def test_unknown_prompt_falls_back_to_base_date(self):
        """알 수 없는 prompt → base_date로 대체 (today가 아닌)."""
        base = date(2026, 4, 4)
        result = self._resolve_one("unknown_keyword", base)
        assert result == "2026-04-04"

    def test_none_base_date_uses_today(self):
        """base_date=None → date.today() 사용."""
        resolver = BuiltinDynamicResolver(base_date=None)
        declarations = {
            "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="today"),
        }
        result = resolver.resolve(declarations)
        assert result["d"] == date.today().isoformat()


# ── resolve_plan_dynamic_params + input_files 통합 테스트 ──


class TestResolvePlanWithInputFiles:
    """resolve_plan_dynamic_params에 input_files를 전달한 통합 테스트."""

    def _make_plan(self, dynamic_params=None, step_params=None):
        plan_dict = {
            "task_id": "test",
            "description": "test",
            "created_at": "2026-03-30T09:00:00+09:00",
            "metadata": {"has_dynamic_params": bool(dynamic_params)},
            "dynamic_params": dynamic_params or {},
            "inputs": {"raw": {"description": "test", "expected_format": "xlsx"}},
            "steps": [
                {
                    "step": 1,
                    "action": "COPY_RANGE",
                    "description": "test step",
                    "params": step_params or {"file": "test"},
                }
            ],
            "final_output": {"type": "file", "description": "test"},
        }
        return ExecutionPlan.model_validate(plan_dict)

    def test_input_file_date_overrides_today(self):
        """input 파일명 날짜가 dynamic_params의 today를 덮어쓴다."""
        plan = self._make_plan(
            dynamic_params={
                "exec_date": {"type": "date", "prompt": "today"},
            },
            step_params={"date": "{{dynamic:exec_date}}"},
        )
        result = resolve_plan_dynamic_params(
            plan,
            input_files={"raw_data": "20260404_생산실적.xlsx"},
        )
        assert result.steps[0].params["date"] == "2026-04-04"

    def test_input_file_date_affects_derived_dates(self):
        """파일명 날짜가 파생 날짜(week_monday 등)에도 반영된다."""
        plan = self._make_plan(
            dynamic_params={
                "exec_date": {"type": "date", "prompt": "today"},
                "week_start": {"type": "date", "prompt": "this_week_monday"},
            },
            step_params={
                "date": "{{dynamic:exec_date}}",
                "week": "{{dynamic:week_start}}",
            },
        )
        result = resolve_plan_dynamic_params(
            plan,
            input_files={"raw_data": "20260404_생산실적.xlsx"},
        )
        # 2026-04-04 = 토요일 → 월요일 = 2026-03-30
        assert result.steps[0].params["date"] == "2026-04-04"
        assert result.steps[0].params["week"] == "2026-03-30"

    def test_no_date_in_filename_uses_today(self):
        """파일명에 날짜 없으면 기존대로 today 사용."""
        plan = self._make_plan(
            dynamic_params={
                "exec_date": {"type": "date", "prompt": "today"},
            },
            step_params={"date": "{{dynamic:exec_date}}"},
        )
        result = resolve_plan_dynamic_params(
            plan,
            input_files={"ref": "기준표.xlsx"},
        )
        assert result.steps[0].params["date"] == date.today().isoformat()

    def test_no_input_files_uses_today(self):
        """input_files=None이면 기존대로 today 사용."""
        plan = self._make_plan(
            dynamic_params={
                "exec_date": {"type": "date", "prompt": "today"},
            },
            step_params={"date": "{{dynamic:exec_date}}"},
        )
        result = resolve_plan_dynamic_params(plan)
        assert result.steps[0].params["date"] == date.today().isoformat()

    def test_conflicting_dates_fall_back_to_today(self):
        """서로 다른 날짜 파일 → today로 폴백."""
        plan = self._make_plan(
            dynamic_params={
                "exec_date": {"type": "date", "prompt": "today"},
            },
            step_params={"date": "{{dynamic:exec_date}}"},
        )
        result = resolve_plan_dynamic_params(
            plan,
            input_files={
                "raw1": "20260404_생산.xlsx",
                "raw2": "20260405_품질.xlsx",
            },
        )
        assert result.steps[0].params["date"] == date.today().isoformat()
