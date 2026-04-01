"""Resolver 테스트 (Builtin, Chain, plan 전체 해소)."""

from __future__ import annotations

from datetime import date, timedelta
from unittest.mock import MagicMock, patch

from autooffice.engine.resolvers.builtin_resolver import BuiltinDynamicResolver
from autooffice.engine.resolvers.chain import (
    ChainResolver,
    resolve_plan_dynamic_params,
)
from autooffice.models.execution_plan import (
    DynamicParamSpec,
    DynamicParamType,
    ExecutionPlan,
)


class TestBuiltinDynamicResolver:
    """BuiltinDynamicResolver: type=date를 로컬에서 해소."""

    def test_date_type_resolves_to_today(self):
        """type=date → date.today().isoformat()."""
        resolver = BuiltinDynamicResolver()
        declarations = {
            "today_date": DynamicParamSpec(
                type=DynamicParamType.DATE,
                prompt="오늘 날짜",
                format="YYYY-MM-DD",
            ),
        }
        result = resolver.resolve(declarations)
        assert result["today_date"] == date.today().isoformat()

    def test_lookup_type_not_resolved(self):
        """type=lookup은 builtin에서 해소하지 않음."""
        resolver = BuiltinDynamicResolver()
        declarations = {
            "target": DynamicParamSpec(
                type=DynamicParamType.LOOKUP,
                prompt="Sheet1 위치 계산",
            ),
        }
        result = resolver.resolve(declarations)
        assert "target" not in result

    def test_mixed_types(self):
        """date만 해소, 나머지는 미해소."""
        resolver = BuiltinDynamicResolver()
        declarations = {
            "today": DynamicParamSpec(type=DynamicParamType.DATE, prompt="오늘"),
            "location": DynamicParamSpec(type=DynamicParamType.LOOKUP, prompt="위치"),
            "note": DynamicParamSpec(type=DynamicParamType.TEXT, prompt="메모"),
        }
        result = resolver.resolve(declarations)
        assert "today" in result
        assert "location" not in result
        assert "note" not in result

    def test_prompt_today(self):
        """prompt='today' → 오늘 날짜."""
        resolver = BuiltinDynamicResolver()
        declarations = {
            "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="today"),
        }
        result = resolver.resolve(declarations)
        assert result["d"] == date.today().isoformat()

    def test_prompt_yesterday(self):
        """prompt='yesterday' → 어제 날짜."""
        resolver = BuiltinDynamicResolver()
        declarations = {
            "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="yesterday"),
        }
        result = resolver.resolve(declarations)
        expected = (date.today() - timedelta(days=1)).isoformat()
        assert result["d"] == expected

    def test_prompt_this_week_monday(self):
        """prompt='this_week_monday' → 이번 주 월요일."""
        resolver = BuiltinDynamicResolver()
        declarations = {
            "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="this_week_monday"),
        }
        result = resolver.resolve(declarations)
        today = date.today()
        monday = today - timedelta(days=today.weekday())
        assert result["d"] == monday.isoformat()

    def test_prompt_this_week_sunday(self):
        """prompt='this_week_sunday' → 이번 주 일요일."""
        resolver = BuiltinDynamicResolver()
        declarations = {
            "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="this_week_sunday"),
        }
        result = resolver.resolve(declarations)
        today = date.today()
        sunday = today - timedelta(days=today.weekday()) + timedelta(days=6)
        assert result["d"] == sunday.isoformat()

    def test_prompt_last_week_monday(self):
        """prompt='last_week_monday' → 지난주 월요일."""
        resolver = BuiltinDynamicResolver()
        declarations = {
            "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="last_week_monday"),
        }
        result = resolver.resolve(declarations)
        today = date.today()
        last_monday = today - timedelta(days=today.weekday() + 7)
        assert result["d"] == last_monday.isoformat()

    def test_prompt_last_week_sunday(self):
        """prompt='last_week_sunday' → 지난주 일요일."""
        resolver = BuiltinDynamicResolver()
        declarations = {
            "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="last_week_sunday"),
        }
        result = resolver.resolve(declarations)
        today = date.today()
        last_sunday = today - timedelta(days=today.weekday() + 1)
        assert result["d"] == last_sunday.isoformat()

    def test_prompt_this_month_start(self):
        """prompt='this_month_start' → 이번 달 1일."""
        resolver = BuiltinDynamicResolver()
        declarations = {
            "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="this_month_start"),
        }
        result = resolver.resolve(declarations)
        assert result["d"] == date.today().replace(day=1).isoformat()

    def test_prompt_this_month_end(self):
        """prompt='this_month_end' → 이번 달 마지막 날."""
        resolver = BuiltinDynamicResolver()
        declarations = {
            "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="this_month_end"),
        }
        result = resolver.resolve(declarations)
        today = date.today()
        if today.month == 12:
            next_month_start = date(today.year + 1, 1, 1)
        else:
            next_month_start = date(today.year, today.month + 1, 1)
        expected = (next_month_start - timedelta(days=1)).isoformat()
        assert result["d"] == expected

    def test_prompt_last_month_start(self):
        """prompt='last_month_start' → 지난달 1일."""
        resolver = BuiltinDynamicResolver()
        declarations = {
            "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="last_month_start"),
        }
        result = resolver.resolve(declarations)
        this_month_start = date.today().replace(day=1)
        last_month_end = this_month_start - timedelta(days=1)
        assert result["d"] == last_month_end.replace(day=1).isoformat()

    def test_prompt_last_month_end(self):
        """prompt='last_month_end' → 지난달 마지막 날."""
        resolver = BuiltinDynamicResolver()
        declarations = {
            "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="last_month_end"),
        }
        result = resolver.resolve(declarations)
        this_month_start = date.today().replace(day=1)
        expected = (this_month_start - timedelta(days=1)).isoformat()
        assert result["d"] == expected

    def test_unknown_prompt_falls_back_to_today(self):
        """알 수 없는 prompt → today 대체."""
        resolver = BuiltinDynamicResolver()
        declarations = {
            "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="unknown_keyword"),
        }
        result = resolver.resolve(declarations)
        assert result["d"] == date.today().isoformat()

    def test_prompt_case_insensitive(self):
        """prompt 대소문자 무시."""
        resolver = BuiltinDynamicResolver()
        declarations = {
            "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="This_Week_Monday"),
        }
        result = resolver.resolve(declarations)
        today = date.today()
        monday = today - timedelta(days=today.weekday())
        assert result["d"] == monday.isoformat()


class TestResolvePlanDynamicParams:
    """resolve_plan_dynamic_params: plan 전체 해소 흐름."""

    def _make_plan(self, dynamic_params=None, step_params=None):
        """테스트용 plan 생성 헬퍼."""
        plan_dict = {
            "task_id": "test",
            "description": "test",
            "created_at": "2026-03-30T09:00:00+09:00",
            "metadata": {
                "has_dynamic_params": bool(dynamic_params),
            },
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

    def test_no_dynamic_params_returns_same_plan(self):
        """dynamic_params 없으면 원본 plan 그대로 반환."""
        plan = self._make_plan()
        result = resolve_plan_dynamic_params(plan)
        assert result is plan

    def test_date_param_resolved_without_llm(self):
        """type=date는 LLM 없이 해소된다."""
        plan = self._make_plan(
            dynamic_params={
                "today_date": {
                    "type": "date",
                    "prompt": "오늘 날짜",
                    "format": "YYYY-MM-DD",
                },
            },
            step_params={"date": "{{dynamic:today_date}}"},
        )
        # BuiltinResolver만 사용하는 resolver
        builtin = BuiltinDynamicResolver()
        result = resolve_plan_dynamic_params(plan, resolver=builtin)
        assert result.steps[0].params["date"] == date.today().isoformat()

    def test_original_plan_not_mutated(self):
        """원본 plan은 변경되지 않는다."""
        plan = self._make_plan(
            dynamic_params={
                "today_date": {"type": "date", "prompt": "오늘"},
            },
            step_params={"date": "{{dynamic:today_date}}"},
        )
        original_param = plan.steps[0].params["date"]
        resolve_plan_dynamic_params(plan, resolver=BuiltinDynamicResolver())
        assert plan.steps[0].params["date"] == original_param

    def test_lookup_with_mock_llm(self):
        """type=lookup은 LLM resolver를 통해 해소 (mock)."""
        plan = self._make_plan(
            dynamic_params={
                "today_target": {
                    "type": "lookup",
                    "prompt": "Sheet1 BT열부터 1일1열",
                    "format": "json",
                },
            },
            step_params={
                "target_column": "{{dynamic:today_target.column}}",
                "target_row": "{{dynamic:today_target.date_row}}",
            },
        )

        # Mock resolver: LLM 응답을 시뮬레이션
        mock_resolver = MagicMock()
        mock_resolver.resolve.return_value = {
            "today_target": {
                "date": "2026-03-30",
                "column": "CQ",
                "date_row": 3,
            },
        }

        result = resolve_plan_dynamic_params(plan, resolver=mock_resolver)
        assert result.steps[0].params["target_column"] == "CQ"
        assert result.steps[0].params["target_row"] == 3


class TestRunnerValidateDynamicMarkers:
    """PlanRunner.validate()가 미해소 동적 마커를 감지하는지 확인."""

    def test_unresolved_marker_reported(self):
        """미해소 {{dynamic:...}} 마커가 있으면 validation 오류."""
        from autooffice.engine.runner import PlanRunner

        plan = ExecutionPlan.model_validate({
            "task_id": "test",
            "description": "test",
            "created_at": "2026-03-30T09:00:00+09:00",
            "inputs": {"raw": {"description": "t", "expected_format": "xlsx"}},
            "steps": [
                {
                    "step": 1,
                    "action": "LOG",
                    "description": "test",
                    "params": {"message": "{{dynamic:unresolved_key}}"},
                }
            ],
            "final_output": {"type": "file", "description": "t"},
        })

        runner = PlanRunner(action_registry={"LOG": MagicMock()})
        errors = runner.validate(plan)
        assert any("미해소" in e for e in errors)

    def test_resolved_plan_passes_validation(self):
        """해소된 plan은 validation 통과."""
        from autooffice.engine.runner import PlanRunner

        plan = ExecutionPlan.model_validate({
            "task_id": "test",
            "description": "test",
            "created_at": "2026-03-30T09:00:00+09:00",
            "inputs": {"raw": {"description": "t", "expected_format": "xlsx"}},
            "steps": [
                {
                    "step": 1,
                    "action": "LOG",
                    "description": "test",
                    "params": {"message": "plain text, no markers"},
                }
            ],
            "final_output": {"type": "file", "description": "t"},
        })

        runner = PlanRunner(action_registry={"LOG": MagicMock()})
        errors = runner.validate(plan)
        assert not any("미해소 동적 파라미터" in e for e in errors)
