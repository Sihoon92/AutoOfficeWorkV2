"""사용자 파일 경로(user_file_paths) 기능 테스트.

{{placeholder}} 플레이스홀더를 사용자 지정 경로로 해소하는 기능을 검증한다.
"""

from __future__ import annotations

import pytest

from autooffice.engine.context import EngineContext


class TestResolveUserFilePaths:
    """EngineContext.resolve()에서 {{placeholder}} 해소 테스트."""

    def test_resolve_simple_placeholder(self):
        """단일 {{placeholder}}가 user_file_paths에서 해소된다."""
        ctx = EngineContext(
            data_dir=".",
            app=None,
            user_file_paths={"raw_data": "/home/user/data/my_raw.xlsx"},
        )
        ctx._app = None  # Excel 앱 초기화 방지
        result = ctx.resolve("{{raw_data}}")
        assert result == "/home/user/data/my_raw.xlsx"

    def test_resolve_placeholder_in_string(self):
        """문자열 내 {{placeholder}}가 해소된다."""
        ctx = EngineContext(
            user_file_paths={"output_dir": "C:/Users/me/결과"},
        )
        result = ctx.resolve("{{output_dir}}/report.xlsx")
        assert result == "C:/Users/me/결과/report.xlsx"

    def test_resolve_multiple_placeholders(self):
        """여러 {{placeholder}}가 한 문자열에서 모두 해소된다."""
        ctx = EngineContext(
            user_file_paths={"dir": "/data", "name": "report"},
        )
        result = ctx.resolve("{{dir}}/{{name}}.xlsx")
        assert result == "/data/report.xlsx"

    def test_resolve_missing_placeholder_raises(self):
        """정의되지 않은 {{placeholder}}는 KeyError를 발생시킨다."""
        ctx = EngineContext(user_file_paths={})
        with pytest.raises(KeyError, match="사용자 파일 경로 'missing'"):
            ctx.resolve("{{missing}}")

    def test_resolve_variable_still_works(self):
        """기존 $variable 해소가 정상 동작한다."""
        ctx = EngineContext(user_file_paths={})
        ctx.store("my_var", [1, 2, 3])
        assert ctx.resolve("$my_var") == [1, 2, 3]

    def test_resolve_plain_string_unchanged(self):
        """플레이스홀더가 없는 문자열은 그대로 반환된다."""
        ctx = EngineContext(user_file_paths={"key": "val"})
        assert ctx.resolve("plain_string.xlsx") == "plain_string.xlsx"

    def test_resolve_non_string_unchanged(self):
        """문자열이 아닌 값은 그대로 반환된다."""
        ctx = EngineContext(user_file_paths={})
        assert ctx.resolve(42) == 42
        assert ctx.resolve(True) is True
        assert ctx.resolve(None) is None

    def test_resolve_params_with_placeholders(self):
        """resolve_params에서 {{placeholder}} 해소가 동작한다."""
        ctx = EngineContext(
            user_file_paths={"input_file": "/data/raw.xlsx", "output_file": "/out/result.xlsx"},
        )
        params = {
            "file_path": "{{input_file}}",
            "alias": "raw",
            "save_as": "{{output_file}}",
        }
        resolved = ctx.resolve_params(params)
        assert resolved["file_path"] == "/data/raw.xlsx"
        assert resolved["alias"] == "raw"
        assert resolved["save_as"] == "/out/result.xlsx"


class TestCollectPlaceholders:
    """CLI _collect_placeholders 함수 테스트."""

    def test_collect_from_plan(self):
        from autooffice.cli import _collect_placeholders
        from autooffice.models.execution_plan import ExecutionPlan

        plan_data = {
            "task_id": "test",
            "description": "test",
            "created_at": "2026-03-25T09:00:00+09:00",
            "inputs": {
                "raw": {"description": "원본", "expected_format": "xlsx"},
            },
            "steps": [
                {
                    "step": 1,
                    "action": "OPEN_FILE",
                    "description": "열기",
                    "params": {"file_path": "{{raw}}", "alias": "raw"},
                },
                {
                    "step": 2,
                    "action": "SAVE_FILE",
                    "description": "저장",
                    "params": {"file": "raw", "save_as": "{{output}}"},
                },
            ],
            "final_output": {"type": "file", "description": "결과"},
        }
        plan = ExecutionPlan.model_validate(plan_data)
        placeholders = _collect_placeholders(plan)
        assert placeholders == {"raw", "output"}

    def test_no_placeholders(self):
        from autooffice.cli import _collect_placeholders
        from autooffice.models.execution_plan import ExecutionPlan

        plan_data = {
            "task_id": "test",
            "description": "test",
            "created_at": "2026-03-25T09:00:00+09:00",
            "inputs": {
                "raw": {"description": "원본", "expected_format": "xlsx"},
            },
            "steps": [
                {
                    "step": 1,
                    "action": "OPEN_FILE",
                    "description": "열기",
                    "params": {"file_path": "raw_data.xlsx", "alias": "raw"},
                },
            ],
            "final_output": {"type": "file", "description": "결과"},
        }
        plan = ExecutionPlan.model_validate(plan_data)
        placeholders = _collect_placeholders(plan)
        assert placeholders == set()


class TestResolveUserFilePathsCLI:
    """CLI _resolve_user_file_paths 함수 테스트."""

    def test_file_option_parsing(self):
        from autooffice.cli import _resolve_user_file_paths
        from autooffice.models.execution_plan import ExecutionPlan

        plan_data = {
            "task_id": "test",
            "description": "test",
            "created_at": "2026-03-25T09:00:00+09:00",
            "inputs": {"raw": {"description": "원본", "expected_format": "xlsx"}},
            "steps": [
                {
                    "step": 1,
                    "action": "OPEN_FILE",
                    "description": "열기",
                    "params": {"file_path": "{{raw}}", "alias": "raw"},
                },
            ],
            "final_output": {"type": "file", "description": "결과"},
        }
        plan = ExecutionPlan.model_validate(plan_data)
        result = _resolve_user_file_paths(plan, ("raw=/my/data.xlsx",))
        assert result["raw"] == "/my/data.xlsx"

    def test_default_file_path_from_inputs(self):
        from autooffice.cli import _resolve_user_file_paths
        from autooffice.models.execution_plan import ExecutionPlan

        plan_data = {
            "task_id": "test",
            "description": "test",
            "created_at": "2026-03-25T09:00:00+09:00",
            "inputs": {
                "raw": {
                    "description": "원본",
                    "expected_format": "xlsx",
                    "file_path": "default_raw.xlsx",
                },
            },
            "steps": [
                {
                    "step": 1,
                    "action": "OPEN_FILE",
                    "description": "열기",
                    "params": {"file_path": "{{raw}}", "alias": "raw"},
                },
            ],
            "final_output": {"type": "file", "description": "결과"},
        }
        plan = ExecutionPlan.model_validate(plan_data)
        result = _resolve_user_file_paths(plan, ())
        assert result["raw"] == "default_raw.xlsx"

    def test_cli_option_overrides_default(self):
        from autooffice.cli import _resolve_user_file_paths
        from autooffice.models.execution_plan import ExecutionPlan

        plan_data = {
            "task_id": "test",
            "description": "test",
            "created_at": "2026-03-25T09:00:00+09:00",
            "inputs": {
                "raw": {
                    "description": "원본",
                    "expected_format": "xlsx",
                    "file_path": "default.xlsx",
                },
            },
            "steps": [
                {
                    "step": 1,
                    "action": "OPEN_FILE",
                    "description": "열기",
                    "params": {"file_path": "{{raw}}", "alias": "raw"},
                },
            ],
            "final_output": {"type": "file", "description": "결과"},
        }
        plan = ExecutionPlan.model_validate(plan_data)
        result = _resolve_user_file_paths(plan, ("raw=/override.xlsx",))
        assert result["raw"] == "/override.xlsx"
