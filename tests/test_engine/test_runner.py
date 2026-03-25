"""PlanRunner 테스트."""

from __future__ import annotations

from autooffice.models.execution_plan import ExecutionPlan


class TestPlanValidation:
    """plan의 논리적 검증 테스트."""

    def test_validate_valid_plan(self, runner, sample_plan_dict):
        plan = ExecutionPlan.model_validate(sample_plan_dict)
        errors = runner.validate(plan)
        assert errors == [], f"검증 오류 발생: {errors}"

    def test_validate_undefined_variable_reference(self, runner, sample_plan_dict):
        """정의되지 않은 변수 참조 탐지."""
        sample_plan_dict["steps"][4]["params"]["source"] = "$undefined_var"
        # store_as "raw_data"가 step 3에 있지만 "$undefined_var"는 없음
        plan = ExecutionPlan.model_validate(sample_plan_dict)
        errors = runner.validate(plan)
        assert any("$undefined_var" in e for e in errors)

    def test_validate_step_number_continuity(self, runner, sample_plan_dict):
        """step 번호 연속성 검증."""
        sample_plan_dict["steps"][2]["step"] = 99  # 3 대신 99
        plan = ExecutionPlan.model_validate(sample_plan_dict)
        errors = runner.validate(plan)
        assert any("불연속" in e for e in errors)


class TestPlanExecution:
    """plan 실행 통합 테스트."""

    def test_run_full_plan(
        self, runner, engine_ctx, sample_plan_dict, sample_raw_excel, sample_template_excel
    ):
        """전체 plan을 실행하고 성공 여부 확인."""
        plan = ExecutionPlan.model_validate(sample_plan_dict)
        log = runner.run(plan, engine_ctx)

        assert log.success, f"실행 실패: {log.error}"
        assert log.total_steps == 7
        assert len(log.failed_steps) == 0

    def test_run_stores_variables(
        self, runner, engine_ctx, sample_plan_dict, sample_raw_excel, sample_template_excel
    ):
        """store_as로 변수가 올바르게 저장되는지 확인."""
        plan = ExecutionPlan.model_validate(sample_plan_dict)
        runner.run(plan, engine_ctx)

        assert "raw_data" in engine_ctx.variables
        raw_data = engine_ctx.variables["raw_data"]
        assert isinstance(raw_data, list)
        assert len(raw_data) == 5  # 더미 데이터 5행

    def test_run_creates_result_file(
        self, runner, engine_ctx, sample_plan_dict, sample_raw_excel, sample_template_excel
    ):
        """결과 파일이 생성되는지 확인."""
        plan = ExecutionPlan.model_validate(sample_plan_dict)
        runner.run(plan, engine_ctx)

        result_path = engine_ctx.data_dir / "result.xlsx"
        assert result_path.exists()

    def test_run_stops_on_fail(self, runner, engine_ctx, sample_plan_dict):
        """STOP on_fail 시 실행이 중단되는지 확인 (파일 없는 경우)."""
        plan = ExecutionPlan.model_validate(sample_plan_dict)
        log = runner.run(plan, engine_ctx)

        # raw_data.xlsx 파일이 없으므로 step 1에서 실패
        assert not log.success
        assert log.step_results[0].status.value == "failed"
        # STOP이므로 이후 step은 실행되지 않음
        assert log.total_steps == 1
