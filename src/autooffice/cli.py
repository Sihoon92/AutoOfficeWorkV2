"""AutoOffice CLI - 실행 계획 실행 및 관리.

사용법:
    autooffice run plan.json --data ./data/ --input raw_data=file.xlsx
    autooffice validate plan.json
    autooffice cache list
    autooffice cache run <plan_id> --data ./data/ --input raw_data=file.xlsx
"""

from __future__ import annotations

import json
import logging
import sys
from pathlib import Path

import click

from autooffice.engine.actions import build_default_registry
from autooffice.engine.context import EngineContext
from autooffice.engine.runner import PlanRunner
from autooffice.models.execution_plan import ExecutionPlan, InputSpec


def _parse_inputs(raw_inputs: tuple[str, ...]) -> dict[str, str]:
    """'KEY=VALUE' 형식의 --input 인자를 dict로 변환한다."""
    result: dict[str, str] = {}
    for item in raw_inputs:
        if "=" not in item:
            raise click.BadParameter(
                f"입력 형식 오류: '{item}' (KEY=PATH 형식 필요)",
                param_hint="--input",
            )
        key, value = item.split("=", 1)
        result[key.strip()] = value.strip()
    return result


def _validate_inputs(
    plan_inputs: dict[str, InputSpec],
    input_map: dict[str, str],
    data_dir: Path,
) -> None:
    """plan의 inputs 정의와 실제 입력을 대조 검증한다."""
    missing = set(plan_inputs.keys()) - set(input_map.keys())
    if missing:
        hints = ", ".join(f"--input {k}=파일경로" for k in sorted(missing))
        raise click.UsageError(f"필수 입력 파일 누락: {', '.join(sorted(missing))}\n  사용법: {hints}")

    for key, filename in input_map.items():
        path = data_dir / filename
        if not path.exists():
            raise click.FileError(str(path), hint=f"--input {key}={filename}")


@click.group()
@click.option("--verbose", "-v", is_flag=True, help="상세 로그 출력")
def main(verbose: bool) -> None:
    """AutoOffice - Claude 사고 프로세스 기반 사무 업무 자동화 실행 엔진"""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )


@main.command()
@click.argument("plan_file", type=click.Path(exists=True))
@click.option("--data", "-d", type=click.Path(exists=True), default=".", help="데이터 디렉토리")
@click.option("--input", "-i", "inputs", multiple=True, help="입력 파일 매핑 (KEY=PATH, 반복 가능)")
@click.option("--no-resolve", is_flag=True, default=False, help="동적 파라미터 해소 건너뛰기")
def run(plan_file: str, data: str, inputs: tuple[str, ...], no_resolve: bool) -> None:
    """execution_plan.json을 실행한다."""
    click.echo(f"plan 로드: {plan_file}")

    try:
        plan = ExecutionPlan.from_json_file(plan_file)
    except Exception as e:
        click.echo(f"plan 파싱 실패: {e}", err=True)
        sys.exit(1)

    # 입력 파일 매핑 처리
    data_dir = Path(data)
    input_map = _parse_inputs(inputs)
    if plan.inputs:
        _validate_inputs(plan.inputs, input_map, data_dir)

    # 동적 파라미터 해소 (BuiltinResolver, LLM 불필요)
    if not no_resolve and plan.dynamic_params:
        from autooffice.engine.resolvers.chain import resolve_plan_dynamic_params

        click.echo(f"동적 파라미터 해소 중... ({len(plan.dynamic_params)}개)")
        plan = resolve_plan_dynamic_params(plan, input_files=input_map)
        click.echo("동적 파라미터 해소 완료")

    runner = PlanRunner(build_default_registry())
    ctx = EngineContext(data_dir=data, input_files=input_map)

    # 먼저 검증
    errors = runner.validate(plan)
    if errors:
        click.echo("plan 검증 실패:")
        for err in errors:
            click.echo(f"  - {err}")
        sys.exit(1)

    # 실행
    log = runner.run(plan, ctx)

    click.echo("")
    click.echo(log.summary())
    if not log.success:
        click.echo(f"오류: {log.error}")
        for sr in log.failed_steps:
            click.echo(f"  Step {sr.step} ({sr.action}): {sr.result.error}")
        sys.exit(1)


@main.command()
@click.argument("plan_file", type=click.Path(exists=True))
def validate(plan_file: str) -> None:
    """execution_plan.json을 실행하지 않고 검증만 한다 (dry-run)."""
    click.echo(f"plan 검증: {plan_file}")

    try:
        plan = ExecutionPlan.from_json_file(plan_file)
    except Exception as e:
        click.echo(f"plan 파싱 실패: {e}", err=True)
        sys.exit(1)

    runner = PlanRunner(build_default_registry())
    errors = runner.validate(plan)

    if errors:
        click.echo(f"검증 실패 ({len(errors)}건):")
        for err in errors:
            click.echo(f"  - {err}")
        sys.exit(1)
    else:
        click.echo("검증 통과! plan 구조가 올바릅니다.")
        click.echo(f"  task_id: {plan.task_id}")
        click.echo(f"  steps: {len(plan.steps)}개")
        click.echo(f"  최종 산출물: {plan.final_output.type} - {plan.final_output.description}")


@main.group()
def cache() -> None:
    """캐시된 plan 관리."""
    pass


@cache.command("list")
def cache_list() -> None:
    """캐시된 plan 목록을 출력한다."""
    from autooffice.cache.plan_cache import PlanCache

    pc = PlanCache()
    plans = pc.list_plans()

    if not plans:
        click.echo("캐시된 plan이 없습니다.")
        return

    click.echo(f"캐시된 plan {len(plans)}개:")
    for entry in plans:
        click.echo(
            f"  [{entry['task_type']}] {entry['task_id']} "
            f"(template: {entry['template_hash'][:8]}..., "
            f"생성: {entry['created_at']})"
        )


@cache.command("run")
@click.argument("plan_id")
@click.option("--data", "-d", type=click.Path(exists=True), default=".", help="데이터 디렉토리")
@click.option("--input", "-i", "inputs", multiple=True, help="입력 파일 매핑 (KEY=PATH, 반복 가능)")
def cache_run(plan_id: str, data: str, inputs: tuple[str, ...]) -> None:
    """캐시된 plan을 실행한다."""
    from autooffice.cache.plan_cache import PlanCache

    pc = PlanCache()
    plan = pc.load_plan(plan_id)

    if plan is None:
        click.echo(f"캐시에서 plan '{plan_id}'을 찾을 수 없습니다.", err=True)
        sys.exit(1)

    # 입력 파일 매핑 처리
    data_dir = Path(data)
    input_map = _parse_inputs(inputs)
    if plan.inputs:
        _validate_inputs(plan.inputs, input_map, data_dir)

    # 동적 파라미터 해소
    if plan.dynamic_params:
        from autooffice.engine.resolvers.chain import resolve_plan_dynamic_params

        click.echo(f"동적 파라미터 해소 중... ({len(plan.dynamic_params)}개)")
        plan = resolve_plan_dynamic_params(plan, input_files=input_map)

    runner = PlanRunner(build_default_registry())
    ctx = EngineContext(data_dir=data, input_files=input_map)
    log = runner.run(plan, ctx)

    click.echo(log.summary())
    if not log.success:
        sys.exit(1)


if __name__ == "__main__":
    main()
