"""AutoOffice CLI - 실행 계획 실행 및 관리.

사용법:
    autooffice run plan.json --data ./data/
    autooffice validate plan.json
    autooffice cache list
    autooffice cache run <plan_id> --data ./data/
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
from autooffice.models.execution_plan import ExecutionPlan


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


def _collect_placeholders(plan: ExecutionPlan) -> set[str]:
    """plan의 모든 step params에서 {{placeholder}} 키를 수집한다."""
    import re

    placeholders: set[str] = set()
    for step in plan.steps:
        for val in step.params.values():
            if isinstance(val, str):
                placeholders.update(re.findall(r"\{\{(\w+)\}\}", val))
    return placeholders


def _resolve_user_file_paths(
    plan: ExecutionPlan,
    file_options: tuple[str, ...],
) -> dict[str, str]:
    """CLI --file 옵션과 대화형 입력으로 사용자 파일 경로를 수집한다."""
    user_paths: dict[str, str] = {}

    # 1) --file 옵션에서 파싱 (name=path 형식)
    for item in file_options:
        if "=" not in item:
            click.echo(f"잘못된 --file 형식: '{item}' (name=path 형식이어야 합니다)", err=True)
            sys.exit(1)
        name, path = item.split("=", 1)
        user_paths[name.strip()] = path.strip()

    # 2) plan에서 필요한 플레이스홀더 수집
    placeholders = _collect_placeholders(plan)

    # 3) inputs에 file_path 기본값이 있으면 기본값으로 설정 (사용자가 이미 지정한 것은 유지)
    for input_key, spec in plan.inputs.items():
        if input_key not in user_paths and spec.file_path:
            user_paths[input_key] = spec.file_path

    # 4) 아직 미지정된 플레이스홀더에 대해 대화형 입력 요청
    missing = placeholders - set(user_paths.keys())
    if missing:
        click.echo("")
        click.echo("파일 경로 설정이 필요합니다:")
        for key in sorted(missing):
            desc = ""
            if key in plan.inputs:
                desc = f" ({plan.inputs[key].description})"
            user_input = click.prompt(f"  {key}{desc}", type=str)
            user_paths[key] = user_input
        click.echo("")

    return user_paths


@main.command()
@click.argument("plan_file", type=click.Path(exists=True))
@click.option("--data", "-d", type=click.Path(exists=True), default=".", help="데이터 디렉토리")
@click.option(
    "--file", "-f", "file_paths", multiple=True,
    help="사용자 파일 경로 지정 (name=path 형식, 여러 번 사용 가능)",
)
def run(plan_file: str, data: str, file_paths: tuple[str, ...]) -> None:
    """execution_plan.json을 실행한다.

    plan 내 {{input_key}} 플레이스홀더를 사용자 지정 경로로 대체한다.

    예시:
        autooffice run plan.json --data ./data/ -f raw_data=./내_데이터.xlsx -f output=./결과.xlsx
    """
    click.echo(f"plan 로드: {plan_file}")

    try:
        plan = ExecutionPlan.from_json_file(plan_file)
    except Exception as e:
        click.echo(f"plan 파싱 실패: {e}", err=True)
        sys.exit(1)

    # 사용자 파일 경로 수집
    user_paths = _resolve_user_file_paths(plan, file_paths)
    if user_paths:
        click.echo("파일 경로 매핑:")
        for k, v in user_paths.items():
            click.echo(f"  {k} → {v}")

    runner = PlanRunner(build_default_registry())
    ctx = EngineContext(data_dir=data, user_file_paths=user_paths)

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
@click.option(
    "--file", "-f", "file_paths", multiple=True,
    help="사용자 파일 경로 지정 (name=path 형식, 여러 번 사용 가능)",
)
def cache_run(plan_id: str, data: str, file_paths: tuple[str, ...]) -> None:
    """캐시된 plan을 실행한다."""
    from autooffice.cache.plan_cache import PlanCache

    pc = PlanCache()
    plan = pc.load_plan(plan_id)

    if plan is None:
        click.echo(f"캐시에서 plan '{plan_id}'을 찾을 수 없습니다.", err=True)
        sys.exit(1)

    user_paths = _resolve_user_file_paths(plan, file_paths)
    runner = PlanRunner(build_default_registry())
    ctx = EngineContext(data_dir=data, user_file_paths=user_paths)
    log = runner.run(plan, ctx)

    click.echo(log.summary())
    if not log.success:
        sys.exit(1)


if __name__ == "__main__":
    main()
