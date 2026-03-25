"""테스트 공통 fixture."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
import xlwings as xw

from autooffice.engine.actions import build_default_registry
from autooffice.engine.context import EngineContext
from autooffice.engine.runner import PlanRunner
from autooffice.models.execution_plan import ExecutionPlan


@pytest.fixture(scope="session")
def xw_app():
    """세션 전체에서 공유하는 Excel 앱 인스턴스."""
    app = xw.App(visible=False)
    app.display_alerts = False
    yield app
    app.quit()


@pytest.fixture
def tmp_data_dir(tmp_path: Path) -> Path:
    """임시 데이터 디렉토리."""
    return tmp_path


@pytest.fixture
def engine_ctx(tmp_data_dir: Path, xw_app: xw.App) -> EngineContext:
    """기본 EngineContext (세션 공유 Excel 앱 사용)."""
    ctx = EngineContext(data_dir=tmp_data_dir, app=xw_app)
    yield ctx
    # 테스트 종료 시 열린 워크북 정리 (앱은 세션 fixture가 관리)
    for wb in list(ctx.open_workbooks.values()):
        try:
            wb.close()
        except Exception:
            pass
    ctx.open_workbooks.clear()
    ctx.file_paths.clear()


@pytest.fixture
def runner() -> PlanRunner:
    """기본 PlanRunner (모든 핸들러 등록)."""
    return PlanRunner(build_default_registry())


@pytest.fixture
def sample_raw_excel(tmp_data_dir: Path, xw_app: xw.App) -> Path:
    """더미 raw 데이터 엑셀 파일 생성."""
    wb = xw_app.books.add()
    ws = wb.sheets[0]
    ws.name = "Sheet1"

    # 헤더
    headers = ["날짜", "라인명", "생산수량", "불량수량", "불량유형"]
    ws.range((1, 1)).value = headers

    # 더미 데이터 (5행)
    data = [
        ["2026-03-25", "코팅1라인", 1000, 30, "외관"],
        ["2026-03-25", "코팅2라인", 800, 20, "치수"],
        ["2026-03-25", "조립1라인", 1200, 80, "기능"],
        ["2026-03-25", "조립2라인", 900, 15, "외관"],
        ["2026-03-25", "포장라인", 1500, 10, "기타"],
    ]
    ws.range((2, 1)).value = data

    path = tmp_data_dir / "raw_data.xlsx"
    wb.save(str(path))
    wb.close()
    return path


@pytest.fixture
def sample_template_excel(tmp_data_dir: Path, xw_app: xw.App) -> Path:
    """더미 양식 엑셀 파일 생성."""
    wb = xw_app.books.add()

    # 데이터입력 시트
    ws_input = wb.sheets[0]
    ws_input.name = "데이터입력"
    headers = ["날짜", "라인명", "생산수량", "불량수량", "불량유형"]
    ws_input.range((1, 2)).value = headers  # B1부터 시작

    # 일별현황 시트
    ws_daily = wb.sheets.add("일별현황", after=ws_input)
    daily_headers = ["날짜", "라인명", "생산수량", "불량수량", "불량유형", "불량률"]
    ws_daily.range((1, 1)).value = daily_headers

    # 주별현황, 월별현황 시트
    ws_weekly = wb.sheets.add("주별현황", after=ws_daily)
    wb.sheets.add("월별현황", after=ws_weekly)

    path = tmp_data_dir / "defect_rate_template.xlsx"
    wb.save(str(path))
    wb.close()
    return path


@pytest.fixture
def sample_plan_dict() -> dict:
    """더미 execution_plan JSON dict."""
    return {
        "task_id": "test_defect_rate",
        "description": "테스트용 불량률 산출",
        "created_by": "Claude",
        "created_at": "2026-03-25T09:00:00+09:00",
        "version": "1.0.0",
        "metadata": {
            "task_type": "defect_rate",
            "template_hash": "test_hash_123",
            "reusable": True,
        },
        "inputs": {
            "raw_data": {
                "description": "생산 데이터",
                "expected_format": "xlsx",
                "expected_sheets": ["Sheet1"],
                "expected_columns": ["날짜", "라인명", "생산수량", "불량수량", "불량유형"],
            },
            "template": {
                "description": "불량률 양식",
                "expected_format": "xlsx",
                "expected_sheets": ["데이터입력", "일별현황"],
            },
        },
        "steps": [
            {
                "step": 1,
                "action": "OPEN_FILE",
                "description": "raw 데이터 열기",
                "params": {"file_path": "raw_data.xlsx", "alias": "raw", "data_only": True},
                "on_fail": "STOP",
            },
            {
                "step": 2,
                "action": "OPEN_FILE",
                "description": "양식 열기",
                "params": {"file_path": "defect_rate_template.xlsx", "alias": "template"},
                "on_fail": "STOP",
            },
            {
                "step": 3,
                "action": "READ_COLUMNS",
                "description": "raw 데이터 읽기",
                "params": {
                    "file": "raw",
                    "sheet": "Sheet1",
                    "columns": ["날짜", "라인명", "생산수량", "불량수량", "불량유형"],
                    "start_row": 2,
                },
                "store_as": "raw_data",
                "on_fail": "STOP",
            },
            {
                "step": 4,
                "action": "CLEAR_RANGE",
                "description": "기존 데이터 클리어",
                "params": {"file": "template", "sheet": "데이터입력", "range": "B3:F100"},
                "on_fail": "STOP",
            },
            {
                "step": 5,
                "action": "WRITE_DATA",
                "description": "데이터 쓰기",
                "params": {
                    "source": "$raw_data",
                    "target_file": "template",
                    "target_sheet": "데이터입력",
                    "target_start": "B3",
                    "column_mapping": {
                        "날짜": "B",
                        "라인명": "C",
                        "생산수량": "D",
                        "불량수량": "E",
                        "불량유형": "F",
                    },
                },
                "on_fail": "STOP",
            },
            {
                "step": 6,
                "action": "SAVE_FILE",
                "description": "결과 저장",
                "params": {"file": "template", "save_as": "result.xlsx"},
                "on_fail": "STOP",
            },
            {
                "step": 7,
                "action": "LOG",
                "description": "완료 로그",
                "params": {"message": "테스트 완료", "level": "info"},
                "on_fail": "SKIP",
            },
        ],
        "final_output": {"type": "file", "description": "결과 엑셀 파일"},
    }
