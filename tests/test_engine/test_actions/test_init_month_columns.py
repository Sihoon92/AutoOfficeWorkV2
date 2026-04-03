"""FIND_DATE_COLUMN 월 자동 생성 기능 테스트.

대상 날짜의 월이 scan_range에 없으면 마지막 날짜 다음 열에
해당 월의 일별 헤더를 자동 생성한 뒤 재탐색한다.
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pytest
import xlwings as xw

from autooffice.engine.actions.excel_actions import FindDateColumnHandler
from autooffice.engine.context import EngineContext


@pytest.fixture
def full_april_ctx(tmp_data_dir: Path, xw_app: xw.App, engine_ctx: EngineContext):
    """4월 전체(30일) 헤더가 있는 엑셀 + EngineContext.

    Sheet1 Row 6: 4/1 ~ 4/30 (A6:AD6)
    워크북을 닫지 않고 engine_ctx에 직접 등록한다 (Excel 보호 보기 방지).
    """
    wb = xw_app.books.add()
    ws = wb.sheets[0]
    ws.name = "Sheet1"

    for day in range(1, 31):
        ws.range((6, day)).value = f"4/{day}"

    engine_ctx.register_workbook("test", wb, tmp_data_dir / "full_april.xlsx")
    return engine_ctx


class TestFindDateColumnAutoInitMonth:
    """FIND_DATE_COLUMN이 새 월 헤더를 자동 생성하는 테스트."""

    def test_auto_creates_may_and_finds_date(self, full_april_ctx):
        """5/1 탐색 시 5월 헤더가 없으면 자동 생성 후 정확히 찾는다."""
        handler = FindDateColumnHandler()

        result = handler.execute(
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "scan_range": "A6:ZZ6",
                "date": "2026-05-01",
            },
            full_april_ctx,
        )

        assert result.success is True
        assert result.data["column"] is not None
        # 자동 생성 후 정확 매칭 (추론이 아님)
        assert "자동 생성" in result.message

    def test_auto_creates_may_15(self, full_april_ctx):
        """5/15 탐색 시 5월 전체 헤더 생성 후 5/15 정확 매칭."""
        handler = FindDateColumnHandler()

        result = handler.execute(
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "scan_range": "A6:ZZ6",
                "date": "2026-05-15",
            },
            full_april_ctx,
        )

        assert result.success is True
        assert result.data["column"] is not None

    def test_existing_month_not_recreated(self, full_april_ctx):
        """4월 날짜 탐색 시 이미 존재하므로 자동 생성하지 않는다."""
        handler = FindDateColumnHandler()

        result = handler.execute(
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "scan_range": "A6:ZZ6",
                "date": "2026-04-15",
            },
            full_april_ctx,
        )

        assert result.success is True
        assert "자동 생성" not in result.message

    def test_auto_init_idempotent(self, full_april_ctx):
        """같은 월을 두 번 탐색해도 헤더가 중복 생성되지 않는다."""
        handler = FindDateColumnHandler()

        # 첫 번째 호출: 5월 자동 생성
        result1 = handler.execute(
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "scan_range": "A6:ZZ6",
                "date": "2026-05-01",
            },
            full_april_ctx,
        )
        col1 = result1.data["column"]

        # 두 번째 호출: 같은 5월 날짜 (이미 존재)
        result2 = handler.execute(
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "scan_range": "A6:ZZ6",
                "date": "2026-05-01",
            },
            full_april_ctx,
        )
        col2 = result2.data["column"]

        assert col1 == col2

    def test_may_has_31_columns(self, full_april_ctx):
        """5월 자동 생성 시 5/31까지 찾을 수 있다 (31일)."""
        handler = FindDateColumnHandler()

        result = handler.execute(
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "scan_range": "A6:ZZ6",
                "date": "2026-05-31",
            },
            full_april_ctx,
        )

        assert result.success is True
        assert result.data["column"] is not None

    def test_headers_written_as_date_strings(self, full_april_ctx):
        """자동 생성된 헤더가 M/D 형식 문자열로 기록된다."""
        handler = FindDateColumnHandler()

        result = handler.execute(
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "scan_range": "A6:ZZ6",
                "date": "2026-05-01",
            },
            full_april_ctx,
        )

        wb = full_april_ctx.get_workbook("test")
        ws = wb.sheets["Sheet1"]
        col = result.data["column"]
        cell_val = str(ws.range(f"{col}6").value)
        assert "5" in cell_val
