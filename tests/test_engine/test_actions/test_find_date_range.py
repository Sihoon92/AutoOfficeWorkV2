"""FIND_DATE_RANGE 액션 핸들러 테스트."""

from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

import pytest
import xlwings as xw

from autooffice.engine.actions.excel_actions import FindDateRangeHandler
from autooffice.engine.context import EngineContext


@pytest.fixture
def date_range_excel(tmp_data_dir: Path, xw_app: xw.App) -> Path:
    """FIND_DATE_RANGE 테스트용 엑셀.

    Sheet1 구조 (날짜가 3행에 위치):
      Row 1: (빈칸)
      Row 2: "항목" | "일별"
      Row 3: (빈칸)  | 3/28  | 3/29  | 3/30  | 3/31  | 4/1   | 4/2   | 4/3
      Row 4: "불량률" | 1.0   | 1.1   | 1.2   | 1.5   | 0.8   | 1.1   | 0.9
      Row 5: "생산량" | 100   | 110   | 120   | 150   | 80    | 110   | 90
    """
    wb = xw_app.books.add()
    ws = wb.sheets[0]
    ws.name = "Sheet1"

    ws.range("A2").value = "항목"
    ws.range("B2").value = "일별"

    # 날짜 헤더 (3/28 ~ 4/3) — "M/D" 형식 문자열
    dates = ["3/28", "3/29", "3/30", "3/31", "4/1", "4/2", "4/3"]
    for i, d in enumerate(dates):
        ws.range((3, 2 + i)).value = d

    # 데이터
    ws.range("A4").value = "불량률"
    ws.range((4, 2)).value = [1.0, 1.1, 1.2, 1.5, 0.8, 1.1, 0.9]
    ws.range("A5").value = "생산량"
    ws.range((5, 2)).value = [100, 110, 120, 150, 80, 110, 90]

    path = tmp_data_dir / "date_range_test.xlsx"
    wb.save(str(path))
    wb.close()
    return path


class TestFindDateRange:
    def test_full_week_match(self, engine_ctx: EngineContext, date_range_excel: Path):
        """범위 내 모든 날짜 매칭 (3/30~4/3, 5일 모두 존재)."""
        wb = engine_ctx.app.books.open(str(date_range_excel))
        engine_ctx.register_workbook("test", wb, date_range_excel)

        handler = FindDateRangeHandler()
        result = handler.execute(
            {"workbook": "test", "sheet": "Sheet1", "scan_range": "A1:J10",
             "start_date": "2026-03-30", "end_date": "2026-04-03"},
            engine_ctx,
        )

        assert result.success
        assert result.data["date_row"] == 3
        assert result.data["matched_count"] == 5
        assert len(result.data["columns"]) == 5
        assert result.data["missing_dates"] == []

    def test_partial_match_with_missing(self, engine_ctx: EngineContext, date_range_excel: Path):
        """범위 내 일부 날짜 없음 → missing_dates에 포함."""
        wb = engine_ctx.app.books.open(str(date_range_excel))
        engine_ctx.register_workbook("test", wb, date_range_excel)

        handler = FindDateRangeHandler()
        result = handler.execute(
            {"workbook": "test", "sheet": "Sheet1", "scan_range": "A1:J10",
             "start_date": "2026-03-30", "end_date": "2026-04-05"},
            engine_ctx,
        )

        assert result.success
        assert result.data["matched_count"] == 5
        assert "2026-04-04" in result.data["missing_dates"]
        assert "2026-04-05" in result.data["missing_dates"]

    def test_no_match(self, engine_ctx: EngineContext, date_range_excel: Path):
        """범위 내 날짜 없음 → 실패."""
        wb = engine_ctx.app.books.open(str(date_range_excel))
        engine_ctx.register_workbook("test", wb, date_range_excel)

        handler = FindDateRangeHandler()
        result = handler.execute(
            {"workbook": "test", "sheet": "Sheet1", "scan_range": "A1:J10",
             "start_date": "2026-05-01", "end_date": "2026-05-07"},
            engine_ctx,
        )

        assert not result.success

    def test_single_day_range(self, engine_ctx: EngineContext, date_range_excel: Path):
        """start_date == end_date (1일)."""
        wb = engine_ctx.app.books.open(str(date_range_excel))
        engine_ctx.register_workbook("test", wb, date_range_excel)

        handler = FindDateRangeHandler()
        result = handler.execute(
            {"workbook": "test", "sheet": "Sheet1", "scan_range": "A1:J10",
             "start_date": "2026-04-01", "end_date": "2026-04-01"},
            engine_ctx,
        )

        assert result.success
        assert result.data["matched_count"] == 1
        assert len(result.data["columns"]) == 1
        assert result.data["start_column"] == result.data["end_column"]

    def test_month_boundary_cross(self, engine_ctx: EngineContext, date_range_excel: Path):
        """월 경계 크로스 (3/29 ~ 4/2)."""
        wb = engine_ctx.app.books.open(str(date_range_excel))
        engine_ctx.register_workbook("test", wb, date_range_excel)

        handler = FindDateRangeHandler()
        result = handler.execute(
            {"workbook": "test", "sheet": "Sheet1", "scan_range": "A1:J10",
             "start_date": "2026-03-29", "end_date": "2026-04-02"},
            engine_ctx,
        )

        assert result.success
        assert result.data["matched_count"] == 5  # 3/29, 3/30, 3/31, 4/1, 4/2
        assert result.data["missing_dates"] == []

    def test_start_end_column_order(self, engine_ctx: EngineContext, date_range_excel: Path):
        """start_column < end_column 순서 보장."""
        wb = engine_ctx.app.books.open(str(date_range_excel))
        engine_ctx.register_workbook("test", wb, date_range_excel)

        handler = FindDateRangeHandler()
        result = handler.execute(
            {"workbook": "test", "sheet": "Sheet1", "scan_range": "A1:J10",
             "start_date": "2026-03-28", "end_date": "2026-04-03"},
            engine_ctx,
        )

        assert result.success
        assert result.data["start_column"] == result.data["columns"][0]
        assert result.data["end_column"] == result.data["columns"][-1]

    def test_invalid_date_format(self, engine_ctx: EngineContext, date_range_excel: Path):
        """잘못된 날짜 형식 → 실패."""
        wb = engine_ctx.app.books.open(str(date_range_excel))
        engine_ctx.register_workbook("test", wb, date_range_excel)

        handler = FindDateRangeHandler()
        result = handler.execute(
            {"workbook": "test", "sheet": "Sheet1", "scan_range": "A1:J10",
             "start_date": "not-a-date", "end_date": "2026-04-03"},
            engine_ctx,
        )

        assert not result.success
        assert "형식 오류" in result.error
