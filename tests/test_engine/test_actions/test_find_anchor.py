"""FIND_ANCHOR 액션 핸들러 테스트."""

from __future__ import annotations

from pathlib import Path

import pytest
import xlwings as xw

from autooffice.engine.actions.excel_actions import FindAnchorHandler
from autooffice.engine.context import EngineContext


@pytest.fixture
def anchor_excel(tmp_data_dir: Path, xw_app: xw.App) -> Path:
    """FIND_ANCHOR 테스트용 엑셀.

    Sheet1 구조:
      Row 1: (빈칸)  | (빈칸)   | (빈칸)
      Row 2: (빈칸)  | "일별현황" | "주간"  | "월간"
      Row 3: "항목"  | 3/30     | 3/31   | 4/1    | "14주" | "4월"
      Row 4: "불량률" | 1.2      | 1.5    | 0.8    | (빈칸) | (빈칸)
      Row 5: "생산량" | 100      | 120    | 110    | (빈칸) | (빈칸)
      Row 6: "Sales" | 200      | 250    | 300    | (빈칸) | (빈칸)
    """
    wb = xw_app.books.add()
    ws = wb.sheets[0]
    ws.name = "Sheet1"

    ws.range("B2").value = "일별현황"
    ws.range("C2").value = "주간"
    ws.range("D2").value = "월간"
    ws.range("A3").value = "항목"
    ws.range("B3").value = "3/30"
    ws.range("C3").value = "3/31"
    ws.range("D3").value = "4/1"
    ws.range("E3").value = "14주"
    ws.range("F3").value = "4월"
    ws.range("A4").value = "불량률"
    ws.range("B4").value = 1.2
    ws.range("C4").value = 1.5
    ws.range("D4").value = 0.8
    ws.range("A5").value = "생산량"
    ws.range("B5").value = 100
    ws.range("C5").value = 120
    ws.range("D5").value = 110
    ws.range("A6").value = "Sales"
    ws.range("B6").value = 200
    ws.range("C6").value = 250
    ws.range("D6").value = 300

    path = tmp_data_dir / "anchor_test.xlsx"
    wb.save(str(path))
    wb.close()
    return path


class TestFindAnchor:
    def test_exact_match(self, engine_ctx: EngineContext, anchor_excel: Path):
        """정확한 텍스트 매칭."""
        wb = engine_ctx.app.books.open(str(anchor_excel))
        engine_ctx.register_workbook("test", wb, anchor_excel)

        handler = FindAnchorHandler()
        result = handler.execute(
            {"workbook": "test", "sheet": "Sheet1", "search_value": "항목",
             "scan_range": "A1:F10", "match_type": "exact"},
            engine_ctx,
        )

        assert result.success
        assert result.data["row"] == 3
        assert result.data["column"] == "A"
        assert result.data["cell"] == "A3"

    def test_contains_match(self, engine_ctx: EngineContext, anchor_excel: Path):
        """부분 문자열 매칭 ('일별' → '일별현황')."""
        wb = engine_ctx.app.books.open(str(anchor_excel))
        engine_ctx.register_workbook("test", wb, anchor_excel)

        handler = FindAnchorHandler()
        result = handler.execute(
            {"workbook": "test", "sheet": "Sheet1", "search_value": "일별",
             "scan_range": "A1:F10", "match_type": "contains"},
            engine_ctx,
        )

        assert result.success
        assert result.data["row"] == 2
        assert result.data["column"] == "B"
        assert result.data["value"] == "일별현황"

    def test_starts_with_match(self, engine_ctx: EngineContext, anchor_excel: Path):
        """접두사 매칭."""
        wb = engine_ctx.app.books.open(str(anchor_excel))
        engine_ctx.register_workbook("test", wb, anchor_excel)

        handler = FindAnchorHandler()
        result = handler.execute(
            {"workbook": "test", "sheet": "Sheet1", "search_value": "일별",
             "scan_range": "A1:F10", "match_type": "starts_with"},
            engine_ctx,
        )

        assert result.success
        assert result.data["value"] == "일별현황"

    def test_case_insensitive(self, engine_ctx: EngineContext, anchor_excel: Path):
        """대소문자 무시 매칭."""
        wb = engine_ctx.app.books.open(str(anchor_excel))
        engine_ctx.register_workbook("test", wb, anchor_excel)

        handler = FindAnchorHandler()
        result = handler.execute(
            {"workbook": "test", "sheet": "Sheet1", "search_value": "sales",
             "scan_range": "A1:F10", "match_type": "exact"},
            engine_ctx,
        )

        assert result.success
        assert result.data["row"] == 6
        assert result.data["column"] == "A"

    def test_not_found(self, engine_ctx: EngineContext, anchor_excel: Path):
        """못 찾았을 때 실패."""
        wb = engine_ctx.app.books.open(str(anchor_excel))
        engine_ctx.register_workbook("test", wb, anchor_excel)

        handler = FindAnchorHandler()
        result = handler.execute(
            {"workbook": "test", "sheet": "Sheet1", "search_value": "존재하지않는텍스트",
             "scan_range": "A1:F10", "match_type": "exact"},
            engine_ctx,
        )

        assert not result.success
        assert "찾을 수 없습니다" in result.error

    def test_returns_first_match(self, engine_ctx: EngineContext, anchor_excel: Path):
        """여러 매칭 중 첫 번째(좌상단 우선) 반환."""
        wb = engine_ctx.app.books.open(str(anchor_excel))
        engine_ctx.register_workbook("test", wb, anchor_excel)

        handler = FindAnchorHandler()
        result = handler.execute(
            {"workbook": "test", "sheet": "Sheet1", "search_value": "주",
             "scan_range": "A1:F10", "match_type": "contains"},
            engine_ctx,
        )

        assert result.success
        assert result.data["row"] == 2
        assert result.data["column"] == "C"
        assert result.data["value"] == "주간"

    def test_numeric_value_match(self, engine_ctx: EngineContext, anchor_excel: Path):
        """숫자 값을 문자열로 변환 후 매칭."""
        wb = engine_ctx.app.books.open(str(anchor_excel))
        engine_ctx.register_workbook("test", wb, anchor_excel)

        handler = FindAnchorHandler()
        result = handler.execute(
            {"workbook": "test", "sheet": "Sheet1", "search_value": "100",
             "scan_range": "A1:F10", "match_type": "exact"},
            engine_ctx,
        )

        assert result.success
        assert result.data["row"] == 5
        assert result.data["column"] == "B"

    def test_missing_search_value(self, engine_ctx: EngineContext, anchor_excel: Path):
        """search_value 누락 시 실패."""
        wb = engine_ctx.app.books.open(str(anchor_excel))
        engine_ctx.register_workbook("test", wb, anchor_excel)

        handler = FindAnchorHandler()
        result = handler.execute(
            {"workbook": "test", "sheet": "Sheet1", "scan_range": "A1:F10"},
            engine_ctx,
        )

        assert not result.success
