"""AGGREGATE_RANGE 핸들러 테스트 (xlwings mock 기반)."""

from __future__ import annotations

import sys
import types
from unittest.mock import MagicMock, PropertyMock

import pytest

# xlwings stub (이 환경에서는 실제 Excel이 없으므로 mock 사용)
_xw_stub = types.ModuleType("xlwings")
_xw_stub.App = MagicMock
_xw_stub.Book = MagicMock
sys.modules.setdefault("xlwings", _xw_stub)

from autooffice.engine.actions.excel_actions import AggregateRangeHandler
from autooffice.engine.context import EngineContext


def _make_mock_ctx_with_sheet(sheet_data_2d: list[list], start_row: int, start_col: int):
    """sheet_data_2d를 반환하는 mock EngineContext를 생성한다.

    Args:
        sheet_data_2d: range.value로 반환될 2D 데이터
        start_row: 소스 범위 시작 행
        start_col: 소스 범위 시작 열
    """
    ctx = MagicMock(spec=EngineContext)

    # ws.range(...).value 설정
    ws = MagicMock()
    written_data = {}

    def mock_range(*args):
        rng = MagicMock()
        if len(args) == 2:
            # (start, end) 형태
            r1, c1 = args[0]
            r2, c2 = args[1]
            # 소스 범위 읽기 vs 타겟 범위 쓰기 구분
            if c1 >= start_col and c2 >= start_col and r1 == start_row:
                # 소스 범위 → value 반환
                rng.value = sheet_data_2d
            else:
                # 타겟 범위 → value 쓰기 기록
                def set_value(val):
                    written_data["values"] = val
                type(rng).value = PropertyMock(side_effect=set_value)
        return rng

    ws.range = mock_range

    wb = MagicMock()
    wb.sheets.__getitem__ = MagicMock(return_value=ws)
    ctx.get_workbook.return_value = wb

    return ctx, written_data


class TestAggregateRangeSum:
    """AGGREGATE_RANGE method=sum 테스트."""

    def test_basic_sum(self):
        """3행 × 3열 데이터의 행별 합계."""
        data = [
            [10, 20, 30],
            [5, 8, 12],
            [100, None, 50],
        ]
        ctx, written = _make_mock_ctx_with_sheet(data, start_row=7, start_col=72)

        handler = AggregateRangeHandler()
        result = handler.execute(
            {
                "workbook": "mat",
                "sheet": "Sheet1",
                "source_columns_start": "BT",
                "source_columns_end": "BV",
                "source_start_row": 7,
                "source_end_row": 9,
                "target_column": "BS",
                "target_start_row": 7,
                "method": "sum",
            },
            ctx,
        )

        assert result.success
        assert result.data["rows_aggregated"] == 3
        assert result.data["method"] == "sum"
        # 기록된 값 확인
        assert written["values"] == [[60], [25], [150]]

    def test_sum_with_non_numeric(self):
        """문자열/None 포함 시 숫자만 합산."""
        data = [
            [10, "text", 30],
            [None, None, None],
            [5, 15, "abc"],
        ]
        ctx, written = _make_mock_ctx_with_sheet(data, start_row=7, start_col=72)

        handler = AggregateRangeHandler()
        result = handler.execute(
            {
                "workbook": "mat",
                "sheet": "Sheet1",
                "source_columns_start": "BT",
                "source_columns_end": "BV",
                "source_start_row": 7,
                "source_end_row": 9,
                "target_column": "BS",
                "target_start_row": 7,
                "method": "sum",
            },
            ctx,
        )

        assert result.success
        assert written["values"] == [[40], [0], [20]]

    def test_sum_single_column(self):
        """단일 열 데이터 (1D 리스트로 반환되는 경우)."""
        # xlwings는 단일 열을 1D list로 반환
        data = [10, 20, 30]
        ctx, written = _make_mock_ctx_with_sheet(data, start_row=7, start_col=72)

        handler = AggregateRangeHandler()
        result = handler.execute(
            {
                "workbook": "mat",
                "sheet": "Sheet1",
                "source_columns_start": "BT",
                "source_columns_end": "BT",
                "source_start_row": 7,
                "source_end_row": 9,
                "target_column": "BS",
                "target_start_row": 7,
                "method": "sum",
            },
            ctx,
        )

        assert result.success
        assert result.data["rows_aggregated"] == 3
        assert written["values"] == [[10], [20], [30]]


class TestAggregateRangeAverage:
    """AGGREGATE_RANGE method=average 테스트."""

    def test_basic_average(self):
        """행별 평균 계산."""
        data = [
            [10, 20, 30],
            [4, 8, 12],
        ]
        ctx, written = _make_mock_ctx_with_sheet(data, start_row=7, start_col=72)

        handler = AggregateRangeHandler()
        result = handler.execute(
            {
                "workbook": "mat",
                "sheet": "Sheet1",
                "source_columns_start": "BT",
                "source_columns_end": "BV",
                "source_start_row": 7,
                "source_end_row": 8,
                "target_column": "BS",
                "target_start_row": 7,
                "method": "average",
            },
            ctx,
        )

        assert result.success
        assert result.data["method"] == "average"
        assert written["values"] == [[20.0], [8.0]]


class TestAggregateRangeCount:
    """AGGREGATE_RANGE method=count 테스트."""

    def test_count_numeric_values(self):
        """행별 숫자 값 개수."""
        data = [
            [10, None, 30],
            ["abc", 8, 12],
        ]
        ctx, written = _make_mock_ctx_with_sheet(data, start_row=7, start_col=72)

        handler = AggregateRangeHandler()
        result = handler.execute(
            {
                "workbook": "mat",
                "sheet": "Sheet1",
                "source_columns_start": "BT",
                "source_columns_end": "BV",
                "source_start_row": 7,
                "source_end_row": 8,
                "target_column": "BS",
                "target_start_row": 7,
                "method": "count",
            },
            ctx,
        )

        assert result.success
        assert written["values"] == [[2], [2]]


class TestAggregateRangeValidation:
    """AGGREGATE_RANGE 파라미터 검증 테스트."""

    def test_missing_source_columns(self):
        """source_columns_start/end 누락 시 실패."""
        ctx = MagicMock(spec=EngineContext)
        handler = AggregateRangeHandler()

        result = handler.execute(
            {
                "workbook": "mat",
                "sheet": "Sheet1",
                "source_columns_start": "",
                "source_columns_end": "BV",
                "source_start_row": 7,
                "source_end_row": 9,
                "target_column": "BS",
            },
            ctx,
        )
        assert not result.success
        assert "source_columns_start" in result.error

    def test_missing_target_column(self):
        """target_column 누락 시 실패."""
        ctx = MagicMock(spec=EngineContext)
        handler = AggregateRangeHandler()

        result = handler.execute(
            {
                "workbook": "mat",
                "sheet": "Sheet1",
                "source_columns_start": "BT",
                "source_columns_end": "BV",
                "source_start_row": 7,
                "source_end_row": 9,
                "target_column": "",
            },
            ctx,
        )
        assert not result.success
        assert "target_column" in result.error

    def test_reversed_columns(self):
        """source_columns_end가 start보다 앞에 있으면 실패."""
        ctx = MagicMock(spec=EngineContext)
        wb = MagicMock()
        ws = MagicMock()
        wb.sheets.__getitem__ = MagicMock(return_value=ws)
        ctx.get_workbook.return_value = wb

        handler = AggregateRangeHandler()
        result = handler.execute(
            {
                "workbook": "mat",
                "sheet": "Sheet1",
                "source_columns_start": "CQ",
                "source_columns_end": "BT",
                "source_start_row": 7,
                "source_end_row": 9,
                "target_column": "BS",
            },
            ctx,
        )
        assert not result.success
        assert "앞에 있습니다" in result.error


class TestAggregateRangeRegistry:
    """AGGREGATE_RANGE가 레지스트리에 등록되어 있는지 확인."""

    def test_registered_in_default_registry(self):
        from autooffice.engine.actions import build_default_registry
        registry = build_default_registry()
        assert "AGGREGATE_RANGE" in registry

    def test_action_type_enum_exists(self):
        from autooffice.models.execution_plan import ActionType
        assert ActionType.AGGREGATE_RANGE.value == "AGGREGATE_RANGE"
