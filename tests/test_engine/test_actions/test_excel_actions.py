"""엑셀 ACTION 핸들러 테스트."""

from __future__ import annotations

from datetime import date, datetime, timedelta

import xlwings as xw

from autooffice.engine.actions.excel_actions import (
    ClearRangeHandler,
    FindDateColumnHandler,
    ReadColumnsHandler,
    WriteDataHandler,
)
from autooffice.engine.context import EngineContext


class TestReadColumns:
    def test_read_by_header_name(self, engine_ctx: EngineContext, sample_raw_excel):
        """헤더 이름으로 컬럼 읽기."""
        wb = engine_ctx.app.books.open(str(sample_raw_excel))
        engine_ctx.register_workbook("raw", wb, sample_raw_excel)

        handler = ReadColumnsHandler()
        result = handler.execute(
            {
                "file": "raw",
                "sheet": "Sheet1",
                "columns": ["날짜", "라인명", "생산수량"],
                "start_row": 2,
            },
            engine_ctx,
        )

        assert result.success
        assert len(result.data) == 5
        assert "날짜" in result.data[0]
        assert "라인명" in result.data[0]
        assert result.data[0]["라인명"] == "코팅1라인"

    def test_read_missing_column(self, engine_ctx: EngineContext, sample_raw_excel):
        """존재하지 않는 컬럼 읽기 시 실패."""
        wb = engine_ctx.app.books.open(str(sample_raw_excel))
        engine_ctx.register_workbook("raw", wb, sample_raw_excel)

        handler = ReadColumnsHandler()
        result = handler.execute(
            {
                "file": "raw",
                "sheet": "Sheet1",
                "columns": ["존재하지않는컬럼"],
                "start_row": 2,
            },
            engine_ctx,
        )

        assert not result.success
        assert "찾을 수 없습니다" in result.error


class TestWriteData:
    def test_write_with_column_mapping(self, engine_ctx: EngineContext, tmp_data_dir):
        """컬럼 매핑으로 데이터 쓰기."""
        wb = engine_ctx.app.books.add()
        ws = wb.sheets[0]
        ws.name = "데이터"
        engine_ctx.register_workbook("target", wb, tmp_data_dir / "target.xlsx")

        source_data = [
            {"이름": "A", "값": 100},
            {"이름": "B", "값": 200},
        ]

        handler = WriteDataHandler()
        result = handler.execute(
            {
                "source": source_data,
                "target_file": "target",
                "target_sheet": "데이터",
                "target_start": "B3",
                "column_mapping": {"이름": "B", "값": "C"},
            },
            engine_ctx,
        )

        assert result.success
        assert result.data["rows_written"] == 2
        assert ws.range((3, 2)).value == "A"
        assert ws.range((3, 3)).value == 100
        assert ws.range((4, 2)).value == "B"
        assert ws.range((4, 3)).value == 200


class TestClearRange:
    def test_clear_range(self, engine_ctx: EngineContext, tmp_data_dir):
        """셀 범위 클리어."""
        wb = engine_ctx.app.books.add()
        ws = wb.sheets[0]
        ws.name = "시트"
        ws.range((1, 1)).value = "헤더"
        ws.range((2, 1)).value = "데이터1"
        ws.range((3, 1)).value = "데이터2"
        engine_ctx.register_workbook("test", wb, tmp_data_dir / "test.xlsx")

        handler = ClearRangeHandler()
        result = handler.execute(
            {"file": "test", "sheet": "시트", "range": "A2:A3"},
            engine_ctx,
        )

        assert result.success
        assert ws.range((1, 1)).value == "헤더"  # 보존
        assert ws.range((2, 1)).value is None  # 클리어됨
        assert ws.range((3, 1)).value is None  # 클리어됨


class TestFindDateColumn:
    """FIND_DATE_COLUMN: 시트에서 날짜 패턴을 탐색하여 오늘 날짜 열을 결정."""

    def _make_date_sheet(self, engine_ctx, tmp_data_dir, dates_in_row6):
        """날짜 패턴이 있는 테스트 워크북 생성 헬퍼.

        Args:
            dates_in_row6: row 6에 넣을 날짜 값 리스트.
                           Col A,B는 빈 헤더, Col C부터 날짜 시작.
        """
        wb = engine_ctx.app.books.add()
        ws = wb.sheets[0]
        ws.name = "Sheet1"

        # 행 6, C열부터 날짜 배치
        for i, d in enumerate(dates_in_row6):
            ws.range((6, 3 + i)).value = d  # C=3, D=4, ...

        engine_ctx.register_workbook("target", wb, tmp_data_dir / "target.xlsx")
        return wb, ws

    def test_exact_match_today(self, engine_ctx, tmp_data_dir):
        """오늘 날짜가 이미 존재하면 해당 열을 반환한다."""
        today = date.today()
        d_minus_2 = today - timedelta(days=2)
        d_minus_1 = today - timedelta(days=1)
        dates = [
            f"{d_minus_2.month}/{d_minus_2.day}",
            f"{d_minus_1.month}/{d_minus_1.day}",
            f"{today.month}/{today.day}",
        ]
        self._make_date_sheet(engine_ctx, tmp_data_dir, dates)

        handler = FindDateColumnHandler()
        result = handler.execute(
            {
                "workbook": "target",
                "sheet": "Sheet1",
                "scan_range": "A1:Z10",
                "date": "today",
            },
            engine_ctx,
        )

        assert result.success
        assert result.data["date_row"] == 6
        # C=3번째 열 + 2(오프셋) = E열
        assert result.data["column"] == "E"

    def test_infer_next_column(self, engine_ctx, tmp_data_dir):
        """오늘 날짜가 없으면 마지막 날짜 다음 열을 추론한다."""
        today = date.today()
        d_minus_2 = today - timedelta(days=2)
        d_minus_1 = today - timedelta(days=1)
        dates = [
            f"{d_minus_2.month}/{d_minus_2.day}",
            f"{d_minus_1.month}/{d_minus_1.day}",
        ]
        self._make_date_sheet(engine_ctx, tmp_data_dir, dates)

        handler = FindDateColumnHandler()
        result = handler.execute(
            {
                "workbook": "target",
                "sheet": "Sheet1",
                "scan_range": "A1:Z10",
                "date": "today",
            },
            engine_ctx,
        )

        assert result.success
        assert result.data["date_row"] == 6
        # C, D에 날짜 → 다음 빈 열 E
        assert result.data["column"] == "E"

    def test_explicit_date_param(self, engine_ctx, tmp_data_dir):
        """date 파라미터로 명시적 날짜를 지정할 수 있다."""
        self._make_date_sheet(engine_ctx, tmp_data_dir, ["3/21", "3/22"])

        handler = FindDateColumnHandler()
        result = handler.execute(
            {
                "workbook": "target",
                "sheet": "Sheet1",
                "scan_range": "A1:Z10",
                "date": "2026-03-22",
            },
            engine_ctx,
        )

        assert result.success
        assert result.data["column"] == "D"  # 3/22는 D열 (C=3/21, D=3/22)
        assert result.data["date_row"] == 6

    def test_datetime_cell_values(self, engine_ctx, tmp_data_dir):
        """Excel datetime 객체도 인식한다."""
        today = date.today()
        yesterday = today - timedelta(days=1)
        dates = [
            datetime(yesterday.year, yesterday.month, yesterday.day),
            datetime(today.year, today.month, today.day),
        ]
        self._make_date_sheet(engine_ctx, tmp_data_dir, dates)

        handler = FindDateColumnHandler()
        result = handler.execute(
            {
                "workbook": "target",
                "sheet": "Sheet1",
                "scan_range": "A1:Z10",
                "date": "today",
            },
            engine_ctx,
        )

        assert result.success
        assert result.data["column"] == "D"  # 오늘 날짜가 D열에 있음

    def test_no_dates_found(self, engine_ctx, tmp_data_dir):
        """날짜를 찾을 수 없으면 실패."""
        wb = engine_ctx.app.books.add()
        ws = wb.sheets[0]
        ws.name = "Sheet1"
        ws.range("A1").value = "제목"
        ws.range("B1").value = "항목"
        engine_ctx.register_workbook("target", wb, tmp_data_dir / "target.xlsx")

        handler = FindDateColumnHandler()
        result = handler.execute(
            {
                "workbook": "target",
                "sheet": "Sheet1",
                "scan_range": "A1:Z10",
                "date": "today",
            },
            engine_ctx,
        )

        assert not result.success
