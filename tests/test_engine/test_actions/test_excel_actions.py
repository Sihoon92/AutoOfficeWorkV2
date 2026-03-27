"""엑셀 ACTION 핸들러 테스트."""

from __future__ import annotations

from datetime import date, datetime, timedelta

import xlwings as xw

from autooffice.engine.actions.excel_actions import (
    ClearRangeHandler,
    CopyRangeHandler,
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


class TestCopyRange:
    """COPY_RANGE: 시트 간 범위 복사 (값 붙여넣기)."""

    def test_copy_values_between_sheets(self, engine_ctx, tmp_data_dir):
        """source 시트의 범위를 target 시트의 열에 값으로 복사한다."""
        wb = engine_ctx.app.books.add()

        # 소스 시트: 수식이 있는 D7:D9
        ws_src = wb.sheets[0]
        ws_src.name = "자재계산식"
        ws_src.range("D7").value = 100
        ws_src.range("D8").value = 200
        ws_src.range("D9").value = 300

        # 타겟 시트
        ws_tgt = wb.sheets.add("Sheet1", after=ws_src)

        engine_ctx.register_workbook("target", wb, tmp_data_dir / "target.xlsx")

        handler = CopyRangeHandler()
        result = handler.execute(
            {
                "workbook": "target",
                "source_sheet": "자재계산식",
                "source_range": "D7:D9",
                "target_sheet": "Sheet1",
                "target_column": "F",
                "target_start_row": 6,
                "row_offset": 1,
                "paste_type": "values",
            },
            engine_ctx,
        )

        assert result.success
        assert result.data["rows_copied"] == 3
        # F7, F8, F9에 값이 복사됨 (target_start_row=6 + row_offset=1 = 7)
        assert ws_tgt.range("F7").value == 100
        assert ws_tgt.range("F8").value == 200
        assert ws_tgt.range("F9").value == 300

    def test_copy_overwrites_existing(self, engine_ctx, tmp_data_dir):
        """기존 데이터가 있으면 덮어쓴다."""
        wb = engine_ctx.app.books.add()
        ws_src = wb.sheets[0]
        ws_src.name = "source"
        ws_src.range("A1").value = 999

        ws_tgt = wb.sheets.add("target", after=ws_src)
        ws_tgt.range("C5").value = "old_value"

        engine_ctx.register_workbook("wb", wb, tmp_data_dir / "wb.xlsx")

        handler = CopyRangeHandler()
        result = handler.execute(
            {
                "workbook": "wb",
                "source_sheet": "source",
                "source_range": "A1:A1",
                "target_sheet": "target",
                "target_column": "C",
                "target_start_row": 5,
                "row_offset": 0,
                "paste_type": "values",
            },
            engine_ctx,
        )

        assert result.success
        assert ws_tgt.range("C5").value == 999

    def test_copy_default_row_offset(self, engine_ctx, tmp_data_dir):
        """row_offset 미지정 시 기본값 0."""
        wb = engine_ctx.app.books.add()
        ws_src = wb.sheets[0]
        ws_src.name = "src"
        ws_src.range("B2").value = 42

        ws_tgt = wb.sheets.add("tgt", after=ws_src)
        engine_ctx.register_workbook("wb", wb, tmp_data_dir / "wb.xlsx")

        handler = CopyRangeHandler()
        result = handler.execute(
            {
                "workbook": "wb",
                "source_sheet": "src",
                "source_range": "B2:B2",
                "target_sheet": "tgt",
                "target_column": "A",
                "target_start_row": 10,
                "paste_type": "values",
            },
            engine_ctx,
        )

        assert result.success
        assert ws_tgt.range("A10").value == 42

    def test_copy_multirow_source(self, engine_ctx, tmp_data_dir):
        """다중 행 복사 (D7:D48 같은 42행 시나리오 축소판)."""
        wb = engine_ctx.app.books.add()
        ws_src = wb.sheets[0]
        ws_src.name = "src"
        # 10행 데이터
        for i in range(10):
            ws_src.range((1 + i, 1)).value = (i + 1) * 10

        ws_tgt = wb.sheets.add("tgt", after=ws_src)
        engine_ctx.register_workbook("wb", wb, tmp_data_dir / "wb.xlsx")

        handler = CopyRangeHandler()
        result = handler.execute(
            {
                "workbook": "wb",
                "source_sheet": "src",
                "source_range": "A1:A10",
                "target_sheet": "tgt",
                "target_column": "D",
                "target_start_row": 3,
                "row_offset": 2,
                "paste_type": "values",
            },
            engine_ctx,
        )

        assert result.success
        assert result.data["rows_copied"] == 10
        # D5 ~ D14 (start_row=3 + offset=2 = 5)
        assert ws_tgt.range("D5").value == 10
        assert ws_tgt.range("D14").value == 100


class TestFindDateColumnCopyRangeIntegration:
    """FIND_DATE_COLUMN → COPY_RANGE 통합 시나리오 (dot notation 변수 참조 포함)."""

    def test_full_workflow(self, engine_ctx, runner, tmp_data_dir):
        """날짜 열 탐색 → 시트 간 값 복사 전체 흐름."""
        today = date.today()
        d_minus_2 = today - timedelta(days=2)
        d_minus_1 = today - timedelta(days=1)

        # 워크북 생성: 소스시트(자재계산식) + 타겟시트(Sheet1, 날짜 헤더)
        wb = engine_ctx.app.books.add()
        ws_calc = wb.sheets[0]
        ws_calc.name = "자재계산식"
        # D7:D9에 계산 결과값 넣기
        ws_calc.range("D7").value = 11.5
        ws_calc.range("D8").value = 22.3
        ws_calc.range("D9").value = 33.7

        ws_sheet1 = wb.sheets.add("Sheet1", after=ws_calc)
        # 행 4에 날짜 헤더: C4=전전일, D4=전일
        ws_sheet1.range("C4").value = f"{d_minus_2.month}/{d_minus_2.day}"
        ws_sheet1.range("D4").value = f"{d_minus_1.month}/{d_minus_1.day}"

        path = tmp_data_dir / "material.xlsx"
        engine_ctx.register_workbook("material", wb, path)

        # Execution Plan
        from autooffice.models.execution_plan import ExecutionPlan

        plan_dict = {
            "task_id": "daily_material_append",
            "description": "자재 계산식 결과를 Sheet1 오늘 열에 추가",
            "created_by": "Claude",
            "created_at": "2026-03-27T09:00:00+09:00",
            "inputs": {
                "material": {
                    "description": "자재 관리 파일",
                    "expected_format": "xlsx",
                    "expected_sheets": ["자재계산식", "Sheet1"],
                }
            },
            "steps": [
                {
                    "step": 1,
                    "action": "FIND_DATE_COLUMN",
                    "description": "Sheet1에서 오늘 날짜 열 탐색",
                    "params": {
                        "workbook": "material",
                        "sheet": "Sheet1",
                        "scan_range": "A1:ZZ10",
                        "date": "today",
                    },
                    "store_as": "date_info",
                },
                {
                    "step": 2,
                    "action": "COPY_RANGE",
                    "description": "자재계산식 D7:D9 → Sheet1 오늘 열",
                    "params": {
                        "workbook": "material",
                        "source_sheet": "자재계산식",
                        "source_range": "D7:D9",
                        "target_sheet": "Sheet1",
                        "target_column": "$date_info.column",
                        "target_start_row": "$date_info.date_row",
                        "row_offset": 1,
                        "paste_type": "values",
                    },
                },
                {
                    "step": 3,
                    "action": "SAVE_FILE",
                    "description": "저장",
                    "params": {"file": "material"},
                },
            ],
            "final_output": {"type": "file", "description": "자재 관리 파일 업데이트"},
        }

        plan = ExecutionPlan.model_validate(plan_dict)

        # validate (dry-run)
        errors = runner.validate(plan)
        assert errors == [], f"검증 오류: {errors}"

        # run
        log = runner.run(plan, engine_ctx)
        assert log.success, f"실행 실패: {log.error}"

        # 결과 검증: Sheet1의 E열(오늘 날짜)에 값이 들어있어야 함
        # C4=전전일, D4=전일 → 오늘=E열, 데이터는 E5:E7
        wb2 = engine_ctx.app.books.open(str(path))
        ws = wb2.sheets["Sheet1"]
        assert ws.range("E5").value == 11.5
        assert ws.range("E6").value == 22.3
        assert ws.range("E7").value == 33.7
        wb2.close()
