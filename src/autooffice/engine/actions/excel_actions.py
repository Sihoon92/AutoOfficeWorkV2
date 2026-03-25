"""엑셀 조작 ACTION 핸들러: READ_COLUMNS, READ_RANGE, WRITE_DATA, CLEAR_RANGE, RECALCULATE."""

from __future__ import annotations

import logging
import re
from typing import Any

from autooffice.engine.actions.base import ActionHandler
from autooffice.engine.context import EngineContext
from autooffice.models.action_result import ActionResult

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# 유틸리티 함수 (openpyxl.utils 대체)
# ---------------------------------------------------------------------------

def _col_letter_to_index(col: str) -> int:
    """컬럼 문자를 1-based 인덱스로 변환: 'A' -> 1, 'B' -> 2, 'AA' -> 27."""
    result = 0
    for char in col.upper():
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result


def _parse_coordinate(coord: str) -> tuple[str, int]:
    """셀 좌표를 (컬럼문자, 행번호)로 분리: 'B3' -> ('B', 3)."""
    match = re.match(r"^([A-Za-z]+)(\d+)$", coord)
    if not match:
        raise ValueError(f"잘못된 셀 좌표: {coord}")
    return match.group(1).upper(), int(match.group(2))


# ---------------------------------------------------------------------------
# 핸들러
# ---------------------------------------------------------------------------

class ReadColumnsHandler(ActionHandler):
    """READ_COLUMNS: 특정 시트에서 지정 컬럼의 데이터를 읽는다.

    params:
        file: 워크북 alias
        sheet: 시트명
        columns: 읽을 컬럼 목록 (헤더명 또는 컬럼 문자)
        start_row: 시작 행 (기본 2, 헤더 다음)
        end_row: 종료 행 (선택, 미지정 시 데이터 끝까지)
    """

    def execute(self, params: dict[str, Any], ctx: EngineContext) -> ActionResult:
        alias = params.get("file", "")
        sheet_name = params.get("sheet", "")
        columns = params.get("columns", [])
        start_row = params.get("start_row", 2)
        end_row = params.get("end_row")

        try:
            wb = ctx.get_workbook(alias)
            ws = wb.sheets[sheet_name]

            # 헤더 행에서 컬럼 인덱스 매핑
            header_map: dict[str, int] = {}
            last_col = ws.used_range.last_cell.column
            if last_col > 0:
                header_values = ws.range((1, 1), (1, last_col)).value
                if not isinstance(header_values, list):
                    header_values = [header_values]
                for col_idx, val in enumerate(header_values, 1):
                    if val is not None:
                        header_map[str(val)] = col_idx

            # 컬럼 인덱스 결정
            col_indices: list[int] = []
            for col in columns:
                if col in header_map:
                    col_indices.append(header_map[col])
                elif isinstance(col, str) and col.isalpha() and col.isascii():
                    col_indices.append(_col_letter_to_index(col))
                else:
                    return ActionResult(
                        success=False,
                        error=f"컬럼 '{col}'을 찾을 수 없습니다. 사용 가능: {list(header_map.keys())}",
                    )

            # 데이터 읽기 (컬럼별 배치 읽기로 COM 호출 최소화)
            actual_end = end_row or ws.used_range.last_cell.row
            if actual_end < start_row:
                return ActionResult(success=True, data=[], message="데이터 없음")

            col_data: dict[str, list] = {}
            for col_name, col_idx in zip(columns, col_indices):
                if start_row == actual_end:
                    values = [ws.range((start_row, col_idx)).value]
                else:
                    values = ws.range((start_row, col_idx), (actual_end, col_idx)).value
                    if not isinstance(values, list):
                        values = [values]
                col_data[col_name] = values

            # 행 단위로 조합
            num_rows = actual_end - start_row + 1
            data = []
            for i in range(num_rows):
                row_data = {col: col_data[col][i] for col in columns}
                # 빈 행 스킵 (모든 값이 None)
                if any(v is not None for v in row_data.values()):
                    data.append(row_data)

            return ActionResult(
                success=True,
                data=data,
                message=f"{len(data)}행 읽기 완료 (컬럼: {columns})",
            )
        except Exception as e:
            return ActionResult(success=False, error=f"READ_COLUMNS 실패: {e}")


class ReadRangeHandler(ActionHandler):
    """READ_RANGE: 셀 범위의 데이터를 읽는다.

    params:
        file: 워크북 alias
        sheet: 시트명
        range: 셀 범위 (예: "B3:F100")
    """

    def execute(self, params: dict[str, Any], ctx: EngineContext) -> ActionResult:
        alias = params.get("file", "")
        sheet_name = params.get("sheet", "")
        cell_range = params.get("range", "")

        try:
            wb = ctx.get_workbook(alias)
            ws = wb.sheets[sheet_name]

            rng = ws.range(cell_range)
            raw = rng.value

            # 2D 리스트로 정규화
            if raw is None:
                data = []
            elif not isinstance(raw, list):
                # 단일 셀
                data = [[raw]]
            elif len(raw) > 0 and isinstance(raw[0], list):
                # 이미 2D
                data = raw
            else:
                # 1D 리스트: 단일 행 또는 단일 열 구분
                rows, cols = rng.shape
                if rows > 1 and cols == 1:
                    # 단일 열 → 각 값을 행으로
                    data = [[v] for v in raw]
                else:
                    # 단일 행
                    data = [raw]

            return ActionResult(
                success=True,
                data=data,
                message=f"범위 {cell_range} 읽기 완료 ({len(data)}행)",
            )
        except Exception as e:
            return ActionResult(success=False, error=f"READ_RANGE 실패: {e}")


class WriteDataHandler(ActionHandler):
    """WRITE_DATA: 데이터를 양식의 지정 위치에 쓴다.

    params:
        source: 데이터 소스 ($변수명 참조 또는 직접 리스트)
        target_file: 대상 워크북 alias
        target_sheet: 대상 시트명
        target_start: 시작 셀 (예: "B3")
        column_mapping: 소스 컬럼 → 대상 열 매핑 (선택)
            예: {"원본컬럼A": "B", "원본컬럼B": "C"}
    """

    def execute(self, params: dict[str, Any], ctx: EngineContext) -> ActionResult:
        source = params.get("source", [])
        target_file = params.get("target_file", "")
        target_sheet = params.get("target_sheet", "")
        target_start = params.get("target_start", "A1")
        column_mapping = params.get("column_mapping")

        try:
            wb = ctx.get_workbook(target_file)
            ws = wb.sheets[target_sheet]

            # 시작 셀 파싱
            col_letter, start_row = _parse_coordinate(target_start)
            start_col = _col_letter_to_index(col_letter)

            if not isinstance(source, list) or len(source) == 0:
                return ActionResult(
                    success=False,
                    error="source 데이터가 비어있거나 리스트가 아닙니다.",
                )

            rows_written = 0
            for i, row_data in enumerate(source):
                current_row = start_row + i

                if isinstance(row_data, dict):
                    if column_mapping:
                        # 매핑 기반 쓰기
                        for src_col, tgt_col_letter in column_mapping.items():
                            tgt_col_idx = _col_letter_to_index(tgt_col_letter)
                            ws.range((current_row, tgt_col_idx)).value = row_data.get(
                                src_col
                            )
                    else:
                        # 순서대로 쓰기
                        for j, value in enumerate(row_data.values()):
                            ws.range((current_row, start_col + j)).value = value
                elif isinstance(row_data, (list, tuple)):
                    for j, value in enumerate(row_data):
                        ws.range((current_row, start_col + j)).value = value

                rows_written += 1

            return ActionResult(
                success=True,
                data={"rows_written": rows_written},
                message=f"{rows_written}행 쓰기 완료 → {target_sheet}!{target_start}",
            )
        except Exception as e:
            return ActionResult(success=False, error=f"WRITE_DATA 실패: {e}")


class ClearRangeHandler(ActionHandler):
    """CLEAR_RANGE: 지정 범위의 셀 값을 비운다.

    params:
        file: 워크북 alias
        sheet: 시트명
        range: 셀 범위 (예: "B3:F100")
    """

    def execute(self, params: dict[str, Any], ctx: EngineContext) -> ActionResult:
        alias = params.get("file", "")
        sheet_name = params.get("sheet", "")
        cell_range = params.get("range", "")

        try:
            wb = ctx.get_workbook(alias)
            ws = wb.sheets[sheet_name]

            rng = ws.range(cell_range)

            # 클리어 전 비어있지 않은 셀 수 세기
            raw = rng.value
            cells_cleared = 0
            if raw is not None:
                if isinstance(raw, list):
                    for item in raw:
                        if isinstance(item, list):
                            cells_cleared += sum(1 for v in item if v is not None)
                        elif item is not None:
                            cells_cleared += 1
                else:
                    cells_cleared = 1

            rng.clear_contents()

            return ActionResult(
                success=True,
                data={"cells_cleared": cells_cleared},
                message=f"범위 {cell_range} 클리어 완료 ({cells_cleared}셀)",
            )
        except Exception as e:
            return ActionResult(success=False, error=f"CLEAR_RANGE 실패: {e}")


class RecalculateHandler(ActionHandler):
    """RECALCULATE: 워크북의 수식 재계산을 트리거한다.

    params:
        file: 워크북 alias

    Note: xlwings는 Excel 엔진을 통해 실제 수식 재계산을 수행한다.
    """

    def execute(self, params: dict[str, Any], ctx: EngineContext) -> ActionResult:
        alias = params.get("file", "")

        try:
            wb = ctx.get_workbook(alias)
            wb.app.calculation = "automatic"
            wb.app.calculate()
            return ActionResult(
                success=True,
                message=f"수식 재계산 완료: {alias}",
            )
        except Exception as e:
            return ActionResult(success=False, error=f"RECALCULATE 실패: {e}")
