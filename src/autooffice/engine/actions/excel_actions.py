"""엑셀 조작 ACTION 핸들러: READ_COLUMNS, READ_RANGE, WRITE_DATA, CLEAR_RANGE, RECALCULATE, FIND_DATE_COLUMN, COPY_RANGE, AGGREGATE_RANGE, FIND_ANCHOR."""

from __future__ import annotations

import logging
import re
from datetime import date, datetime
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


def _col_index_to_letter(idx: int) -> str:
    """1-based 인덱스를 컬럼 문자로 변환: 1 -> 'A', 27 -> 'AA'."""
    result: list[str] = []
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        result.append(chr(remainder + ord("A")))
    return "".join(reversed(result))


def _try_parse_date(value: Any, year: int | None = None) -> date | None:
    """셀 값을 date로 파싱 시도. 실패 시 None 반환.

    인접 셀의 형식(type)을 자동 감지하여 다양한 날짜 표현을 인식한다.

    지원 형식:
    - datetime 객체 → date로 변환
    - date 객체 → 그대로 반환
    - float/int (Excel 시리얼 날짜) → date 변환 (1~2958465 범위)
    - 문자열 형식:
      - "M/D", "MM/DD" (슬래시 구분)
      - "M-D", "MM-DD" (하이픈 구분)
      - "M.D", "MM.DD" (점 구분)
      - "YYYY-MM-DD", "YYYY/MM/DD", "YYYY.MM.DD" (연도 포함)
      - "M월D일", "MM월DD일" (한국어)
    """
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    # Excel 시리얼 날짜 (float/int)
    if isinstance(value, (int, float)) and 1 <= value <= 2958465:
        try:
            # Excel epoch: 1899-12-30 (1900 날짜 체계, Lotus 1-2-3 호환 버그 포함)
            from datetime import timedelta
            excel_epoch = date(1899, 12, 30)
            return excel_epoch + timedelta(days=int(value))
        except (ValueError, OverflowError):
            return None

    if isinstance(value, str):
        s = value.strip()
        y = year or date.today().year

        # "YYYY-MM-DD" / "YYYY/MM/DD" / "YYYY.MM.DD"
        m = re.match(r"^(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})$", s)
        if m:
            try:
                return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                return None

        # "M/D" / "M-D" / "M.D" (연도 없는 형식)
        m = re.match(r"^(\d{1,2})[/\-.](\d{1,2})$", s)
        if m:
            try:
                return date(y, int(m.group(1)), int(m.group(2)))
            except ValueError:
                return None

        # "M월D일" / "M월 D일" (한국어 형식)
        m = re.match(r"^(\d{1,2})\s*월\s*(\d{1,2})\s*일$", s)
        if m:
            try:
                return date(y, int(m.group(1)), int(m.group(2)))
            except ValueError:
                return None

    return None


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


class FindDateColumnHandler(ActionHandler):
    """FIND_DATE_COLUMN: 시트에서 날짜 패턴을 탐색하여 대상 열을 결정한다.

    지정된 scan_range 내에서 날짜 값이 있는 셀을 찾고,
    날짜가 가장 많은 행을 date_header_row로 결정한다.
    오늘 날짜와 매칭되는 열이 있으면 해당 열을,
    없으면 마지막 날짜 다음 열을 추론하여 반환한다.

    params:
        workbook: 워크북 alias
        sheet: 시트명
        scan_range: 날짜를 탐색할 셀 범위 (예: "A1:ZZ10")
        date: 대상 날짜 (ISO 형식 "YYYY-MM-DD", 예: "2026-03-29")

    returns (store_as):
        column: 대상 열 문자 (예: "CP")
        date_row: 날짜가 발견된 행 번호
    """

    def execute(self, params: dict[str, Any], ctx: EngineContext) -> ActionResult:
        workbook = params.get("workbook", "")
        sheet_name = params.get("sheet", "")
        scan_range = params.get("scan_range", "A1:ZZ10")
        date_str = params.get("date", "")

        if not date_str:
            return ActionResult(
                success=False,
                error="date 파라미터가 필요합니다. ISO 형식(YYYY-MM-DD)으로 입력하세요.",
            )

        try:
            target_date = date.fromisoformat(date_str)
        except ValueError:
            return ActionResult(
                success=False,
                error=f"date 형식 오류: '{date_str}'. ISO 형식(YYYY-MM-DD)으로 입력하세요.",
            )

        try:
            wb = ctx.get_workbook(workbook)
            ws = wb.sheets[sheet_name]

            # scan_range 읽기
            rng = ws.range(scan_range)
            raw = rng.value
            if raw is None:
                return ActionResult(success=False, error="scan_range가 비어있습니다.")

            # 2D 리스트로 정규화
            if not isinstance(raw, list):
                raw = [[raw]]
            elif not isinstance(raw[0], list):
                raw = [raw]

            start_row = rng.row
            start_col = rng.column

            # 행별 날짜 셀 수집: {row: [(col_idx, date), ...]}
            date_cells_by_row: dict[int, list[tuple[int, date]]] = {}
            for r_offset, row_data in enumerate(raw):
                actual_row = start_row + r_offset
                if row_data is None:
                    continue
                dates_in_row: list[tuple[int, date]] = []
                for c_offset, cell_val in enumerate(row_data):
                    actual_col = start_col + c_offset
                    parsed = _try_parse_date(cell_val, year=target_date.year)
                    if parsed is not None:
                        dates_in_row.append((actual_col, parsed))
                if dates_in_row:
                    date_cells_by_row[actual_row] = dates_in_row

            if not date_cells_by_row:
                return ActionResult(
                    success=False,
                    error="scan_range에서 날짜를 찾을 수 없습니다.",
                )

            # 날짜가 가장 많은 행 = date header row
            date_row = max(date_cells_by_row, key=lambda r: len(date_cells_by_row[r]))
            date_cells = sorted(date_cells_by_row[date_row], key=lambda x: x[0])

            # 오늘 날짜 정확히 매칭
            for col_idx, d in date_cells:
                if d == target_date:
                    col_letter = _col_index_to_letter(col_idx)
                    return ActionResult(
                        success=True,
                        data={"column": col_letter, "date_row": date_row},
                        message=f"날짜({target_date}) 발견: {col_letter}{date_row}",
                    )

            # 매칭 실패 → 패턴 기반 추론
            last_col_idx, last_date = date_cells[-1]
            days_gap = (target_date - last_date).days

            if days_gap <= 0:
                return ActionResult(
                    success=False,
                    error=(
                        f"대상 날짜({target_date})가 마지막 날짜({last_date})보다 "
                        f"이전이거나 같습니다."
                    ),
                )

            # 열 간격 패턴 감지 (기본: 1일 = 1열)
            col_per_day = 1.0
            if len(date_cells) >= 2:
                col_intervals = [
                    date_cells[i][0] - date_cells[i - 1][0]
                    for i in range(1, len(date_cells))
                ]
                day_intervals = [
                    (date_cells[i][1] - date_cells[i - 1][1]).days
                    for i in range(1, len(date_cells))
                ]
                # 가장 빈번한 패턴 사용
                if day_intervals and all(d > 0 for d in day_intervals):
                    typical_col = max(set(col_intervals), key=col_intervals.count)
                    typical_day = max(set(day_intervals), key=day_intervals.count)
                    if typical_day > 0:
                        col_per_day = typical_col / typical_day

            next_col_idx = last_col_idx + round(days_gap * col_per_day)
            col_letter = _col_index_to_letter(next_col_idx)

            if days_gap > 7:
                logger.warning(
                    "마지막 날짜(%s)와 대상 날짜(%s) 간격이 %d일입니다.",
                    last_date, target_date, days_gap,
                )

            return ActionResult(
                success=True,
                data={"column": col_letter, "date_row": date_row},
                message=(
                    f"날짜({target_date}) 열 추론: {col_letter}{date_row} "
                    f"(마지막: {last_date}→{_col_index_to_letter(last_col_idx)}열)"
                ),
            )
        except Exception as e:
            return ActionResult(success=False, error=f"FIND_DATE_COLUMN 실패: {e}")


class CopyRangeHandler(ActionHandler):
    """COPY_RANGE: 시트 간 범위 복사 (값 또는 수식).

    같은 워크북 내에서 소스 시트의 범위를 읽어
    타겟 시트의 지정 열/행에 붙여넣는다.
    paste_type=values이면 수식의 계산 결과(값)만 복사한다.

    params:
        workbook: 워크북 alias
        source_sheet: 소스 시트명
        source_range: 소스 범위 (예: "D7:D48")
        target_sheet: 타겟 시트명
        target_column: 타겟 열 문자 (예: "CP")
        target_start_row: 타겟 기준 행 (date_row 등)
        row_offset: 기준 행에서의 오프셋 (기본: 0)
        paste_type: 붙여넣기 유형 ("values" | "formulas" | "all", 기본: "values")
    """

    def execute(self, params: dict[str, Any], ctx: EngineContext) -> ActionResult:
        workbook = params.get("workbook", "")
        source_sheet = params.get("source_sheet", "")
        source_range = params.get("source_range", "")
        target_sheet = params.get("target_sheet", "")
        target_column = params.get("target_column", "")
        target_start_row = params.get("target_start_row", 1)
        row_offset = params.get("row_offset", 0)
        paste_type = params.get("paste_type", "values")

        try:
            wb = ctx.get_workbook(workbook)
            ws_src = wb.sheets[source_sheet]
            ws_tgt = wb.sheets[target_sheet]

            src_rng = ws_src.range(source_range)

            # 소스 데이터 읽기
            if paste_type == "formulas":
                raw = src_rng.formula
            else:
                raw = src_rng.value  # 계산된 값

            # 소스 범위 행/열 수 파악
            src_rows, src_cols = src_rng.shape

            # 데이터를 2D 리스트로 정규화
            # xlwings는 단일 셀 → scalar, 단일 열 → 1D list, 다중 열 → 2D list 반환
            if not isinstance(raw, list):
                # 단일 셀
                values_2d: list[list] = [[raw]]
            elif src_rows == 1 and src_cols > 1:
                # 단일 행: 1D list → 1행 2D
                values_2d = [raw]
            elif src_cols == 1:
                # 단일 열: 1D list → 열 벡터 2D
                values_2d = [[v] for v in raw]
            else:
                # 이미 2D
                values_2d = raw

            num_rows = len(values_2d)
            num_cols_out = len(values_2d[0]) if values_2d else 1

            # 타겟 시작 위치 계산
            actual_start_row = int(target_start_row) + int(row_offset)
            target_col_idx = _col_letter_to_index(target_column)

            # 타겟 범위 설정
            tgt_rng = ws_tgt.range(
                (actual_start_row, target_col_idx),
                (actual_start_row + num_rows - 1, target_col_idx + num_cols_out - 1),
            )

            if paste_type == "formulas":
                tgt_rng.formula = values_2d
            else:
                tgt_rng.value = values_2d

            return ActionResult(
                success=True,
                data={"rows_copied": num_rows},
                message=(
                    f"복사 완료: {source_sheet}!{source_range} → "
                    f"{target_sheet}!{target_column}{actual_start_row} "
                    f"({num_rows}행, {paste_type})"
                ),
            )
        except Exception as e:
            return ActionResult(success=False, error=f"COPY_RANGE 실패: {e}")


class AggregateRangeHandler(ActionHandler):
    """AGGREGATE_RANGE: 여러 열의 데이터를 행별로 집계하여 대상 열에 쓴다.

    지정된 열 범위(source_columns_start~source_columns_end)의 데이터를
    행별로 집계(합계/평균/개수)하여 target_column에 기록한다.
    주간 누계, 월간 누계 등 다수 열 데이터의 행별 집계에 사용한다.

    params:
        workbook: 워크북 alias
        sheet: 시트명
        source_columns_start: 집계 시작 열 문자 (예: "BT")
        source_columns_end: 집계 끝 열 문자 (예: "CQ")
        source_start_row: 데이터 시작 행 (예: 7)
        source_end_row: 데이터 끝 행 (예: 48)
        target_column: 결과를 쓸 열 문자 (예: "BS")
        target_start_row: 결과 시작 행 (기본: source_start_row과 동일)
        method: 집계 방법 ("sum" | "average" | "count", 기본: "sum")

    returns (store_as):
        rows_aggregated: 집계된 행 수
        method: 사용된 집계 방법
    """

    def execute(self, params: dict[str, Any], ctx: EngineContext) -> ActionResult:
        workbook = params.get("workbook", "")
        sheet_name = params.get("sheet", "")
        src_col_start = params.get("source_columns_start", "")
        src_col_end = params.get("source_columns_end", "")
        src_start_row = params.get("source_start_row", 1)
        src_end_row = params.get("source_end_row", 1)
        target_column = params.get("target_column", "")
        target_start_row = params.get("target_start_row", src_start_row)
        method = params.get("method", "sum")

        if not src_col_start or not src_col_end:
            return ActionResult(
                success=False,
                error="source_columns_start, source_columns_end 파라미터가 필요합니다.",
            )
        if not target_column:
            return ActionResult(
                success=False,
                error="target_column 파라미터가 필요합니다.",
            )

        try:
            wb = ctx.get_workbook(workbook)
            ws = wb.sheets[sheet_name]

            start_col_idx = _col_letter_to_index(src_col_start)
            end_col_idx = _col_letter_to_index(src_col_end)

            if end_col_idx < start_col_idx:
                return ActionResult(
                    success=False,
                    error=(
                        f"source_columns_end({src_col_end})가 "
                        f"source_columns_start({src_col_start})보다 앞에 있습니다."
                    ),
                )

            src_start_row = int(src_start_row)
            src_end_row = int(src_end_row)
            target_start_row = int(target_start_row)
            num_rows = src_end_row - src_start_row + 1

            if num_rows <= 0:
                return ActionResult(
                    success=False,
                    error=f"행 범위가 올바르지 않습니다: {src_start_row}~{src_end_row}",
                )

            # 소스 범위 일괄 읽기
            src_rng = ws.range(
                (src_start_row, start_col_idx),
                (src_end_row, end_col_idx),
            )
            raw = src_rng.value

            # 2D 리스트로 정규화
            num_cols = end_col_idx - start_col_idx + 1
            if not isinstance(raw, list):
                raw = [[raw]]
            elif num_rows == 1 and num_cols > 1:
                raw = [raw]
            elif num_cols == 1:
                raw = [[v] for v in raw]

            # 행별 집계
            results: list[list[Any]] = []
            for row_data in raw:
                # 숫자 값만 필터링
                nums = [v for v in row_data if isinstance(v, (int, float))]
                if method == "sum":
                    agg = sum(nums) if nums else 0
                elif method == "average":
                    agg = sum(nums) / len(nums) if nums else 0
                elif method == "count":
                    agg = len(nums)
                else:
                    agg = sum(nums) if nums else 0
                results.append([agg])

            # 결과 쓰기
            tgt_col_idx = _col_letter_to_index(target_column)
            tgt_rng = ws.range(
                (target_start_row, tgt_col_idx),
                (target_start_row + num_rows - 1, tgt_col_idx),
            )
            tgt_rng.value = results

            return ActionResult(
                success=True,
                data={"rows_aggregated": num_rows, "method": method},
                message=(
                    f"집계 완료: {sheet_name}!"
                    f"{src_col_start}{src_start_row}:{src_col_end}{src_end_row} → "
                    f"{target_column}{target_start_row} "
                    f"({num_rows}행, {method})"
                ),
            )
        except Exception as e:
            return ActionResult(success=False, error=f"AGGREGATE_RANGE 실패: {e}")


class FindAnchorHandler(ActionHandler):
    """FIND_ANCHOR: scan_range 내에서 텍스트를 검색하여 셀 위치를 반환한다.

    하드코딩 좌표 대신 텍스트 기준점을 찾아 후속 스텝에서
    $anchor.column, $anchor.row로 참조할 수 있도록 한다.

    params:
        workbook: 워크북 alias
        sheet: 시트명
        search_value: 찾을 텍스트
        scan_range: 탐색 범위 (기본: "A1:ZZ100")
        match_type: 매칭 방식 ("exact" | "contains" | "starts_with", 기본: "exact")

    returns (store_as):
        row: 발견된 행 번호 (1-based)
        column: 발견된 열 문자 (예: "B")
        cell: 셀 주소 (예: "B3")
        value: 매칭된 셀의 실제 값
    """

    def execute(self, params: dict[str, Any], ctx: EngineContext) -> ActionResult:
        workbook = params.get("workbook", "")
        sheet_name = params.get("sheet", "")
        search_value = params.get("search_value", "")
        scan_range = params.get("scan_range", "A1:ZZ100")
        match_type = params.get("match_type", "exact")

        if not search_value:
            return ActionResult(
                success=False,
                error="search_value 파라미터가 필요합니다.",
            )

        try:
            wb = ctx.get_workbook(workbook)
            ws = wb.sheets[sheet_name]

            rng = ws.range(scan_range)
            raw = rng.value
            if raw is None:
                return ActionResult(
                    success=False,
                    error="scan_range가 비어있습니다.",
                )

            # 2D 리스트로 정규화
            if not isinstance(raw, list):
                raw = [[raw]]
            elif not isinstance(raw[0], list):
                raw = [raw]

            start_row = rng.row
            start_col = rng.column
            search_lower = search_value.lower()

            for r_offset, row_data in enumerate(raw):
                if row_data is None:
                    continue
                for c_offset, cell_val in enumerate(row_data):
                    if cell_val is None:
                        continue

                    cell_str = str(cell_val).strip()
                    # 숫자의 .0 제거 (xlwings는 정수도 float로 반환)
                    if isinstance(cell_val, float) and cell_val == int(cell_val):
                        cell_str = str(int(cell_val))

                    cell_lower = cell_str.lower()

                    matched = False
                    if match_type == "exact":
                        matched = cell_lower == search_lower
                    elif match_type == "contains":
                        matched = search_lower in cell_lower
                    elif match_type == "starts_with":
                        matched = cell_lower.startswith(search_lower)

                    if matched:
                        actual_row = start_row + r_offset
                        actual_col = start_col + c_offset
                        col_letter = _col_index_to_letter(actual_col)
                        return ActionResult(
                            success=True,
                            data={
                                "row": actual_row,
                                "column": col_letter,
                                "cell": f"{col_letter}{actual_row}",
                                "value": cell_str,
                            },
                            message=(
                                f"앵커 발견: '{search_value}' → "
                                f"{col_letter}{actual_row} (값: {cell_str})"
                            ),
                        )

            return ActionResult(
                success=False,
                error=(
                    f"'{search_value}'을(를) {scan_range} 범위에서 "
                    f"찾을 수 없습니다. (match_type: {match_type})"
                ),
            )
        except Exception as e:
            return ActionResult(success=False, error=f"FIND_ANCHOR 실패: {e}")
