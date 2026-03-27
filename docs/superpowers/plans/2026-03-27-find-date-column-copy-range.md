# FIND_DATE_COLUMN + COPY_RANGE Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add date-pattern-aware column detection and inter-sheet value copy to the execution engine, enabling automated daily data append workflows.

**Architecture:** Two new action handlers (FIND_DATE_COLUMN, COPY_RANGE) plus EngineContext dot-notation variable resolution. FIND_DATE_COLUMN scans a sheet area for date patterns and returns `{"column": "CP", "date_row": 6}`. COPY_RANGE reads computed values from a source range and writes them to a target column at a row position relative to `date_row`. Dot notation (`$date_info.column`) enables dict field access in variable references.

**Tech Stack:** Python 3.11+, xlwings, Pydantic v2, pytest

---

## File Structure

| File | Action | Responsibility |
|------|--------|----------------|
| `src/autooffice/engine/context.py` | Modify | `resolve()` dot notation support |
| `src/autooffice/engine/runner.py` | Modify | `validate()` dot notation in `$var` reference check |
| `src/autooffice/engine/actions/excel_actions.py` | Modify | Add `_col_index_to_letter`, `_try_parse_date`, `FindDateColumnHandler`, `CopyRangeHandler` |
| `src/autooffice/engine/actions/__init__.py` | Modify | Register 2 new handlers |
| `src/autooffice/models/execution_plan.py` | Modify | Add 2 entries to `ActionType` enum |
| `schemas/execution_plan.schema.json` | Modify | Add 2 entries to action enum |
| `tests/test_engine/test_context.py` | Create | Dot notation resolve tests |
| `tests/test_engine/test_actions/test_excel_actions.py` | Modify | FindDateColumn + CopyRange tests |

---

### Task 1: EngineContext Dot Notation Variable Resolution

**Files:**
- Modify: `src/autooffice/engine/context.py:52-58` (resolve method)
- Modify: `src/autooffice/engine/runner.py:111-118` (validate variable check)
- Create: `tests/test_engine/test_context.py`

- [ ] **Step 1: Write failing tests for dot notation resolve**

Create `tests/test_engine/test_context.py`:

```python
"""EngineContext 변수 참조 해소 테스트."""

from __future__ import annotations

import pytest

from autooffice.engine.context import EngineContext


class TestResolveDotNotation:
    """$var.field 형식의 점 표기법 변수 참조 해소."""

    def test_simple_variable_still_works(self, tmp_path):
        """기존 $var 참조는 그대로 동작한다."""
        ctx = EngineContext(data_dir=tmp_path)
        ctx.store("name", "hello")
        assert ctx.resolve("$name") == "hello"

    def test_dot_notation_dict_field(self, tmp_path):
        """$var.field로 dict 필드에 접근한다."""
        ctx = EngineContext(data_dir=tmp_path)
        ctx.store("date_info", {"column": "CP", "date_row": 6})
        assert ctx.resolve("$date_info.column") == "CP"
        assert ctx.resolve("$date_info.date_row") == 6

    def test_dot_notation_missing_field(self, tmp_path):
        """존재하지 않는 필드 접근 시 KeyError."""
        ctx = EngineContext(data_dir=tmp_path)
        ctx.store("info", {"column": "CP"})
        with pytest.raises(KeyError, match="no_such_field"):
            ctx.resolve("$info.no_such_field")

    def test_dot_notation_on_non_dict(self, tmp_path):
        """dict가 아닌 변수에 점 표기법 사용 시 TypeError."""
        ctx = EngineContext(data_dir=tmp_path)
        ctx.store("count", 42)
        with pytest.raises(TypeError):
            ctx.resolve("$count.field")

    def test_non_variable_passthrough(self, tmp_path):
        """$ 없는 일반 문자열은 그대로 반환."""
        ctx = EngineContext(data_dir=tmp_path)
        assert ctx.resolve("plain_text") == "plain_text"
        assert ctx.resolve(123) == 123

    def test_resolve_params_with_dot_notation(self, tmp_path):
        """resolve_params에서 점 표기법이 동작한다."""
        ctx = EngineContext(data_dir=tmp_path)
        ctx.store("date_info", {"column": "CP", "date_row": 6})
        params = {
            "target_column": "$date_info.column",
            "target_row": "$date_info.date_row",
            "fixed_value": "hello",
        }
        resolved = ctx.resolve_params(params)
        assert resolved["target_column"] == "CP"
        assert resolved["target_row"] == 6
        assert resolved["fixed_value"] == "hello"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_engine/test_context.py -v`
Expected: FAIL — `test_dot_notation_dict_field` fails with `KeyError: "변수 'date_info.column'이(가) 정의되지 않았습니다."` because resolve() treats `date_info.column` as the entire variable name.

- [ ] **Step 3: Implement dot notation in resolve()**

Modify `src/autooffice/engine/context.py`, replace the `resolve` method (lines 52-58):

```python
def resolve(self, value: Any) -> Any:
    """값이 $variable_name 또는 $variable_name.field 형식이면 변수 저장소에서 참조를 해소한다."""
    if isinstance(value, str) and value.startswith("$"):
        var_path = value[1:]
        parts = var_path.split(".", 1)
        var_name = parts[0]
        if var_name not in self.variables:
            raise KeyError(f"변수 '{var_name}'이(가) 정의되지 않았습니다.")
        result = self.variables[var_name]
        if len(parts) > 1:
            field = parts[1]
            if not isinstance(result, dict):
                raise TypeError(
                    f"변수 '{var_name}'은(는) dict가 아니므로 "
                    f"'{field}' 필드에 접근할 수 없습니다."
                )
            if field not in result:
                raise KeyError(
                    f"변수 '{var_name}'에 '{field}' 필드가 없습니다."
                )
            return result[field]
        return result
    return value
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_engine/test_context.py -v`
Expected: All 7 tests PASS.

- [ ] **Step 5: Update runner validate() for dot notation**

Modify `src/autooffice/engine/runner.py`, lines 112-114. Change the variable reference extraction to strip the dot-notation field:

```python
            for key, val in step.params.items():
                if isinstance(val, str) and val.startswith("$"):
                    var_name = val[1:].split(".", 1)[0]
                    if var_name not in defined_vars:
                        errors.append(
                            f"Step {step.step}: 정의되지 않은 변수 '${var_name}' 참조 "
                            f"(param: {key})"
                        )
```

- [ ] **Step 6: Run full test suite to verify no regressions**

Run: `pytest tests/ -v`
Expected: All existing tests pass, plus new context tests pass.

- [ ] **Step 7: Commit**

```bash
git add src/autooffice/engine/context.py src/autooffice/engine/runner.py tests/test_engine/test_context.py
git commit -m "feat: add dot notation variable resolution ($var.field) to EngineContext"
```

---

### Task 2: FindDateColumnHandler

**Files:**
- Modify: `src/autooffice/engine/actions/excel_actions.py`
- Modify: `tests/test_engine/test_actions/test_excel_actions.py`

- [ ] **Step 1: Write failing tests for FindDateColumnHandler**

Add to `tests/test_engine/test_actions/test_excel_actions.py`:

```python
from datetime import date, datetime

from autooffice.engine.actions.excel_actions import (
    ClearRangeHandler,
    FindDateColumnHandler,
    ReadColumnsHandler,
    WriteDataHandler,
)


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
        dates = [
            f"{today.month}/{today.day - 2}",
            f"{today.month}/{today.day - 1}",
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
        dates = [
            f"{today.month}/{today.day - 2}",
            f"{today.month}/{today.day - 1}",
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
        dates = [
            datetime(today.year, today.month, today.day - 1),
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_engine/test_actions/test_excel_actions.py::TestFindDateColumn -v`
Expected: FAIL — `ImportError: cannot import name 'FindDateColumnHandler'`

- [ ] **Step 3: Implement utility functions and FindDateColumnHandler**

Add to `src/autooffice/engine/actions/excel_actions.py`.

First, add imports at the top (after existing imports):

```python
from datetime import date, datetime
```

Add utility functions after `_parse_coordinate()`:

```python
def _col_index_to_letter(idx: int) -> str:
    """1-based 인덱스를 컬럼 문자로 변환: 1 -> 'A', 27 -> 'AA'."""
    result: list[str] = []
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        result.append(chr(remainder + ord("A")))
    return "".join(reversed(result))


def _try_parse_date(value: Any, year: int | None = None) -> date | None:
    """셀 값을 date로 파싱 시도. 실패 시 None 반환.

    지원 형식:
    - datetime 객체 → date로 변환
    - date 객체 → 그대로 반환
    - 문자열 "M/D" 또는 "MM/DD" → 올해(또는 지정 연도) 기준 date 생성
    """
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        match = re.match(r"^(\d{1,2})/(\d{1,2})$", value.strip())
        if match:
            month, day = int(match.group(1)), int(match.group(2))
            y = year or date.today().year
            try:
                return date(y, month, day)
            except ValueError:
                return None
    return None
```

Add the handler class after `RecalculateHandler`:

```python
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
        date: 대상 날짜 ("today" 또는 "YYYY-MM-DD")

    returns (store_as):
        column: 대상 열 문자 (예: "CP")
        date_row: 날짜가 발견된 행 번호
    """

    def execute(self, params: dict[str, Any], ctx: EngineContext) -> ActionResult:
        workbook = params.get("workbook", "")
        sheet_name = params.get("sheet", "")
        scan_range = params.get("scan_range", "A1:ZZ10")
        date_str = params.get("date", "today")

        try:
            wb = ctx.get_workbook(workbook)
            ws = wb.sheets[sheet_name]

            # 대상 날짜 결정
            target_date = date.today() if date_str == "today" else date.fromisoformat(date_str)

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
            col_per_day = 1
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_engine/test_actions/test_excel_actions.py::TestFindDateColumn -v`
Expected: All 5 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add src/autooffice/engine/actions/excel_actions.py tests/test_engine/test_actions/test_excel_actions.py
git commit -m "feat: add FindDateColumnHandler with date pattern detection"
```

---

### Task 3: CopyRangeHandler

**Files:**
- Modify: `src/autooffice/engine/actions/excel_actions.py`
- Modify: `tests/test_engine/test_actions/test_excel_actions.py`

- [ ] **Step 1: Write failing tests for CopyRangeHandler**

Add to `tests/test_engine/test_actions/test_excel_actions.py`:

```python
from autooffice.engine.actions.excel_actions import (
    ClearRangeHandler,
    CopyRangeHandler,
    FindDateColumnHandler,
    ReadColumnsHandler,
    WriteDataHandler,
)


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
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_engine/test_actions/test_excel_actions.py::TestCopyRange -v`
Expected: FAIL — `ImportError: cannot import name 'CopyRangeHandler'`

- [ ] **Step 3: Implement CopyRangeHandler**

Add to `src/autooffice/engine/actions/excel_actions.py`, after `FindDateColumnHandler`:

```python
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

            # 단일 셀 → 리스트로 정규화
            if not isinstance(raw, list):
                values = [raw]
            else:
                # 단일 열(1D 리스트)인 경우 그대로 사용
                # 2D 리스트(다중 열)인 경우 그대로 사용
                values = raw

            # 타겟 시작 위치 계산
            actual_start_row = int(target_start_row) + int(row_offset)
            target_col_idx = _col_letter_to_index(target_column)

            # 데이터 쓰기
            if isinstance(values[0], list):
                # 다중 열: 2D 데이터
                num_rows = len(values)
                num_cols = len(values[0])
                tgt_rng = ws_tgt.range(
                    (actual_start_row, target_col_idx),
                    (actual_start_row + num_rows - 1, target_col_idx + num_cols - 1),
                )
            else:
                # 단일 열: 1D 데이터
                num_rows = len(values)
                tgt_rng = ws_tgt.range(
                    (actual_start_row, target_col_idx),
                    (actual_start_row + num_rows - 1, target_col_idx),
                )

            if paste_type == "formulas":
                tgt_rng.formula = values
            else:
                tgt_rng.value = values

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
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_engine/test_actions/test_excel_actions.py::TestCopyRange -v`
Expected: All 4 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add src/autooffice/engine/actions/excel_actions.py tests/test_engine/test_actions/test_excel_actions.py
git commit -m "feat: add CopyRangeHandler for inter-sheet value copy"
```

---

### Task 4: Registry, Enum, Schema Integration + Integration Test

**Files:**
- Modify: `src/autooffice/models/execution_plan.py:15-30` (ActionType enum)
- Modify: `src/autooffice/engine/actions/__init__.py` (registry + imports)
- Modify: `schemas/execution_plan.schema.json:100-113` (action enum)
- Modify: `tests/test_engine/test_actions/test_excel_actions.py` (integration test)

- [ ] **Step 1: Add to ActionType enum**

Modify `src/autooffice/models/execution_plan.py`, add two entries to `ActionType` after `RECALCULATE`:

```python
class ActionType(str, Enum):
    """실행 엔진이 지원하는 ACTION 타입."""

    OPEN_FILE = "OPEN_FILE"
    READ_COLUMNS = "READ_COLUMNS"
    READ_RANGE = "READ_RANGE"
    WRITE_DATA = "WRITE_DATA"
    CLEAR_RANGE = "CLEAR_RANGE"
    RECALCULATE = "RECALCULATE"
    FIND_DATE_COLUMN = "FIND_DATE_COLUMN"
    COPY_RANGE = "COPY_RANGE"
    SAVE_FILE = "SAVE_FILE"
    VALIDATE = "VALIDATE"
    FORMAT_MESSAGE = "FORMAT_MESSAGE"
    SEND_MESSENGER = "SEND_MESSENGER"
    SEND_EMAIL = "SEND_EMAIL"
    GENERATE_PPT = "GENERATE_PPT"
    LOG = "LOG"
```

- [ ] **Step 2: Register handlers in __init__.py**

Modify `src/autooffice/engine/actions/__init__.py`:

Update the import:
```python
from autooffice.engine.actions.excel_actions import (
    CopyRangeHandler,
    FindDateColumnHandler,
    ReadColumnsHandler,
    ReadRangeHandler,
    WriteDataHandler,
    ClearRangeHandler,
    RecalculateHandler,
)
```

Update the registry function:
```python
def build_default_registry() -> dict[str, ActionHandler]:
    """기본 ACTION 핸들러 레지스트리를 생성한다."""
    return {
        "OPEN_FILE": OpenFileHandler(),
        "SAVE_FILE": SaveFileHandler(),
        "READ_COLUMNS": ReadColumnsHandler(),
        "READ_RANGE": ReadRangeHandler(),
        "WRITE_DATA": WriteDataHandler(),
        "CLEAR_RANGE": ClearRangeHandler(),
        "RECALCULATE": RecalculateHandler(),
        "FIND_DATE_COLUMN": FindDateColumnHandler(),
        "COPY_RANGE": CopyRangeHandler(),
        "VALIDATE": ValidateHandler(),
        "FORMAT_MESSAGE": FormatMessageHandler(),
        "SEND_MESSENGER": SendMessengerHandler(),
        "LOG": LogHandler(),
    }
```

- [ ] **Step 3: Update JSON schema**

Modify `schemas/execution_plan.schema.json`, add to the action enum array (after `"RECALCULATE"`):

```json
"enum": [
    "OPEN_FILE",
    "READ_COLUMNS",
    "READ_RANGE",
    "WRITE_DATA",
    "CLEAR_RANGE",
    "RECALCULATE",
    "FIND_DATE_COLUMN",
    "COPY_RANGE",
    "SAVE_FILE",
    "VALIDATE",
    "FORMAT_MESSAGE",
    "SEND_MESSENGER",
    "SEND_EMAIL",
    "GENERATE_PPT",
    "LOG"
]
```

- [ ] **Step 4: Write integration test — full plan with FIND_DATE_COLUMN + COPY_RANGE**

Add to `tests/test_engine/test_actions/test_excel_actions.py`:

```python
class TestFindDateColumnCopyRangeIntegration:
    """FIND_DATE_COLUMN → COPY_RANGE 통합 시나리오 (dot notation 변수 참조 포함)."""

    def test_full_workflow(self, engine_ctx, runner, tmp_data_dir):
        """날짜 열 탐색 → 시트 간 값 복사 전체 흐름."""
        today = date.today()

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
        ws_sheet1.range("C4").value = f"{today.month}/{today.day - 2}"
        ws_sheet1.range("D4").value = f"{today.month}/{today.day - 1}"

        path = tmp_data_dir / "material.xlsx"
        wb.save(str(path))
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
```

- [ ] **Step 5: Run all tests**

Run: `pytest tests/ -v`
Expected: All tests PASS including the integration test.

- [ ] **Step 6: Commit**

```bash
git add src/autooffice/models/execution_plan.py src/autooffice/engine/actions/__init__.py schemas/execution_plan.schema.json tests/test_engine/test_actions/test_excel_actions.py
git commit -m "feat: integrate FIND_DATE_COLUMN + COPY_RANGE into action registry and schema"
```

---

## Execution Plan Example (for Claude skill reference)

After all code changes, a typical execution plan using the new actions:

```json
{
  "task_id": "daily_material_update",
  "description": "raw 데이터 복사 후 자재 계산식 결과를 오늘 날짜 열에 추가",
  "steps": [
    {"step": 1, "action": "OPEN_FILE", "params": {"file_path": "material.xlsx", "alias": "mat"}},
    {"step": 2, "action": "OPEN_FILE", "params": {"file_path": "raw.xlsx", "alias": "raw"}},
    {"step": 3, "action": "READ_RANGE", "params": {"file": "raw", "sheet": "Sheet1", "range": "A2:E100"}, "store_as": "raw_data"},
    {"step": 4, "action": "WRITE_DATA", "params": {"source": "$raw_data", "target_file": "mat", "target_sheet": "입력", "target_start": "H5"}},
    {"step": 5, "action": "RECALCULATE", "params": {"file": "mat"}},
    {
      "step": 6,
      "action": "FIND_DATE_COLUMN",
      "description": "Sheet1에서 오늘 날짜 열 탐색",
      "params": {
        "workbook": "mat",
        "sheet": "Sheet1",
        "scan_range": "A1:ZZ10",
        "date": "today"
      },
      "store_as": "date_info"
    },
    {
      "step": 7,
      "action": "COPY_RANGE",
      "description": "자재 계산식 D7:D48 결과 → Sheet1 오늘 날짜 열에 값 붙여넣기",
      "params": {
        "workbook": "mat",
        "source_sheet": "자재 계산식",
        "source_range": "D7:D48",
        "target_sheet": "Sheet1",
        "target_column": "$date_info.column",
        "target_start_row": "$date_info.date_row",
        "row_offset": 1,
        "paste_type": "values"
      }
    },
    {"step": 8, "action": "SAVE_FILE", "params": {"file": "mat"}},
    {"step": 9, "action": "LOG", "params": {"message": "일일 자재 업데이트 완료", "level": "info"}}
  ]
}
```
