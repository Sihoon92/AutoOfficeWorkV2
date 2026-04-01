# FIND_ANCHOR + FIND_DATE_RANGE + BuiltinResolver 확장 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 하드코딩 좌표를 텍스트 앵커 기반 탐색으로 대체하고, 날짜 범위 탐색과 LLM 없는 날짜 해소를 추가한다.

**Architecture:** 기존 ActionHandler ABC 패턴으로 두 핸들러(FindAnchorHandler, FindDateRangeHandler)를 추가하고, BuiltinDynamicResolver에 DATE_FUNCTIONS 딕셔너리를 도입하여 prompt 키워드 기반 날짜 계산을 로컬 처리한다.

**Tech Stack:** Python 3.11+, Pydantic v2, xlwings, pytest

---

### Task 1: BuiltinResolver 확장 — DATE_FUNCTIONS 딕셔너리

**Files:**
- Modify: `src/autooffice/engine/resolvers/builtin_resolver.py`
- Test: `tests/test_engine/test_resolvers/test_resolvers.py`

- [ ] **Step 1: 기존 테스트 통과 확인**

Run: `pytest tests/test_engine/test_resolvers/test_resolvers.py -v`
Expected: 모든 기존 테스트 PASS

- [ ] **Step 2: 새 테스트 작성**

`tests/test_engine/test_resolvers/test_resolvers.py`의 `TestBuiltinDynamicResolver` 클래스에 추가:

```python
def test_prompt_today(self):
    """prompt='today' → 오늘 날짜."""
    resolver = BuiltinDynamicResolver()
    declarations = {
        "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="today"),
    }
    result = resolver.resolve(declarations)
    assert result["d"] == date.today().isoformat()

def test_prompt_yesterday(self):
    """prompt='yesterday' → 어제 날짜."""
    resolver = BuiltinDynamicResolver()
    declarations = {
        "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="yesterday"),
    }
    result = resolver.resolve(declarations)
    expected = (date.today() - timedelta(days=1)).isoformat()
    assert result["d"] == expected

def test_prompt_this_week_monday(self):
    """prompt='this_week_monday' → 이번 주 월요일."""
    resolver = BuiltinDynamicResolver()
    declarations = {
        "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="this_week_monday"),
    }
    result = resolver.resolve(declarations)
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    assert result["d"] == monday.isoformat()

def test_prompt_this_week_sunday(self):
    """prompt='this_week_sunday' → 이번 주 일요일."""
    resolver = BuiltinDynamicResolver()
    declarations = {
        "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="this_week_sunday"),
    }
    result = resolver.resolve(declarations)
    today = date.today()
    sunday = today - timedelta(days=today.weekday()) + timedelta(days=6)
    assert result["d"] == sunday.isoformat()

def test_prompt_last_week_monday(self):
    """prompt='last_week_monday' → 지난주 월요일."""
    resolver = BuiltinDynamicResolver()
    declarations = {
        "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="last_week_monday"),
    }
    result = resolver.resolve(declarations)
    today = date.today()
    last_monday = today - timedelta(days=today.weekday() + 7)
    assert result["d"] == last_monday.isoformat()

def test_prompt_last_week_sunday(self):
    """prompt='last_week_sunday' → 지난주 일요일."""
    resolver = BuiltinDynamicResolver()
    declarations = {
        "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="last_week_sunday"),
    }
    result = resolver.resolve(declarations)
    today = date.today()
    last_sunday = today - timedelta(days=today.weekday() + 1)
    assert result["d"] == last_sunday.isoformat()

def test_prompt_this_month_start(self):
    """prompt='this_month_start' → 이번 달 1일."""
    resolver = BuiltinDynamicResolver()
    declarations = {
        "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="this_month_start"),
    }
    result = resolver.resolve(declarations)
    assert result["d"] == date.today().replace(day=1).isoformat()

def test_prompt_this_month_end(self):
    """prompt='this_month_end' → 이번 달 마지막 날."""
    resolver = BuiltinDynamicResolver()
    declarations = {
        "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="this_month_end"),
    }
    result = resolver.resolve(declarations)
    today = date.today()
    # 다음 달 1일 - 1일 = 이번 달 마지막
    if today.month == 12:
        next_month_start = date(today.year + 1, 1, 1)
    else:
        next_month_start = date(today.year, today.month + 1, 1)
    expected = (next_month_start - timedelta(days=1)).isoformat()
    assert result["d"] == expected

def test_prompt_last_month_start(self):
    """prompt='last_month_start' → 지난달 1일."""
    resolver = BuiltinDynamicResolver()
    declarations = {
        "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="last_month_start"),
    }
    result = resolver.resolve(declarations)
    this_month_start = date.today().replace(day=1)
    last_month_end = this_month_start - timedelta(days=1)
    assert result["d"] == last_month_end.replace(day=1).isoformat()

def test_prompt_last_month_end(self):
    """prompt='last_month_end' → 지난달 마지막 날."""
    resolver = BuiltinDynamicResolver()
    declarations = {
        "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="last_month_end"),
    }
    result = resolver.resolve(declarations)
    this_month_start = date.today().replace(day=1)
    expected = (this_month_start - timedelta(days=1)).isoformat()
    assert result["d"] == expected

def test_unknown_prompt_falls_back_to_today(self):
    """알 수 없는 prompt → today 대체."""
    resolver = BuiltinDynamicResolver()
    declarations = {
        "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="unknown_keyword"),
    }
    result = resolver.resolve(declarations)
    assert result["d"] == date.today().isoformat()

def test_prompt_case_insensitive(self):
    """prompt 대소문자 무시."""
    resolver = BuiltinDynamicResolver()
    declarations = {
        "d": DynamicParamSpec(type=DynamicParamType.DATE, prompt="This_Week_Monday"),
    }
    result = resolver.resolve(declarations)
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    assert result["d"] == monday.isoformat()
```

테스트 파일 상단 import에 `from datetime import timedelta` 추가 필요.

- [ ] **Step 3: 테스트 실패 확인**

Run: `pytest tests/test_engine/test_resolvers/test_resolvers.py::TestBuiltinDynamicResolver -v`
Expected: 새 테스트들 FAIL (기존 로직은 prompt를 무시하고 항상 today 반환)

- [ ] **Step 4: BuiltinResolver 구현**

`src/autooffice/engine/resolvers/builtin_resolver.py` 전체를 다음으로 교체:

```python
"""LLM 호출 없이 내장 함수로 해소 가능한 동적 파라미터 처리.

type=date는 prompt 키워드에 따라 다양한 날짜를 로컬에서 즉시 해소한다.
"""

from __future__ import annotations

import logging
from collections.abc import Callable
from datetime import date, timedelta
from typing import Any

from autooffice.engine.resolvers.base import DynamicResolver
from autooffice.models.execution_plan import DynamicParamSpec, DynamicParamType

logger = logging.getLogger(__name__)


def _this_week_monday() -> date:
    today = date.today()
    return today - timedelta(days=today.weekday())


def _this_month_start() -> date:
    return date.today().replace(day=1)


def _next_month_start() -> date:
    today = date.today()
    if today.month == 12:
        return date(today.year + 1, 1, 1)
    return date(today.year, today.month + 1, 1)


DATE_FUNCTIONS: dict[str, Callable[[], date]] = {
    "today": lambda: date.today(),
    "yesterday": lambda: date.today() - timedelta(days=1),
    "this_week_monday": _this_week_monday,
    "this_week_sunday": lambda: _this_week_monday() + timedelta(days=6),
    "last_week_monday": lambda: _this_week_monday() - timedelta(days=7),
    "last_week_sunday": lambda: _this_week_monday() - timedelta(days=1),
    "this_month_start": _this_month_start,
    "this_month_end": lambda: _next_month_start() - timedelta(days=1),
    "last_month_start": lambda: (_this_month_start() - timedelta(days=1)).replace(day=1),
    "last_month_end": lambda: _this_month_start() - timedelta(days=1),
}


class BuiltinDynamicResolver(DynamicResolver):
    """내장 해소기: LLM 없이 로컬에서 해소 가능한 파라미터를 처리한다.

    해소 가능한 유형:
    - type=date: prompt 키워드에 따라 날짜 계산
      (today, yesterday, this_week_monday, this_week_sunday,
       last_week_monday, last_week_sunday, this_month_start,
       this_month_end, last_month_start, last_month_end)

    해소하지 못한 파라미터는 결과에 포함하지 않는다 (체인의 다음 resolver에 위임).
    """

    def resolve(
        self, declarations: dict[str, DynamicParamSpec]
    ) -> dict[str, Any]:
        resolved: dict[str, Any] = {}

        for key, spec in declarations.items():
            if spec.type == DynamicParamType.DATE:
                func_key = spec.prompt.strip().lower()
                if func_key in DATE_FUNCTIONS:
                    value = DATE_FUNCTIONS[func_key]().isoformat()
                else:
                    value = date.today().isoformat()
                    logger.warning(
                        "알 수 없는 date prompt '%s', today로 대체", func_key
                    )
                resolved[key] = value
                logger.info(
                    "내장 해소: %s → %s (prompt: %s)", key, value, func_key
                )

        return resolved
```

- [ ] **Step 5: 테스트 통과 확인**

Run: `pytest tests/test_engine/test_resolvers/test_resolvers.py -v`
Expected: 기존 테스트 + 새 테스트 모두 PASS

- [ ] **Step 6: 커밋**

```bash
git add src/autooffice/engine/resolvers/builtin_resolver.py tests/test_engine/test_resolvers/test_resolvers.py
git commit -m "feat: extend BuiltinResolver with DATE_FUNCTIONS for week/month date calculation"
```

---

### Task 2: ActionType enum에 FIND_ANCHOR, FIND_DATE_RANGE 추가

**Files:**
- Modify: `src/autooffice/models/execution_plan.py:15-33`
- Modify: `skills/execution-plan-generator/references/pydantic_model.py:15-33`

- [ ] **Step 1: execution_plan.py의 ActionType enum에 추가**

`src/autooffice/models/execution_plan.py`의 `ActionType` 클래스에 추가:

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
    FIND_DATE_RANGE = "FIND_DATE_RANGE"
    FIND_ANCHOR = "FIND_ANCHOR"
    COPY_RANGE = "COPY_RANGE"
    AGGREGATE_RANGE = "AGGREGATE_RANGE"
    SAVE_FILE = "SAVE_FILE"
    VALIDATE = "VALIDATE"
    FORMAT_MESSAGE = "FORMAT_MESSAGE"
    SEND_MESSENGER = "SEND_MESSENGER"
    SEND_EMAIL = "SEND_EMAIL"
    GENERATE_PPT = "GENERATE_PPT"
    LOG = "LOG"
```

- [ ] **Step 2: skills 레퍼런스의 pydantic_model.py도 동일하게 동기화**

`skills/execution-plan-generator/references/pydantic_model.py`의 `ActionType` 클래스에 동일하게 `FIND_DATE_RANGE = "FIND_DATE_RANGE"`, `FIND_ANCHOR = "FIND_ANCHOR"` 추가.

- [ ] **Step 3: 기존 테스트 통과 확인**

Run: `pytest tests/ -v --timeout=60`
Expected: 모든 기존 테스트 PASS

- [ ] **Step 4: 커밋**

```bash
git add src/autooffice/models/execution_plan.py skills/execution-plan-generator/references/pydantic_model.py
git commit -m "feat: add FIND_ANCHOR and FIND_DATE_RANGE to ActionType enum"
```

---

### Task 3: FindAnchorHandler 구현

**Files:**
- Modify: `src/autooffice/engine/actions/excel_actions.py` (클래스 추가)
- Modify: `src/autooffice/engine/actions/__init__.py` (import + registry)
- Create: `tests/test_engine/test_actions/test_find_anchor.py`

- [ ] **Step 1: 테스트 파일 작성**

`tests/test_engine/test_actions/test_find_anchor.py`:

```python
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
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "search_value": "항목",
                "scan_range": "A1:F10",
                "match_type": "exact",
            },
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
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "search_value": "일별",
                "scan_range": "A1:F10",
                "match_type": "contains",
            },
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
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "search_value": "일별",
                "scan_range": "A1:F10",
                "match_type": "starts_with",
            },
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
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "search_value": "sales",
                "scan_range": "A1:F10",
                "match_type": "exact",
            },
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
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "search_value": "존재하지않는텍스트",
                "scan_range": "A1:F10",
                "match_type": "exact",
            },
            engine_ctx,
        )

        assert not result.success
        assert "찾을 수 없습니다" in result.error

    def test_returns_first_match(self, engine_ctx: EngineContext, anchor_excel: Path):
        """여러 매칭 중 첫 번째(좌상단 우선) 반환."""
        wb = engine_ctx.app.books.open(str(anchor_excel))
        engine_ctx.register_workbook("test", wb, anchor_excel)

        # "주" 가 포함된 셀: "주간"(C2), "14주"(E3)
        handler = FindAnchorHandler()
        result = handler.execute(
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "search_value": "주",
                "scan_range": "A1:F10",
                "match_type": "contains",
            },
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
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "search_value": "100",
                "scan_range": "A1:F10",
                "match_type": "exact",
            },
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
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "scan_range": "A1:F10",
            },
            engine_ctx,
        )

        assert not result.success
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `pytest tests/test_engine/test_actions/test_find_anchor.py -v`
Expected: FAIL — `FindAnchorHandler` import 실패

- [ ] **Step 3: FindAnchorHandler 구현**

`src/autooffice/engine/actions/excel_actions.py` 파일 끝에 추가:

```python
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
```

- [ ] **Step 4: __init__.py에 등록**

`src/autooffice/engine/actions/__init__.py`의 import에 `FindAnchorHandler` 추가하고 `build_default_registry()`에 `"FIND_ANCHOR": FindAnchorHandler()` 추가.

- [ ] **Step 5: 테스트 통과 확인**

Run: `pytest tests/test_engine/test_actions/test_find_anchor.py -v`
Expected: 모든 테스트 PASS

- [ ] **Step 6: 커밋**

```bash
git add src/autooffice/engine/actions/excel_actions.py src/autooffice/engine/actions/__init__.py tests/test_engine/test_actions/test_find_anchor.py
git commit -m "feat: add FIND_ANCHOR action handler for text-based cell search"
```

---

### Task 4: FindDateRangeHandler 구현

**Files:**
- Modify: `src/autooffice/engine/actions/excel_actions.py` (클래스 추가)
- Modify: `src/autooffice/engine/actions/__init__.py` (import + registry)
- Create: `tests/test_engine/test_actions/test_find_date_range.py`

- [ ] **Step 1: 테스트 파일 작성**

`tests/test_engine/test_actions/test_find_date_range.py`:

```python
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
    def test_full_week_match(
        self, engine_ctx: EngineContext, date_range_excel: Path
    ):
        """월~일 전체 매칭 (3/30 ~ 4/5, 실제 데이터는 3/30~4/3만 존재)."""
        wb = engine_ctx.app.books.open(str(date_range_excel))
        engine_ctx.register_workbook("test", wb, date_range_excel)

        handler = FindDateRangeHandler()
        result = handler.execute(
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "scan_range": "A1:J10",
                "start_date": "2026-03-30",
                "end_date": "2026-04-03",
            },
            engine_ctx,
        )

        assert result.success
        assert result.data["date_row"] == 3
        assert result.data["matched_count"] == 5
        assert len(result.data["columns"]) == 5
        assert result.data["missing_dates"] == []

    def test_partial_match_with_missing(
        self, engine_ctx: EngineContext, date_range_excel: Path
    ):
        """범위 내 일부 날짜 없음 → missing_dates에 포함."""
        wb = engine_ctx.app.books.open(str(date_range_excel))
        engine_ctx.register_workbook("test", wb, date_range_excel)

        handler = FindDateRangeHandler()
        # 3/30 ~ 4/5: 4/4, 4/5는 엑셀에 없음
        result = handler.execute(
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "scan_range": "A1:J10",
                "start_date": "2026-03-30",
                "end_date": "2026-04-05",
            },
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
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "scan_range": "A1:J10",
                "start_date": "2026-05-01",
                "end_date": "2026-05-07",
            },
            engine_ctx,
        )

        assert not result.success

    def test_single_day_range(
        self, engine_ctx: EngineContext, date_range_excel: Path
    ):
        """start_date == end_date (1일)."""
        wb = engine_ctx.app.books.open(str(date_range_excel))
        engine_ctx.register_workbook("test", wb, date_range_excel)

        handler = FindDateRangeHandler()
        result = handler.execute(
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "scan_range": "A1:J10",
                "start_date": "2026-04-01",
                "end_date": "2026-04-01",
            },
            engine_ctx,
        )

        assert result.success
        assert result.data["matched_count"] == 1
        assert len(result.data["columns"]) == 1
        assert result.data["start_column"] == result.data["end_column"]

    def test_month_boundary_cross(
        self, engine_ctx: EngineContext, date_range_excel: Path
    ):
        """월 경계 크로스 (3/29 ~ 4/2)."""
        wb = engine_ctx.app.books.open(str(date_range_excel))
        engine_ctx.register_workbook("test", wb, date_range_excel)

        handler = FindDateRangeHandler()
        result = handler.execute(
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "scan_range": "A1:J10",
                "start_date": "2026-03-29",
                "end_date": "2026-04-02",
            },
            engine_ctx,
        )

        assert result.success
        assert result.data["matched_count"] == 5  # 3/29, 3/30, 3/31, 4/1, 4/2
        assert result.data["missing_dates"] == []

    def test_start_end_column_order(
        self, engine_ctx: EngineContext, date_range_excel: Path
    ):
        """start_column < end_column 순서 보장."""
        wb = engine_ctx.app.books.open(str(date_range_excel))
        engine_ctx.register_workbook("test", wb, date_range_excel)

        handler = FindDateRangeHandler()
        result = handler.execute(
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "scan_range": "A1:J10",
                "start_date": "2026-03-28",
                "end_date": "2026-04-03",
            },
            engine_ctx,
        )

        assert result.success
        assert result.data["start_column"] == result.data["columns"][0]
        assert result.data["end_column"] == result.data["columns"][-1]

    def test_invalid_date_format(
        self, engine_ctx: EngineContext, date_range_excel: Path
    ):
        """잘못된 날짜 형식 → 실패."""
        wb = engine_ctx.app.books.open(str(date_range_excel))
        engine_ctx.register_workbook("test", wb, date_range_excel)

        handler = FindDateRangeHandler()
        result = handler.execute(
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "scan_range": "A1:J10",
                "start_date": "not-a-date",
                "end_date": "2026-04-03",
            },
            engine_ctx,
        )

        assert not result.success
        assert "형식 오류" in result.error
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `pytest tests/test_engine/test_actions/test_find_date_range.py -v`
Expected: FAIL — `FindDateRangeHandler` import 실패

- [ ] **Step 3: FindDateRangeHandler 구현**

`src/autooffice/engine/actions/excel_actions.py` 파일 끝에 추가:

```python
class FindDateRangeHandler(ActionHandler):
    """FIND_DATE_RANGE: 날짜 범위에 해당하는 열들을 한번에 탐색한다.

    scan_range 내에서 start_date~end_date에 해당하는 모든 날짜 열을 찾는다.
    FIND_DATE_COLUMN의 날짜 스캔 로직(_try_parse_date, 날짜 행 감지)을 재사용한다.

    params:
        workbook: 워크북 alias
        sheet: 시트명
        scan_range: 날짜 탐색 범위 (기본: "A1:ZZ10")
        start_date: 시작 날짜 ISO (예: "2026-03-30")
        end_date: 종료 날짜 ISO (예: "2026-04-05")

    returns (store_as):
        start_column: 매칭된 첫 번째 열 문자
        end_column: 매칭된 마지막 열 문자
        date_row: 날짜 헤더가 있는 행 번호
        columns: 매칭된 모든 열 문자 리스트
        matched_count: 매칭된 날짜 수
        missing_dates: 범위 내 못 찾은 날짜 ISO 리스트
    """

    def execute(self, params: dict[str, Any], ctx: EngineContext) -> ActionResult:
        workbook = params.get("workbook", "")
        sheet_name = params.get("sheet", "")
        scan_range = params.get("scan_range", "A1:ZZ10")
        start_date_str = params.get("start_date", "")
        end_date_str = params.get("end_date", "")

        # 날짜 파싱
        try:
            start_date = date.fromisoformat(start_date_str)
        except (ValueError, TypeError):
            return ActionResult(
                success=False,
                error=f"start_date 형식 오류: '{start_date_str}'. ISO 형식(YYYY-MM-DD)으로 입력하세요.",
            )
        try:
            end_date = date.fromisoformat(end_date_str)
        except (ValueError, TypeError):
            return ActionResult(
                success=False,
                error=f"end_date 형식 오류: '{end_date_str}'. ISO 형식(YYYY-MM-DD)으로 입력하세요.",
            )

        if end_date < start_date:
            return ActionResult(
                success=False,
                error=f"end_date({end_date})가 start_date({start_date})보다 이전입니다.",
            )

        try:
            wb = ctx.get_workbook(workbook)
            ws = wb.sheets[sheet_name]

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

            # 행별 날짜 셀 수집 (FIND_DATE_COLUMN과 동일한 로직)
            date_cells_by_row: dict[int, list[tuple[int, date]]] = {}
            for r_offset, row_data in enumerate(raw):
                actual_row = start_row + r_offset
                if row_data is None:
                    continue
                dates_in_row: list[tuple[int, date]] = []
                for c_offset, cell_val in enumerate(row_data):
                    actual_col = start_col + c_offset
                    parsed = _try_parse_date(cell_val, year=start_date.year)
                    if parsed is not None:
                        dates_in_row.append((actual_col, parsed))
                if dates_in_row:
                    date_cells_by_row[actual_row] = dates_in_row

            if not date_cells_by_row:
                return ActionResult(
                    success=False,
                    error="scan_range에서 날짜를 찾을 수 없습니다.",
                )

            # 날짜가 가장 많은 행 = date_row
            date_row = max(date_cells_by_row, key=lambda r: len(date_cells_by_row[r]))
            date_cells = sorted(date_cells_by_row[date_row], key=lambda x: x[0])

            # 범위 내 날짜 필터
            matched: list[tuple[int, date]] = [
                (col_idx, d)
                for col_idx, d in date_cells
                if start_date <= d <= end_date
            ]

            if not matched:
                return ActionResult(
                    success=False,
                    error=(
                        f"{start_date}~{end_date} 범위의 날짜를 "
                        f"찾을 수 없습니다."
                    ),
                )

            columns = [_col_index_to_letter(col_idx) for col_idx, _ in matched]
            matched_dates = {d for _, d in matched}

            # missing_dates 계산
            all_dates_in_range: set[date] = set()
            current = start_date
            while current <= end_date:
                all_dates_in_range.add(current)
                current += timedelta(days=1)
            missing = sorted(all_dates_in_range - matched_dates)

            return ActionResult(
                success=True,
                data={
                    "start_column": columns[0],
                    "end_column": columns[-1],
                    "date_row": date_row,
                    "columns": columns,
                    "matched_count": len(columns),
                    "missing_dates": [d.isoformat() for d in missing],
                },
                message=(
                    f"날짜 범위 탐색 완료: {start_date}~{end_date} → "
                    f"{columns[0]}~{columns[-1]} ({len(columns)}열 매칭, "
                    f"{len(missing)}개 누락)"
                ),
            )
        except Exception as e:
            return ActionResult(success=False, error=f"FIND_DATE_RANGE 실패: {e}")
```

`from datetime import timedelta` import를 파일 상단에 추가 필요 (이미 `from datetime import date, datetime`이 있으므로 `timedelta` 추가).

- [ ] **Step 4: __init__.py에 등록**

`src/autooffice/engine/actions/__init__.py`의 import에 `FindDateRangeHandler` 추가하고 `build_default_registry()`에 `"FIND_DATE_RANGE": FindDateRangeHandler()` 추가.

- [ ] **Step 5: 테스트 통과 확인**

Run: `pytest tests/test_engine/test_actions/test_find_date_range.py -v`
Expected: 모든 테스트 PASS

- [ ] **Step 6: 커밋**

```bash
git add src/autooffice/engine/actions/excel_actions.py src/autooffice/engine/actions/__init__.py tests/test_engine/test_actions/test_find_date_range.py
git commit -m "feat: add FIND_DATE_RANGE action handler for date range column search"
```

---

### Task 5: skills 레퍼런스 문서 업데이트

**Files:**
- Modify: `skills/execution-plan-generator/references/action_reference.md`

- [ ] **Step 1: action_reference.md에 새 액션 문서 추가**

파일 끝에 다음 섹션 추가:

```markdown
## FIND_ANCHOR

scan_range 내에서 텍스트를 검색하여 셀 위치를 반환한다.
하드코딩 좌표 대신 텍스트 기준점으로 동적 위치를 결정할 때 사용한다.

### params
| 이름 | 타입 | 필수 | 기본값 | 설명 |
|------|------|------|--------|------|
| workbook | str | O | | 워크북 alias |
| sheet | str | O | | 시트명 |
| search_value | str | O | | 찾을 텍스트 |
| scan_range | str | | "A1:ZZ100" | 탐색 범위 |
| match_type | str | | "exact" | "exact", "contains", "starts_with" |

### store_as 반환값
| 필드 | 타입 | 설명 |
|------|------|------|
| row | int | 발견된 행 번호 |
| column | str | 발견된 열 문자 |
| cell | str | 셀 주소 (예: "B3") |
| value | str | 매칭된 셀의 실제 값 |

### 사용 예시
```json
{
    "action": "FIND_ANCHOR",
    "params": {
        "workbook": "defect",
        "sheet": "Sheet1",
        "search_value": "주간",
        "match_type": "contains"
    },
    "store_as": "weekly_anchor"
}
```

## FIND_DATE_RANGE

날짜 범위에 해당하는 열들을 한번에 탐색한다.
FIND_DATE_COLUMN의 날짜 스캔 로직을 재사용하여 여러 날짜를 일괄 매칭한다.

### params
| 이름 | 타입 | 필수 | 기본값 | 설명 |
|------|------|------|--------|------|
| workbook | str | O | | 워크북 alias |
| sheet | str | O | | 시트명 |
| scan_range | str | | "A1:ZZ10" | 날짜 탐색 범위 |
| start_date | str | O | | 시작 날짜 (ISO: YYYY-MM-DD) |
| end_date | str | O | | 종료 날짜 (ISO: YYYY-MM-DD) |

### store_as 반환값
| 필드 | 타입 | 설명 |
|------|------|------|
| start_column | str | 매칭된 첫 번째 열 문자 |
| end_column | str | 매칭된 마지막 열 문자 |
| date_row | int | 날짜 헤더 행 번호 |
| columns | list[str] | 매칭된 모든 열 문자 리스트 |
| matched_count | int | 매칭된 날짜 수 |
| missing_dates | list[str] | 범위 내 못 찾은 날짜 ISO 리스트 |

### 사용 예시
```json
{
    "action": "FIND_DATE_RANGE",
    "params": {
        "workbook": "defect",
        "sheet": "Sheet1",
        "start_date": "{{dynamic:week_start}}",
        "end_date": "{{dynamic:week_end}}"
    },
    "store_as": "week_cols"
}
```
```

- [ ] **Step 2: 커밋**

```bash
git add skills/execution-plan-generator/references/action_reference.md skills/execution-plan-generator/references/pydantic_model.py
git commit -m "docs: add FIND_ANCHOR and FIND_DATE_RANGE to action reference"
```

---

### Task 6: 전체 회귀 테스트

- [ ] **Step 1: 전체 테스트 실행**

Run: `pytest tests/ -v --timeout=120`
Expected: 모든 테스트 PASS

- [ ] **Step 2: 실패 시 수정 후 재실행**

실패한 테스트가 있으면 원인 파악 후 수정. 주로 enum 변경에 따른 기존 테스트 호환성 확인.

- [ ] **Step 3: 최종 커밋 (수정 사항이 있는 경우만)**

```bash
git add -A
git commit -m "fix: resolve regression issues from new action handlers"
```
