# INIT_MONTH_COLUMNS Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 매월 1일에 새 월의 일별 날짜 헤더 컬럼을 시트 오른쪽에 자동 추가하는 INIT_MONTH_COLUMNS action 구현

**Architecture:** 기존 ActionHandler 패턴을 따라 `InitMonthColumnsHandler`를 구현한다. scan_range에서 마지막 날짜 헤더를 찾고, 다음 열부터 해당 월의 일수만큼 날짜 헤더를 쓴다. 이미 존재하면 `already_exists: true`를 반환하여 중복 방지. 기존 `_try_parse_date`, `_col_index_to_letter` 유틸리티를 재사용한다.

**Tech Stack:** Python 3.12, xlwings, Pydantic v2, pytest

---

### Task 1: ActionType enum에 INIT_MONTH_COLUMNS 추가

**Files:**
- Modify: `src/autooffice/models/execution_plan.py:17-36` (ActionType enum)

- [ ] **Step 1: enum 값 추가**

`src/autooffice/models/execution_plan.py`의 `ActionType` enum에 추가:

```python
class ActionType(str, Enum):
    # ... 기존 항목 ...
    EXTRACT_DATE = "EXTRACT_DATE"
    INIT_MONTH_COLUMNS = "INIT_MONTH_COLUMNS"  # 추가
    SAVE_FILE = "SAVE_FILE"
```

`EXTRACT_DATE`와 `SAVE_FILE` 사이에 삽입한다.

- [ ] **Step 2: 확인**

Run: `python -c "from autooffice.models.execution_plan import ActionType; print(ActionType.INIT_MONTH_COLUMNS.value)"`
Expected: `INIT_MONTH_COLUMNS`

- [ ] **Step 3: Commit**

```bash
git add src/autooffice/models/execution_plan.py
git commit -m "feat: add INIT_MONTH_COLUMNS to ActionType enum"
```

---

### Task 2: InitMonthColumnsHandler 테스트 작성 (TDD red)

**Files:**
- Create: `tests/test_engine/test_actions/test_init_month_columns.py`

- [ ] **Step 1: 테스트 fixture 작성**

```python
"""INIT_MONTH_COLUMNS 액션 핸들러 테스트."""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pytest
import xlwings as xw

from autooffice.engine.actions.excel_actions import InitMonthColumnsHandler
from autooffice.engine.context import EngineContext


@pytest.fixture
def month_excel(tmp_data_dir: Path, xw_app: xw.App) -> Path:
    """INIT_MONTH_COLUMNS 테스트용 엑셀.

    Sheet1 구조 (헤더 Row 6):
      M1..M12 | W1..W4 | 4/1 | 4/2 | ... | 4/30
      (테스트 간소화: M1, W1, 4/1~4/5 만 생성)
    """
    wb = xw_app.books.add()
    ws = wb.sheets[0]
    ws.name = "Sheet1"

    # 월별/주별 헤더 (간소화)
    ws.range("A6").value = "M1"
    ws.range("B6").value = "W1"

    # 일별 헤더: 4/1 ~ 4/5 (C6:G6)
    ws.range("C6").value = "4/1"
    ws.range("D6").value = "4/2"
    ws.range("E6").value = "4/3"
    ws.range("F6").value = "4/4"
    ws.range("G6").value = "4/5"

    path = tmp_data_dir / "month_test.xlsx"
    wb.save(str(path))
    wb.close()
    return path


@pytest.fixture
def full_april_excel(tmp_data_dir: Path, xw_app: xw.App) -> Path:
    """4월 전체(30일) 헤더가 있는 테스트 엑셀."""
    wb = xw_app.books.add()
    ws = wb.sheets[0]
    ws.name = "Sheet1"

    # 일별 헤더: 4/1 ~ 4/30 (A6:AD6)
    for day in range(1, 31):
        col_idx = day  # 1-based
        ws.range((6, col_idx)).value = f"4/{day}"

    path = tmp_data_dir / "full_april_test.xlsx"
    wb.save(str(path))
    wb.close()
    return path
```

- [ ] **Step 2: 기본 동작 테스트 작성**

```python
class TestInitMonthColumns:
    def test_creates_new_month_headers(
        self, engine_ctx: EngineContext, month_excel: Path
    ):
        """새 월 헤더를 마지막 날짜 다음 열에 생성한다."""
        engine_ctx.open_workbook("test", month_excel)
        handler = InitMonthColumnsHandler()

        result = handler.execute(
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "header_row": 6,
                "scan_range": "A6:ZZ6",
                "month_date": "2026-05-01",
                "date_format": "M/D",
            },
            engine_ctx,
        )

        assert result.success is True
        assert result.data["days_created"] == 31
        assert result.data["already_exists"] is False
        assert "start_column" in result.data
        assert "end_column" in result.data

        # 실제 셀 확인: 첫 번째 새 헤더 = "5/1"
        wb = engine_ctx.get_workbook("test")
        ws = wb.sheets["Sheet1"]
        start_col = result.data["start_column"]
        cell_value = ws.range(f"{start_col}6").value
        assert "5" in str(cell_value) and "1" in str(cell_value)

    def test_already_exists_returns_true(
        self, engine_ctx: EngineContext, month_excel: Path
    ):
        """이미 해당 월의 날짜가 존재하면 already_exists=True."""
        engine_ctx.open_workbook("test", month_excel)
        handler = InitMonthColumnsHandler()

        # 4월 헤더가 이미 있으므로 4월 요청 → already_exists
        result = handler.execute(
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "header_row": 6,
                "scan_range": "A6:ZZ6",
                "month_date": "2026-04-01",
                "date_format": "M/D",
            },
            engine_ctx,
        )

        assert result.success is True
        assert result.data["already_exists"] is True
        assert result.data["days_created"] == 0

    def test_returns_correct_column_range(
        self, engine_ctx: EngineContext, full_april_excel: Path
    ):
        """반환된 start_column~end_column이 정확히 해당 월 일수만큼."""
        engine_ctx.open_workbook("test", full_april_excel)
        handler = InitMonthColumnsHandler()

        result = handler.execute(
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "header_row": 6,
                "scan_range": "A6:ZZ6",
                "month_date": "2026-05-01",
                "date_format": "M/D",
            },
            engine_ctx,
        )

        assert result.success is True
        assert result.data["days_created"] == 31  # 5월 = 31일

    def test_february_28_days(
        self, engine_ctx: EngineContext, month_excel: Path
    ):
        """2월(비윤년) → 28일 생성."""
        engine_ctx.open_workbook("test", month_excel)
        handler = InitMonthColumnsHandler()

        result = handler.execute(
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "header_row": 6,
                "scan_range": "A6:ZZ6",
                "month_date": "2026-02-01",
                "date_format": "M/D",
            },
            engine_ctx,
        )

        assert result.success is True
        assert result.data["days_created"] == 28

    def test_february_29_days_leap_year(
        self, engine_ctx: EngineContext, month_excel: Path
    ):
        """2월(윤년) → 29일 생성."""
        engine_ctx.open_workbook("test", month_excel)
        handler = InitMonthColumnsHandler()

        result = handler.execute(
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "header_row": 6,
                "scan_range": "A6:ZZ6",
                "month_date": "2028-02-01",  # 2028 = 윤년
                "date_format": "M/D",
            },
            engine_ctx,
        )

        assert result.success is True
        assert result.data["days_created"] == 29

    def test_missing_workbook_fails(self, engine_ctx: EngineContext):
        """존재하지 않는 workbook → 실패."""
        handler = InitMonthColumnsHandler()
        result = handler.execute(
            {
                "workbook": "nonexistent",
                "sheet": "Sheet1",
                "header_row": 6,
                "scan_range": "A6:ZZ6",
                "month_date": "2026-05-01",
            },
            engine_ctx,
        )
        assert result.success is False

    def test_invalid_month_date_fails(
        self, engine_ctx: EngineContext, month_excel: Path
    ):
        """잘못된 날짜 형식 → 실패."""
        engine_ctx.open_workbook("test", month_excel)
        handler = InitMonthColumnsHandler()

        result = handler.execute(
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "header_row": 6,
                "scan_range": "A6:ZZ6",
                "month_date": "invalid-date",
            },
            engine_ctx,
        )
        assert result.success is False

    def test_empty_scan_range_creates_from_next_column(
        self, engine_ctx: EngineContext, tmp_data_dir: Path, xw_app: xw.App
    ):
        """scan_range에 날짜가 없으면 scan_range 시작 열부터 생성."""
        wb = xw_app.books.add()
        ws = wb.sheets[0]
        ws.name = "Sheet1"
        path = tmp_data_dir / "empty_test.xlsx"
        wb.save(str(path))
        wb.close()

        engine_ctx.open_workbook("test", path)
        handler = InitMonthColumnsHandler()

        result = handler.execute(
            {
                "workbook": "test",
                "sheet": "Sheet1",
                "header_row": 6,
                "scan_range": "C6:ZZ6",
                "month_date": "2026-05-01",
                "date_format": "M/D",
            },
            engine_ctx,
        )

        assert result.success is True
        assert result.data["days_created"] == 31
        assert result.data["start_column"] == "C"
```

- [ ] **Step 3: 테스트 실행 (실패 확인)**

Run: `pytest tests/test_engine/test_actions/test_init_month_columns.py -v`
Expected: FAIL (ImportError - `InitMonthColumnsHandler` 미존재)

- [ ] **Step 4: Commit**

```bash
git add tests/test_engine/test_actions/test_init_month_columns.py
git commit -m "test: add INIT_MONTH_COLUMNS handler tests (red)"
```

---

### Task 3: InitMonthColumnsHandler 구현 (TDD green)

**Files:**
- Modify: `src/autooffice/engine/actions/excel_actions.py` (하단에 클래스 추가)

- [ ] **Step 1: 핸들러 구현**

`excel_actions.py` 하단에 추가:

```python
class InitMonthColumnsHandler(ActionHandler):
    """INIT_MONTH_COLUMNS: 새 월의 일별 날짜 헤더 컬럼을 오른쪽에 추가한다.

    scan_range에서 기존 날짜 헤더의 마지막 열을 찾고,
    다음 열부터 해당 월의 일수만큼 날짜 헤더를 쓴다.
    해당 월의 날짜가 이미 존재하면 already_exists=True를 반환한다.

    params:
        workbook: 워크북 alias
        sheet: 시트명
        header_row: 날짜 헤더가 있는 행 번호
        scan_range: 기존 날짜를 탐색할 셀 범위 (예: "A6:ZZ6")
        month_date: 생성할 월의 날짜 (ISO "YYYY-MM-DD", 일은 무시됨)
        date_format: 헤더 날짜 표시 형식 (기본 "M/D")

    returns (store_as):
        start_column: 새 월 첫 번째 열 문자
        end_column: 새 월 마지막 열 문자
        days_created: 생성된 일수 (이미 존재하면 0)
        already_exists: 해당 월 날짜가 이미 존재하는지 여부
    """

    def execute(self, params: dict[str, Any], ctx: EngineContext) -> ActionResult:
        workbook = params.get("workbook", "")
        sheet_name = params.get("sheet", "")
        header_row = params.get("header_row", 1)
        scan_range = params.get("scan_range", "A1:ZZ1")
        month_date_str = params.get("month_date", "")
        date_format = params.get("date_format", "M/D")

        # month_date 파싱
        try:
            month_date = date.fromisoformat(month_date_str)
        except (ValueError, TypeError):
            return ActionResult(
                success=False,
                error=f"month_date 형식 오류: '{month_date_str}'. ISO 형식(YYYY-MM-DD)으로 입력하세요.",
            )

        # 해당 월의 일수 계산
        target_year = month_date.year
        target_month = month_date.month
        if target_month == 12:
            days_in_month = (date(target_year + 1, 1, 1) - date(target_year, 12, 1)).days
        else:
            days_in_month = (date(target_year, target_month + 1, 1) - date(target_year, target_month, 1)).days

        try:
            wb = ctx.get_workbook(workbook)
            ws = wb.sheets[sheet_name]
        except Exception as e:
            return ActionResult(success=False, error=f"워크북/시트 접근 실패: {e}")

        # scan_range에서 기존 날짜 탐색
        try:
            rng = ws.range(scan_range)
            raw = rng.value
        except Exception as e:
            return ActionResult(success=False, error=f"scan_range 읽기 실패: {e}")

        # 1D 리스트로 정규화 (단일 행)
        if raw is None:
            cells = []
        elif not isinstance(raw, list):
            cells = [raw]
        else:
            cells = raw

        start_col_idx = rng.column
        last_date_col_idx = start_col_idx - 1  # 날짜 없으면 scan 시작 직전

        # 기존 날짜 탐색 및 해당 월 존재 여부 확인
        existing_months: set[tuple[int, int]] = set()  # (year, month)
        for c_offset, cell_val in enumerate(cells):
            parsed = _try_parse_date(cell_val, year=target_year)
            if parsed is not None:
                actual_col = start_col_idx + c_offset
                last_date_col_idx = max(last_date_col_idx, actual_col)
                existing_months.add((parsed.year, parsed.month))

        # 이미 해당 월 존재 여부 체크
        if (target_year, target_month) in existing_months:
            return ActionResult(
                success=True,
                data={
                    "start_column": "",
                    "end_column": "",
                    "days_created": 0,
                    "already_exists": True,
                },
                message=f"{target_year}년 {target_month}월 헤더가 이미 존재합니다.",
            )

        # 새 열 시작 위치
        new_start_col_idx = last_date_col_idx + 1
        if last_date_col_idx < start_col_idx:
            # 날짜가 하나도 없으면 scan_range 시작 열부터
            new_start_col_idx = start_col_idx

        # 날짜 헤더 쓰기
        headers = []
        for day in range(1, days_in_month + 1):
            d = date(target_year, target_month, day)
            if date_format == "M/D":
                headers.append(f"{d.month}/{d.day}")
            elif date_format == "MM/DD":
                headers.append(f"{d.month:02d}/{d.day:02d}")
            elif date_format == "YYYY-MM-DD":
                headers.append(d.isoformat())
            else:
                headers.append(f"{d.month}/{d.day}")

        # 한 행에 가로로 쓰기
        start_cell = f"{_col_index_to_letter(new_start_col_idx)}{header_row}"
        ws.range(start_cell).value = [headers]  # [[...]] = 1행 N열

        new_end_col_idx = new_start_col_idx + days_in_month - 1
        start_col_letter = _col_index_to_letter(new_start_col_idx)
        end_col_letter = _col_index_to_letter(new_end_col_idx)

        logger.info(
            "월 헤더 생성: %s년 %s월, %s~%s (%d일)",
            target_year, target_month,
            start_col_letter, end_col_letter, days_in_month,
        )

        return ActionResult(
            success=True,
            data={
                "start_column": start_col_letter,
                "end_column": end_col_letter,
                "days_created": days_in_month,
                "already_exists": False,
            },
            message=f"{target_year}년 {target_month}월 헤더 생성 완료: {start_col_letter}~{end_col_letter} ({days_in_month}일)",
        )
```

- [ ] **Step 2: 테스트 실행 (통과 확인)**

Run: `pytest tests/test_engine/test_actions/test_init_month_columns.py -v`
Expected: ALL PASS

- [ ] **Step 3: Commit**

```bash
git add src/autooffice/engine/actions/excel_actions.py
git commit -m "feat: implement InitMonthColumnsHandler"
```

---

### Task 4: 레지스트리 등록

**Files:**
- Modify: `src/autooffice/engine/actions/__init__.py`

- [ ] **Step 1: import 및 레지스트리 추가**

`__init__.py`에서:

import 추가:
```python
from autooffice.engine.actions.excel_actions import (
    # ... 기존 import ...
    InitMonthColumnsHandler,
)
```

`build_default_registry()`에 추가:
```python
"INIT_MONTH_COLUMNS": InitMonthColumnsHandler(),
```

`EXTRACT_DATE` 항목 아래에 배치.

- [ ] **Step 2: 등록 확인 테스트**

Run: `python -c "from autooffice.engine.actions import build_default_registry; r = build_default_registry(); print('INIT_MONTH_COLUMNS' in r)"`
Expected: `True`

- [ ] **Step 3: 전체 테스트 실행**

Run: `pytest tests/test_engine/test_actions/test_init_month_columns.py -v`
Expected: ALL PASS

- [ ] **Step 4: Commit**

```bash
git add src/autooffice/engine/actions/__init__.py
git commit -m "feat: register INIT_MONTH_COLUMNS in action registry"
```

---

### Task 5: 스킬 문서 업데이트

**Files:**
- Modify: `skills/execution-plan-generator/references/action_reference.md`
- Modify: `skills/execution-plan-generator/SKILL.md`

- [ ] **Step 1: action_reference.md에 INIT_MONTH_COLUMNS 섹션 추가**

EXTRACT_DATE 섹션 아래에 추가:

```markdown
## INIT_MONTH_COLUMNS

새 월의 일별 날짜 헤더 컬럼을 시트 오른쪽에 추가한다.
scan_range에서 기존 날짜 헤더의 마지막 열을 찾고, 다음 열부터 해당 월 일수만큼 헤더를 생성한다.
해당 월 헤더가 이미 존재하면 생성하지 않는다 (멱등성 보장).

### params
| 이름 | 필수 | 기본값 | 설명 |
|------|------|--------|------|
| workbook | O | | 워크북 alias |
| sheet | O | | 시트명 |
| header_row | O | | 날짜 헤더가 있는 행 번호 |
| scan_range | O | | 기존 날짜를 탐색할 범위 (예: "BT6:ZZZ6") |
| month_date | O | | 생성할 월 (ISO "YYYY-MM-DD", 일은 무시됨) |
| date_format | - | "M/D" | 헤더 표시 형식: "M/D", "MM/DD", "YYYY-MM-DD" |

### store_as 반환값
| 필드 | 타입 | 설명 |
|------|------|------|
| start_column | str | 새 월 첫 번째 열 문자 |
| end_column | str | 새 월 마지막 열 문자 |
| days_created | int | 생성된 일수 (이미 존재하면 0) |
| already_exists | bool | 해당 월 날짜가 이미 존재하는지 |

### 사용 예시
\```json
{
    "step": 2,
    "action": "INIT_MONTH_COLUMNS",
    "description": "5월 일별 헤더 추가",
    "params": {
        "workbook": "report",
        "sheet": "Sheet1",
        "header_row": 6,
        "scan_range": "BT6:ZZZ6",
        "month_date": "{{dynamic:execution_date}}",
        "date_format": "M/D"
    },
    "store_as": "new_month"
}
\```

이후 step에서 참조:
- `$new_month.start_column` → 새 월 첫 번째 열
- `$new_month.end_column` → 새 월 마지막 열
- `$new_month.already_exists` → 이미 존재 여부 (VALIDATE로 분기 가능)
```

- [ ] **Step 2: SKILL.md의 ActionType 목록 및 self-check 업데이트**

SKILL.md에서 action 수를 19개 → 20개로, ActionType 목록에 `INIT_MONTH_COLUMNS` 추가.
Phase 3의 패턴 매칭에 추가:
```
- "다음 달 일별 데이터 준비" → INIT_MONTH_COLUMNS로 새 월 헤더 생성
- "월별 보고서 초기화" → INIT_MONTH_COLUMNS + CLEAR_RANGE 조합
```

- [ ] **Step 3: Commit**

```bash
git add skills/execution-plan-generator/references/action_reference.md skills/execution-plan-generator/SKILL.md
git commit -m "docs: add INIT_MONTH_COLUMNS to skill references"
```

---

### Task 6: 기존 테스트 전체 통과 확인

- [ ] **Step 1: 전체 테스트 실행**

Run: `pytest --timeout=30 -v`
Expected: 기존 테스트 + 신규 테스트 모두 PASS (runner.py의 기존 Excel COM 실패 2개는 무관)

- [ ] **Step 2: Commit (최종)**

```bash
git add -A
git commit -m "feat: INIT_MONTH_COLUMNS action - auto-create monthly date headers"
```
