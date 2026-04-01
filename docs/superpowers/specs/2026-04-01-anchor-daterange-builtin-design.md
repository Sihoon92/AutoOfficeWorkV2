# FIND_ANCHOR + FIND_DATE_RANGE + BuiltinResolver 확장 설계

## 배경

현재 실행계획에는 셀 좌표가 하드코딩된다 (`"scan_range": "A1:ZZ10"`, `"source_range": "D7:D9"`).
사용자가 엑셀 헤더 위치를 밀거나 양식을 수정하면 계획과 실제가 불일치하여 실행이 실패한다.

또한 날짜 관련 동적 변수(오늘, 이번 주, 이번 달)를 모두 LLM에 요청하고 있으나,
이는 Python 표준 라이브러리만으로 충분히 계산 가능하다. LLM 호출을 제거하고
BuiltinResolver를 확장하여 비용 없이 로컬에서 해소한다.

## 목표

1. **FIND_ANCHOR**: 텍스트 기반 기준점 탐색으로 하드코딩 좌표 탈피
2. **FIND_DATE_RANGE**: 날짜 범위에 해당하는 열들을 한번에 탐색
3. **BuiltinResolver 확장**: 주/월 날짜 계산을 LLM 없이 로컬 해소

## 변경 범위

| 파일 | 변경 내용 |
|------|----------|
| `src/autooffice/models/execution_plan.py` | ActionType enum에 `FIND_ANCHOR`, `FIND_DATE_RANGE` 추가 |
| `src/autooffice/engine/actions/excel_actions.py` | `FindAnchorHandler`, `FindDateRangeHandler` 구현 |
| `src/autooffice/engine/actions/__init__.py` | import + registry 등록 |
| `src/autooffice/engine/resolvers/builtin_resolver.py` | DATE_FUNCTIONS 딕셔너리 기반으로 확장 |
| `skills/.../pydantic_model.py` | ActionType enum 동기화 |
| `skills/.../action_reference.md` | 새 액션 문서 추가 |
| `tests/test_engine/test_actions/test_find_anchor.py` | FIND_ANCHOR 테스트 |
| `tests/test_engine/test_actions/test_find_date_range.py` | FIND_DATE_RANGE 테스트 |
| `tests/test_engine/test_resolvers/test_builtin_resolver.py` | BuiltinResolver 확장 테스트 |

---

## 1. FIND_ANCHOR

### 목적

scan_range 내에서 특정 텍스트를 검색하여 셀 위치를 반환한다.
후속 스텝이 이 위치를 `$anchor.column`, `$anchor.row`로 참조하여
하드코딩 좌표 없이 동적으로 작업 범위를 결정한다.

### 인터페이스

```
params:
    workbook: str          — 워크북 alias
    sheet: str             — 시트명
    search_value: str      — 찾을 텍스트 (예: "자재코드", "일별", "4월")
    scan_range: str        — 탐색 범위 (기본: "A1:ZZ100")
    match_type: str        — 매칭 방식 (기본: "exact")
                             "exact": 정확히 일치
                             "contains": 부분 문자열 포함
                             "starts_with": 접두사 일치

returns (store_as):
    row: int               — 발견된 행 번호 (1-based)
    column: str            — 발견된 열 문자 (예: "B")
    cell: str              — 셀 주소 (예: "B3")
    value: str             — 매칭된 셀의 실제 값
```

### 핵심 로직

```python
def execute(self, params, ctx) -> ActionResult:
    # 1. scan_range를 2D 배열로 읽기
    # 2. 셀 순회 (행 우선, 왼쪽→오른쪽, 위→아래)
    # 3. 셀 값을 str로 변환 후 match_type에 따라 비교
    #    - exact:       str(cell_value).strip() == search_value
    #    - contains:    search_value in str(cell_value)
    #    - starts_with: str(cell_value).strip().startswith(search_value)
    # 4. 첫 번째 매칭 셀의 좌표를 반환
    # 5. 못 찾으면 ActionResult(success=False)
```

### 매칭 시 고려사항

- 셀 값이 숫자/날짜인 경우 `str()`로 변환 후 비교
- 대소문자 구분 없이 비교 (한국어는 해당 없으나, 영문 헤더 대비)
- None 셀은 건너뛰기
- 첫 번째 매칭만 반환 (단일 앵커 원칙)

### 사용 예시

```json
{
    "step": 1,
    "action": "FIND_ANCHOR",
    "description": "일별 섹션 시작점 찾기",
    "params": {
        "workbook": "defect",
        "sheet": "Sheet1",
        "search_value": "일별",
        "scan_range": "A1:ZZ10",
        "match_type": "contains"
    },
    "store_as": "daily_anchor",
    "expect": {"condition": "success", "value": true},
    "on_fail": "STOP"
}
```

---

## 2. FIND_DATE_RANGE

### 목적

scan_range 내에서 시작일~종료일 범위에 해당하는 모든 날짜 열을 한번에 찾는다.
`FIND_DATE_COLUMN`의 날짜 스캔 로직(`_try_parse_date`, 날짜 행 자동 감지)을 재사용한다.

### 인터페이스

```
params:
    workbook: str          — 워크북 alias
    sheet: str             — 시트명
    scan_range: str        — 날짜 탐색 범위 (기본: "A1:ZZ10")
    start_date: str        — 시작 날짜 ISO (예: "2026-03-30")
    end_date: str          — 종료 날짜 ISO (예: "2026-04-05")

returns (store_as):
    start_column: str      — 매칭된 첫 번째 열 문자 (예: "BU")
    end_column: str        — 매칭된 마지막 열 문자 (예: "CA")
    date_row: int          — 날짜 헤더가 있는 행 번호
    columns: list[str]     — 매칭된 모든 열 문자 리스트
    matched_count: int     — 매칭된 날짜 수
    missing_dates: list[str] — 범위 내 못 찾은 날짜 ISO 리스트
```

### 핵심 로직

```python
def execute(self, params, ctx) -> ActionResult:
    # 1. start_date, end_date 파싱 및 검증
    # 2. scan_range를 2D 배열로 읽기
    # 3. _try_parse_date로 모든 셀 스캔 → 행별 날짜 셀 수집
    #    (FIND_DATE_COLUMN과 동일한 로직)
    # 4. 날짜가 가장 많은 행 = date_row (FIND_DATE_COLUMN과 동일)
    # 5. date_row에서 start_date <= 날짜 <= end_date 인 열만 필터
    # 6. 범위 내 모든 날짜를 생성하여 missing_dates 계산
    # 7. 매칭된 열이 0개이면 success=False
```

### _try_parse_date 재사용

기존 `FIND_DATE_COLUMN`에서 사용 중인 `_try_parse_date()` 함수를 그대로 사용한다.
지원 형식: datetime, date, Excel 시리얼 날짜, `"M/D"`, `"M-D"`, `"M.D"`,
`"YYYY-MM-DD"`, `"M월D일"` 등.

### 사용 예시

```json
{
    "step": 2,
    "action": "FIND_DATE_RANGE",
    "description": "이번 주 날짜 열 범위 찾기",
    "params": {
        "workbook": "defect",
        "sheet": "Sheet1",
        "scan_range": "A1:ZZ10",
        "start_date": "{{dynamic:week_start}}",
        "end_date": "{{dynamic:week_end}}"
    },
    "store_as": "week_cols"
}
```

### AGGREGATE_RANGE와의 연계

```json
{
    "step": 3,
    "action": "AGGREGATE_RANGE",
    "params": {
        "workbook": "defect",
        "sheet": "Sheet1",
        "source_columns_start": "$week_cols.start_column",
        "source_columns_end": "$week_cols.end_column",
        "source_start_row": 7,
        "source_end_row": 48,
        "target_column": "$weekly_anchor.column",
        "method": "sum"
    }
}
```

---

## 3. BuiltinResolver 확장

### 목적

현재 BuiltinResolver는 `type=date`일 때 `date.today().isoformat()`만 반환한다.
주/월 시작일, 종료일 등 Python으로 계산 가능한 날짜를 모두 로컬에서 해소하여
LLM 호출을 제거한다.

### 설계: prompt 기반 DATE_FUNCTIONS 딕셔너리

`DynamicParamSpec.prompt` 필드를 함수 키로 사용한다.

```python
DATE_FUNCTIONS: dict[str, Callable[[], date]] = {
    "today":              lambda: date.today(),
    "yesterday":          lambda: date.today() - timedelta(days=1),
    "this_week_monday":   lambda: date.today() - timedelta(days=date.today().weekday()),
    "this_week_sunday":   lambda: _this_week_monday() + timedelta(days=6),
    "last_week_monday":   lambda: _this_week_monday() - timedelta(days=7),
    "last_week_sunday":   lambda: _this_week_monday() - timedelta(days=1),
    "this_month_start":   lambda: date.today().replace(day=1),
    "this_month_end":     lambda: _next_month_start() - timedelta(days=1),
    "last_month_start":   lambda: (_this_month_start() - timedelta(days=1)).replace(day=1),
    "last_month_end":     lambda: _this_month_start() - timedelta(days=1),
}
```

### 해소 로직 변경

```python
def resolve(self, declarations):
    resolved = {}
    for key, spec in declarations.items():
        if spec.type == DynamicParamType.DATE:
            func_key = spec.prompt.strip().lower()
            if func_key in DATE_FUNCTIONS:
                value = DATE_FUNCTIONS[func_key]().isoformat()
            else:
                # 알 수 없는 prompt → 기본값 today
                value = date.today().isoformat()
                logger.warning("알 수 없는 date prompt '%s', today로 대체", func_key)
            resolved[key] = value
            logger.info("내장 해소: %s → %s (prompt: %s)", key, value, func_key)
    return resolved
```

### 실행계획에서의 사용

```json
{
    "dynamic_params": {
        "today": {
            "type": "date",
            "prompt": "today",
            "description": "오늘 날짜"
        },
        "week_start": {
            "type": "date",
            "prompt": "this_week_monday",
            "description": "이번 주 월요일"
        },
        "week_end": {
            "type": "date",
            "prompt": "this_week_sunday",
            "description": "이번 주 일요일"
        },
        "month_start": {
            "type": "date",
            "prompt": "this_month_start",
            "description": "이번 달 1일"
        },
        "month_end": {
            "type": "date",
            "prompt": "this_month_end",
            "description": "이번 달 마지막 날"
        }
    }
}
```

### LLM 호출 제거 효과

변경 전: 모든 `type=date`가 동일하게 `today` 반환, `type=lookup`은 LLM에 위임
변경 후: `type=date`에서 주/월 계산까지 처리 → 기존에 `type=lookup`으로
LLM에 보내던 날짜 계산이 `type=date`로 전환 가능

```
Before: date(today) + lookup(week_start→LLM) + lookup(week_end→LLM) = 1 API 호출
After:  date(today) + date(week_start) + date(week_end) = 0 API 호출
```

---

## 4. 전체 연계: 주간 불량률 합산 시나리오

### 사용자 요청

> "이번 주 일별 불량률 데이터를 합산해서 주간 칼럼에 넣어줘"

### Claude가 생성하는 실행계획

```json
{
    "task_id": "weekly_defect_aggregation",
    "description": "이번 주 일별 불량률 합산 → 주간 칼럼 기록",
    "created_by": "Claude",
    "created_at": "2026-04-01T09:00:00+09:00",
    "version": "1.0.0",
    "dynamic_params": {
        "week_start": {
            "type": "date",
            "prompt": "this_week_monday",
            "description": "이번 주 월요일"
        },
        "week_end": {
            "type": "date",
            "prompt": "this_week_sunday",
            "description": "이번 주 일요일"
        }
    },
    "inputs": {
        "defect_report": {
            "description": "불량률 관리 엑셀",
            "expected_format": "xlsx",
            "expected_sheets": ["Sheet1"]
        }
    },
    "steps": [
        {
            "step": 1,
            "action": "OPEN_FILE",
            "description": "불량률 엑셀 열기",
            "params": {"path": "불량률.xlsx", "alias": "defect"}
        },
        {
            "step": 2,
            "action": "FIND_DATE_RANGE",
            "description": "이번 주 날짜에 해당하는 열 범위 찾기",
            "params": {
                "workbook": "defect",
                "sheet": "Sheet1",
                "scan_range": "A1:ZZ10",
                "start_date": "{{dynamic:week_start}}",
                "end_date": "{{dynamic:week_end}}"
            },
            "store_as": "week_cols",
            "expect": {"condition": "success", "value": true},
            "on_fail": "STOP"
        },
        {
            "step": 3,
            "action": "FIND_ANCHOR",
            "description": "주간 합계 칼럼 위치 찾기",
            "params": {
                "workbook": "defect",
                "sheet": "Sheet1",
                "search_value": "주간",
                "scan_range": "A1:ZZ10",
                "match_type": "contains"
            },
            "store_as": "weekly_anchor",
            "expect": {"condition": "success", "value": true},
            "on_fail": "STOP"
        },
        {
            "step": 4,
            "action": "AGGREGATE_RANGE",
            "description": "이번 주 일별 데이터 행별 합산 → 주간 칼럼에 기록",
            "params": {
                "workbook": "defect",
                "sheet": "Sheet1",
                "source_columns_start": "$week_cols.start_column",
                "source_columns_end": "$week_cols.end_column",
                "source_start_row": "$week_cols.date_row",
                "source_end_row": 48,
                "target_column": "$weekly_anchor.column",
                "target_start_row": "$week_cols.date_row",
                "method": "sum"
            },
            "store_as": "agg_result"
        },
        {
            "step": 5,
            "action": "SAVE_FILE",
            "description": "결과 저장",
            "params": {"file": "defect"}
        }
    ],
    "final_output": {
        "type": "excel",
        "description": "주간 합산 결과가 기록된 불량률 엑셀"
    }
}
```

### 실행 흐름

```
Phase 1 (계획 생성) — Claude, 데이터 없음
│
│  빈 템플릿 구조 분석 → 실행계획 JSON 생성
│  좌표 하드코딩 대신 FIND_ANCHOR/FIND_DATE_RANGE 사용
│
▼
Phase 1.5 (동적 파라미터 해소) — BuiltinResolver, LLM 불필요
│
│  week_start: "this_week_monday" → "2026-03-30"
│  week_end:   "this_week_sunday" → "2026-04-05"
│  (0 API 호출)
│
▼
Phase 2 (실행) — Python 엔진, 실제 데이터
│
│  Step 1: OPEN_FILE
│  Step 2: FIND_DATE_RANGE(3/30 ~ 4/5)
│          → {start_column: "BU", end_column: "CA", date_row: 3,
│             columns: ["BU","BV","BW","BX","BY","BZ","CA"],
│             matched_count: 7, missing_dates: []}
│  Step 3: FIND_ANCHOR("주간")
│          → {row: 3, column: "BS", cell: "BS3"}
│  Step 4: AGGREGATE_RANGE(BU~CA, row 3~48, sum → BS)
│  Step 5: SAVE_FILE
```

---

## 5. 테스트 계획

### FIND_ANCHOR 테스트

| 테스트명 | 검증 내용 |
|---------|----------|
| `test_find_anchor_exact_match` | 정확한 텍스트로 찾기 |
| `test_find_anchor_contains` | 부분 문자열 매칭 ("일별현황" 에서 "일별") |
| `test_find_anchor_starts_with` | 접두사 매칭 |
| `test_find_anchor_case_insensitive` | 대소문자 무시 매칭 |
| `test_find_anchor_not_found` | 못 찾았을 때 `success=False` |
| `test_find_anchor_shifted_header` | 헤더가 밀린 상황에서도 정상 탐색 |
| `test_find_anchor_numeric_value` | 숫자 값을 문자열로 변환 후 매칭 |
| `test_find_anchor_returns_first_match` | 여러 매칭 중 첫 번째(좌상단) 반환 |

### FIND_DATE_RANGE 테스트

| 테스트명 | 검증 내용 |
|---------|----------|
| `test_find_date_range_full_week` | 월~일 7일 전부 매칭 |
| `test_find_date_range_partial_match` | 일부 날짜만 존재 시 missing_dates 확인 |
| `test_find_date_range_no_match` | 범위 내 날짜 없음 → `success=False` |
| `test_find_date_range_single_day` | start_date == end_date (1일 범위) |
| `test_find_date_range_month_boundary` | 월 경계 (3/29 ~ 4/3) 크로스 매칭 |
| `test_find_date_range_various_formats` | "3/30", "2026-03-30", "3월30일" 혼재 |
| `test_find_date_range_with_aggregate` | FIND_DATE_RANGE → AGGREGATE_RANGE 통합 |

### BuiltinResolver 테스트

| 테스트명 | 검증 내용 |
|---------|----------|
| `test_builtin_today` | prompt="today" → 오늘 날짜 |
| `test_builtin_yesterday` | prompt="yesterday" → 어제 날짜 |
| `test_builtin_this_week_monday` | prompt="this_week_monday" → 이번 주 월요일 |
| `test_builtin_this_week_sunday` | prompt="this_week_sunday" → 이번 주 일요일 |
| `test_builtin_last_week` | prompt="last_week_monday"/"last_week_sunday" |
| `test_builtin_this_month` | prompt="this_month_start"/"this_month_end" |
| `test_builtin_last_month` | prompt="last_month_start"/"last_month_end" |
| `test_builtin_unknown_prompt_fallback` | 알 수 없는 prompt → today 대체 + 경고 |
| `test_builtin_no_llm_called` | date 타입만 있으면 LLM 호출 0회 확인 |

### 통합 테스트

| 테스트명 | 검증 내용 |
|---------|----------|
| `test_weekly_aggregation_e2e` | 전체 시나리오: 동적 파라미터 해소 → FIND_DATE_RANGE → FIND_ANCHOR → AGGREGATE_RANGE → 결과 검증 |

---

## 6. 테스트용 엑셀 생성 fixture

FIND_ANCHOR, FIND_DATE_RANGE 테스트를 위한 엑셀 fixture 설계:

```python
# conftest.py에 추가

@pytest.fixture
def sample_weekly_excel(tmp_data_dir, xw_app):
    """주간 합산 시나리오용 엑셀.

    Sheet1 구조:
      Row 1: (빈칸)
      Row 2: (빈칸) | (빈칸)  | "일별" |  ...   | "주간" | "월간"
      Row 3: (빈칸) | "항목"  |  3/30  | 3/31 | 4/1 | 4/2 | 4/3 | "14주" | "4월"
      Row 4: (빈칸) | "불량률" |  1.2  | 1.5  | 0.8 | 1.1 | 0.9 | (빈칸) | (빈칸)
      Row 5: (빈칸) | "생산량" |  100  | 120  | 110 | 130 | 105 | (빈칸) | (빈칸)

    특징:
    - 헤더가 1행이 아닌 2~3행에 위치 (밀린 상황 시뮬레이션)
    - 날짜가 "M/D" 형식
    - "일별", "주간", "월간" 텍스트 앵커 존재
    - 주간/월간 칼럼은 비어있어 합산 결과를 쓸 수 있음
    """
```

---

## 7. 구현하지 않는 것

- **LLM_DECIDE 액션**: 현재 유스케이스에서 필요한 시나리오 없음. YAGNI 원칙.
- **퍼지 매칭**: FIND_ANCHOR의 `contains`로 대부분 커버됨.
  "자재코드"→"원자재코드" 같은 케이스도 `contains`로 해결.
  유사도 기반 매칭은 실제 필요 발생 시 추가.
- **LLMResolver 삭제**: 기존 `type=lookup`/`type=text` 사용하는 계획이 있을 수 있으므로 유지.
  다만 `type=date`의 주/월 계산이 Builtin으로 이동하여 LLM 호출 빈도는 감소.
