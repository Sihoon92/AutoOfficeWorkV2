---
name: execution-plan-generator
description: 반복 사무업무를 자동화하기 위한 execution_plan.json을 생성하는 스킬. 사용자가 빈 엑셀 양식과 작업 프로세스를 설명하면, 사내 실행 엔진(autooffice)이 순차 실행할 수 있는 표준화된 실행 계획서를 생성한다. 이 스킬은 사용자가 반복 사무업무 자동화를 요청하거나, 엑셀 기반 업무 프로세스를 설명하거나, execution plan 생성을 요청할 때 사용한다.
---

# Execution Plan Generator

## 핵심 원칙

**"Thinking은 Claude, Doing은 사내 엔진"**

이 스킬의 목적은 Claude의 사고 능력으로 사내 실행 엔진이 실수 없이 따라할 수 있는 **완벽한 지시서(execution_plan.json)**를 생성하는 것이다.

**보안 경계**:
- Claude에게 가는 것: 빈 양식 구조, 작업 프로세스 설명 (기밀 아님)
- 사내에 남는 것: 실제 데이터 값, 계산 결과, 개인정보 (기밀)

절대 사용자에게 실제 데이터를 요청하지 않는다. 구조와 프로세스만으로 계획을 생성한다.

---

## 입력 수집

사용자로부터 반드시 **2가지**를 확보한다:

### 1. 양식 파일 (빈 템플릿)
- 실제 데이터가 없는 빈 양식 엑셀 파일
- 첨부된 파일을 분석하여 시트 구조, 컬럼 배치, 수식 위치, 데이터 입력 영역을 파악한다
- 양식 파일이 없으면 양식의 구조를 텍스트로라도 설명해달라고 요청한다

### 2. 작업 프로세스 설명 (자연어)
- 어떤 데이터를 어디에 넣는지, 어떤 컬럼을 매핑하는지
- 최종 산출물이 무엇인지 (메신저 발송, PPT 생성, 이메일, 파일 저장 등)

두 가지 중 하나라도 빠지면 정중하게 요청한다. **"데이터 파일을 보여달라"고 하지 않는다.**

---

## 사고 프로세스 (5 Phase)

5단계를 순서대로 수행하며, 각 단계의 사고 과정을 thinking_log.txt에 기록한다.

### Phase 1: Intent Parsing (의도 분석)

사용자의 자연어 요청을 분해하여 하위 목표를 도출한다.

- 전체 작업을 독립적 하위 목표로 분해 (예: 데이터 추출 → 양식 삽입 → 결과 산출 → 전파)
- 각 하위 목표에 필요한 ACTION 능력 식별 (파일 I/O, 엑셀 조작, 메신저 API 등)
- **암묵적 요구사항 추론** — 이것이 핵심이다:
  - "붙여넣어" → 기존 데이터 CLEAR_RANGE 필요
  - "결과를 보내줘" → FORMAT_MESSAGE + SEND_MESSENGER + require_confirm 필요
  - "양식에 넣어" → 양식에 수식이 있으면 RECALCULATE 필요
  - "파일로 저장" → SAVE_FILE + save_as 경로 결정
  - "오늘 날짜 열에 추가" → `dynamic_params`로 날짜를 선언하고 `FIND_DATE_COLUMN` step으로 열 위치를 탐색. 또는 `EXTRACT_DATE`로 파일명에서 날짜를 추출하여 파생 정보와 함께 사용
  - "계산식 결과를 복사" → RECALCULATE 먼저, COPY_RANGE paste_type: "values"
  - "주간/월간 누계 작성" → AGGREGATE_RANGE method: "sum". `FIND_DATE_RANGE`로 주간/월간 범위를, `FIND_ANCHOR`로 집계 열을 탐색

### Phase 2: Context Acquisition (맥락 수집)

양식 파일을 분석하여 구조적 맥락을 수집한다.

- 시트 목록과 각 시트의 역할 파악
- 데이터 입력 영역 식별 (시작 셀, 컬럼 순서, 데이터 타입)
- 수식 영역 식별 (어떤 셀에 수식이 있고, 어떤 입력에 의존하는지)
- 결과 출력 영역 식별 (최종 산출물이 어디에 나타나는지)
- raw 데이터의 예상 스키마 추론 (사용자 설명 기반)

### Phase 3: Mapping Strategy (매핑 전략 수립)

raw 데이터 → 양식 간의 매핑 규칙을 확정한다.

- 컬럼 대 컬럼 매핑 테이블 작성
- 제외해야 할 컬럼 명시
- 데이터 타입 변환 필요 여부 (datetime → date, string → integer 등)
- 삽입 시작 위치, 삽입 방향 (행 방향/열 방향)
- 기존 데이터 처리 전략 (클리어 후 삽입 vs 추가 삽입)

### Phase 4: Execution Planning (실행 계획 수립)

실행 엔진이 한 step씩 따라할 수 있는 구체적 계획을 만든다.

- 각 step에 하나의 원자적 ACTION 배정
- 각 step에 expect 조건 (성공 기준) 명시
- 각 step에 on_fail 행동 (실패 시 대응) 명시
- step 간 데이터 흐름 연결 (store_as → $변수 참조, dict 결과는 $변수.필드 점 표기법)
- 외부 발송, 파일 덮어쓰기 등에 require_confirm 설정

**on_fail 전략 가이드라인**:
- 데이터 읽기/쓰기 등 핵심 step → `STOP`
- 수식 재계산 → `WARN_AND_CONTINUE`
- 경고성 검증 → `WARN_AND_CONTINUE`
- 로그 기록 → `SKIP`

### Phase 5: Validation Design (검증 설계)

실행 결과의 정합성을 검증할 로직을 설계한다.

- Sanity check 규칙 (값 범위, 합계 일치, 행 수 비교)
- 이상치 탐지 기준 (임계값 기반 알림)
- VALIDATE step을 plan에 포함시켜 엔진이 자동 검증하도록 한다

---

## 출력물 생성

3가지 파일을 생성한다.

### 1. execution_plan.json

사내 실행 엔진이 Pydantic으로 파싱하므로, **필드명이 정확히 일치해야 한다**.
Pydantic 모델 원본은 `references/pydantic_model.py` 참조.
ACTION 유형과 params 상세는 `references/action_reference.md` 참조.
완성된 예시는 `references/sample_plan.json` 참조.

#### 필수 스키마 (정확한 필드명)

**절대 필드명을 변경하지 않는다. 아래 필드명을 그대로 사용해야 한다.**

```json
{
  "task_id": "snake_case_식별자",
  "description": "작업 한 줄 설명",
  "created_by": "Claude",
  "created_at": "2026-03-26T09:00:00+09:00",
  "version": "1.0.0",
  "metadata": {
    "task_type": "작업유형",
    "template_hash": "양식해시",
    "reusable": true,
    "has_dynamic_params": false
  },
  "dynamic_params": {},
  "inputs": {
    "입력키": {
      "description": "설명",
      "expected_format": "xlsx",
      "expected_sheets": ["시트1"],
      "expected_columns": ["컬럼1", "컬럼2"]
    }
  },
  "steps": [
    {
      "step": 1,
      "action": "OPEN_FILE",
      "description": "step 설명",
      "params": { "file_path": "파일.xlsx", "alias": "별칭" },
      "expect": { "condition": "not_empty", "value": null },
      "on_fail": "STOP",
      "store_as": "변수명",
      "require_confirm": false
    }
  ],
  "final_output": {
    "type": "file",
    "description": "최종 산출물 설명"
  }
}
```

#### 필드 상세 (틀리기 쉬운 항목)

**최상위 필드** — 반드시 아래 이름 사용:

| 필드명 | 타입 | 필수 | 주의사항 |
|--------|------|------|----------|
| `task_id` | string | **필수** | ~~plan_id~~, ~~id~~ 아님. snake_case |
| `description` | string | **필수** | |
| `created_by` | string | 선택 | 값은 항상 `"Claude"` |
| `created_at` | string | **필수** | ISO 8601 형식 (예: `"2026-03-26T09:00:00+09:00"`) |
| `version` | string | 선택 | `"1.0.0"` |
| `metadata` | object | 선택 | |
| `inputs` | object | **필수** | dict[str, InputSpec] |
| `steps` | array | **필수** | 최소 1개 |
| `final_output` | object | **필수** | |

**metadata 필드** — 반드시 아래 이름 사용:

| 필드명 | 타입 | 주의사항 |
|--------|------|----------|
| `task_type` | string | ~~type~~, ~~category~~ 아님 |
| `template_hash` | string | |
| `reusable` | boolean | |

**inputs 값 (InputSpec)** — 반드시 아래 이름 사용:

| 필드명 | 타입 | 필수 | 주의사항 |
|--------|------|------|----------|
| `description` | string | **필수** | |
| `expected_format` | string | **필수** | `"xlsx"`, `"xls"`, `"csv"`, `"json"`, `"txt"` 중 하나 |
| `expected_sheets` | string[] | 선택 | |
| `expected_columns` | string[] | 선택 | |

**steps 배열 원소 (Step)** — 반드시 아래 이름 사용:

| 필드명 | 타입 | 필수 | 주의사항 |
|--------|------|------|----------|
| `step` | integer | **필수** | 1부터 시작, 연속. ~~step_number~~, ~~order~~ 아님 |
| `action` | string | **필수** | ActionType enum 값만 허용 (아래 목록) |
| `description` | string | **필수** | |
| `params` | object | **필수** | ACTION별 파라미터 |
| `expect` | object/null | 선택 | `{"condition": "...", "value": ...}` |
| `on_fail` | string | 선택 | `"STOP"`, `"SKIP"`, `"RETRY"`, `"WARN_AND_CONTINUE"` 중 하나. 기본 `"STOP"` |
| `store_as` | string/null | 선택 | 이후 step에서 `$변수명` 또는 `$변수명.필드명`으로 참조 |
| `require_confirm` | boolean | 선택 | 기본 false |

**final_output** — 반드시 아래 이름 사용:

| 필드명 | 타입 | 필수 | 주의사항 |
|--------|------|------|----------|
| `type` | string | **필수** | `"file"`, `"messenger"`, `"email"`, `"ppt"`, `"multiple"` 중 하나 |
| `description` | string | **필수** | |

#### action 허용 값 (ActionType enum)

다음 19개 값만 사용 가능하다. **정확히 대문자로 표기해야 한다**:

```
OPEN_FILE, READ_COLUMNS, READ_RANGE, WRITE_DATA, CLEAR_RANGE,
RECALCULATE, FIND_DATE_COLUMN, FIND_DATE_RANGE, FIND_ANCHOR,
COPY_RANGE, AGGREGATE_RANGE, EXTRACT_DATE,
SAVE_FILE, VALIDATE, FORMAT_MESSAGE, SEND_MESSENGER, SEND_EMAIL,
GENERATE_PPT, LOG
```

#### 절대 하지 말아야 할 것 (흔한 오류)

아래는 실제 파싱 실패를 일으킨 사례다. **절대 이렇게 생성하지 않는다.**

**오류 1: `inputs` 필드 누락**
```json
// ❌ 잘못됨 — inputs 필드가 아예 없음
{ "task_id": "...", "steps": [...], "final_output": {...} }

// ✅ 올바름 — inputs 필수
{ "task_id": "...", "inputs": { "raw_data": { "description": "...", "expected_format": "xlsx" } }, "steps": [...], "final_output": {...} }
```

**오류 2: step에 `description` 누락**
```json
// ❌ 잘못됨 — description 없음
{ "step": 1, "action": "OPEN_FILE", "params": {...}, "on_fail": "STOP" }

// ✅ 올바름 — description 필수
{ "step": 1, "action": "OPEN_FILE", "description": "raw 데이터 파일 열기", "params": {...}, "on_fail": "STOP" }
```

**오류 3: `expect`를 문자열로 작성**
```json
// ❌ 잘못됨 — 문자열 사용
"expect": "logged"
"expect": "file_opened AND file_extension IN [.xlsx]"

// ✅ 올바름 — 반드시 객체 {"condition": "...", "value": ...}
"expect": { "condition": "not_empty", "value": null }
"expect": { "condition": "row_count_gt", "value": 0 }
// 또는 expect가 불필요하면 필드 자체를 생략
```

**오류 4: `on_fail`에 임의 문자열 사용**
```json
// ❌ 잘못됨 — 엔진이 파싱할 수 없는 값
"on_fail": "CONTINUE"
"on_fail": "ABORT with error: 소스 파일을 열 수 없습니다"
"on_fail": "ABORT"
"on_fail": "ERROR"
"on_fail": "WARN"

// ✅ 올바름 — 4개 enum 값만 허용
"on_fail": "STOP"
"on_fail": "SKIP"
"on_fail": "RETRY"
"on_fail": "WARN_AND_CONTINUE"
```
에러 메시지, 설명 등의 자유 텍스트는 on_fail에 넣지 않는다. 에러 설명이 필요하면 step의 `description`에 적는다.

**오류 5: `task_id` 대신 다른 이름 사용**
```json
// ❌ 잘못됨
"plan_id": "my_plan"
"id": "my_plan"

// ✅ 올바름
"task_id": "my_plan"
```

**오류 6: FIND_DATE_COLUMN의 `date`에 `"today"` 사용**
```json
// ❌ 잘못됨 — 엔진이 "today" 문자열을 날짜로 인식하지 못함
"params": { "date": "today" }

// ✅ 올바름 — 반드시 실제 날짜를 ISO 형식으로 입력
"params": { "date": "2026-03-29" }
```
실행 시점의 오늘 날짜를 `YYYY-MM-DD` 형식으로 계산하여 넣는다. 양식의 날짜 표시 형식(예: `"3/29"`)과 무관하게, **date 파라미터는 항상 ISO 형식**이다. 엔진이 셀의 다양한 날짜 형식을 자동 감지하여 매칭한다.

#### 생성 전 자가 검증 체크리스트

execution_plan.json 출력 전 **반드시** 아래를 하나씩 확인한다. 하나라도 실패하면 수정 후 출력한다:

1. 최상위에 `task_id` (~~plan_id~~ 아님) 필드가 있는가?
2. 최상위에 `inputs` 필드가 있는가? (비어있어도 `{}` 형태로 존재해야 함)
3. `inputs`의 각 값에 `description`과 `expected_format`이 있는가?
4. `created_at`이 ISO 8601 datetime 문자열인가?
5. `steps`가 비어있지 않은 배열인가?
6. **모든** step에 `step`, `action`, `description`, `params` 4개 필수 필드가 있는가?
7. `step` 번호가 1부터 시작하여 빠짐없이 연속인가?
8. 모든 `action` 값이 19개 ActionType enum 중 하나인가?
9. 모든 `on_fail` 값이 **정확히** `"STOP"`, `"SKIP"`, `"RETRY"`, `"WARN_AND_CONTINUE"` 4개 중 하나인가? (자유 텍스트 절대 불가)
10. 모든 `expect` 값이 `{"condition": "...", "value": ...}` 객체인가? (문자열 절대 불가)
11. `store_as`로 저장한 변수가 `$변수명`으로 참조되기 전에 먼저 정의되는가?
12. `final_output`에 `type`과 `description`이 있는가?
13. JSON 전체가 유효한 문법인가? (trailing comma, 주석 없음)
14. FIND_DATE_COLUMN의 `date` 파라미터가 `"today"` 같은 문자열이 아닌, 실제 날짜 ISO 형식(`"YYYY-MM-DD"`)인가?
15. `dynamic_params`에 선언된 모든 key가 step params의 `{{dynamic:key}}`에서 실제로 참조되는가?
16. `{{dynamic:key}}`로 참조된 모든 key가 `dynamic_params`에 선언되어 있는가?
17. `dynamic_params`가 비어있지 않으면 `metadata.has_dynamic_params`가 `true`인가?
18. `dynamic_params`의 모든 항목에 `type`과 `prompt` 필드가 있는가? (`type`은 `"date"` 만 허용, `prompt`는 16개 함수 키워드 중 하나)
19. type=date의 `prompt`가 16개 허용 키워드 중 하나인가? (today, yesterday, this_week_monday, this_week_sunday, last_week_monday, last_week_sunday, this_month_start, this_month_end, last_month_start, last_month_end, today_yyyymmdd, today_yyyy_mm_dd, week_number, month_number, year, quarter)

### 2. thinking_log.txt

5 Phase 사고 과정을 사람이 읽을 수 있는 형태로 기록한다. 형식:

```
═══════════════════════════════════════════════════
[Phase N] 단계 제목
═══════════════════════════════════════════════════
- 판단 내용 (무엇을 왜 이렇게 결정했는지)
- 주의사항 / 발견된 이슈
```

thinking_log.txt에는 **실제 데이터 값을 포함하지 않는다**. 구조와 로직만 기록한다.

### 3. test_plan.py

execution_plan의 논리적 정합성을 실제 데이터 없이 검증하는 pytest 코드.

검증 항목:
- plan JSON이 스키마를 준수하는지
- step 번호가 1부터 연속인지
- store_as 변수가 참조 전에 정의되는지
- 존재하지 않는 ACTION이 사용되지 않았는지
- on_fail이 유효한 값인지

---

## 동적 파라미터 (Dynamic Parameters)

plan이 **반복 실행**될 때 실행 시점마다 바뀌는 날짜 값이 있다면,
`dynamic_params`에 선언하고 step params에서 `{{dynamic:key}}` 마커로 참조한다.
실행 엔진의 BuiltinResolver가 런타임에 로컬에서 해소한다 (LLM 불필요, API 비용 0).

### 선언 형식

```json
"dynamic_params": {
  "target_date": {
    "type": "date",
    "prompt": "today",
    "description": "실행 기준 오늘 날짜"
  },
  "week_start": {
    "type": "date",
    "prompt": "this_week_monday",
    "description": "이번 주 월요일"
  }
}
```

**필드 상세:**

| 필드 | 필수 | 설명 |
|------|------|------|
| `type` | O | `"date"` (유일한 유형) |
| `prompt` | **O** | BuiltinResolver 함수 키워드. 아래 16개 중 하나 |
| `format` | - | 기대 출력 형식 (`"YYYY-MM-DD"` 등) |
| `description` | - | 사람용 설명 |

**prompt 허용 키워드 (16개):**

| 키워드 | 반환 예시 (2026-04-02 기준) |
|--------|---------------------------|
| `today` | `"2026-04-02"` |
| `yesterday` | `"2026-04-01"` |
| `this_week_monday` | `"2026-03-30"` |
| `this_week_sunday` | `"2026-04-05"` |
| `last_week_monday` | `"2026-03-23"` |
| `last_week_sunday` | `"2026-03-29"` |
| `this_month_start` | `"2026-04-01"` |
| `this_month_end` | `"2026-04-30"` |
| `last_month_start` | `"2026-03-01"` |
| `last_month_end` | `"2026-03-31"` |
| `today_yyyymmdd` | `"20260402"` (파일명 생성용) |
| `today_yyyy_mm_dd` | `"2026_04_02"` |
| `week_number` | `"14"` |
| `month_number` | `"4"` |
| `year` | `"2026"` |
| `quarter` | `"2"` |

**오류 예시:**
```json
// ❌ 잘못됨 — prompt 누락 → Pydantic 파싱 에러
{"type": "date", "description": "오늘 날짜"}

// ❌ 잘못됨 — type: "lookup" 은 지원하지 않음 (LLM 미사용)
{"type": "lookup", "prompt": "Sheet1에서 오늘 날짜 열 계산..."}

// ✅ 올바름
{"type": "date", "prompt": "today", "description": "오늘 날짜"}
```

### 참조 형식

step params에서 `{{dynamic:key}}`로 참조:

```json
{
  "step": 3,
  "action": "FIND_DATE_COLUMN",
  "params": {
    "workbook": "tpl",
    "sheet": "Sheet1",
    "date": "{{dynamic:target_date}}"
  },
  "store_as": "day_col"
}
```

- 부분 문자열도 가능: `"report_{{dynamic:target_date}}.xlsx"`

### 위치 결정 전략: 내장 액션 사용

**날짜/위치 결정은 반드시 내장 액션 조합으로 해결한다. LLM에 위임하지 않는다.**

| 필요한 정보 | 해결 방법 |
|------------|----------|
| 오늘/어제/이번주 등 날짜 값 | `dynamic_params` `type: "date"` |
| 파일명에서 날짜 추출 + 파생 정보 | `EXTRACT_DATE` step |
| 특정 날짜의 컬럼 위치 | `FIND_DATE_COLUMN` step |
| 날짜 범위의 컬럼들 (주간/월간) | `FIND_DATE_RANGE` step |
| 텍스트 라벨의 셀 위치 ("주간", "4월" 등) | `FIND_ANCHOR` step |
| 여러 열의 행별 합계/평균 | `AGGREGATE_RANGE` step |

### 예시: 내장 액션 조합으로 일/주/월 컬럼 결정

```json
{
  "dynamic_params": {
    "target_date": {"type": "date", "prompt": "today"},
    "week_start": {"type": "date", "prompt": "this_week_monday"}
  },
  "steps": [
    {"step": 3, "action": "FIND_DATE_COLUMN",
     "description": "오늘 날짜의 일별 컬럼 찾기",
     "params": {"workbook": "tpl", "sheet": "Sheet1", "date": "{{dynamic:target_date}}"},
     "store_as": "day_col"},

    {"step": 4, "action": "FIND_DATE_RANGE",
     "description": "이번 주 월~오늘 컬럼 범위 찾기",
     "params": {"workbook": "tpl", "sheet": "Sheet1",
                "start_date": "{{dynamic:week_start}}", "end_date": "{{dynamic:target_date}}"},
     "store_as": "week_range"},

    {"step": 5, "action": "FIND_ANCHOR",
     "description": "주간 집계 컬럼 찾기",
     "params": {"workbook": "tpl", "sheet": "Sheet1",
                "search_value": "주간", "match_type": "contains"},
     "store_as": "week_agg"},

    {"step": 6, "action": "AGGREGATE_RANGE",
     "description": "주간 누계 계산",
     "params": {
       "workbook": "tpl", "sheet": "Sheet1",
       "source_columns_start": "$week_range.start_column",
       "source_columns_end": "$week_range.end_column",
       "source_start_row": 7, "source_end_row": 48,
       "target_column": "$week_agg.column",
       "method": "sum"
     }}
  ]
}
```

### EXTRACT_DATE를 활용한 파일명 기반 전체 흐름

raw data 파일명에 날짜가 포함된 경우 (`YYYYMMDD_xxx.xlsx`),
EXTRACT_DATE가 `week_monday`, `month`, `quarter` 등을 한번에 계산하므로 dynamic_params가 불필요할 수 있다:

```json
{
  "steps": [
    {"step": 1, "action": "OPEN_FILE",
     "params": {"input_key": "raw_data", "alias": "raw"}},

    {"step": 2, "action": "EXTRACT_DATE",
     "description": "파일명에서 날짜 + 파생 정보 추출",
     "params": {"source": "$input.raw_data", "pattern": "YYYYMMDD"},
     "store_as": "file_date"},

    {"step": 3, "action": "FIND_DATE_COLUMN",
     "params": {"workbook": "tpl", "sheet": "Sheet1", "date": "$file_date.date"},
     "store_as": "day_col"},

    {"step": 4, "action": "FIND_DATE_RANGE",
     "params": {"workbook": "tpl", "sheet": "Sheet1",
                "start_date": "$file_date.week_monday", "end_date": "$file_date.date"},
     "store_as": "week_range"},

    {"step": 5, "action": "FIND_ANCHOR",
     "params": {"workbook": "tpl", "sheet": "Sheet1",
                "search_value": "주간", "match_type": "contains"},
     "store_as": "week_agg"}
  ]
}
```

---

## 품질 기준

생성된 execution_plan이 반드시 충족해야 하는 기준:

- **원자성**: 각 step은 하나의 ACTION만 수행한다
- **방어성**: 모든 step에 on_fail이 있다. 핵심 step에는 expect도 있다
- **추적성**: step 간 데이터 흐름이 store_as → $참조로 명확히 연결된다
- **재현성**: 같은 구조의 데이터에 반복 실행하면 항상 같은 결과
- **안전성**: 외부 발송, 파일 덮어쓰기에 require_confirm이 설정된다

---

## 입력 파일 매핑

plan의 `inputs` 섹션에 선언된 입력 파일은 실행 시 `--input` 옵션으로 사용자가 지정한다.

### 실행 방법

```bash
autooffice run plan.json --data ./data/ \
    --input raw_data=20260402_production.xlsx \
    --input template=defect_template.xlsx
```

### OPEN_FILE에서의 참조

입력 파일을 여는 step에서는 `input_key`를 사용한다:

```json
{"step": 1, "action": "OPEN_FILE",
 "params": {"input_key": "raw_data", "alias": "raw"}}
```

`input_key`는 `inputs` 섹션의 키와 일치해야 한다. 파일명이 고정된 경우(양식 등)에도 `input_key`를 사용하는 것을 권장한다.

### 파일명에서 날짜 추출

raw data 파일명에 날짜가 포함된 경우 (`YYYYMMDD_xxx.xlsx`), EXTRACT_DATE로 날짜와 파생 정보를 추출한다:

```json
{"step": 3, "action": "EXTRACT_DATE",
 "params": {"source": "$input.raw_data", "pattern": "YYYYMMDD"},
 "store_as": "file_date"}
```

이후 `$file_date.date`, `$file_date.week_monday`, `$file_date.month` 등으로 일/주/월별 컬럼 위치를 결정한다.

---

## 재사용 안내

생성 완료 후 사용자에게 안내한다:

> 이 계획은 동일한 양식 + 동일한 프로세스에 대해 재사용 가능합니다.
> `autooffice cache list`로 캐시된 계획을 확인하고,
> `autooffice cache run <plan_id> --data ./data/ --input raw_data=파일명.xlsx --input template=양식명.xlsx`로 Claude 호출 없이 반복 실행할 수 있습니다.
