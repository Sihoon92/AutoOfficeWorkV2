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
- step 간 데이터 흐름 연결 (store_as → $변수 참조)
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
    "reusable": true
  },
  "inputs": {
    "입력키": {
      "description": "설명",
      "expected_format": "xlsx",
      "file_path": "기본_파일경로.xlsx",
      "expected_sheets": ["시트1"],
      "expected_columns": ["컬럼1", "컬럼2"]
    }
  },
  "steps": [
    {
      "step": 1,
      "action": "OPEN_FILE",
      "description": "step 설명",
      "params": { "file_path": "{{입력키}}", "alias": "별칭" },
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
| `file_path` | string | 선택 | 기본 파일 경로 (사용자가 `--file` 옵션으로 재지정 가능) |
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
| `store_as` | string/null | 선택 | 이후 step에서 `$변수명`으로 참조 |
| `require_confirm` | boolean | 선택 | 기본 false |

**final_output** — 반드시 아래 이름 사용:

| 필드명 | 타입 | 필수 | 주의사항 |
|--------|------|------|----------|
| `type` | string | **필수** | `"file"`, `"messenger"`, `"email"`, `"ppt"`, `"multiple"` 중 하나 |
| `description` | string | **필수** | |

#### action 허용 값 (ActionType enum)

다음 13개 값만 사용 가능하다. **정확히 대문자로 표기해야 한다**:

```
OPEN_FILE, READ_COLUMNS, READ_RANGE, WRITE_DATA, CLEAR_RANGE,
RECALCULATE, SAVE_FILE, VALIDATE, FORMAT_MESSAGE,
SEND_MESSENGER, SEND_EMAIL, GENERATE_PPT, LOG
```

#### 사용자 ��일 경로 플레��스홀더 (`{{input_key}}`)

실행 계획의 ��일 경로는 사용자 환경마다 다를 수 있다. **파일 경로를 하드코딩하지 말고**, `inputs` 키를 `{{key}}` 형태로 참조한다. 실행 시 사용자가 자신의 실제 경로를 지정할 수 있다.

**inputs에 기본 경로 설정:**
```json
"inputs": {
  "raw_data": {
    "description": "생산 데이터 원본",
    "expected_format": "xlsx",
    "file_path": "raw_data.xlsx"
  },
  "output": {
    "description": "결과 저장 파일",
    "expected_format": "xlsx",
    "file_path": "result.xlsx"
  }
}
```

**step params에서 플레이스홀더 사용:**
```json
{"step": 1, "action": "OPEN_FILE", "params": {"file_path": "{{raw_data}}", "alias": "raw"}}
{"step": 8, "action": "SAVE_FILE", "params": {"file": "template", "save_as": "{{output}}"}}
```

**사용자 실행 시 경로 지정:**
```bash
autooffice run plan.json --data ./data/ -f raw_data=./���_데이터.xlsx -f output=./결과.xlsx
```

- `-f` 옵션을 생략하면 대화형으로 경로를 물어본다
- `inputs`의 `file_path`에 기본값이 있으면 자동 사용된다
- **OPEN_FILE의 `file_path`, SAVE_FILE의 `save_as`에는 반드시 `{{input_key}}` 플레이스홀더를 사용한다**

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

**오류 6: OPEN_FILE/SAVE_FILE에 파일 경로 하드코딩**
```json
// ❌ 잘못됨 — 경로가 Claude 환경에 종속됨
"params": { "file_path": "/home/claude/data/raw_data.xlsx" }
"params": { "file_path": "C:\\Users\\claude\\raw_data.xlsx" }
"params": { "save_as": "/tmp/result.xlsx" }

// ✅ 올바름 — {{placeholder}}로 사용자가 경로를 지정할 수 있도록 함
"params": { "file_path": "{{raw_data}}" }
"params": { "save_as": "{{output}}" }
```

**오류 5: `task_id` 대신 다른 이름 사용**
```json
// ❌ 잘못됨
"plan_id": "my_plan"
"id": "my_plan"

// ✅ 올바름
"task_id": "my_plan"
```

#### 생성 전 자가 검증 체크리스트

execution_plan.json 출력 전 **반드시** 아래를 하나씩 확인한다. 하나라도 실패하면 수정 후 출력한다:

1. 최상위에 `task_id` (~~plan_id~~ 아님) 필드가 있는가?
2. 최상위에 `inputs` 필드가 있는가? (비어있어도 `{}` 형태로 존재해야 함)
3. `inputs`의 각 값에 `description`과 `expected_format`이 있는가?
4. `created_at`이 ISO 8601 datetime 문자열인가?
5. `steps`가 비어있지 않은 배열인가?
6. **모든** step에 `step`, `action`, `description`, `params` 4개 필수 필드가 있는가?
7. `step` 번호가 1부터 시작하여 빠짐없이 연속인가?
8. 모든 `action` 값이 13개 ActionType enum 중 하나인가?
9. 모든 `on_fail` 값이 **정확히** `"STOP"`, `"SKIP"`, `"RETRY"`, `"WARN_AND_CONTINUE"` 4개 중 하나인가? (자유 텍스트 절대 불가)
10. 모든 `expect` 값이 `{"condition": "...", "value": ...}` 객체인가? (문자열 절대 불가)
11. `store_as`로 저장한 변수가 `$변수명`으로 참조되기 전에 먼저 정의되는가?
12. `final_output`에 `type`과 `description`이 있는가?
13. JSON 전체가 유효한 문법인가? (trailing comma, 주석 없음)
14. OPEN_FILE의 `file_path`와 SAVE_FILE의 `save_as`에 `{{input_key}}` 플레이스홀더를 사용하고 있는가? (하드코딩된 절대/상대 경로 대신)
15. `{{placeholder}}`에 사용된 키가 `inputs`에 정의되어 있는가?

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

## 품질 기준

생성된 execution_plan이 반드시 충족해야 하는 기준:

- **원자성**: 각 step은 하나의 ACTION만 수행한다
- **방어성**: 모든 step에 on_fail이 있다. 핵심 step에는 expect도 있다
- **추적성**: step 간 데이터 흐름이 store_as → $참조로 명확히 연결된다
- **재현성**: 같은 구조의 데이터에 반복 실행하면 항상 같은 결과
- **안전성**: 외부 발송, 파일 덮어쓰기에 require_confirm이 설정된다

---

## 재사용 안내

생성 완료 후 사용자에게 안내한다:

> 이 계획은 동일한 양식 + 동일한 프로세스에 대해 재사용 가능합니다.
> `autooffice cache list`로 캐시된 계획을 확인하고,
> `autooffice cache run <plan_id> --data ./data/`로 Claude 호출 없이 반복 실행할 수 있습니다.
