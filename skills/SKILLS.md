# Execution Plan Generator

## 핵심 원칙: "Thinking은 Claude, Doing은 사내 API"

이 스킬의 목적은 Claude의 에이전트적 사고 능력을 사내 환경에 이식하는 것이다. Claude는 **데이터를 보지 않는다**. 대신, 아무것도 모르는 실행 엔진이 실수 없이 따라할 수 있는 완벽한 지시서를 만든다.

보안 경계가 명확하다:

- Claude에게 가는 것: 빈 양식 구조, 작업 프로세스 설명 (기밀 아님)
- 사내에 남는 것: 실제 데이터 값, 계산 결과, 개인정보 (기밀)

---

## 입력 요구사항

사용자로부터 반드시 받아야 하는 2가지:

### 1. 양식 파일 (빈 템플릿)

- 실제 데이터가 없는 빈 양식 엑셀 파일
- Claude는 이 파일을 분석하여 시트 구조, 컬럼 배치, 수식 위치, 데이터 입력 영역을 파악한다
- 양식 파일이 첨부되지 않은 경우, 양식의 구조를 텍스트로라도 설명해달라고 요청한다

### 2. 작업 프로세스 설명 (자연어)

- "raw_data.xlsx의 Sheet1 데이터를 양식의 '데이터입력' 시트 B3부터 붙여넣어" 같은 구체적 지시
- 어떤 컬럼을 어디에 매핑하는지, 어떤 컬럼은 제외하는지
- 최종 산출물이 무엇인지 (메신저 발송, PPT 생성, 이메일 등)

사용자가 이 두 가지 중 하나라도 빠뜨리면, 정중하게 요청한다. 단, "데이터 파일을 보여달라"고 하지 않는다. 구조와 프로세스만 있으면 된다.

---

## 사고 프로세스 (5 Phase)

Claude는 아래 5단계 사고 과정을 거친다. 이 과정 자체가 thinking_log.txt로 기록된다.

### Phase 1: Intent Parsing (의도 분석)

사용자의 자연어 요청을 분해하여 하위 목표를 도출한다.

수행 내용:

- 전체 작업을 독립적 하위 목표로 분해 (예: 데이터 추출 → 양식 삽입 → 결과 산출 → 전파)
- 각 하위 목표에 필요한 능력 식별 (파일 I/O, 엑셀 조작, 메신저 API 등)
- 암묵적 요구사항 추론 (예: "붙여넣어" → 기존 데이터 클리어 필요)

### Phase 2: Context Acquisition (맥락 수집)

양식 파일을 분석하여 구조적 맥락을 수집한다.

수행 내용:

- 양식의 시트 목록, 각 시트의 역할 파악
- 데이터 입력 영역 식별 (시작 셀, 컬럼 순서, 데이터 타입)
- 수식 영역 식별 (어떤 셀에 수식이 있고, 어떤 입력에 의존하는지)
- 결과 출력 영역 식별 (최종 산출물이 어디에 나타나는지)
- raw 데이터의 예상 스키마 추론 (사용자 설명 기반)

### Phase 3: Mapping Strategy (매핑 전략 수립)

raw 데이터 → 양식 간의 매핑 규칙을 확정한다.

수행 내용:

- 컬럼 대 컬럼 매핑 테이블 작성
- 제외해야 할 컬럼 명시
- 데이터 타입 변환 필요 여부 (datetime → date, string → integer 등)
- 삽입 시작 위치, 삽입 방향 (행 방향/열 방향)
- 기존 데이터 처리 전략 (클리어 후 삽입 vs 추가 삽입)

### Phase 4: Execution Planning (실행 계획 수립)

바보 API가 한 줄씩 따라할 수 있는 구체적 step-by-step 계획을 만든다.

수행 내용:

- 각 step에 하나의 원자적 ACTION 배정
- 각 step에 expect 조건 (성공 기준) 명시
- 각 step에 on_fail 행동 (실패 시 대응) 명시
- step 간 데이터 흐름 연결 (store_as → 참조)
- 사용자 확인이 필요한 step 표시 (require_confirm)

### Phase 5: Validation Design (검증 설계)

실행 결과의 정합성을 검증할 로직을 설계한다.

수행 내용:

- Sanity check 규칙 (값 범위, 합계 일치, 행 수 비교)
- 이상치 탐지 기준 (임계값 기반 알림)
- 더미 데이터 기반 단위 테스트 작성

---

## 출력물 3가지

모든 출력물은 `/mnt/user-data/outputs/` 에 저장한다.

### 출력물 1: thinking_log.txt

Claude의 전체 사고 과정을 사람이 읽을 수 있는 형태로 기록한다.

형식:

```
═══════════════════════════════════════════════════
[STEP N] 단계 제목
═══════════════════════════════════════════════════
- 판단 내용 (무엇을 왜 이렇게 결정했는지)
- ACTION: 실행할 도구 호출 (있는 경우)
- 결과 확인 내용 (있는 경우)
- 주의사항 / 발견된 이슈
```

thinking_log.txt에는 실제 데이터 값이 포함되지 않는다. 구조와 로직만 기록한다.

### 출력물 2: execution_plan.json

사내 실행 엔진이 파싱하여 순차 실행할 수 있는 구조화된 계획서.

사용 가능한 ACTION 유형:

| ACTION | 설명 | 필수 params |
|--------|------|-------------|
| OPEN_FILE | 파일 열기 | file_path |
| READ_COLUMNS | 특정 컬럼 데이터 읽기 | file, sheet, columns, start_row |
| READ_RANGE | 셀 범위 읽기 | file, sheet, range |
| WRITE_DATA | 데이터 쓰기 | source, target_file, target_sheet, target_start, column_mapping |
| CLEAR_RANGE | 셀 범위 비우기 | file, sheet, range |
| RECALCULATE | 수식 재계산 트리거 | file |
| SAVE_FILE | 파일 저장 | file, save_as (optional) |
| VALIDATE | 데이터 검증 | check, source, logic |
| FORMAT_MESSAGE | 텍스트 메시지 구성 | template, data_source |
| SEND_MESSENGER | 메신저 발송 | to, message, attachment (optional) |
| SEND_EMAIL | 이메일 발송 | to, subject, body, attachment (optional) |
| GENERATE_PPT | PPT 생성 | data_source, template (optional), output_path |
| LOG | 로그 기록 | message, level (info/warn/error) |

### 출력물 3: test_plan.py

실행 계획의 논리적 정합성을 실제 데이터 없이 검증하는 Python 테스트 코드.

---

## 품질 기준

좋은 execution_plan의 기준:

- **원자성**: 각 step은 하나의 ACTION만 수행한다.
- **방어성**: 모든 step에 expect와 on_fail이 있다.
- **추적성**: step 간 데이터 흐름이 store_as → 참조로 명확히 연결된다.
- **재현성**: 같은 구조의 데이터에 대해 반복 실행하면 항상 같은 결과가 나온다.
- **안전성**: require_confirm이 적절히 사용된다 (외부 발송, 파일 덮어쓰기 등).

---

## 참고: 반복 작업 시 Plan 재사용

같은 양식 + 같은 프로세스의 반복 작업이라면, 한 번 생성된 execution_plan.json을 캐싱하여 Claude 호출 없이 재사용할 수 있다. 실행 엔진의 `autooffice cache` 명령으로 관리한다.
