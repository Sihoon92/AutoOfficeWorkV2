# 동적 파라미터 해소 프레임워크

- **커밋**: `f33036e`
- **날짜**: 2026-03-30

---

## 문제 정의

기존 시스템은 Claude가 plan을 생성하고, Python 엔진이 실행하는 2단계 구조이다.
매일 반복 실행되는 업무(예: "자재Loss 일보")에서 두 가지 값이 실행 시점에 결정되어야 한다:

- **날짜**: 매일 바뀌는 "오늘 날짜"를 plan에 하드코딩할 수 없음
- **위치**: 엑셀에서 오늘 날짜에 해당하는 열(column)이 매일 달라짐 (BT→BU→BV...)

`"today"` 매직 문자열은 이미 제거했으므로(002), 실행 시점에 동적으로 값을 해소하는 체계가 필요하다.

---

## 설계

### 핵심 아이디어: 마커 + 선언 + 해소

```
Plan 생성 시 (Claude)              실행 시 (Python Engine)
┌─────────────────────────┐       ┌──────────────────────────┐
│ dynamic_params:         │       │ BuiltinResolver (무료)    │
│   today_date:           │       │   type=date → 즉시 해소   │
│     type: date          │──────▶│                          │
│   today_target:         │       │ LLMResolver (API 1회)     │
│     type: lookup        │──────▶│   type=lookup → LLM 호출  │
│     prompt: "BT열부터…" │       │   날짜+위치 함께 반환      │
└─────────────────────────┘       └──────────────────────────┘
         │
         ▼
  steps[].params에서 참조:
    "date": "{{dynamic:today_date}}"
    "target_column": "{{dynamic:today_target.column}}"
```

3가지 구성요소:

| 구성요소 | 역할 |
|---------|------|
| `dynamic_params` (선언) | plan 최상위에 변수 이름, 타입, 프롬프트를 선언 |
| `{{dynamic:key}}` (마커) | step params 안에서 변수를 참조 |
| Resolver (해소기) | 실행 시점에 마커를 실제 값으로 치환 |

### 기존 시스템과의 분리

| 시스템 | 문법 | 시점 | 용도 |
|--------|------|------|------|
| 동적 파라미터 | `{{dynamic:key}}` | 실행 전 (plan 해소) | 매일 바뀌는 외부 값 |
| 변수 참조 | `$variable` | 실행 중 (step 간 전달) | step 간 데이터 전달 |

역할이 다르므로 문법도 분리했다. `dynamic_params`의 기본값이 `{}`이므로 기존 plan은 수정 없이 동작한다.

### FIND_DATE_COLUMN과의 공존

| 상황 | 사용할 것 |
|------|----------|
| 매일 반복, 열 패턴이 규칙적 (1일1열) | `dynamic_params` (type=lookup) |
| 1회성 실행, 또는 날짜 배치가 불규칙 | `FIND_DATE_COLUMN` step |

엑셀 시트마다 날짜 배열 규칙이 다르다. 규칙적인 경우 LLM이 계산할 수 있지만, 불규칙한 경우(빈 열, 병합 셀 등)는 실제 데이터를 읽어야 하므로 FIND_DATE_COLUMN이 필요하다.

---

## 구현 상세

### Phase A: 모델 확장

**파일**: `src/autooffice/models/execution_plan.py`

```python
class DynamicParamType(str, Enum):
    DATE = "date"      # 로컬에서 해소 (무료)
    LOOKUP = "lookup"  # LLM이 JSON으로 해소
    TEXT = "text"      # LLM이 텍스트로 해소

class DynamicParamSpec(BaseModel):
    type: DynamicParamType
    prompt: str          # resolver에게 전달할 지시문
    format: str          # 기대 출력 형식
    default: str | None  # 해소 실패 시 기본값
```

3가지 타입을 분리한 이유:
- `date`: 매일 바뀌지만 계산이 단순 → API 호출 불필요, 비용 절약
- `lookup`: 날짜+위치를 함께 계산해야 함 → LLM이 구조화된 JSON 반환
- `text`: 단순 텍스트 응답이 필요한 경우

### Phase B: 치환 엔진

**파일**: `src/autooffice/engine/resolvers/substitution.py`

```python
DYNAMIC_PATTERN = re.compile(r"\{\{dynamic:(\w+(?:\.\w+)?)\}\}")
```

`substitute_params(params, resolved)` 함수의 동작:

```python
# 입력
params = {"target_column": "{{dynamic:today_target.column}}"}
resolved = {"today_target": {"column": "CQ", "date_row": 3}}

# 출력
{"target_column": "CQ"}
```

타입 보존 규칙:
- 마커가 문자열 전체를 차지하면 원래 타입(int, dict 등)을 그대로 반환
- 부분 문자열이면 str로 변환하여 문자열 결합

Excel 행 번호 같은 값은 int로 유지해야 엔진이 올바르게 처리하기 때문이다.

`extract_dynamic_keys(params_list)` 함수: step params에서 참조된 동적 키를 추출하여, 실제 사용되는 키만 해소하도록 필터링한다.

### Phase C: Resolver 체인

**파일**: `src/autooffice/engine/resolvers/builtin_resolver.py`, `llm_resolver.py`, `chain.py`

```
ChainResolver
├── BuiltinDynamicResolver  (1순위: type=date 즉시 해소)
└── LLMDynamicResolver      (2순위: 나머지를 LLM API 1회로 해소)
```

체인 패턴을 사용한 이유:
- `type=date`는 `date.today().isoformat()`로 0원에 해소 가능
- LLM API는 비용이 발생하므로, 꼭 필요한 것만 보냄
- 매일 실행되는 자동화에서 불필요한 API 호출을 최소화

LLM Resolver 동작:
- 시스템 프롬프트에 오늘 날짜를 주입
- 모든 미해소 선언을 하나의 프롬프트로 묶어 1회 호출
- JSON 응답으로 날짜+열 위치를 함께 반환
- langchain_openai ChatOpenAI 사용 (사내 LLM API 연동)
- 해소 실패 시 `default` 값으로 폴백

### Phase D: Plan 해소 함수

**파일**: `src/autooffice/engine/resolvers/chain.py`

`resolve_plan_dynamic_params(plan, resolver)` 전체 흐름:

```
1. plan.dynamic_params 확인 → 없으면 원본 그대로 반환 (하위 호환)
2. step params에서 실제 참조된 키만 필터링 (불필요한 해소 방지)
3. Resolver 체인 실행
4. plan을 deep copy (원본 불변 보장)
5. 복사본의 모든 step params에서 마커를 치환
6. 해소된 plan 반환
```

deep copy를 사용한 이유: 캐시된 plan을 재사용할 때, 원본이 변경되면 다음 실행에 영향을 준다. 원본 불변성이 안전한 반복 실행의 핵심이다.

### Phase E: Runner 검증 + CLI 통합

**파일**: `src/autooffice/engine/runner.py`, `src/autooffice/cli.py`

Runner에 추가된 검증:
```python
_UNRESOLVED_DYNAMIC = re.compile(r"\{\{dynamic:\w+(?:\.\w+)?\}\}")
```
해소가 실패하거나 빠졌을 때, 실행 전에 오류를 잡아낸다.

CLI에 추가된 옵션:
- `--no-resolve`: 동적 해소 건너뛰기 (디버깅용)
- `--llm-model`: LLM 모델 지정

### Phase F: 스킬 문서 업데이트

**파일**: `skills/execution-plan-generator/SKILL.md`, `skills/execution-plan-generator/references/action_reference.md`

Claude가 plan을 생성할 때 동적 파라미터를 올바르게 사용하도록:
- 동적 파라미터 섹션, 타입별 사용 기준, 프롬프트 작성 가이드 추가
- "방법 A: 동적 파라미터 (권장)" vs "방법 B: FIND_DATE_COLUMN" 선택 기준 추가
- 체크리스트에 동적 파라미터 관련 검증 항목 3개 추가

---

## 파일 목록

### 신규 파일

| 파일 | 설명 |
|------|------|
| `src/autooffice/engine/resolvers/__init__.py` | 패키지 init, 주요 클래스 export |
| `src/autooffice/engine/resolvers/base.py` | `DynamicResolver` 추상 클래스 |
| `src/autooffice/engine/resolvers/builtin_resolver.py` | `BuiltinDynamicResolver` (type=date 해소) |
| `src/autooffice/engine/resolvers/llm_resolver.py` | `LLMDynamicResolver` (ChatOpenAI 연동) |
| `src/autooffice/engine/resolvers/chain.py` | `ChainResolver` + `resolve_plan_dynamic_params()` |
| `src/autooffice/engine/resolvers/substitution.py` | 마커 치환 유틸리티 |
| `tests/test_engine/test_resolvers/__init__.py` | 테스트 패키지 init |
| `tests/test_engine/test_resolvers/test_substitution.py` | 치환 유틸리티 테스트 |
| `tests/test_engine/test_resolvers/test_resolvers.py` | Resolver + plan 해소 테스트 |

### 수정 파일

| 파일 | 변경 내용 |
|------|----------|
| `src/autooffice/models/execution_plan.py` | DynamicParamType, DynamicParamSpec, dynamic_params 필드 추가 |
| `src/autooffice/models/__init__.py` | 새 모델 export 추가 |
| `src/autooffice/engine/runner.py` | 미해소 마커 검증 로직 추가 |
| `src/autooffice/cli.py` | --no-resolve, --llm-model 옵션, 해소 흐름 통합 |
| `pyproject.toml` | langchain-openai 선택적 의존성 추가 |
| `skills/execution-plan-generator/SKILL.md` | 동적 파라미터 가이드 전면 추가 |
| `skills/execution-plan-generator/references/action_reference.md` | 방법 A/B 선택 가이드 추가 |
| `skills/execution-plan-generator/references/pydantic_model.py` | DynamicParamSpec 모델 동기화 |

---

## 테스트

```bash
# 전체 테스트 실행
pytest tests/test_engine/test_resolvers/

# 개별 테스트
pytest tests/test_engine/test_resolvers/test_substitution.py -v
pytest tests/test_engine/test_resolvers/test_resolvers.py -v
```

테스트 커버리지:
- 치환: 단순 치환, 점 표기법, 부분 문자열, 중첩 dict/list, 미해소 마커 보존, 원본 불변
- Resolver: date 타입 해소, lookup 미해소 확인, 혼합 타입, plan 해소 전체 흐름, mock LLM
- Runner: 미해소 마커 검증, 해소된 plan 통과
