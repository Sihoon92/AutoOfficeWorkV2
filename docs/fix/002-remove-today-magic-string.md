# "today" 매직 문자열 제거, ISO 날짜 필수화

- **커밋**: `e2a9489`
- **날짜**: 2026-03-30

---

## 문제

FIND_DATE_COLUMN의 `date` 파라미터에 `"today"`라는 텍스트를 전달하는 방식이 사용되고 있었다.
Python 엔진은 `"today"`라는 문자열을 날짜로 인식할 수 없으므로, 이 방식은 근본적으로 동작하지 않는다.

Claude(execution-plan-generator)가 plan을 생성할 때 `"today"`를 넣는 것 자체가 잘못된 설계였다.

## 해결

### 엔진 코드 수정

- `FindDateColumnHandler`에서 `"today"` 문자열 처리 로직을 완전히 제거
- `date` 파라미터를 필수값으로 변경
- ISO 형식(`YYYY-MM-DD`) 검증 로직 추가, 형식이 맞지 않으면 명확한 에러 메시지 반환

```python
# 제거된 코드
if date_str == "today":
    target_date = date.today()

# 추가된 검증
try:
    target_date = date.fromisoformat(date_str)
except ValueError:
    return ActionResult(success=False, message="date 파라미터는 ISO 형식(YYYY-MM-DD)이어야 합니다")
```

### 스킬 문서 수정

- `SKILL.md`에 에러 케이스 6번 추가: `"today"` 사용 금지 명시
- `action_reference.md`에서 `date` 파라미터 설명을 ISO 형식 필수로 변경
- 테스트 코드의 모든 `"date": "today"`를 `date.today().isoformat()`으로 교체

### 수정 파일

- `src/autooffice/engine/actions/excel_actions.py` — `FindDateColumnHandler`
- `skills/execution-plan-generator/SKILL.md` — 에러 케이스 추가
- `skills/execution-plan-generator/references/action_reference.md` — date 파라미터 설명
- `tests/test_engine/test_actions/test_excel_actions.py` — 테스트 수정

## 후속 조치

이 변경으로 인해, 매일 바뀌는 날짜를 plan에 어떻게 넣을 것인가라는 근본적인 문제가 남았다.
이는 다음 커밋(003)의 동적 파라미터 프레임워크로 해결한다.
