# 날짜 형식 자동 감지 개선 + row_count_gt expect 수정

- **커밋**: `3e3ceb9`
- **날짜**: 2026-03-30

---

## 1. FIND_DATE_COLUMN 날짜 형식 자동 감지

### 문제

FIND_DATE_COLUMN이 엑셀 셀에서 날짜를 읽을 때, `YYYY-MM-DD` 형식만 인식했다.
실제 엑셀 파일에는 `3-30`, `3.30`, `3월30일`, Excel 시리얼 번호(46110) 등 다양한 형식이 존재하여 날짜 매칭에 실패했다.

### 해결

`_try_parse_date()` 함수를 개선하여 6가지 날짜 형식을 자동 감지하도록 확장했다.

| 형식 | 예시 | 설명 |
|------|------|------|
| `M-D` | `3-30` | 하이픈 구분 (월-일) |
| `M.D` | `3.30` | 점 구분 |
| `YYYY-MM-DD` | `2026-03-30` | ISO 표준 |
| `YYYY/MM/DD` | `2026/03/30` | 슬래시 구분 |
| `M월D일` | `3월30일` | 한국어 형식 |
| Excel 시리얼 | `46110` | float/int (epoch: 1899-12-30) |

### 수정 파일

- `src/autooffice/engine/actions/excel_actions.py` — `_try_parse_date()` 함수
- `tests/test_engine/test_actions/test_excel_actions.py` — `TestTryParseDate`, `TestFindDateColumnFormats`

---

## 2. row_count_gt expect 조건 dict 결과 처리

### 문제

COPY_RANGE 액션이 dict 형태의 결과(`{"rows_copied": 42, ...}`)를 반환할 때,
`row_count_gt` expect 조건이 이를 처리하지 못하고 경고를 발생시켰다.

### 해결

`_check_expect()` 함수에서 `row_count_gt` 조건 평가 시, 결과가 dict인 경우
`rows_copied` 또는 `rows_written` 키의 값을 사용하도록 확장했다.

```python
# 기존: int만 처리
if isinstance(data, int) and data > threshold: ...

# 수정: dict 결과도 처리
if isinstance(data, dict):
    count = data.get("rows_copied") or data.get("rows_written")
```

### 수정 파일

- `src/autooffice/engine/runner.py` — `_check_expect()` 함수
- `tests/test_engine/test_actions/test_excel_actions.py` — `TestCheckExpectRowCountGt`
