# dynamic_params 작성 가이드

## 핵심 규칙

`dynamic_params`의 `type`은 **`"date"` 만 허용**된다.
`"lookup"`, `"text"` 등 다른 타입은 Pydantic ValidationError로 plan 실행이 즉시 실패한다.

---

## 잘못된 패턴과 올바른 대체

### ❌ 패턴 1: 구조화된 JSON을 하나의 파라미터로 요청

```json
"week_info": {
  "type": "text",
  "prompt": "오늘이 몇째 주인지, 주간 시작일과 종료일을 JSON으로 반환",
  "format": "json"
}
```

**문제**: `type: "text"` 는 엔진에 존재하지 않는다.

### ✅ 대체: 필요한 날짜를 개별 파라미터로 쪼갠다

```json
"week_monday": {
  "type": "date", "prompt": "this_week_monday",
  "description": "이번 주 월요일"
},
"week_sunday": {
  "type": "date", "prompt": "this_week_sunday",
  "description": "이번 주 일요일"
},
"week_num": {
  "type": "date", "prompt": "week_number",
  "description": "주차 번호"
}
```

---

### ❌ 패턴 2: LLM에게 위치 계산을 요청

```json
"today_layout": {
  "type": "lookup",
  "prompt": "Sheet1 6행 헤더에서 오늘 날짜 열과 주간/월간 합계 열을 JSON으로 반환",
  "format": "json"
}
```

**문제**: `type: "lookup"` 은 엔진에 존재하지 않는다.

### ✅ 대체: 날짜는 dynamic_params, 위치 탐색은 builtin action

```json
"dynamic_params": {
  "execution_date": {
    "type": "date", "prompt": "today",
    "description": "실행일 날짜"
  }
},
"steps": [
  {
    "step": 3,
    "action": "FIND_DATE_COLUMN",
    "description": "오늘 날짜 열 탐색",
    "params": {
      "workbook": "report",
      "sheet": "Sheet1",
      "scan_range": "BT6:ZZ6",
      "date": "{{dynamic:execution_date}}"
    },
    "store_as": "today_col"
  }
]
```

---

### ❌ 패턴 3: prompt 누락 또는 자유 텍스트

```json
"execution_date": {
  "type": "date",
  "format": "YYYY-MM-DD",
  "description": "실행일 날짜 (자동 해소)"
}
```

**문제**: `prompt` 필드 누락 → `Field required` 에러.

```json
"execution_date": {
  "type": "date",
  "prompt": "오늘 날짜를 YYYY-MM-DD로 반환해줘"
}
```

**문제**: `prompt`는 자유 텍스트가 아니라 16개 키워드 중 하나여야 한다.

### ✅ 올바른 형태

```json
"execution_date": {
  "type": "date",
  "prompt": "today",
  "description": "실행일 날짜"
}
```

---

## 허용되는 prompt 키워드 (16개)

| 키워드 | 반환 예시 (2026-04-02 기준) | 용도 |
|--------|---------------------------|------|
| `today` | `2026-04-02` | 실행일 날짜 |
| `yesterday` | `2026-04-01` | 전일 날짜 |
| `this_week_monday` | `2026-03-30` | 주간 시작일 |
| `this_week_sunday` | `2026-04-05` | 주간 종료일 |
| `last_week_monday` | `2026-03-23` | 지난주 시작일 |
| `last_week_sunday` | `2026-03-29` | 지난주 종료일 |
| `this_month_start` | `2026-04-01` | 월간 시작일 |
| `this_month_end` | `2026-04-30` | 월간 종료일 |
| `last_month_start` | `2026-03-01` | 전월 시작일 |
| `last_month_end` | `2026-03-31` | 전월 종료일 |
| `today_yyyymmdd` | `20260402` | 파일명 생성용 |
| `today_yyyy_mm_dd` | `2026_04_02` | 파일명 생성용 |
| `week_number` | `14` | 주차 번호 |
| `month_number` | `4` | 월 번호 |
| `year` | `2026` | 연도 |
| `quarter` | `2` | 분기 |

---

## 변환 요약표

| 필요한 정보 | ❌ 잘못된 접근 | ✅ 올바른 접근 |
|------------|--------------|--------------|
| 오늘 날짜 | `type: "text"`, prompt에 자유 텍스트 | `type: "date"`, `prompt: "today"` |
| 주간 시작~종료 | `type: "text"`, JSON으로 통째 요청 | `this_week_monday` + `this_week_sunday` 개별 선언 |
| 월간 시작~종료 | `type: "lookup"`, JSON으로 통째 요청 | `this_month_start` + `this_month_end` 개별 선언 |
| 주차/월 번호 | `type: "text"`, 주차 정보 요청 | `week_number`, `month_number` prompt 사용 |
| 날짜 열 위치 | `type: "lookup"`, 열 계산 요청 | `FIND_DATE_COLUMN` action |
| 날짜 범위 | `type: "lookup"`, 범위 계산 요청 | `FIND_DATE_RANGE` action |
| 특정 값 위치 | `type: "text"`, 위치 계산 요청 | `FIND_ANCHOR` action |
| 구조화된 JSON | `format: "json"` 으로 통째 요청 | builtin action 여러 개 + `store_as` + `$변수.필드` |

---

## 완성 예시: 주간/월간 보고서 자동화

```json
{
  "dynamic_params": {
    "execution_date": {
      "type": "date", "prompt": "today",
      "description": "실행일 날짜"
    },
    "week_monday": {
      "type": "date", "prompt": "this_week_monday",
      "description": "이번 주 월요일"
    },
    "week_sunday": {
      "type": "date", "prompt": "this_week_sunday",
      "description": "이번 주 일요일"
    },
    "month_start": {
      "type": "date", "prompt": "this_month_start",
      "description": "이번 달 시작일"
    },
    "month_end": {
      "type": "date", "prompt": "this_month_end",
      "description": "이번 달 종료일"
    },
    "week_num": {
      "type": "date", "prompt": "week_number",
      "description": "주차 번호"
    },
    "month_num": {
      "type": "date", "prompt": "month_number",
      "description": "월 번호"
    }
  },
  "steps": [
    {
      "step": 1,
      "action": "OPEN_FILE",
      "description": "보고서 파일 열기",
      "params": {
        "workbook": "report",
        "file_path": "월간보고서.xlsx"
      }
    },
    {
      "step": 2,
      "action": "FIND_DATE_COLUMN",
      "description": "오늘 날짜 열 탐색",
      "params": {
        "workbook": "report",
        "sheet": "Sheet1",
        "scan_range": "BT6:ZZ6",
        "date": "{{dynamic:execution_date}}"
      },
      "store_as": "today_col"
    },
    {
      "step": 3,
      "action": "FIND_DATE_RANGE",
      "description": "이번 주 날짜 범위 탐색",
      "params": {
        "workbook": "report",
        "sheet": "Sheet1",
        "scan_range": "BT6:ZZ6",
        "start_date": "{{dynamic:week_monday}}",
        "end_date": "{{dynamic:execution_date}}"
      },
      "store_as": "weekly_range"
    },
    {
      "step": 4,
      "action": "FIND_ANCHOR",
      "description": "주간합계 열 탐색",
      "params": {
        "workbook": "report",
        "sheet": "Sheet1",
        "scan_range": "BT6:ZZ6",
        "anchor_value": "W",
        "search_after": "$today_col.column"
      },
      "store_as": "weekly_col"
    },
    {
      "step": 5,
      "action": "AGGREGATE_RANGE",
      "description": "이번 주 일별 데이터를 주간 누계",
      "params": {
        "workbook": "report",
        "sheet": "Sheet1",
        "source_columns_start": "$weekly_range.start_column",
        "source_columns_end": "$weekly_range.end_column",
        "source_start_row": 7,
        "source_end_row": 48,
        "target_column": "$weekly_col.column",
        "target_start_row": 7,
        "method": "sum"
      }
    }
  ]
}
```

**패턴 요약**:
- `dynamic_params` → 날짜 값만 (type: "date", prompt: 키워드)
- `FIND_DATE_COLUMN` → 날짜로 열 위치 탐색
- `FIND_DATE_RANGE` → 날짜 범위로 열 범위 탐색
- `FIND_ANCHOR` → 특정 값(W, M 등)으로 열 위치 탐색
- `AGGREGATE_RANGE` → 탐색된 범위를 집계
- 각 action의 결과는 `store_as` → `$변수.필드`로 다음 step에 전달
