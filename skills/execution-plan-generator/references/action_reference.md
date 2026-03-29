# ACTION Reference

사내 실행 엔진이 지원하는 ACTION 유형과 params 상세.

## 파일 I/O

### OPEN_FILE
파일을 열어 엔진 컨텍스트에 등록한다.

```json
{
  "action": "OPEN_FILE",
  "params": {
    "file_path": "raw_data.xlsx",
    "alias": "raw",
    "data_only": true
  }
}
```

| param | 필수 | 설명 |
|-------|------|------|
| file_path | O | 파일 경로 (상대경로면 --data 기준) |
| alias | - | 이후 step에서 참조할 이름 (기본: 파일명) |
| data_only | - | true면 수식 대신 계산된 값으로 읽기 (기본 false) |

### SAVE_FILE
워크북을 파일로 저장한다.

```json
{
  "action": "SAVE_FILE",
  "params": {
    "file": "template",
    "save_as": "result.xlsx"
  }
}
```

| param | 필수 | 설명 |
|-------|------|------|
| file | O | 컨텍스트에 등록된 워크북 alias |
| save_as | - | 다른 이름으로 저장 (미지정 시 원본 경로) |

---

## 엑셀 읽기

### READ_COLUMNS
특정 시트에서 지정 컬럼의 데이터를 읽는다.

```json
{
  "action": "READ_COLUMNS",
  "params": {
    "file": "raw",
    "sheet": "Sheet1",
    "columns": ["날짜", "라인명", "생산수량"],
    "start_row": 2,
    "end_row": null
  },
  "store_as": "raw_data"
}
```

| param | 필수 | 설명 |
|-------|------|------|
| file | O | 워크북 alias |
| sheet | O | 시트명 |
| columns | O | 읽을 컬럼 목록 (헤더명 또는 컬럼 문자 "B") |
| start_row | O | 시작 행 (보통 2, 헤더 다음) |
| end_row | - | 종료 행 (미지정 시 데이터 끝까지) |

**반환값**: `[{"컬럼1": 값, "컬럼2": 값}, ...]` 형태의 dict 리스트

### READ_RANGE
셀 범위의 데이터를 읽는다.

```json
{
  "action": "READ_RANGE",
  "params": {
    "file": "result",
    "sheet": "일별현황",
    "range": "A2:G10"
  },
  "store_as": "daily_summary"
}
```

| param | 필수 | 설명 |
|-------|------|------|
| file | O | 워크북 alias |
| sheet | O | 시트명 |
| range | O | 셀 범위 (예: "B3:F100") |

**반환값**: 2D 리스트 `[[값, 값, ...], ...]`

---

## 엑셀 쓰기

### WRITE_DATA
데이터를 양식의 지정 위치에 쓴다.

```json
{
  "action": "WRITE_DATA",
  "params": {
    "source": "$raw_data",
    "target_file": "template",
    "target_sheet": "데이터입력",
    "target_start": "B3",
    "column_mapping": {
      "날짜": "B",
      "라인명": "C",
      "생산수량": "D"
    }
  }
}
```

| param | 필수 | 설명 |
|-------|------|------|
| source | O | 데이터 ($변수 참조 또는 직접 리스트) |
| target_file | O | 대상 워크북 alias |
| target_sheet | O | 대상 시트명 |
| target_start | O | 시작 셀 (예: "B3") |
| column_mapping | - | 소스 컬럼 → 대상 열 매핑 (예: {"이름": "B"}) |

column_mapping 미지정 시 소스 dict의 값 순서대로 target_start부터 기록한다.

### CLEAR_RANGE
지정 범위의 셀 값을 비운다. 데이터 삽입 전 기존 데이터 제거용.

```json
{
  "action": "CLEAR_RANGE",
  "params": {
    "file": "template",
    "sheet": "데이터입력",
    "range": "B3:F10000"
  }
}
```

| param | 필수 | 설명 |
|-------|------|------|
| file | O | 워크북 alias |
| sheet | O | 시트명 |
| range | O | 클리어할 셀 범위 |

### RECALCULATE
워크북의 수식 재계산을 트리거한다. xlwings를 통해 Excel 엔진이 실제 재계산 수행.

```json
{
  "action": "RECALCULATE",
  "params": {
    "file": "template"
  }
}
```

| param | 필수 | 설명 |
|-------|------|------|
| file | O | 워크북 alias |

### FIND_DATE_COLUMN
시트에서 날짜 패턴을 탐색하여 대상 열과 날짜 행을 결정한다. 날짜가 가장 많은 행을 자동으로 date header row로 인식한다.

```json
{
  "action": "FIND_DATE_COLUMN",
  "params": {
    "workbook": "material",
    "sheet": "Sheet1",
    "scan_range": "A1:ZZ10",
    "date": "2026-03-29"
  },
  "store_as": "date_info"
}
```

| param | 필수 | 설명 |
|-------|------|------|
| workbook | O | 워크북 alias |
| sheet | O | 시트명 |
| scan_range | - | 날짜를 탐색할 셀 범위 (기본: "A1:ZZ10") |
| date | **O** | 대상 날짜. 반드시 ISO 형식 `"YYYY-MM-DD"` (예: `"2026-03-29"`). ~~`"today"` 사용 불가~~ |

**반환값** (store_as로 저장): `{"column": "CP", "date_row": 6}`
- `column`: 대상 날짜의 열 문자
- `date_row`: 날짜가 발견된 행 번호

**날짜 인식 형식** (인접 셀의 형식을 자동 감지):
- Excel datetime/date 객체
- Excel 시리얼 날짜 (숫자)
- 문자열: `"M/D"`, `"M-D"`, `"M.D"` (연도 없는 형식, 올해 기준)
- 문자열: `"YYYY-MM-DD"`, `"YYYY/MM/DD"`, `"YYYY.MM.DD"` (연도 포함)
- 한국어: `"M월D일"`, `"M월 D일"`

**동작 방식**:
1. scan_range 내 날짜 셀 스캔
2. 날짜가 가장 많은 행 = date header row
3. 대상 날짜와 정확히 일치하는 열이 있으면 → 해당 열 반환
4. 없으면 → 마지막 날짜 다음 열을 패턴 기반 추론 (열/일 간격 비율 계산)

**이후 step에서 참조**: `$date_info.column` → `"CP"`, `$date_info.date_row` → `6` (점 표기법)

### COPY_RANGE
같은 워크북 내에서 소스 시트의 범위를 타겟 시트의 지정 열/행에 복사한다.

```json
{
  "action": "COPY_RANGE",
  "params": {
    "workbook": "material",
    "source_sheet": "자재 계산식",
    "source_range": "D7:D48",
    "target_sheet": "Sheet1",
    "target_column": "$date_info.column",
    "target_start_row": "$date_info.date_row",
    "row_offset": 1,
    "paste_type": "values"
  }
}
```

| param | 필수 | 설명 |
|-------|------|------|
| workbook | O | 워크북 alias |
| source_sheet | O | 소스 시트명 |
| source_range | O | 소스 범위 (예: "D7:D48") |
| target_sheet | O | 타겟 시트명 |
| target_column | O | 타겟 열 문자 (예: "CP", 변수 참조 가능) |
| target_start_row | O | 타겟 기준 행 (변수 참조 가능) |
| row_offset | - | 기준 행에서의 오프셋 (기본: 0). 실제 시작행 = target_start_row + row_offset |
| paste_type | - | `"values"` (기본), `"formulas"`, `"all"` |

**paste_type 설명**:
- `values`: 수식의 계산 결과(값)만 복사. 소스에 수식이 있어도 타겟에는 값만 들어감
- `formulas`: 수식 문자열 자체를 복사
- `all`: values와 동일 (향후 서식 복사 지원 예정)

**반환값**: `{"rows_copied": 42}`

---

## 날짜 열 탐색 + 시트 간 복사 패턴

FIND_DATE_COLUMN과 COPY_RANGE를 조합하면, 날짜 기반으로 열이 확장되는 엑셀에 자동으로 데이터를 추가할 수 있다.

```json
{
  "step": 5,
  "action": "FIND_DATE_COLUMN",
  "description": "Sheet1에서 오늘 날짜 열 탐색",
  "params": {
    "workbook": "mat",
    "sheet": "Sheet1",
    "scan_range": "A1:ZZ10",
    "date": "2026-03-29"
  },
  "store_as": "date_info"
},
{
  "step": 6,
  "action": "COPY_RANGE",
  "description": "자재 계산식 결과를 오늘 날짜 열에 값으로 붙여넣기",
  "params": {
    "workbook": "mat",
    "source_sheet": "자재 계산식",
    "source_range": "D7:D48",
    "target_sheet": "Sheet1",
    "target_column": "$date_info.column",
    "target_start_row": "$date_info.date_row",
    "row_offset": 1,
    "paste_type": "values"
  }
}
```

**핵심 포인트**:
- `$date_info.column`과 `$date_info.date_row`는 **점 표기법**으로 dict 필드를 참조한다
- `row_offset: 1`은 날짜 헤더 행 바로 아래부터 데이터를 쓰겠다는 의미
- `paste_type: "values"`는 소스의 수식이 아닌 계산된 값만 복사

---

## 검증

### VALIDATE
데이터에 대한 sanity check를 수행한다. 5가지 체크 유형 지원.

**check: row_count** — 행 수 범위 검증
```json
{
  "action": "VALIDATE",
  "params": {
    "check": "row_count",
    "source": "$raw_data",
    "logic": { "min": 1, "max": 10000 }
  }
}
```

**check: value_range** — 컬럼 값 범위 검증
```json
{
  "action": "VALIDATE",
  "params": {
    "check": "value_range",
    "source": "$daily_summary",
    "logic": { "column": "불량률", "min": 0, "max": 100 }
  }
}
```

**check: not_empty** — 비어있지 않은지 검증
```json
{
  "action": "VALIDATE",
  "params": {
    "check": "not_empty",
    "source": "$raw_data",
    "logic": { "columns": ["날짜", "라인명"] }
  }
}
```

**check: sum_equals** — 합계 일치 검증
```json
{
  "action": "VALIDATE",
  "params": {
    "check": "sum_equals",
    "source": "$data",
    "logic": { "column": "수량", "expected": 1000 }
  }
}
```

**check: column_exists** — 필수 컬럼 존재 검증
```json
{
  "action": "VALIDATE",
  "params": {
    "check": "column_exists",
    "source": "$raw_data",
    "logic": { "columns": ["날짜", "라인명", "생산수량"] }
  }
}
```

---

## 메시지 / 전송

### FORMAT_MESSAGE
템플릿 기반으로 텍스트 메시지를 구성한다.

```json
{
  "action": "FORMAT_MESSAGE",
  "params": {
    "template": "[품질현황 보고]\n결과 파일이 첨부되었습니다.",
    "data_source": "$daily_summary"
  },
  "store_as": "messenger_message"
}
```

### SEND_MESSENGER
메신저 채널/사용자에게 메시지를 발송한다. **require_confirm: true 권장**.

```json
{
  "action": "SEND_MESSENGER",
  "params": {
    "to": "#품질관리팀",
    "message": "$messenger_message",
    "attachment": "result.xlsx"
  },
  "require_confirm": true
}
```

### SEND_EMAIL
이메일을 발송한다. **require_confirm: true 권장**.

```json
{
  "action": "SEND_EMAIL",
  "params": {
    "to": "team@company.com",
    "subject": "일별 보고서",
    "body": "$email_body",
    "attachment": "result.xlsx"
  },
  "require_confirm": true
}
```

### GENERATE_PPT
데이터 기반으로 PowerPoint를 생성한다.

```json
{
  "action": "GENERATE_PPT",
  "params": {
    "data_source": "$summary",
    "template": "report_template.pptx",
    "output_path": "monthly_report.pptx"
  }
}
```

---

## 로그

### LOG
실행 로그를 기록한다.

```json
{
  "action": "LOG",
  "params": {
    "message": "작업 완료",
    "level": "info"
  }
}
```

| level | 용도 |
|-------|------|
| info | 정상 진행 기록 |
| warn | 경고 (계속 진행) |
| error | 오류 기록 |

---

## Step 공통 필드

모든 step에서 사용 가능한 필드:

| 필드 | 필수 | 설명 |
|------|------|------|
| step | O | 1부터 시작하는 연속 번호 |
| action | O | ACTION 유형 |
| description | O | 이 step이 하는 일 설명 |
| params | O | ACTION별 파라미터 |
| expect | - | 성공 기준 `{"condition": "...", "value": ...}` |
| on_fail | - | 실패 시 행동: STOP, SKIP, RETRY, WARN_AND_CONTINUE (기본 STOP) |
| store_as | - | 결과를 변수로 저장 (이후 step에서 `$변수명` 또는 `$변수명.필드명`으로 참조) |
| require_confirm | - | true면 사용자 확인 후 실행 (외부 발송 시 권장) |
