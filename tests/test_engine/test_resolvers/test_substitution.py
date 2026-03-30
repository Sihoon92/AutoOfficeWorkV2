"""{{dynamic:...}} 치환 유틸리티 테스트."""

from __future__ import annotations

from autooffice.engine.resolvers.substitution import (
    extract_dynamic_keys,
    substitute_params,
)


class TestSubstituteParams:
    """substitute_params: 동적 마커를 해소된 값으로 치환."""

    def test_simple_string_replacement(self):
        """문자열 전체가 마커인 경우 치환."""
        resolved = {"today_date": "2026-03-30"}
        params = {"date": "{{dynamic:today_date}}"}
        result = substitute_params(params, resolved)
        assert result["date"] == "2026-03-30"

    def test_dot_notation_field_access(self):
        """점 표기법으로 JSON 응답 필드 접근."""
        resolved = {
            "today_target": {"date": "2026-03-30", "column": "CQ", "date_row": 3}
        }
        params = {
            "target_column": "{{dynamic:today_target.column}}",
            "target_start_row": "{{dynamic:today_target.date_row}}",
        }
        result = substitute_params(params, resolved)
        assert result["target_column"] == "CQ"
        assert result["target_start_row"] == 3  # int 타입 보존

    def test_partial_string_substitution(self):
        """부분 문자열 내 마커 치환."""
        resolved = {"today_date": "2026-03-30"}
        params = {"filename": "report_{{dynamic:today_date}}.xlsx"}
        result = substitute_params(params, resolved)
        assert result["filename"] == "report_2026-03-30.xlsx"

    def test_no_markers_passthrough(self):
        """마커 없는 params는 그대로 반환."""
        params = {"file": "test.xlsx", "sheet": "Sheet1", "count": 42}
        result = substitute_params(params, {})
        assert result == params

    def test_nested_dict(self):
        """중첩 dict 내 마커 치환."""
        resolved = {"today_target": {"column": "CQ"}}
        params = {"mapping": {"target": "{{dynamic:today_target.column}}"}}
        result = substitute_params(params, resolved)
        assert result["mapping"]["target"] == "CQ"

    def test_nested_list(self):
        """리스트 내 마커 치환."""
        resolved = {"v": "hello"}
        params = {"items": ["{{dynamic:v}}", "static"]}
        result = substitute_params(params, resolved)
        assert result["items"] == ["hello", "static"]

    def test_unresolved_marker_preserved(self):
        """미해소 마커는 원본 유지."""
        params = {"val": "{{dynamic:unknown_key}}"}
        result = substitute_params(params, {})
        assert result["val"] == "{{dynamic:unknown_key}}"

    def test_unresolved_field_preserved(self):
        """존재하지 않는 필드 접근 시 마커 유지."""
        resolved = {"target": {"column": "CQ"}}
        params = {"val": "{{dynamic:target.nonexistent}}"}
        result = substitute_params(params, resolved)
        assert result["val"] == "{{dynamic:target.nonexistent}}"

    def test_mixed_static_and_dynamic(self):
        """정적 + 동적 params 혼합."""
        resolved = {"today_date": "2026-03-30"}
        params = {
            "workbook": "material",
            "date": "{{dynamic:today_date}}",
            "sheet": "Sheet1",
        }
        result = substitute_params(params, resolved)
        assert result["workbook"] == "material"
        assert result["date"] == "2026-03-30"
        assert result["sheet"] == "Sheet1"

    def test_original_params_not_mutated(self):
        """원본 params는 변경되지 않는다."""
        resolved = {"v": "new_value"}
        params = {"key": "{{dynamic:v}}"}
        original_val = params["key"]
        substitute_params(params, resolved)
        assert params["key"] == original_val


class TestExtractDynamicKeys:
    """extract_dynamic_keys: step params에서 동적 키 추출."""

    def test_basic_extraction(self):
        """기본 키 추출."""
        params_list = [
            {"date": "{{dynamic:today_date}}"},
            {"col": "{{dynamic:today_target.column}}"},
        ]
        keys = extract_dynamic_keys(params_list)
        assert keys == {"today_date", "today_target"}

    def test_no_markers(self):
        """마커 없으면 빈 set."""
        params_list = [{"file": "test.xlsx"}]
        assert extract_dynamic_keys(params_list) == set()

    def test_dedup(self):
        """같은 키 여러 번 참조해도 하나만."""
        params_list = [
            {"a": "{{dynamic:k}}", "b": "{{dynamic:k.field}}"},
        ]
        assert extract_dynamic_keys(params_list) == {"k"}
