"""EngineContext 변수 참조 해소 테스트."""

from __future__ import annotations

import pytest

from autooffice.engine.context import EngineContext


class TestResolveDotNotation:
    """$var.field 형식의 점 표기법 변수 참조 해소."""

    def test_simple_variable_still_works(self, tmp_path):
        """기존 $var 참조는 그대로 동작한다."""
        ctx = EngineContext(data_dir=tmp_path)
        ctx.store("name", "hello")
        assert ctx.resolve("$name") == "hello"

    def test_dot_notation_dict_field(self, tmp_path):
        """$var.field로 dict 필드에 접근한다."""
        ctx = EngineContext(data_dir=tmp_path)
        ctx.store("date_info", {"column": "CP", "date_row": 6})
        assert ctx.resolve("$date_info.column") == "CP"
        assert ctx.resolve("$date_info.date_row") == 6

    def test_dot_notation_missing_field(self, tmp_path):
        """존재하지 않는 필드 접근 시 KeyError."""
        ctx = EngineContext(data_dir=tmp_path)
        ctx.store("info", {"column": "CP"})
        with pytest.raises(KeyError, match="no_such_field"):
            ctx.resolve("$info.no_such_field")

    def test_dot_notation_on_non_dict(self, tmp_path):
        """dict가 아닌 변수에 점 표기법 사용 시 TypeError."""
        ctx = EngineContext(data_dir=tmp_path)
        ctx.store("count", 42)
        with pytest.raises(TypeError):
            ctx.resolve("$count.field")

    def test_non_variable_passthrough(self, tmp_path):
        """$ 없는 일반 문자열은 그대로 반환."""
        ctx = EngineContext(data_dir=tmp_path)
        assert ctx.resolve("plain_text") == "plain_text"
        assert ctx.resolve(123) == 123

    def test_resolve_params_with_dot_notation(self, tmp_path):
        """resolve_params에서 점 표기법이 동작한다."""
        ctx = EngineContext(data_dir=tmp_path)
        ctx.store("date_info", {"column": "CP", "date_row": 6})
        params = {
            "target_column": "$date_info.column",
            "target_row": "$date_info.date_row",
            "fixed_value": "hello",
        }
        resolved = ctx.resolve_params(params)
        assert resolved["target_column"] == "CP"
        assert resolved["target_row"] == 6
        assert resolved["fixed_value"] == "hello"
