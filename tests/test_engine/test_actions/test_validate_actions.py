"""VALIDATE ACTION 핸들러 테스트."""

from __future__ import annotations

from autooffice.engine.actions.validate_actions import ValidateHandler
from autooffice.engine.context import EngineContext


class TestValidateHandler:
    def setup_method(self):
        self.handler = ValidateHandler()
        self.ctx = EngineContext()

    def test_row_count_pass(self):
        data = [{"a": 1}, {"a": 2}, {"a": 3}]
        result = self.handler.execute(
            {"check": "row_count", "source": data, "logic": {"min": 1, "max": 10}},
            self.ctx,
        )
        assert result.success

    def test_row_count_fail(self):
        data = []
        result = self.handler.execute(
            {"check": "row_count", "source": data, "logic": {"min": 1, "max": 10}},
            self.ctx,
        )
        assert not result.success

    def test_value_range_pass(self):
        data = [{"불량률": 3.0}, {"불량률": 4.5}, {"불량률": 1.2}]
        result = self.handler.execute(
            {"check": "value_range", "source": data, "logic": {"column": "불량률", "min": 0, "max": 5}},
            self.ctx,
        )
        assert result.success

    def test_value_range_fail(self):
        data = [{"불량률": 3.0}, {"불량률": 7.5}]
        result = self.handler.execute(
            {"check": "value_range", "source": data, "logic": {"column": "불량률", "min": 0, "max": 5}},
            self.ctx,
        )
        assert not result.success
        assert len(result.data) == 1  # 7.5만 위반

    def test_not_empty_pass(self):
        result = self.handler.execute(
            {"check": "not_empty", "source": [{"a": 1}], "logic": {}},
            self.ctx,
        )
        assert result.success

    def test_not_empty_fail(self):
        result = self.handler.execute(
            {"check": "not_empty", "source": [], "logic": {}},
            self.ctx,
        )
        assert not result.success

    def test_column_exists_pass(self):
        data = [{"날짜": "2026-01-01", "라인명": "A", "수량": 100}]
        result = self.handler.execute(
            {"check": "column_exists", "source": data, "logic": {"columns": ["날짜", "라인명"]}},
            self.ctx,
        )
        assert result.success

    def test_column_exists_fail(self):
        data = [{"날짜": "2026-01-01"}]
        result = self.handler.execute(
            {"check": "column_exists", "source": data, "logic": {"columns": ["날짜", "라인명"]}},
            self.ctx,
        )
        assert not result.success
        assert "라인명" in result.error
