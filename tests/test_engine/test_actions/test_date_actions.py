"""EXTRACT_DATE 핸들러 테스트."""

from __future__ import annotations

from datetime import date

import pytest

from autooffice.engine.actions.date_actions import (
    ExtractDateHandler,
    _parse_date,
    derive_date_info,
)
from autooffice.engine.context import EngineContext
from autooffice.models.action_result import ActionResult


class TestParseDate:
    """_parse_date: 다양한 패턴에서 날짜를 추출한다."""

    def test_yyyymmdd_from_filename(self):
        assert _parse_date("20260402_production.xlsx", "YYYYMMDD") == date(2026, 4, 2)

    def test_yyyymmdd_embedded(self):
        assert _parse_date("report_20261231_final.xlsx", "YYYYMMDD") == date(2026, 12, 31)

    def test_yyyy_mm_dd_hyphen(self):
        assert _parse_date("data_2026-04-02.csv", "YYYY-MM-DD") == date(2026, 4, 2)

    def test_yyyy_mm_dd_underscore(self):
        assert _parse_date("data_2026_04_02.csv", "YYYY_MM_DD") == date(2026, 4, 2)

    def test_yyyy_mm_dd_slash(self):
        assert _parse_date("2026/04/02 report", "YYYY/MM/DD") == date(2026, 4, 2)

    def test_auto_yyyymmdd(self):
        assert _parse_date("20260402_prod.xlsx", "auto") == date(2026, 4, 2)

    def test_auto_hyphen(self):
        assert _parse_date("report-2026-04-02.xlsx", "auto") == date(2026, 4, 2)

    def test_auto_underscore(self):
        assert _parse_date("report_2026_04_02.xlsx", "auto") == date(2026, 4, 2)

    def test_no_match_returns_none(self):
        assert _parse_date("no_date_here.xlsx", "YYYYMMDD") is None

    def test_invalid_date_returns_none(self):
        assert _parse_date("20261332_bad.xlsx", "YYYYMMDD") is None

    def test_unknown_pattern_returns_none(self):
        assert _parse_date("20260402", "DD/MM/YYYY") is None


class TestDeriveDateInfo:
    """derive_date_info: 날짜에서 파생 정보를 계산한다."""

    def test_basic_fields(self):
        d = date(2026, 4, 2)
        info = derive_date_info(d)
        assert info["date"] == "2026-04-02"
        assert info["yyyymmdd"] == "20260402"
        assert info["year"] == 2026
        assert info["month"] == 4
        assert info["day"] == 2

    def test_weekday(self):
        # 2026-04-02 = Thursday
        info = derive_date_info(date(2026, 4, 2))
        assert info["weekday"] == "Thu"

    def test_week_number(self):
        info = derive_date_info(date(2026, 4, 2))
        assert info["week_number"] == date(2026, 4, 2).isocalendar()[1]

    def test_week_monday_sunday(self):
        info = derive_date_info(date(2026, 4, 2))
        # 2026-04-02(Thu)의 주: Mon=3/30, Sun=4/5
        assert info["week_monday"] == "2026-03-30"
        assert info["week_sunday"] == "2026-04-05"

    def test_month_start_end(self):
        info = derive_date_info(date(2026, 4, 2))
        assert info["month_start"] == "2026-04-01"
        assert info["month_end"] == "2026-04-30"

    def test_quarter(self):
        assert derive_date_info(date(2026, 1, 15))["quarter"] == 1
        assert derive_date_info(date(2026, 4, 2))["quarter"] == 2
        assert derive_date_info(date(2026, 7, 1))["quarter"] == 3
        assert derive_date_info(date(2026, 12, 31))["quarter"] == 4

    def test_december_month_end(self):
        info = derive_date_info(date(2026, 12, 15))
        assert info["month_end"] == "2026-12-31"


class TestExtractDateHandler:
    """ExtractDateHandler: EXTRACT_DATE 액션 핸들러."""

    @pytest.fixture()
    def handler(self):
        return ExtractDateHandler()

    @pytest.fixture()
    def ctx(self, tmp_path):
        return EngineContext(data_dir=tmp_path)

    def test_basic_extraction(self, handler, ctx):
        result = handler.execute({"source": "20260402_prod.xlsx", "pattern": "YYYYMMDD"}, ctx)
        assert result.success
        assert result.data["date"] == "2026-04-02"
        assert result.data["month"] == 4
        assert result.data["week_monday"] == "2026-03-30"

    def test_auto_pattern(self, handler, ctx):
        result = handler.execute({"source": "report-2026-04-02.xlsx", "pattern": "auto"}, ctx)
        assert result.success
        assert result.data["date"] == "2026-04-02"

    def test_default_pattern_is_yyyymmdd(self, handler, ctx):
        result = handler.execute({"source": "20260402_prod.xlsx"}, ctx)
        assert result.success
        assert result.data["date"] == "2026-04-02"

    def test_failure_on_no_date(self, handler, ctx):
        result = handler.execute({"source": "no_date.xlsx", "pattern": "YYYYMMDD"}, ctx)
        assert not result.success
        assert "날짜 추출 실패" in result.error

    def test_failure_on_empty_source(self, handler, ctx):
        result = handler.execute({"source": "", "pattern": "YYYYMMDD"}, ctx)
        assert not result.success

    def test_with_input_variable(self, tmp_path):
        """$input.raw_data 참조를 통한 파일명 추출 시뮬레이션."""
        ctx = EngineContext(data_dir=tmp_path, input_files={"raw_data": "20260402_prod.xlsx"})
        # resolve $input.raw_data → "20260402_prod.xlsx"
        source = ctx.resolve("$input.raw_data")
        handler = ExtractDateHandler()
        result = handler.execute({"source": source, "pattern": "YYYYMMDD"}, ctx)
        assert result.success
        assert result.data["date"] == "2026-04-02"
