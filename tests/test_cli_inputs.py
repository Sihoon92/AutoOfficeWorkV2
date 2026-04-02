"""CLI --input 옵션 파싱 및 검증 테스트."""

from __future__ import annotations

from pathlib import Path

import click
import pytest

from autooffice.cli import _parse_inputs, _validate_inputs
from autooffice.models.execution_plan import InputSpec


class TestParseInputs:
    """_parse_inputs: KEY=VALUE 형식 파싱."""

    def test_single_input(self):
        result = _parse_inputs(("raw_data=20260402_prod.xlsx",))
        assert result == {"raw_data": "20260402_prod.xlsx"}

    def test_multiple_inputs(self):
        result = _parse_inputs((
            "raw_data=20260402_prod.xlsx",
            "template=defect_template.xlsx",
        ))
        assert result == {
            "raw_data": "20260402_prod.xlsx",
            "template": "defect_template.xlsx",
        }

    def test_empty_inputs(self):
        result = _parse_inputs(())
        assert result == {}

    def test_value_with_equals(self):
        """값에 = 기호가 포함된 경우."""
        result = _parse_inputs(("key=path/to=file.xlsx",))
        assert result == {"key": "path/to=file.xlsx"}

    def test_invalid_format_raises(self):
        with pytest.raises(click.BadParameter, match="입력 형식 오류"):
            _parse_inputs(("no_equals_sign",))

    def test_whitespace_stripped(self):
        result = _parse_inputs((" raw_data = file.xlsx ",))
        assert result == {"raw_data": "file.xlsx"}


class TestValidateInputs:
    """_validate_inputs: plan inputs와 실제 입력 대조 검증."""

    def test_all_inputs_provided(self, tmp_path):
        """모든 필수 입력이 제공되면 통과."""
        (tmp_path / "data.xlsx").write_bytes(b"")
        (tmp_path / "tpl.xlsx").write_bytes(b"")

        plan_inputs = {
            "raw": InputSpec(description="raw", expected_format="xlsx"),
            "tpl": InputSpec(description="tpl", expected_format="xlsx"),
        }
        # 예외 없이 통과
        _validate_inputs(plan_inputs, {"raw": "data.xlsx", "tpl": "tpl.xlsx"}, tmp_path)

    def test_missing_input_raises(self, tmp_path):
        """필수 입력 누락 시 UsageError."""
        plan_inputs = {
            "raw": InputSpec(description="raw", expected_format="xlsx"),
            "tpl": InputSpec(description="tpl", expected_format="xlsx"),
        }
        with pytest.raises(click.UsageError, match="필수 입력 파일 누락"):
            _validate_inputs(plan_inputs, {"raw": "data.xlsx"}, tmp_path)

    def test_file_not_found_raises(self, tmp_path):
        """파일이 존재하지 않으면 FileError."""
        plan_inputs = {
            "raw": InputSpec(description="raw", expected_format="xlsx"),
        }
        with pytest.raises(click.FileError):
            _validate_inputs(plan_inputs, {"raw": "nonexistent.xlsx"}, tmp_path)

    def test_extra_inputs_allowed(self, tmp_path):
        """plan에 없는 추가 입력은 허용한다 (에러 아님)."""
        (tmp_path / "data.xlsx").write_bytes(b"")
        (tmp_path / "extra.xlsx").write_bytes(b"")

        plan_inputs = {
            "raw": InputSpec(description="raw", expected_format="xlsx"),
        }
        # extra 입력이 있어도 에러 아님
        _validate_inputs(
            plan_inputs,
            {"raw": "data.xlsx", "extra": "extra.xlsx"},
            tmp_path,
        )

    def test_empty_plan_inputs(self, tmp_path):
        """plan에 inputs가 비어있으면 누락 검증 통과 (파일 존재 검증만)."""
        (tmp_path / "data.xlsx").write_bytes(b"")
        _validate_inputs({}, {"raw": "data.xlsx"}, tmp_path)
