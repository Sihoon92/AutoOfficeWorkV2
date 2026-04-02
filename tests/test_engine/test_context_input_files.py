"""EngineContext input_files 관련 테스트."""

from __future__ import annotations

import pytest

from autooffice.engine.context import EngineContext


class TestInputFiles:
    """EngineContext: input_files 매핑 및 $input 변수 참조."""

    def test_input_files_stored(self, tmp_path):
        """input_files가 ctx에 저장된다."""
        ctx = EngineContext(data_dir=tmp_path, input_files={"raw": "data.xlsx"})
        assert ctx.input_files == {"raw": "data.xlsx"}

    def test_input_available_as_variable(self, tmp_path):
        """$input.KEY로 입력 파일명을 참조할 수 있다."""
        ctx = EngineContext(data_dir=tmp_path, input_files={"raw_data": "20260402_prod.xlsx"})
        assert ctx.resolve("$input.raw_data") == "20260402_prod.xlsx"

    def test_input_multiple_keys(self, tmp_path):
        """여러 입력 키를 동시에 참조할 수 있다."""
        ctx = EngineContext(
            data_dir=tmp_path,
            input_files={"raw_data": "20260402.xlsx", "template": "tpl.xlsx"},
        )
        assert ctx.resolve("$input.raw_data") == "20260402.xlsx"
        assert ctx.resolve("$input.template") == "tpl.xlsx"

    def test_input_resolve_params(self, tmp_path):
        """resolve_params에서 $input 참조가 동작한다."""
        ctx = EngineContext(data_dir=tmp_path, input_files={"raw": "file.xlsx"})
        resolved = ctx.resolve_params({"source": "$input.raw", "fixed": "hello"})
        assert resolved["source"] == "file.xlsx"
        assert resolved["fixed"] == "hello"

    def test_no_input_files(self, tmp_path):
        """input_files 미지정 시 빈 dict."""
        ctx = EngineContext(data_dir=tmp_path)
        assert ctx.input_files == {}
        assert "input" not in ctx.variables

    def test_missing_input_key_raises(self, tmp_path):
        """존재하지 않는 input key 접근 시 KeyError."""
        ctx = EngineContext(data_dir=tmp_path, input_files={"raw": "file.xlsx"})
        with pytest.raises(KeyError):
            ctx.resolve("$input.nonexistent")

    def test_input_does_not_conflict_with_store_as(self, tmp_path):
        """store_as 변수와 input 변수가 공존한다."""
        ctx = EngineContext(data_dir=tmp_path, input_files={"raw": "file.xlsx"})
        ctx.store("my_data", [1, 2, 3])
        assert ctx.resolve("$input.raw") == "file.xlsx"
        assert ctx.resolve("$my_data") == [1, 2, 3]
