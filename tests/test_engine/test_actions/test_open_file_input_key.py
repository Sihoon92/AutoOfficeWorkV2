"""OpenFileHandler input_key 지원 테스트."""

from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from autooffice.engine.actions.file_actions import OpenFileHandler
from autooffice.engine.context import EngineContext


def _make_ctx_with_mock_app(tmp_path, input_files=None):
    """mock Excel app이 주입된 EngineContext를 생성한다."""
    mock_app = MagicMock()
    mock_wb = MagicMock()
    mock_wb.sheets = [MagicMock(name="Sheet1")]
    mock_app.books.open.return_value = mock_wb

    ctx = EngineContext(data_dir=tmp_path, input_files=input_files or {}, app=mock_app)
    return ctx, mock_app


class TestOpenFileInputKey:
    """OpenFileHandler: input_key 방식과 기존 file_path 하위 호환."""

    def test_input_key_resolves_to_file(self, tmp_path):
        """input_key로 input_files에서 파일 경로를 가져온다."""
        (tmp_path / "20260402_prod.xlsx").write_bytes(b"")
        ctx, mock_app = _make_ctx_with_mock_app(
            tmp_path, input_files={"raw_data": "20260402_prod.xlsx"}
        )
        handler = OpenFileHandler()

        result = handler.execute({"input_key": "raw_data", "alias": "raw"}, ctx)

        assert result.success
        mock_app.books.open.assert_called_once()

    def test_input_key_missing_returns_error(self, tmp_path):
        """존재하지 않는 input_key는 에러를 반환한다."""
        ctx = EngineContext(data_dir=tmp_path, input_files={})
        handler = OpenFileHandler()
        result = handler.execute({"input_key": "nonexistent", "alias": "raw"}, ctx)
        assert not result.success
        assert "미지정" in result.error

    def test_file_path_still_works(self, tmp_path):
        """기존 file_path 방식도 여전히 동작한다."""
        (tmp_path / "template.xlsx").write_bytes(b"")
        ctx, mock_app = _make_ctx_with_mock_app(tmp_path)
        handler = OpenFileHandler()

        result = handler.execute({"file_path": "template.xlsx", "alias": "tpl"}, ctx)

        assert result.success
        mock_app.books.open.assert_called_once()

    def test_input_key_takes_precedence_over_file_path(self, tmp_path):
        """input_key와 file_path가 둘 다 있으면 input_key가 우선한다."""
        (tmp_path / "from_input.xlsx").write_bytes(b"")
        ctx, mock_app = _make_ctx_with_mock_app(
            tmp_path, input_files={"raw": "from_input.xlsx"}
        )
        handler = OpenFileHandler()

        result = handler.execute(
            {"input_key": "raw", "file_path": "other.xlsx", "alias": "raw"}, ctx
        )

        assert result.success
        call_path = mock_app.books.open.call_args[0][0]
        assert "from_input.xlsx" in call_path
