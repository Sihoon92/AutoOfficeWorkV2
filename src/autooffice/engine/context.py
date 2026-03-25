"""런타임 컨텍스트 - step 간 상태 공유.

EngineContext는 실행 중인 plan의 런타임 상태를 관리한다:
- Excel 앱 인스턴스 관리
- 열린 워크북 핸들
- step 간 데이터 전달용 변수 저장소
- 실행 이력 로그
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

import xlwings as xw

logger = logging.getLogger(__name__)


class EngineContext:
    """실행 엔진의 런타임 컨텍스트.

    모든 ACTION 핸들러는 이 컨텍스트를 공유하여 상태를 주고받는다.
    store_as로 저장한 값은 variables에서 참조 가능하다.
    """

    def __init__(self, data_dir: str | Path = ".", app: xw.App | None = None) -> None:
        self.data_dir = Path(data_dir)
        self._app = app
        self._owns_app = app is None
        self.open_workbooks: dict[str, xw.Book] = {}
        self.file_paths: dict[str, Path] = {}
        self.variables: dict[str, Any] = {}
        self.log_messages: list[str] = []
        self.dry_run: bool = False

    @property
    def app(self) -> xw.App:
        """Excel 앱 인스턴스 (lazy initialization)."""
        if self._app is None:
            self._app = xw.App(visible=False)
            self._app.display_alerts = False
            self._owns_app = True
        return self._app

    def store(self, key: str, value: Any) -> None:
        """변수 저장소에 값을 저장한다."""
        self.variables[key] = value
        logger.debug("변수 저장: %s = %s", key, type(value).__name__)

    def resolve(self, value: Any) -> Any:
        """값이 $variable_name 형식이면 변수 저장소에서 참조를 해소한다."""
        if isinstance(value, str) and value.startswith("$"):
            var_name = value[1:]
            if var_name not in self.variables:
                raise KeyError(f"변수 '{var_name}'이(가) 정의되지 않았습니다.")
            return self.variables[var_name]
        return value

    def resolve_params(self, params: dict[str, Any]) -> dict[str, Any]:
        """params 딕셔너리의 모든 값에 대해 변수 참조를 해소한다."""
        resolved = {}
        for k, v in params.items():
            resolved[k] = self.resolve(v)
        return resolved

    def register_workbook(self, alias: str, workbook: xw.Book, path: Path) -> None:
        """워크북을 컨텍스트에 등록한다."""
        self.open_workbooks[alias] = workbook
        self.file_paths[alias] = path
        logger.info("워크북 등록: %s -> %s", alias, path)

    def get_workbook(self, alias: str) -> xw.Book:
        """등록된 워크북을 가져온다."""
        if alias not in self.open_workbooks:
            raise KeyError(f"워크북 '{alias}'이(가) 열려있지 않습니다.")
        return self.open_workbooks[alias]

    def close_all(self) -> None:
        """모든 워크북을 닫고, 자체 생성한 Excel 앱이면 종료한다."""
        for alias, wb in self.open_workbooks.items():
            try:
                wb.close()
                logger.debug("워크북 닫힘: %s", alias)
            except Exception:
                logger.warning("워크북 닫기 실패: %s", alias)
        self.open_workbooks.clear()
        self.file_paths.clear()
        if self._owns_app and self._app is not None:
            try:
                self._app.quit()
                logger.debug("Excel 앱 종료")
            except Exception:
                logger.warning("Excel 앱 종료 실패")
            self._app = None
