"""ACTION 핸들러 패키지.

각 ACTION 타입은 ActionHandler를 구현하여 action_registry에 등록된다.
"""

from autooffice.engine.actions.base import ActionHandler
from autooffice.engine.actions.file_actions import OpenFileHandler, SaveFileHandler
from autooffice.engine.actions.excel_actions import (
    AggregateRangeHandler,
    CopyRangeHandler,
    FindAnchorHandler,
    FindDateColumnHandler,
    FindDateRangeHandler,
    ReadColumnsHandler,
    ReadRangeHandler,
    WriteDataHandler,
    ClearRangeHandler,
    RecalculateHandler,
)
from autooffice.engine.actions.date_actions import ExtractDateHandler
from autooffice.engine.actions.validate_actions import ValidateHandler
from autooffice.engine.actions.format_actions import FormatMessageHandler
from autooffice.engine.actions.messenger_actions import SendMessengerHandler
from autooffice.engine.actions.log_actions import LogHandler


def build_default_registry() -> dict[str, ActionHandler]:
    """기본 ACTION 핸들러 레지스트리를 생성한다."""
    return {
        "OPEN_FILE": OpenFileHandler(),
        "SAVE_FILE": SaveFileHandler(),
        "READ_COLUMNS": ReadColumnsHandler(),
        "READ_RANGE": ReadRangeHandler(),
        "WRITE_DATA": WriteDataHandler(),
        "CLEAR_RANGE": ClearRangeHandler(),
        "RECALCULATE": RecalculateHandler(),
        "FIND_DATE_COLUMN": FindDateColumnHandler(),
        "COPY_RANGE": CopyRangeHandler(),
        "AGGREGATE_RANGE": AggregateRangeHandler(),
        "FIND_ANCHOR": FindAnchorHandler(),
        "FIND_DATE_RANGE": FindDateRangeHandler(),
        "EXTRACT_DATE": ExtractDateHandler(),
        "VALIDATE": ValidateHandler(),
        "FORMAT_MESSAGE": FormatMessageHandler(),
        "SEND_MESSENGER": SendMessengerHandler(),
        "LOG": LogHandler(),
    }


__all__ = [
    "ActionHandler",
    "build_default_registry",
]
