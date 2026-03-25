"""로깅 ACTION 핸들러: LOG."""

from __future__ import annotations

import logging
from typing import Any

from autooffice.engine.actions.base import ActionHandler
from autooffice.engine.context import EngineContext
from autooffice.models.action_result import ActionResult

logger = logging.getLogger(__name__)


class LogHandler(ActionHandler):
    """LOG: 실행 로그를 기록한다.

    params:
        message: 로그 메시지
        level: 로그 레벨 (info/warn/error, 기본 info)
    """

    def execute(self, params: dict[str, Any], ctx: EngineContext) -> ActionResult:
        message = params.get("message", "")
        level = params.get("level", "info").lower()

        log_func = {
            "info": logger.info,
            "warn": logger.warning,
            "warning": logger.warning,
            "error": logger.error,
        }.get(level, logger.info)

        log_func("[LOG] %s", message)
        ctx.log_messages.append(f"[{level.upper()}] {message}")

        return ActionResult(
            success=True,
            message=message,
        )
