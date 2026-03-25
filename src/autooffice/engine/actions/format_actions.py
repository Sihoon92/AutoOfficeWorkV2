"""메시지 포맷팅 ACTION 핸들러: FORMAT_MESSAGE."""

from __future__ import annotations

import logging
from typing import Any

from autooffice.engine.actions.base import ActionHandler
from autooffice.engine.context import EngineContext
from autooffice.models.action_result import ActionResult

logger = logging.getLogger(__name__)


class FormatMessageHandler(ActionHandler):
    """FORMAT_MESSAGE: 템플릿 문자열에 데이터를 주입하여 메시지를 구성한다.

    params:
        template: 메시지 템플릿 (Python format 문법 또는 Jinja2)
            예: "[{date}] 품질현황\\n불량률: {defect_rate}%"
        data_source: 데이터 소스 ($변수명 참조 또는 직접 dict)
        use_jinja: True면 Jinja2 렌더링 (기본 False)
    """

    def execute(self, params: dict[str, Any], ctx: EngineContext) -> ActionResult:
        template = params.get("template", "")
        data_source = params.get("data_source", {})
        use_jinja = params.get("use_jinja", False)

        try:
            if use_jinja:
                message = self._render_jinja(template, data_source)
            else:
                message = self._render_format(template, data_source)

            return ActionResult(
                success=True,
                data=message,
                message=f"메시지 포맷팅 완료 ({len(message)}자)",
            )
        except Exception as e:
            return ActionResult(success=False, error=f"FORMAT_MESSAGE 실패: {e}")

    def _render_format(self, template: str, data: Any) -> str:
        """Python str.format()으로 렌더링."""
        if isinstance(data, dict):
            return template.format(**data)
        elif isinstance(data, list) and len(data) > 0:
            # 리스트면 첫 번째 항목으로 렌더링
            if isinstance(data[0], dict):
                return template.format(**data[0])
        return template

    def _render_jinja(self, template: str, data: Any) -> str:
        """Jinja2로 렌더링."""
        from jinja2 import Template

        tmpl = Template(template)
        if isinstance(data, dict):
            return tmpl.render(**data)
        elif isinstance(data, list):
            return tmpl.render(rows=data)
        return tmpl.render(data=data)
