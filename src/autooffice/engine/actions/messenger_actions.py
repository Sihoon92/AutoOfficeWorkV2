"""메신저 발송 ACTION 핸들러: SEND_MESSENGER."""

from __future__ import annotations

import logging
from typing import Any

from autooffice.engine.actions.base import ActionHandler
from autooffice.engine.context import EngineContext
from autooffice.models.action_result import ActionResult

logger = logging.getLogger(__name__)


class SendMessengerHandler(ActionHandler):
    """SEND_MESSENGER: 메신저로 메시지를 발송한다.

    params:
        to: 수신자/채널 (채널ID, 사용자ID, 또는 이름)
        message: 발송할 메시지 ($변수명 참조 가능)
        attachment: 첨부 파일 경로 (선택)

    Note: 이 핸들러는 사내 메신저 API에 맞게 커스터마이즈해야 한다.
    현재는 로그만 남기는 스텁 구현이다. 실제 환경에서는
    requests 등으로 사내 메신저 API를 호출하도록 수정한다.
    """

    def execute(self, params: dict[str, Any], ctx: EngineContext) -> ActionResult:
        to = params.get("to", "")
        message = params.get("message", "")
        attachment = params.get("attachment")

        if not to:
            return ActionResult(success=False, error="수신자(to)가 지정되지 않았습니다.")
        if not message:
            return ActionResult(success=False, error="메시지(message)가 비어있습니다.")

        # === 사내 메신저 API 호출 지점 ===
        # 실제 환경에서는 아래 주석을 해제하고 사내 API에 맞게 수정하세요.
        #
        # import requests
        # response = requests.post(
        #     "https://internal-messenger.company.com/api/send",
        #     json={"to": to, "message": message, "attachment": attachment},
        #     headers={"Authorization": f"Bearer {os.environ['MESSENGER_TOKEN']}"},
        # )
        # if response.status_code != 200:
        #     return ActionResult(success=False, error=f"메신저 API 오류: {response.text}")

        logger.info("📨 메신저 발송 (스텁): to=%s, message=%s자", to, len(message))
        if attachment:
            logger.info("   첨부파일: %s", attachment)

        return ActionResult(
            success=True,
            data={"to": to, "message_length": len(message)},
            message=f"메신저 발송 완료 (to: {to})",
        )
