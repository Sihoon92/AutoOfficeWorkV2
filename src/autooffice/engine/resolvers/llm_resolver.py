"""LLM 기반 동적 파라미터 해소기.

langchain_openai의 ChatOpenAI를 사용하여 사내 API에
동적 파라미터 해소를 요청한다.
"""

from __future__ import annotations

import json
import logging
from datetime import date
from typing import Any

from autooffice.engine.resolvers.base import DynamicResolver
from autooffice.models.execution_plan import DynamicParamSpec

logger = logging.getLogger(__name__)

SYSTEM_PROMPT = """\
당신은 엑셀 자동화 파라미터 해소기입니다.
오늘 날짜: {today}

각 파라미터의 지시에 따라 값을 계산하여 JSON 객체로 반환하세요.
반드시 유효한 JSON만 출력하세요. 설명이나 마크다운 없이 JSON만 반환하세요.

응답 형식:
{{"param_key1": "resolved_value1", "param_key2": {{"nested": "value"}}}}
"""


class LLMDynamicResolver(DynamicResolver):
    """사내 LLM API를 통해 동적 파라미터를 해소하는 resolver.

    langchain_openai ChatOpenAI를 사용하여 모든 미해소 파라미터를
    하나의 프롬프트로 일괄 처리한다 (API 호출 최소화).
    """

    def __init__(
        self,
        model_name: str = "gpt-4o-mini",
        api_key: str | None = None,
        base_url: str | None = None,
        temperature: float = 0,
    ) -> None:
        self._model_name = model_name
        self._api_key = api_key
        self._base_url = base_url
        self._temperature = temperature
        self._llm = None

    def _get_llm(self):
        """ChatOpenAI 인스턴스를 lazy 생성한다."""
        if self._llm is None:
            from langchain_openai import ChatOpenAI

            kwargs: dict[str, Any] = {
                "model": self._model_name,
                "temperature": self._temperature,
            }
            if self._api_key:
                kwargs["api_key"] = self._api_key
            if self._base_url:
                kwargs["base_url"] = self._base_url
            self._llm = ChatOpenAI(**kwargs)
        return self._llm

    def resolve(
        self, declarations: dict[str, DynamicParamSpec]
    ) -> dict[str, Any]:
        if not declarations:
            return {}

        today = date.today().isoformat()
        system_msg = SYSTEM_PROMPT.format(today=today)

        # 사용자 프롬프트: 각 파라미터의 해소 지시
        param_lines = []
        for key, spec in declarations.items():
            line = f"- {key}"
            if spec.format:
                line += f" (출력 형식: {spec.format})"
            line += f": {spec.prompt}"
            param_lines.append(line)

        user_msg = "다음 파라미터를 해소하세요:\n\n" + "\n".join(param_lines)

        logger.info("LLM 해소 요청: %d개 파라미터", len(declarations))
        logger.debug("프롬프트:\n%s", user_msg)

        try:
            llm = self._get_llm()
            response = llm.invoke([
                {"role": "system", "content": system_msg},
                {"role": "user", "content": user_msg},
            ])

            raw_text = response.content.strip()
            # JSON 블록 추출 (```json ... ``` 래핑 대응)
            if raw_text.startswith("```"):
                raw_text = raw_text.split("\n", 1)[1].rsplit("```", 1)[0].strip()

            resolved = json.loads(raw_text)
            logger.info("LLM 해소 완료: %s", list(resolved.keys()))
            return resolved

        except Exception as e:
            logger.error("LLM 해소 실패: %s", e)
            # 폴백: default 값 사용
            fallback = {}
            for key, spec in declarations.items():
                if spec.default is not None:
                    fallback[key] = spec.default
                    logger.warning("폴백 사용: %s → %s", key, spec.default)
            return fallback
