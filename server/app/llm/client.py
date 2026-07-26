from __future__ import annotations

from types import TracebackType
from typing import Any

import httpx

from .config import ModelConnection
from .errors import (
    LLMConnectError,
    LLMHTTPStatusError,
    LLMResponseError,
    LLMTimeoutError,
    LLMTransportError,
)


class OpenAICompatibleClient:
    """Minimal adapter for OpenAI-compatible chat/completions services."""

    def __init__(self, connection: ModelConnection, *, timeout: float) -> None:
        self.connection = connection
        self.timeout = timeout
        self._client: httpx.AsyncClient | None = None

    async def __aenter__(self) -> OpenAICompatibleClient:
        self._client = httpx.AsyncClient(timeout=self.timeout)
        return self

    async def __aexit__(
        self,
        exc_type: type[BaseException] | None,
        exc_value: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        if self._client is not None:
            await self._client.aclose()
            self._client = None

    def _headers(self) -> dict[str, str]:
        headers = {"Content-Type": "application/json"}
        if self.connection.api_key:
            headers["Authorization"] = f"Bearer {self.connection.api_key}"
        return headers

    async def chat_completions(
        self,
        *,
        messages: list[dict[str, Any]],
        max_tokens: int | None = None,
        tools: list[dict[str, Any]] | None = None,
        tool_choice: str | None = None,
    ) -> dict[str, Any]:
        if self._client is None:
            raise RuntimeError("OpenAICompatibleClient 必须在 async with 中使用")
        payload: dict[str, Any] = {
            "model": self.connection.model,
            "messages": messages,
        }
        if max_tokens is not None:
            payload["max_tokens"] = max_tokens
        if tools is not None:
            payload["tools"] = tools
        if tool_choice is not None:
            payload["tool_choice"] = tool_choice

        try:
            response = await self._client.post(
                f"{self.connection.base_url}/chat/completions",
                headers=self._headers(),
                json=payload,
            )
            response.raise_for_status()
        except httpx.TimeoutException as error:
            raise LLMTimeoutError("模型响应超时") from error
        except httpx.ConnectError as error:
            raise LLMConnectError("无法连接模型服务") from error
        except httpx.HTTPStatusError as error:
            raise LLMHTTPStatusError(
                error.response.status_code,
                error.response.text,
            ) from error
        except httpx.HTTPError as error:
            raise LLMTransportError(
                f"模型服务连接失败：{type(error).__name__}"
            ) from error

        try:
            payload = response.json()
        except ValueError as error:
            raise LLMResponseError("模型服务没有返回有效 JSON") from error
        if not isinstance(payload, dict):
            raise LLMResponseError("模型服务返回的 JSON 顶层不是对象")
        return payload
