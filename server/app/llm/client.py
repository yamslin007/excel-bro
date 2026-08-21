from __future__ import annotations

import asyncio
import random
from types import TracebackType
from typing import Any

import httpx

from ..capabilities import capabilities, capability_float
from .config import ModelConnection
from .errors import (
    LLMConnectError,
    LLMError,
    LLMHTTPStatusError,
    LLMResponseError,
    LLMTimeoutError,
    LLMTransportError,
)


def _parse_retry_after(value: str | None) -> float | None:
    """Parse a numeric Retry-After header (seconds). HTTP-date form is ignored."""
    if not value:
        return None
    try:
        seconds = float(value.strip())
    except ValueError:
        return None
    return seconds if seconds >= 0 else None


def _capability_nonneg_int(section: str, key: str) -> int:
    value = capabilities()[section][key]
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise RuntimeError(f"能力配置 {section}.{key} 必须是非负整数")
    return value


class OpenAICompatibleClient:
    """Minimal adapter for OpenAI-compatible chat/completions services."""

    def __init__(
        self,
        connection: ModelConnection,
        *,
        timeout: float,
        max_retries: int | None = None,
        retry_base_delay: float | None = None,
        retry_max_delay: float | None = None,
    ) -> None:
        self.connection = connection
        self.timeout = timeout
        self.max_retries = (
            max_retries
            if max_retries is not None
            else _capability_nonneg_int("llm", "maxRetries")
        )
        self.retry_base_delay = (
            retry_base_delay
            if retry_base_delay is not None
            else capability_float("llm", "retryBaseDelaySeconds")
        )
        self.retry_max_delay = (
            retry_max_delay
            if retry_max_delay is not None
            else capability_float("llm", "retryMaxDelaySeconds")
        )
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
        temperature: float | None = None,
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
        if temperature is not None:
            payload["temperature"] = temperature
        if tools is not None:
            payload["tools"] = tools
        if tool_choice is not None:
            payload["tool_choice"] = tool_choice

        attempt = 0
        while True:
            try:
                return await self._request_once(payload)
            except LLMError as error:
                retryable = self._is_retryable(error)
                if not retryable or attempt >= self.max_retries:
                    raise
                delay = self._retry_delay(attempt, error)
                attempt += 1
                await asyncio.sleep(delay)

    async def _request_once(self, payload: dict[str, Any]) -> dict[str, Any]:
        assert self._client is not None
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
            status_error = LLMHTTPStatusError(
                error.response.status_code,
                error.response.text,
            )
            status_error.retry_after = _parse_retry_after(
                error.response.headers.get("Retry-After")
            )
            raise status_error from error
        except httpx.HTTPError as error:
            raise LLMTransportError(
                f"模型服务连接失败：{type(error).__name__}"
            ) from error

        try:
            body = response.json()
        except ValueError as error:
            raise LLMResponseError("模型服务没有返回有效 JSON") from error
        if not isinstance(body, dict):
            raise LLMResponseError("模型服务返回的 JSON 顶层不是对象")
        return body

    @staticmethod
    def _is_retryable(error: LLMError) -> bool:
        if isinstance(error, LLMHTTPStatusError):
            return error.retryable
        # 网络抖动与超时可重试；LLMResponseError 是响应损坏，重试无意义。
        return isinstance(error, (LLMTimeoutError, LLMConnectError, LLMTransportError))

    def _retry_delay(self, attempt: int, error: LLMError) -> float:
        retry_after = getattr(error, "retry_after", None)
        if isinstance(retry_after, (int, float)) and retry_after > 0:
            return min(float(retry_after), self.retry_max_delay)
        backoff = self.retry_base_delay * (2**attempt)
        jitter = backoff * random.uniform(0.0, 0.25)
        return min(backoff + jitter, self.retry_max_delay)
