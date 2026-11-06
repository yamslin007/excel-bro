from __future__ import annotations

from .client import OpenAICompatibleClient
from .config import ModelConnection
from .errors import LLMResponseError, LLMTimeoutError
from ..capabilities import capability_float, capability_int
import asyncio


async def test_model_connection(connection: ModelConnection) -> None:
    """Send a tiny request without persisting the draft connection."""
    timeout = capability_float("modelConnectionTest", "timeoutSeconds")
    try:
        async with asyncio.timeout(timeout):
            async with OpenAICompatibleClient(connection, timeout=timeout, max_retries=0) as client:
                payload, _ = await client.chat_completions(
                    messages=[{"role": "user", "content": "Reply with OK."}],
                    max_tokens=capability_int("modelConnectionTest", "maxOutputTokens"),
                )
    except TimeoutError as error:
        raise LLMTimeoutError("模型连接测试超时。") from error
    choices = payload.get("choices")
    if not isinstance(choices, list) or not choices:
        raise LLMResponseError("模型服务已响应，但没有返回有效的 choices")
    choice = choices[0]
    message = choice.get("message") if isinstance(choice, dict) else None
    content = message.get("content") if isinstance(message, dict) else None
    if not isinstance(content, str) or not content.strip():
        raise LLMResponseError("模型服务已响应，但没有返回有效文本。")
    if choice.get("finish_reason") not in {None, "stop"}:
        raise LLMResponseError("模型服务已响应，但回复不完整，请检查模型输出配置。")
