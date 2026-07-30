from __future__ import annotations

from .client import OpenAICompatibleClient
from .config import ModelConnection
from .errors import LLMResponseError


async def test_model_connection(connection: ModelConnection) -> None:
    """Send a tiny request without persisting the draft connection."""
    async with OpenAICompatibleClient(connection, timeout=20) as client:
        payload = await client.chat_completions(
            messages=[
                {
                    "role": "user",
                    "content": "Reply with OK.",
                }
            ],
            max_tokens=4,
        )
    choices = payload.get("choices")
    if not isinstance(choices, list) or not choices:
        raise LLMResponseError("模型服务已响应，但没有返回有效的 choices")
