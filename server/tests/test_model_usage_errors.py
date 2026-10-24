import asyncio

import httpx
import pytest

from server.app.llm.client import OpenAICompatibleClient
from server.app.llm.config import ModelConnection
from server.app.llm.errors import LLMHTTPStatusError
from server.app.main import _map_turn_error


@pytest.mark.parametrize("status,body,code,retryable", [
    (402, "Insufficient Balance", "MODEL_QUOTA_EXHAUSTED", False),
    (429, '{"error":{"code":"insufficient_quota","message":"PRIVATE_KEY"}}', "MODEL_QUOTA_EXHAUSTED", False),
    (429, '{"error":{"type":"billing_hard_limit_reached"}}', "MODEL_QUOTA_EXHAUSTED", False),
    (400, '{"error":{"message":"Your credit balance is too low"}}', "MODEL_QUOTA_EXHAUSTED", False),
    (400, '{"error":{"code":"context_length_exceeded"}}', "MODEL_CONTEXT_LIMIT", False),
    (429, '{"error":{"code":"rate_limit_exceeded"}}', "MODEL_OVERLOADED", True),
    (429, '{"request":{"prompt":"insufficient_quota"},"error":{"message":"Too many requests"}}', "MODEL_OVERLOADED", True),
    (401, "invalid token", "MODEL_HTTP_ERROR", False),
    (500, "insufficient_quota", "MODEL_HTTP_ERROR", True),
])
def test_limit_classification_is_explicit_and_safe(status, body, code, retryable):
    error = LLMHTTPStatusError(status, body)
    assert error.retryable is retryable
    _, actual_code, message, actual_retryable = _map_turn_error(error)
    assert (actual_code, actual_retryable) == (code, retryable)
    assert "PRIVATE_KEY" not in message
    if code == "MODEL_QUOTA_EXHAUSTED":
        assert "余额不足" in message and "模型设置" in message
    if code == "MODEL_OVERLOADED":
        assert "余额不足" not in message


def test_quota_failure_is_not_automatically_retried(monkeypatch):
    calls = []

    def handler(request):
        calls.append(request)
        return httpx.Response(429, json={"error": {"code": "insufficient_quota"}})

    original_client = httpx.AsyncClient
    monkeypatch.setattr(httpx, "AsyncClient", lambda **kw: original_client(transport=httpx.MockTransport(handler), **kw))
    connection = ModelConnection(base_url="https://example.invalid/v1", model="test", api_key="", supports_vision=False)

    async def run():
        async with OpenAICompatibleClient(connection, timeout=1, max_retries=3, retry_base_delay=0) as client:
            await client.chat_completions(messages=[])

    with pytest.raises(LLMHTTPStatusError):
        asyncio.run(run())
    assert len(calls) == 1


def test_usage_code_contract_matches_both_hosts():
    import re
    from pathlib import Path
    from typing import get_args
    from server.app.models import ModelUsageErrorCode

    core = Path(__file__).resolve().parents[2] / "packages/core/src/modelUsageNotice.ts"
    declared = core.read_text(encoding="utf-8").split("const notices:")[0]
    assert set(re.findall(r'"(MODEL_[A-Z_]+)"', declared)) == set(get_args(ModelUsageErrorCode))
