from __future__ import annotations

import asyncio
import json

import httpx
import pytest

from server.app.llm import (
    LLMHTTPStatusError,
    LLMTimeoutError,
    ModelConnection,
    OpenAICompatibleClient,
    selected_model_config,
)


@pytest.fixture(autouse=True)
def clear_model_environment(monkeypatch: pytest.MonkeyPatch) -> None:
    for name in (
        "AI_API_KEY",
        "AI_BASE_URL",
        "AI_MODEL",
        "AI_MODELS",
        "AI_VISION_MODELS",
        "OPENAI_API_KEY",
        "OPENAI_BASE_URL",
        "OPENAI_MODEL",
    ):
        monkeypatch.delenv(name, raising=False)


def test_kimi_configuration_is_isolated_in_model_connection(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("AI_BASE_URL", "https://api.moonshot.cn/v1/")
    monkeypatch.setenv("AI_API_KEY", "private-key")
    monkeypatch.setenv("AI_MODEL", "kimi-k3")
    monkeypatch.setenv("AI_VISION_MODELS", "kimi-k3")

    connection = selected_model_config()

    assert connection == ModelConnection(
        base_url="https://api.moonshot.cn/v1",
        model="kimi-k3",
        api_key="private-key",
        supports_vision=True,
    )


def test_openai_compatible_client_owns_auth_and_request_shape(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    async def handler(request: httpx.Request) -> httpx.Response:
        captured["url"] = str(request.url)
        captured["authorization"] = request.headers.get("authorization")
        captured["body"] = json.loads(request.content)
        return httpx.Response(
            200,
            json={"choices": [{"message": {"content": "ok"}}]},
        )

    transport = httpx.MockTransport(handler)
    original_client = httpx.AsyncClient
    monkeypatch.setattr(
        httpx,
        "AsyncClient",
        lambda **kwargs: original_client(transport=transport, **kwargs),
    )
    connection = ModelConnection(
        base_url="https://api.example.test/v1",
        model="test-model",
        api_key="secret",
        supports_vision=False,
    )

    async def run() -> dict[str, object]:
        async with OpenAICompatibleClient(
            connection,
            timeout=12,
        ) as client:
            return await client.chat_completions(
                messages=[{"role": "user", "content": "hello"}],
                max_tokens=100,
            )

    response = asyncio.run(run())

    assert response["choices"][0]["message"]["content"] == "ok"
    assert captured["url"] == "https://api.example.test/v1/chat/completions"
    assert captured["authorization"] == "Bearer secret"
    assert captured["body"] == {
        "model": "test-model",
        "messages": [{"role": "user", "content": "hello"}],
        "max_tokens": 100,
    }


def test_adapter_normalizes_http_status_errors(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    transport = httpx.MockTransport(
        lambda _request: httpx.Response(429, text="rate limited")
    )
    original_client = httpx.AsyncClient
    monkeypatch.setattr(
        httpx,
        "AsyncClient",
        lambda **kwargs: original_client(transport=transport, **kwargs),
    )
    connection = ModelConnection(
        base_url="https://api.example.test/v1",
        model="test-model",
        api_key="",
        supports_vision=False,
    )

    async def run() -> None:
        async with OpenAICompatibleClient(
            connection,
            timeout=12,
        ) as client:
            await client.chat_completions(messages=[])

    with pytest.raises(LLMHTTPStatusError) as captured:
        asyncio.run(run())

    assert captured.value.status_code == 429
    assert captured.value.retryable is True
    assert captured.value.body == "rate limited"


def test_adapter_normalizes_timeouts(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    async def handler(request: httpx.Request) -> httpx.Response:
        raise httpx.ReadTimeout("slow", request=request)

    transport = httpx.MockTransport(handler)
    original_client = httpx.AsyncClient
    monkeypatch.setattr(
        httpx,
        "AsyncClient",
        lambda **kwargs: original_client(transport=transport, **kwargs),
    )
    connection = ModelConnection(
        base_url="https://api.example.test/v1",
        model="test-model",
        api_key="",
        supports_vision=False,
    )

    async def run() -> None:
        async with OpenAICompatibleClient(
            connection,
            timeout=1,
        ) as client:
            await client.chat_completions(messages=[])

    with pytest.raises(LLMTimeoutError):
        asyncio.run(run())
