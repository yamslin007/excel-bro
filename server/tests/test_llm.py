from __future__ import annotations

import asyncio
import json
from pathlib import Path

import httpx
import pytest

import server.app.llm.config as llm_config
from server.app.llm import (
    LLMHTTPStatusError,
    LLMTimeoutError,
    ModelConnection,
    OpenAICompatibleClient,
    selected_model_config,
)


@pytest.fixture(autouse=True)
def clear_model_environment(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
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
    monkeypatch.setattr(
        llm_config,
        "_config_path",
        lambda: tmp_path / "config.env",
    )


def test_model_timeout_prefers_env_override_then_config(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from server.app.capabilities import capability_float, model_timeout_seconds

    monkeypatch.delenv("AI_TIMEOUT_SECONDS", raising=False)
    assert model_timeout_seconds() == capability_float("llm", "timeoutSeconds")

    monkeypatch.setenv("AI_TIMEOUT_SECONDS", "35")
    assert model_timeout_seconds() == 35.0

    # 非法或非正的环境值不得中断请求，回退到配置默认。
    monkeypatch.setenv("AI_TIMEOUT_SECONDS", "not-a-number")
    assert model_timeout_seconds() == capability_float("llm", "timeoutSeconds")
    monkeypatch.setenv("AI_TIMEOUT_SECONDS", "0")
    assert model_timeout_seconds() == capability_float("llm", "timeoutSeconds")


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


def test_openai_compatible_client_sends_temperature_when_provided(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """可选 temperature 参数应进入请求体；不传时保持向后兼容（不出现该字段）。"""
    captured: dict[str, object] = {}

    async def handler(request: httpx.Request) -> httpx.Response:
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
        api_key="",
        supports_vision=False,
    )

    async def run() -> None:
        async with OpenAICompatibleClient(connection, timeout=12) as client:
            await client.chat_completions(
                messages=[{"role": "user", "content": "hello"}],
                temperature=0.1,
            )

    asyncio.run(run())
    assert captured["body"]["temperature"] == 0.1


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
            max_retries=0,
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
            max_retries=0,
        ) as client:
            await client.chat_completions(messages=[])

    with pytest.raises(LLMTimeoutError):
        asyncio.run(run())


def _retry_connection() -> ModelConnection:
    return ModelConnection(
        base_url="https://api.example.test/v1",
        model="test-model",
        api_key="",
        supports_vision=False,
    )


def _mock_transport(monkeypatch: pytest.MonkeyPatch, handler) -> None:
    transport = httpx.MockTransport(handler)
    original_client = httpx.AsyncClient
    monkeypatch.setattr(
        httpx,
        "AsyncClient",
        lambda **kwargs: original_client(transport=transport, **kwargs),
    )


def test_retries_overloaded_then_succeeds(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls = {"count": 0}
    sleeps: list[float] = []

    def handler(_request: httpx.Request) -> httpx.Response:
        calls["count"] += 1
        if calls["count"] <= 2:
            return httpx.Response(
                429,
                json={"error": {"message": "overloaded"}},
            )
        return httpx.Response(
            200,
            json={"choices": [{"message": {"content": "ok"}}]},
        )

    _mock_transport(monkeypatch, handler)

    async def fake_sleep(delay: float) -> None:
        sleeps.append(delay)

    monkeypatch.setattr("server.app.llm.client.asyncio.sleep", fake_sleep)

    async def run() -> dict[str, object]:
        async with OpenAICompatibleClient(
            _retry_connection(),
            timeout=12,
            max_retries=2,
            retry_base_delay=1.0,
            retry_max_delay=20.0,
        ) as client:
            return await client.chat_completions(messages=[])

    response = asyncio.run(run())

    assert response["choices"][0]["message"]["content"] == "ok"
    assert calls["count"] == 3
    assert len(sleeps) == 2


def test_retries_exhausted_raises_last_error(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls = {"count": 0}

    def handler(_request: httpx.Request) -> httpx.Response:
        calls["count"] += 1
        return httpx.Response(429, text="still overloaded")

    _mock_transport(monkeypatch, handler)

    async def fake_sleep(_delay: float) -> None:
        return None

    monkeypatch.setattr("server.app.llm.client.asyncio.sleep", fake_sleep)

    async def run() -> None:
        async with OpenAICompatibleClient(
            _retry_connection(),
            timeout=12,
            max_retries=2,
            retry_base_delay=0.01,
            retry_max_delay=1.0,
        ) as client:
            await client.chat_completions(messages=[])

    with pytest.raises(LLMHTTPStatusError) as captured:
        asyncio.run(run())

    assert captured.value.status_code == 429
    assert calls["count"] == 3  # 首次 + 2 次重试


def test_non_retryable_status_is_not_retried(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls = {"count": 0}

    def handler(_request: httpx.Request) -> httpx.Response:
        calls["count"] += 1
        return httpx.Response(401, text="unauthorized")

    _mock_transport(monkeypatch, handler)

    async def run() -> None:
        async with OpenAICompatibleClient(
            _retry_connection(),
            timeout=12,
            max_retries=3,
        ) as client:
            await client.chat_completions(messages=[])

    with pytest.raises(LLMHTTPStatusError) as captured:
        asyncio.run(run())

    assert captured.value.status_code == 401
    assert calls["count"] == 1  # 4xx 非 429 不重试


def test_retry_after_header_controls_delay(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls = {"count": 0}
    sleeps: list[float] = []

    def handler(_request: httpx.Request) -> httpx.Response:
        calls["count"] += 1
        if calls["count"] == 1:
            return httpx.Response(
                429,
                text="slow down",
                headers={"Retry-After": "7"},
            )
        return httpx.Response(
            200,
            json={"choices": [{"message": {"content": "ok"}}]},
        )

    _mock_transport(monkeypatch, handler)

    async def fake_sleep(delay: float) -> None:
        sleeps.append(delay)

    monkeypatch.setattr("server.app.llm.client.asyncio.sleep", fake_sleep)

    async def run() -> dict[str, object]:
        async with OpenAICompatibleClient(
            _retry_connection(),
            timeout=12,
            max_retries=2,
            retry_base_delay=1.0,
            retry_max_delay=20.0,
        ) as client:
            return await client.chat_completions(messages=[])

    asyncio.run(run())

    assert sleeps == [7.0]
