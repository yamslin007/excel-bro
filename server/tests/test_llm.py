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
            json={
                "choices": [{"message": {"content": "ok"}}],
                "usage": {
                    "prompt_tokens": 7,
                    "completion_tokens": 2,
                    "total_tokens": 9,
                },
            },
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

    async def run() -> tuple[dict[str, object], dict[str, int]]:
        async with OpenAICompatibleClient(
            connection,
            timeout=12,
        ) as client:
            return await client.chat_completions(
                messages=[{"role": "user", "content": "hello"}],
                max_tokens=100,
            )

    response, usage = asyncio.run(run())

    assert response["choices"][0]["message"]["content"] == "ok"
    assert usage == {
        "prompt_tokens": 7,
        "completion_tokens": 2,
        "total_tokens": 9,
    }
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
            json={
                "choices": [{"message": {"content": "ok"}}],
                "usage": {
                    "prompt_tokens": 7,
                    "completion_tokens": 2,
                    "total_tokens": 9,
                },
            },
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


def test_provider_default_output_budget_is_absent_from_real_request(monkeypatch):
    captured = []
    async def handler(request):
        captured.append(json.loads(request.content))
        return httpx.Response(200, json={"choices": [{"message": {"content": "ok"}}],
            "usage": {"completion_tokens": 9000, "completion_tokens_details": {"reasoning_tokens": 8500}}})
    original_client = httpx.AsyncClient
    monkeypatch.setattr(httpx, "AsyncClient", lambda **kwargs: original_client(transport=httpx.MockTransport(handler), **kwargs))
    connection = ModelConnection(base_url="https://api.example.test/v1", model="generic", api_key="", supports_vision=False)
    async def run():
        async with OpenAICompatibleClient(connection, timeout=12) as client:
            return await client.chat_completions(messages=[{"role": "user", "content": "hello"}], max_tokens=None)
    _, usage = asyncio.run(run())
    assert captured[0]["messages"]  # Not an empty-body privacy/budget assertion.
    assert "max_tokens" not in captured[0]
    assert "max_completion_tokens" not in captured[0]
    assert usage["reasoning_tokens"] == 8500


@pytest.mark.parametrize("value", [0, -1, True, "4000", 1.5])
def test_optional_output_budget_rejects_invalid_configuration(monkeypatch, value):
    import server.app.capabilities as limits
    monkeypatch.setattr(limits, "capabilities", lambda: {"agent": {"maxOutputTokens": value}})
    with pytest.raises(RuntimeError):
        limits.capability_optional_int("agent", "maxOutputTokens")


@pytest.mark.parametrize("value", [None, 8192])
def test_optional_output_budget_accepts_default_or_positive_integer(monkeypatch, value):
    import server.app.capabilities as limits
    monkeypatch.setattr(limits, "capabilities", lambda: {"agent": {"maxOutputTokens": value}})
    assert limits.capability_optional_int("agent", "maxOutputTokens") == value


def test_intent_and_planning_real_requests_omit_output_caps(monkeypatch):
    from server.app.intent import check_intent
    from server.app.excel_agent import run_excel_agent
    from server.app.models import IntentCheckRequest, PlanRequest
    captured = []
    async def handler(request):
        body = json.loads(request.content)
        captured.append(body)
        message = ({"content": json.dumps({"kind": "proceed", "summary": "清空授权范围", "confirmedPrompt": "清空 B2"})}
                   if "tools" not in body else
                   {"content": None, "tool_calls": [{"id": "answer", "type": "function", "function": {
                       "name": "submit_answer", "arguments": json.dumps({"message": "需要用户确认，未写入"})}}]})
        return httpx.Response(200, json={"choices": [{"message": message, "finish_reason": "stop"}]})
    original_client = httpx.AsyncClient
    monkeypatch.setattr(httpx, "AsyncClient", lambda **kwargs: original_client(transport=httpx.MockTransport(handler), **kwargs))
    connection = ModelConnection(base_url="https://api.example.test/v1", model="generic", api_key="", supports_vision=False)
    intent = IntentCheckRequest.model_validate({"prompt": "清空 B2", "scope": {
        "workbookName": "test.xlsx", "sourceMode": "workbook", "selectionMode": "auto",
        "activeWorksheet": "Sheet1", "totalWorksheetCount": 1, "worksheetNames": ["Sheet1"],
        "sheets": [{"name": "Sheet1", "rowCount": 2, "columnCount": 2, "headers": ["类别", "结果"]}]}})
    plan = PlanRequest.model_validate({"prompt": "清空 B2", "planningReadPolicy": "none", "workbook": {
        "name": "test.xlsx", "capturedAt": "2026-10-03T00:00:00Z", "activeWorksheet": "Sheet1",
        "worksheets": [{"name": "Sheet1", "usedRange": "A1:B2", "rowCount": 2, "columnCount": 2,
            "headers": ["类别", "结果"], "dataRows": [["LOCAL_PRIVATE_SENTINEL", 3]], "truncated": False}]}})
    assert "LOCAL_PRIVATE_SENTINEL" in str(plan.workbook.worksheets[0].dataRows)
    async def run():
        assert (await check_intent(intent, config=connection)).kind == "proceed"
        assert (await run_excel_agent(plan, connection=connection, timeout=5)).kind == "answer"
    asyncio.run(run())
    assert len(captured) == 2
    assert all(body["messages"] for body in captured)
    assert "tools" not in captured[0] and captured[1]["tools"]
    assert all("max_tokens" not in body for body in captured)
    assert "LOCAL_PRIVATE_SENTINEL" not in json.dumps(captured)


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

    async def run() -> tuple[dict[str, object], dict[str, int]]:
        async with OpenAICompatibleClient(
            _retry_connection(),
            timeout=12,
            max_retries=2,
            retry_base_delay=1.0,
            retry_max_delay=20.0,
        ) as client:
            return await client.chat_completions(messages=[])

    response, usage = asyncio.run(run())

    assert response["choices"][0]["message"]["content"] == "ok"
    assert usage == {
        "prompt_tokens": 0,
        "completion_tokens": 0,
        "total_tokens": 0,
    }
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


def test_actionable_statuses_get_a_chinese_hint_instead_of_raw_body() -> None:
    """402 之类只回英文 JSON 的状态码要给用户可行动的提示。"""
    insufficient_balance = LLMHTTPStatusError(
        402,
        '{"error":{"message":"Insufficient Balance","code":"invalid_request_error"}}',
    )

    assert "余额不足" in str(insufficient_balance)
    assert "切换" in str(insufficient_balance)
    assert "HTTP 402" in str(insufficient_balance)
    # 原始 body 仍保留在属性上供诊断，只是不再直接糊到用户面前。
    assert "Insufficient Balance" in insufficient_balance.body
    assert "Insufficient Balance" not in str(insufficient_balance)

    assert "API Key" in str(LLMHTTPStatusError(401, "unauthorized"))
    assert "baseUrl" in str(LLMHTTPStatusError(404, "not found"))


def test_unmapped_status_still_passes_through_the_upstream_body() -> None:
    """没映射的状态码不能吞掉原始信息，否则排查无从下手。"""
    error = LLMHTTPStatusError(418, "i am a teapot")

    assert "模型服务返回 HTTP 418" in str(error)
    assert "i am a teapot" in str(error)


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

    async def run() -> tuple[dict[str, object], dict[str, int]]:
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
