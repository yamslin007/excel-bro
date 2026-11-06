"""Wire-level protocol tests: one connectivity check and shared formula caller."""
import asyncio
import json

import httpx
import pytest
from fastapi.testclient import TestClient

from server.app import main, rule_generator
from server.app.llm import config
from server.app.llm.client import OpenAICompatibleClient
from server.app.llm.connection_test import test_model_connection as probe
from server.app.llm.errors import LLMResponseError, LLMTimeoutError, LLMUsageLimitError
from server.app.rule_generator import GenerateFormulaRequest


def reply(protocol, text="OK", **changes):
    if protocol == "anthropic":
        return {"type": "message", "role": "assistant", "content": [{"type": "text", "text": text}],
                "stop_reason": "end_turn", "usage": {"input_tokens": 5, "output_tokens": 2}, **changes}
    return {"choices": [{"message": {"content": text}, "finish_reason": "stop"}], **changes}


@pytest.fixture
def wire(monkeypatch, tmp_path):
    original = httpx.AsyncClient
    captured = []
    state = {"body": None, "status": 200}
    def handler(request):
        captured.append(request)
        return httpx.Response(state["status"], json=state["body"])
    monkeypatch.setattr(httpx, "AsyncClient", lambda **kwargs: original(transport=httpx.MockTransport(handler), **kwargs))
    monkeypatch.setattr(config, "_config_path", lambda: tmp_path / ".env")
    monkeypatch.setattr(config, "load_model_settings", lambda: None)
    monkeypatch.setattr(config, "_claude_sdk_connection", lambda: None)
    monkeypatch.setattr(config, "_claude_sdk_state_ready", lambda: False)
    return captured, state


@pytest.mark.parametrize("protocol,base,path,auth", [
    ("openai_compatible", "https://provider.test/v1", "/v1/chat/completions", "authorization"),
    ("openai_compatible", "https://provider.test/custom/chat/completions", "/custom/chat/completions", "authorization"),
    ("anthropic", "https://provider.test", "/v1/messages", "x-api-key"),
    ("anthropic", "https://provider.test/v1", "/v1/messages", "x-api-key"),
    ("anthropic", "https://provider.test/anthropic", "/anthropic/v1/messages", "x-api-key"),
    ("anthropic", "https://provider.test/custom/messages", "/custom/messages", "x-api-key"),
])
def test_unified_http_probe_works_without_sdk_or_tools_and_never_saves(wire, protocol, base, path, auth):
    captured, state = wire
    state["body"] = reply(protocol)
    response = TestClient(main.app).post("/api/settings/model/connections/test", json={
        "label": "Draft", "baseUrl": base, "modelId": "chosen-model", "apiKey": "private-key", "protocol": protocol,
    })
    assert response.status_code == 200, response.text
    assert response.json() == {"ok": True, "testedFor": "connection", "message": "连接测试通过，模型已正常回复。"}
    assert len(captured) == 1 and captured[0].url.path == path
    assert "private-key" in captured[0].headers[auth]
    payload = json.loads(captured[0].content)
    assert payload["model"] == "chosen-model"
    content = [{"type": "text", "text": "Reply with OK."}] if protocol == "anthropic" else "Reply with OK."
    assert payload["messages"] == [{"role": "user", "content": content}]
    assert "tools" not in payload and "tool_choice" not in payload
    assert not config._connection_store_path().exists()


@pytest.mark.parametrize("protocol", ["openai_compatible", "anthropic"])
@pytest.mark.parametrize("body_kind", ["empty", "reasoning_only", "error", "truncated"])
def test_http_200_is_not_enough_without_a_valid_text_reply(wire, protocol, body_kind):
    captured, state = wire
    if body_kind == "error":
        state["body"] = {"error": {"message": "private-provider-error"}}
    elif protocol == "anthropic":
        state["body"] = reply(protocol, "" if body_kind == "empty" else "OK",
            **({"content": [{"type": "thinking", "thinking": "internal"}]} if body_kind == "reasoning_only" else
               {"stop_reason": "max_tokens"} if body_kind == "truncated" else {}))
    else:
        state["body"] = {"choices": [{"message": {"content": "" if body_kind != "truncated" else "OK", "reasoning_content": "internal"},
                                      "finish_reason": "length" if body_kind == "truncated" else "stop"}]}
    response = TestClient(main.app).post("/api/settings/model/connections/test", json={
        "label": "Draft", "baseUrl": "https://provider.test/v1", "modelId": "m", "protocol": protocol,
    })
    assert response.status_code != 200
    assert "private-provider-error" not in response.text
    assert len(captured) == 1


@pytest.mark.parametrize("protocol", ["openai_compatible", "anthropic"])
def test_formula_generation_uses_saved_protocol_and_existing_safety_pipeline(wire, monkeypatch, protocol):
    captured, state = wire
    state["body"] = reply(protocol, json.dumps({"modernFormula": "=SUM(A1:A2)", "modernExplanation": "求和",
                                               "compatFormula": "=SUM(A1:A2)", "compatExplanation": "求和"}))
    connection = config.ModelConnection("https://provider.test/v1", "custom-model", "formula-key", False,
        "claude_agent_sdk" if protocol == "anthropic" else "openai_compatible")
    monkeypatch.setattr(rule_generator, "formula_model_config", lambda: connection)
    # Exercise the standard formula text path, avoiding the optional compiler path.
    import server.app.formula_tools_integration as integration
    async def skip_compiler(**kwargs):
        raise ValueError("exercise text generation")
    monkeypatch.setattr(integration, "generate_formula_with_tool_call", skip_compiler)
    request = GenerateFormulaRequest(activeCell="B1", description="合计 A1:A2", modelId="connection:chosen")
    result = asyncio.run(rule_generator.generate_formula(request))
    assert result.modernFormula == "=SUM(A1:A2)" and result.compatFormula == "=SUM(A1:A2)"
    body = json.loads(captured[0].content)
    assert body["model"] == "custom-model"
    assert captured[0].url.path == ("/v1/messages" if protocol == "anthropic" else "/v1/chat/completions")
    if protocol == "anthropic":
        assert body["system"] and all(m["role"] != "system" for m in body["messages"])


def test_anthropic_text_blocks_usage_and_tool_receipt_are_normalized(wire):
    captured, state = wire
    state["body"] = reply("anthropic", content=[{"type": "text", "text": "A"}, {"type": "text", "text": "B"}],
        usage={"input_tokens": 5, "output_tokens": 2, "cache_read_input_tokens": 3})
    async def run():
        async with OpenAICompatibleClient(config.ModelConnection("https://provider.test", "m", "k", False, "claude_agent_sdk"), timeout=5) as client:
            return await client.chat_completions(messages=[{"role": "system", "content": "system"},
                {"role": "assistant", "content": None, "tool_calls": [{"id": "t1", "function": {"name": "query_table", "arguments": "{}"}}]},
                {"role": "tool", "tool_call_id": "t1", "content": "safe receipt"}])
    result, usage = asyncio.run(run())
    assert result["choices"][0]["message"]["content"] == "AB"
    assert usage == {"prompt_tokens": 8, "completion_tokens": 2, "total_tokens": 10}
    assert json.loads(captured[0].content)["messages"][-1]["content"][0]["tool_use_id"] == "t1"
    assert captured[0].headers["anthropic-version"] == "2023-06-01"
    assert "authorization" not in captured[0].headers


def test_probe_total_timeout_has_a_handled_error(monkeypatch):
    import server.app.llm.connection_test as module
    monkeypatch.setattr(module, "capability_float", lambda *args: 0.01)
    async def stalled(*args, **kwargs):
        await asyncio.sleep(10)
    monkeypatch.setattr(OpenAICompatibleClient, "chat_completions", stalled)
    with pytest.raises(LLMTimeoutError):
        asyncio.run(probe(config.ModelConnection("http://localhost:1/v1", "m", "", False)))


def test_text_probe_adapts_explicit_openai_output_parameter_rejection(monkeypatch):
    calls = []
    original = httpx.AsyncClient
    def handler(request):
        body = json.loads(request.content)
        calls.append(body)
        if "max_tokens" in body:
            return httpx.Response(400, json={"error": {"code": "unsupported_parameter", "param": "max_tokens"}})
        return httpx.Response(200, json=reply("openai_compatible"))
    monkeypatch.setattr(httpx, "AsyncClient", lambda **kwargs: original(transport=httpx.MockTransport(handler), **kwargs))
    asyncio.run(probe(config.ModelConnection("https://provider.test/v1", "reasoning-model", "", False)))
    assert len(calls) == 2
    assert calls[1]["max_completion_tokens"] == calls[0]["max_tokens"]
    assert "max_tokens" not in calls[1]
    assert calls[1]["model"] == calls[0]["model"]
