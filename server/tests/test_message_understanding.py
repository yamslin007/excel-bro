from __future__ import annotations

import asyncio
import json
from types import SimpleNamespace

import pytest
from fastapi.testclient import TestClient
from pydantic import ValidationError

from server.app import intent, main, claude_sdk_agent
from server.app.llm import LLMResponseError, ModelNotConfiguredError
from server.app.models import MessageUnderstandingRequest, MessageConversationResponse


def request(**updates):
    return MessageUnderstandingRequest.model_validate({"prompt": "谢谢帮助", **updates})


class Client:
    def __init__(self, responses):
        self.responses = responses
        self.calls = []

    async def __aenter__(self):
        return self

    async def __aexit__(self, *_args):
        return False

    async def chat_completions(self, **kwargs):
        self.calls.append(kwargs)
        return {"choices": [self.responses[len(self.calls) - 1]]}, {}


def response(payload, finish_reason="stop"):
    return {"message": {"content": json.dumps(payload)}, "finish_reason": finish_reason}


@pytest.mark.parametrize("field", ["scope", "workbook", "priorResult", "priorIntent", "dataResults", "images"])
def test_unscoped_request_rejects_workbook_payloads(field):
    with pytest.raises(ValidationError):
        request(**{field: {"private": "RAW_DATA_SENTINEL"}})


@pytest.mark.parametrize("payload", [
    {"kind": "conversation", "message": "不客气！"},
    {"kind": "needs_workbook_context", "summary": "检查当前表"},
])
def test_message_only_model_route(monkeypatch, payload):
    client = Client([response(payload)])
    monkeypatch.setattr(intent, "OpenAICompatibleClient", lambda *_a, **_k: client)
    result = asyncio.run(intent.check_message_understanding(
        request(conversation=[{"role": "user", "text": "检查当前表"}]), config=SimpleNamespace()
    ))
    assert result.kind == payload["kind"]
    outbound = json.loads(client.calls[0]["messages"][1]["content"])
    assert set(outbound) == {"prompt", "imageCount", "conversation"}
    assert outbound["conversation"][0]["text"] == "检查当前表"


def test_images_are_not_claimed_seen(monkeypatch):
    client = Client([response({"kind": "conversation", "message": "我看到了图片"})])
    monkeypatch.setattr(intent, "OpenAICompatibleClient", lambda *_a, **_k: client)
    result = asyncio.run(intent.check_message_understanding(request(imageCount=1), config=SimpleNamespace()))
    assert result.kind == "needs_workbook_context"


@pytest.mark.parametrize("payload", [
    {"kind": "proceed", "confirmedPrompt": "读取整表"},
    {"kind": "conversation", "message": "你好", "plan": {}},
    {"kind": "needs_workbook_context", "summary": ""},
    {"summary": "读取字段"},
    {"kind": "conversation", "message": "   "},
])
def test_invalid_routing_does_not_default_to_workbook_access(monkeypatch, payload):
    client = Client([response(payload), response(payload)])
    monkeypatch.setattr(intent, "OpenAICompatibleClient", lambda *_a, **_k: client)
    with pytest.raises(LLMResponseError, match="尚未读取"):
        asyncio.run(intent.check_message_understanding(request(), config=SimpleNamespace()))
    assert len(client.calls) == 2


def test_truncated_route_is_not_trusted(monkeypatch):
    client = Client([response({"kind": "needs_workbook_context", "summary": "读取字段"}, "length")])
    monkeypatch.setattr(intent, "OpenAICompatibleClient", lambda *_a, **_k: client)
    with pytest.raises(LLMResponseError):
        asyncio.run(intent.check_message_understanding(request(), config=SimpleNamespace()))
    assert len(client.calls) == 1


def test_no_model_is_not_implicit_permission():
    with pytest.raises(ModelNotConfiguredError):
        asyncio.run(intent.check_message_understanding(request(), config=None))


def test_sdk_uses_same_message_only_payload(monkeypatch):
    captured = []
    sdk = object()
    monkeypatch.setattr(claude_sdk_agent, "_sdk", lambda: sdk)
    monkeypatch.setattr(claude_sdk_agent, "_options", lambda *_a, **_k: "restricted-options")

    async def result(_sdk, prompt, options, timeout):
        captured.append((json.loads(prompt), options))
        return json.dumps({"kind": "conversation", "message": "不客气！"})

    monkeypatch.setattr(claude_sdk_agent, "_result", result)
    result = asyncio.run(intent.check_message_understanding(request(), config=SimpleNamespace(adapter="claude_agent_sdk")))
    assert result.kind == "conversation"
    assert set(captured[0][0]) == {"prompt", "imageCount", "conversation"}
    assert captured[0][1] == "restricted-options"


def test_sdk_invalid_route_is_not_permission(monkeypatch):
    monkeypatch.setattr(claude_sdk_agent, "_sdk", lambda: object())
    monkeypatch.setattr(claude_sdk_agent, "_options", lambda *_a, **_k: "restricted-options")

    async def invalid(*_args):
        return '{"kind":"proceed","confirmedPrompt":"读字段"}'

    monkeypatch.setattr(claude_sdk_agent, "_result", invalid)
    with pytest.raises(LLMResponseError, match="尚未读取"):
        asyncio.run(intent.check_message_understanding(request(), config=SimpleNamespace(adapter="claude_agent_sdk")))


def test_endpoint_has_no_scope_or_turn_registry_requirement(monkeypatch):
    received = []
    monkeypatch.setattr(main, "selected_model_config", lambda _id: None)

    async def understand(value, *, config):
        received.append(value.model_dump())
        return MessageConversationResponse(message="不客气！")

    monkeypatch.setattr(main, "check_message_understanding", understand)
    client = TestClient(main.app)
    assert client.post("/api/intent/message", json={"prompt": "谢谢帮助"}).json()["kind"] == "conversation"
    assert "scope" not in received[0]
    assert client.post("/api/intent/message", json={"prompt": "谢谢帮助", "scope": {}}).status_code == 422
