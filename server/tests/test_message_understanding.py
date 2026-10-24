from __future__ import annotations

import asyncio
import json
from dataclasses import replace

import pytest
from fastapi.testclient import TestClient
from pydantic import ValidationError

from server.app import intent, main, claude_sdk_agent
from server.app.llm import LLMResponseError, ModelNotConfiguredError
from server.app.models import MessageUnderstandingRequest, MessageConversationResponse
from server.tests.sdk_text_harness import install_text_transport
from server.tests.test_claude_sdk_agent import _connection


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
    install_text_transport(monkeypatch, client)
    result = asyncio.run(intent.check_message_understanding(
        request(conversation=[{"role": "user", "text": "检查当前表"}]), config=_connection()
    ))
    assert result.kind == payload["kind"]
    outbound = json.loads(client.calls[0]["messages"][1]["content"])
    assert set(outbound) == {"prompt", "imageCount", "conversation"}
    assert outbound["conversation"][0]["text"] == "检查当前表"


def test_images_are_not_claimed_seen(monkeypatch):
    client = Client([response({"kind": "conversation", "message": "我看到了图片"})])
    install_text_transport(monkeypatch, client)
    result = asyncio.run(intent.check_message_understanding(request(imageCount=1), config=replace(_connection(), supports_vision=True)))
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
    install_text_transport(monkeypatch, client)
    with pytest.raises(LLMResponseError, match="尚未读取"):
        asyncio.run(intent.check_message_understanding(request(), config=_connection()))
    assert len(client.calls) == 2


def test_truncated_route_is_not_trusted(monkeypatch):
    client = Client([response({"kind": "needs_workbook_context", "summary": "读取字段"}, "length")])
    install_text_transport(monkeypatch, client)
    with pytest.raises(LLMResponseError):
        asyncio.run(intent.check_message_understanding(request(), config=_connection()))
    assert len(client.calls) == 1


def test_no_model_is_not_implicit_permission():
    with pytest.raises(ModelNotConfiguredError):
        asyncio.run(intent.check_message_understanding(request(), config=None))


def test_model_routes_continue_to_offered_task_without_workbook_payload(monkeypatch):
    client = Client([response({"kind": "task_recovery", "turnId": "paused-original"})])
    install_text_transport(monkeypatch, client)
    result = asyncio.run(intent.check_message_understanding(
        request(prompt="接着刚才的分析做", recoveryTasks=["paused-original"]), config=_connection()
    ))
    assert result.kind == "task_recovery"
    assert result.turnId == "paused-original"
    outbound = json.loads(client.calls[0]["messages"][1]["content"])
    assert outbound["recoveryTasks"] == ["paused-original"]
    assert set(outbound) == {"prompt", "conversation", "imageCount", "recoveryTasks"}


@pytest.mark.parametrize("tasks,images", [([], 0), (["other-task"], 0), (["paused-original"], 1)])
def test_recovery_cannot_invent_identity_or_ignore_attachment(tasks, images):
    with pytest.raises(ValueError):
        intent.parse_message_understanding(
            '{"kind":"task_recovery","turnId":"paused-original"}',
            request(prompt="继续", recoveryTasks=tasks, imageCount=images),
        )


def test_ambiguous_recovery_can_ask_without_selecting_a_task():
    result = intent.parse_message_understanding(
        '{"kind":"conversation","message":"要继续哪一个暂停任务？"}',
        request(prompt="继续", recoveryTasks=["first", "second"]),
    )
    assert result.kind == "conversation"


@pytest.mark.parametrize("action", ["inspect", "continue"])
def test_recovery_preserves_explicit_action(action):
    result = intent.parse_message_understanding(
        json.dumps({"kind": "task_recovery", "turnId": "paused", "action": action}),
        request(recoveryTasks=["paused"]),
    )
    assert result.action == action


def test_legacy_recovery_without_action_is_read_only():
    result = intent.parse_message_understanding(
        '{"kind":"task_recovery","turnId":"paused"}', request(recoveryTasks=["paused"])
    )
    assert result.action == "inspect"


def test_recovery_rejects_unknown_action():
    with pytest.raises(ValueError):
        intent.parse_message_understanding(
            '{"kind":"task_recovery","turnId":"paused","action":"execute"}',
            request(recoveryTasks=["paused"]),
        )


def test_sdk_uses_same_message_only_payload(monkeypatch):
    captured = []
    sdk = object()
    monkeypatch.setattr(claude_sdk_agent, "_sdk", lambda: sdk)
    monkeypatch.setattr(claude_sdk_agent, "_options", lambda *_a, **_k: "restricted-options")

    async def result(_sdk, prompt, options, timeout):
        captured.append((json.loads(prompt), options))
        return json.dumps({"kind": "conversation", "message": "不客气！"})

    monkeypatch.setattr(claude_sdk_agent, "_result", result)
    result = asyncio.run(intent.check_message_understanding(request(), config=_connection()))
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
        asyncio.run(intent.check_message_understanding(request(), config=_connection()))


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
