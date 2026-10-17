from __future__ import annotations

import asyncio
import json

import pytest
from pydantic import ValidationError

from server.app import claude_sdk_agent
from server.app.models import IntentCheckRequest, PlanRequest
from server.app.task_context import task_context_payload
from server.app import excel_agent, intent as intent_module
from server.tests.test_agent_integration import _FakeClient as _PlanningClient, _connection as _ordinary_connection, _tool_payload, _usage
from server.tests.test_intent_read_only_routing import _FakeClient as _IntentClient
from server.tests.test_claude_sdk_agent import _FakeSdk, _connection, _install_fake, _intent_request, _plan_request


CONTEXT = {"schemaVersion": 1, "originalPrompt": "根据类别填写等级", "confirmations": [
    {"questionId": "rank", "label": "多项优先级", "value": "已明确按高到低取一个"}
]}


def test_ordinary_intent_and_real_planning_loop_receive_the_same_context(monkeypatch):
    client = _IntentClient([{"kind": "proceed", "summary": "填写", "confirmedPrompt": "根据类别填写等级"}])
    monkeypatch.setattr(intent_module, "OpenAICompatibleClient", lambda *_args, **_kwargs: client)
    request = IntentCheckRequest.model_validate({**_intent_request().model_dump(), "taskContext": CONTEXT})
    result = asyncio.run(intent_module.check_intent(request, config=_ordinary_connection()))
    assert result.kind == "proceed"
    serialized = json.dumps(CONTEXT, ensure_ascii=False)
    assert serialized in client.calls[0][1]["content"]
    payload = _tool_payload("submit_answer", {"message": "完整规则已收到"}, "answer-1")
    planner = _PlanningClient(lambda _turn: (payload, _usage()))
    monkeypatch.setattr(excel_agent, "OpenAICompatibleClient", lambda *_args, **_kwargs: planner)
    plan = PlanRequest.model_validate({**_plan_request().model_dump(), "taskContext": CONTEXT})
    result = asyncio.run(excel_agent.run_excel_agent(plan, connection=_ordinary_connection(), timeout=5))
    assert result.kind == "answer"
    sent = json.dumps(planner.message_snapshots, ensure_ascii=False)
    assert serialized in planner.message_snapshots[0][1]["content"]
    assert "PRIVATE_DATA_ROW" not in sent


def test_context_is_bounded_and_cannot_carry_rows_or_authorization():
    for request in (_intent_request(), _plan_request()):
        data = request.model_dump()
        data["taskContext"] = {**CONTEXT, "authorized": True, "rows": [["PRIVATE_ROW"]]}
        with pytest.raises(ValidationError):
            type(request).model_validate(data)
    data = _intent_request().model_dump()
    data["taskContext"] = {**CONTEXT, "confirmations": [{"label": "规则", "value": "甲" * 1001}]}
    with pytest.raises(ValidationError):
        IntentCheckRequest.model_validate(data)


def test_same_question_keeps_only_latest_answer():
    data = _intent_request().model_dump()
    data["taskContext"] = {**CONTEXT, "confirmations": [
        {"questionId": "mode", "label": "口径", "value": "汇总"},
        {"questionId": "mode", "label": "口径", "value": "逐行"},
    ]}
    request = IntentCheckRequest.model_validate(data)
    assert task_context_payload(request)["confirmations"] == [{"questionId": "mode", "label": "口径", "value": "逐行"}]


def test_sdk_intent_and_planner_receive_same_confirmations_without_history(monkeypatch, tmp_path):
    fake = _FakeSdk(response=json.dumps({"kind": "proceed", "summary": "填写", "confirmedPrompt": "根据类别填写等级"}, ensure_ascii=False))
    _install_fake(monkeypatch, tmp_path, fake)
    intent = IntentCheckRequest.model_validate({**_intent_request().model_dump(), "taskContext": CONTEXT})
    asyncio.run(claude_sdk_agent.check_sdk_intent(intent, _connection()))
    sent = json.loads(fake.calls[0][0])
    assert sent["taskContext"] == CONTEXT
    assert "PRIVATE_CONVERSATION_ROW" not in fake.calls[0][0]
    assert "PRIVATE_RESULT_ROW" not in fake.calls[0][0]
    fake.response = json.dumps({"kind": "answer", "message": "规则已保留"}, ensure_ascii=False)
    plan = PlanRequest.model_validate({**_plan_request().model_dump(), "taskContext": CONTEXT})
    asyncio.run(claude_sdk_agent.run_sdk_plan(plan, _connection()))
    assert json.loads(fake.calls[-1][0])["taskContext"] == CONTEXT
    assert "PRIVATE_DATA_ROW" not in fake.calls[-1][0]
