from __future__ import annotations

import asyncio
import json
from types import SimpleNamespace

import pytest

from server.app import claude_sdk_agent, intent
from server.app.assistant_identity import PRODUCT_IDENTITY_PROMPT, with_runtime_model_identity
from server.app.excel_agent import AGENT_SYSTEM_PROMPT
from server.app.models import MessageUnderstandingRequest


@pytest.mark.parametrize("prompt", [intent.MESSAGE_UNDERSTANDING_PROMPT, intent.INTENT_SYSTEM_PROMPT, AGENT_SYSTEM_PROMPT])
def test_all_stages_share_product_identity_separate_from_internal_task(prompt):
    assert prompt.startswith(PRODUCT_IDENTITY_PROMPT)
    assert "不要自称消息理解器" in prompt
    assert "Claude Agent SDK 是调用适配方式" in prompt
    assert "经用户另行确认才执行" not in prompt
    assert "写入需用户另行确认" not in prompt
    assert "不得要求执行前另行确认" in prompt
    assert "不重复执行" in prompt


@pytest.mark.parametrize("model", [None, True, "", "  ", "unsafe\nmodel", "x" * 201])
def test_missing_or_malformed_model_identity_is_not_guessed(model):
    projected = with_runtime_model_identity("product prompt", SimpleNamespace(model=model))
    assert json.loads(projected.split("：", 1)[1]) == {"configuredModel": None}


def test_runtime_identity_projection_never_serializes_connection_secrets():
    connection = SimpleNamespace(model="configured-gateway-model", api_key="PRIVATE_KEY_SENTINEL",
                                 base_url="https://PRIVATE_HOST_SENTINEL", adapter="claude_agent_sdk")
    projected = with_runtime_model_identity(PRODUCT_IDENTITY_PROMPT, connection)
    assert '"configuredModel": "configured-gateway-model"' in projected
    assert "PRIVATE_KEY_SENTINEL" not in projected
    assert "PRIVATE_HOST_SENTINEL" not in projected
    assert "不证明网关背后的实际模型" in projected


@pytest.mark.parametrize("adapter", ["openai_compatible", "claude_agent_sdk"])
@pytest.mark.parametrize("question", ["你是谁", "你用的是什么模型", "你是 Claude 吗"])
def test_identity_questions_still_use_model_and_both_backends_receive_same_identity(monkeypatch, adapter, question):
    from server.tests.test_message_understanding import Client, response
    captured = []
    reply = {"kind": "conversation", "message": "我是 Excel Bro，你的 AI 表格助手。"}
    from server.app.llm import ModelConnection
    connection = ModelConnection(model="configured-model", adapter=adapter,
        api_key="PRIVATE_KEY", base_url="https://api.deepseek.com/v1", supports_vision=False)
    fake_sdk = SimpleNamespace(ClaudeAgentOptions=lambda **kwargs: SimpleNamespace(**kwargs))
    monkeypatch.setattr(claude_sdk_agent, "_sdk", lambda: fake_sdk)
    monkeypatch.setattr(claude_sdk_agent, "_state_dir", lambda: "synthetic-outside-repo")
    async def sdk_result(_sdk, prompt, options, _timeout):
        captured.append((prompt, options))
        return json.dumps(reply)
    monkeypatch.setattr(claude_sdk_agent, "_result", sdk_result)
    result = asyncio.run(intent.check_message_understanding(MessageUnderstandingRequest(prompt=question), config=connection))
    assert result.kind == "conversation"
    assert len(captured) == 1
    payload = json.loads(captured[0][0])
    options = captured[0][1]
    system = options.system_prompt
    assert options.tools == [] and options.allowed_tools == []
    assert options.max_turns == 1
    assert system.startswith(PRODUCT_IDENTITY_PROMPT)
    assert '"configuredModel": "configured-model"' in system
    assert "PRIVATE_KEY" not in system and "PRIVATE_HOST" not in system
    assert payload["prompt"] == question
    assert set(payload) == {"prompt", "conversation", "imageCount"}
