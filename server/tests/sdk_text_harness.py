"""Stub only the SDK result transport; production routing/parsing stay real."""
import json

from server.app import claude_sdk_agent
from server.app.llm import LLMResponseError


def install_text_transport(monkeypatch, client):
    monkeypatch.setattr(claude_sdk_agent, "_sdk", lambda: object())
    monkeypatch.setattr(claude_sdk_agent, "_options", lambda *_a, **_k: None)

    async def result(_sdk, prompt, options, timeout, *, connection=None):
        payload, _ = await client.chat_completions(
            messages=[{"role": "system", "content": "SDK restricted test transport"},
                      {"role": "user", "content": prompt}], max_tokens=None,
        )
        choice = payload["choices"][0]
        if choice.get("finish_reason") == "length":
            raise LLMResponseError("SDK 未能完成本轮请求（截断）；尚未读取工作表。")
        return choice["message"]["content"]
    monkeypatch.setattr(claude_sdk_agent, "_result", result)
