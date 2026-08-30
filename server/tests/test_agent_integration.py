"""Agent 循环保护与主循环的集成测试。"""

from __future__ import annotations

import asyncio
import copy
import json
from typing import Any

import pytest

import server.app.excel_agent as excel_agent
from server.app.agent_loop_protection import LoopConfig
from server.app.llm import ModelConnection
from server.app.models import PlanRequest


def _request() -> PlanRequest:
    """构造只授权一个工作表的最小规划请求。"""
    return PlanRequest.model_validate(
        {
            "prompt": "分析得分并给出计划",
            "workbook": {
                "name": "scores.xlsx",
                "capturedAt": "2026-08-30T00:00:00Z",
                "activeWorksheet": "Sheet1",
                "worksheets": [
                    {
                        "name": "Sheet1",
                        "usedRange": "A1:B3",
                        "rowCount": 3,
                        "columnCount": 2,
                        "headers": ["人员", "得分"],
                        "dataRows": [["甲", 10], ["乙", 20]],
                        "truncated": False,
                    }
                ],
            },
        }
    )


def _connection() -> ModelConnection:
    return ModelConnection(
        base_url="https://api.example.test/v1",
        model="test-model",
        api_key="",
        supports_vision=False,
    )


def _tool_payload(name: str, arguments: dict[str, Any], call_id: str) -> dict[str, Any]:
    return {
        "choices": [
            {
                "message": {
                    "content": None,
                    "tool_calls": [
                        {
                            "id": call_id,
                            "type": "function",
                            "function": {
                                "name": name,
                                "arguments": json.dumps(
                                    arguments, ensure_ascii=False
                                ),
                            },
                        }
                    ],
                }
            }
        ]
    }


def _usage(total_tokens: int = 0) -> dict[str, int]:
    return {
        "prompt_tokens": total_tokens,
        "completion_tokens": 0,
        "total_tokens": total_tokens,
    }


class _FakeClient:
    """按回调生成每轮模型响应，并保留传入消息快照。"""

    def __init__(self, responder):
        self.responder = responder
        self.call_count = 0
        self.message_snapshots: list[list[dict[str, Any]]] = []

    async def __aenter__(self):
        return self

    async def __aexit__(self, exc_type, exc, tb):
        return False

    async def chat_completions(
        self, *, messages: list[dict[str, Any]], **_kwargs: Any
    ) -> tuple[dict[str, Any], dict[str, int]]:
        self.call_count += 1
        self.message_snapshots.append(copy.deepcopy(messages))
        return self.responder(self.call_count)


def _install_client(
    monkeypatch: pytest.MonkeyPatch, client: _FakeClient
) -> None:
    monkeypatch.setattr(
        excel_agent,
        "OpenAICompatibleClient",
        lambda *_args, **_kwargs: client,
    )


def test_consecutive_invalid_plans_trigger_safe_degradation(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    client = _FakeClient(
        lambda call_count: (
            _tool_payload("submit_plan", {"plan": {}}, f"bad-{call_count}"),
            _usage(),
        )
    )
    _install_client(monkeypatch, client)

    response = asyncio.run(
        excel_agent.run_excel_agent(
            _request(),
            connection=_connection(),
            timeout=10,
            config=LoopConfig(max_wall_time_seconds=3600),
        )
    )

    assert response.kind == "answer"
    assert response.provider == "system"
    assert "连续 3 次" in response.message
    assert "格式不正确" in response.message
    assert client.call_count == 3


def test_total_token_budget_triggers_safe_degradation(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    client = _FakeClient(
        lambda _call_count: (
            _tool_payload("get_workbook_context", {}, "context-call"),
            _usage(20_000),
        )
    )
    _install_client(monkeypatch, client)

    response = asyncio.run(
        excel_agent.run_excel_agent(
            _request(),
            connection=_connection(),
            timeout=10,
            config=LoopConfig(max_wall_time_seconds=3600),
        )
    )

    assert response.kind == "answer"
    assert response.provider == "system"
    assert "20000 tokens" in response.message
    assert client.call_count == 1


def test_repeated_cached_query_injects_warning(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def responder(call_count: int):
        if call_count <= 4:
            return (
                _tool_payload(
                    "find_fields",
                    {"query": "得分"},
                    f"find-{call_count}",
                ),
                _usage(),
            )
        return (
            _tool_payload(
                "submit_answer",
                {"message": "已根据现有信息完成分析。"},
                "answer-call",
            ),
            _usage(),
        )

    client = _FakeClient(responder)
    _install_client(monkeypatch, client)

    response = asyncio.run(
        excel_agent.run_excel_agent(
            _request(),
            connection=_connection(),
            timeout=10,
            tool_cache={},
            config=LoopConfig(
                max_turns=6,
                max_wall_time_seconds=3600,
                warning_turn=5,
            ),
        )
    )

    assert response.kind == "answer"
    warning_contents = [
        message["content"]
        for snapshot in client.message_snapshots
        for message in snapshot
        if message.get("role") == "tool"
        and isinstance(message.get("content"), str)
        and "_warning" in message["content"]
    ]
    assert warning_contents
    assert any("第 3 次查询相同内容" in content for content in warning_contents)
