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


def _content_payload(content: str) -> dict[str, Any]:
    return {"choices": [{"message": {"content": content, "tool_calls": []}}]}


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
        self.call_kwargs: list[dict[str, Any]] = []

    async def __aenter__(self):
        return self

    async def __aexit__(self, exc_type, exc, tb):
        return False

    async def chat_completions(
        self, *, messages: list[dict[str, Any]], **_kwargs: Any
    ) -> tuple[dict[str, Any], dict[str, int]]:
        self.call_count += 1
        self.message_snapshots.append(copy.deepcopy(messages))
        self.call_kwargs.append(copy.deepcopy(_kwargs))
        return self.responder(self.call_count)


def _install_client(
    monkeypatch: pytest.MonkeyPatch, client: _FakeClient
) -> None:
    monkeypatch.setattr(
        excel_agent,
        "OpenAICompatibleClient",
        lambda *_args, **_kwargs: client,
    )


def test_progress_events_summarize_tools_without_exposing_raw_rows(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def responder(call_count: int):
        if call_count == 1:
            return (
                _tool_payload(
                    "read_range",
                    {"sheet": "Sheet1", "range": "A1:B3"},
                    "read-call",
                ),
                _usage(),
            )
        return (
            _tool_payload(
                "submit_answer",
                {"message": "已完成受控读取。"},
                "answer-call",
            ),
            _usage(),
        )

    client = _FakeClient(responder)
    _install_client(monkeypatch, client)
    events: list[dict[str, Any]] = []

    async def collect(event: dict[str, Any]) -> None:
        events.append(event)

    response = asyncio.run(
        excel_agent.run_excel_agent(
            _request(),
            connection=_connection(),
            timeout=10,
            on_event=collect,
            config=LoopConfig(max_wall_time_seconds=3600),
        )
    )

    assert response.kind == "answer"
    completed_read = next(
        event
        for event in events
        if event.get("stepId") == "planning.tool.1.1"
        and event.get("status") == "completed"
    )
    assert "3 行 × 2 列" in str(completed_read.get("summary"))
    assert completed_read["trace"]["kind"] == "evidence"
    assert completed_read["toolName"] == "read_range"
    assert {fact["label"] for fact in completed_read["facts"]} >= {
        "工作表",
        "范围",
        "结果",
        "快照",
    }
    started_read = next(
        event
        for event in events
        if event.get("stepId") == "planning.tool.1.1"
        and event.get("status") == "started"
    )
    assert started_read["toolName"] == "read_range"
    assert any(fact["value"] == "A1:B3" for fact in started_read["facts"])
    serialized = json.dumps(events, ensure_ascii=False)
    assert '"values"' not in serialized
    assert "甲" not in serialized


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


def test_invalid_plan_uses_compact_plan_only_repair_round(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    valid_plan = {
        "id": "repaired-plan",
        "title": "创建结果表",
        "summary": "按确认后的口径创建结果表",
        "assumptions": [],
        "warnings": [],
        "actions": [{"type": "createWorksheet", "sheet": "结果"}],
    }

    def responder(call_count: int):
        if call_count == 1:
            return (
                _tool_payload(
                    "submit_plan",
                    {"plan": {"invalid": "x" * 5000}},
                    "bad-plan",
                ),
                _usage(8_000),
            )
        return (
            _tool_payload("submit_plan", {"plan": valid_plan}, "fixed-plan"),
            _usage(8_000),
        )

    client = _FakeClient(responder)
    _install_client(monkeypatch, client)

    response = asyncio.run(
        excel_agent.run_excel_agent(
            _request(),
            connection=_connection(),
            timeout=10,
            config=LoopConfig(max_wall_time_seconds=3600),
        )
    )

    assert response.kind == "plan"
    assert response.plan.id == "repaired-plan"
    assert client.call_kwargs[0]["max_tokens"] == excel_agent.MAX_OUTPUT_TOKENS
    assert client.call_kwargs[1]["max_tokens"] == excel_agent.REPAIR_MAX_OUTPUT_TOKENS
    repair_tools = client.call_kwargs[1]["tools"]
    assert len(repair_tools) == 1
    assert repair_tools[0]["function"]["name"] == "submit_plan"
    prior_assistant = next(
        item
        for item in client.message_snapshots[1]
        if item.get("role") == "assistant" and item.get("tool_calls")
    )
    compact_arguments = prior_assistant["tool_calls"][0]["function"]["arguments"]
    assert compact_arguments == '{"plan": null}'
    assert "x" * 100 not in json.dumps(
        client.message_snapshots[1], ensure_ascii=False
    )


def test_redundant_split_field_is_normalized_without_model_retry() -> None:
    arguments = {
        "plan": {
            "id": "split-plan",
            "title": "按地区拆分",
            "summary": "按地区建表并在表内按产品汇总",
            "assumptions": [],
            "warnings": [],
            "actions": [
                {
                    "type": "splitGroupAggregate",
                    "sheet": "Sheet1",
                    "splitBy": "地区",
                    "groupBy": ["地区", "产品"],
                    "metrics": [
                        {
                            "operation": "countRows",
                            "outputName": "记录数",
                            "ratioOutputName": "占比",
                        }
                    ],
                }
            ],
        }
    }

    response = excel_agent._execute_tool(_request(), "submit_plan", arguments)

    assert response.kind == "plan"
    action = response.plan.actions[0]
    assert action.type == "splitGroupAggregate"
    assert action.splitBy == "地区"
    assert action.groupBy == ["产品"]
    assert arguments["plan"]["actions"][0]["groupBy"] == ["地区", "产品"]
    assert any(
        "拆分字段会自动保留在结果列中" in warning
        for warning in response.plan.warnings
    )


def test_null_default_collections_do_not_trigger_model_retry(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    submitted_plan = {
        "id": "default-collections",
        "title": "创建结果表",
        "summary": "创建一张结果工作表",
        "sourceFingerprintSheets": None,
        "assumptions": None,
        "warnings": None,
        "acceptanceCriteria": None,
        "actions": [{"type": "createWorksheet", "sheet": "结果"}],
    }
    client = _FakeClient(
        lambda _call_count: (
            _tool_payload(
                "submit_plan",
                {"plan": submitted_plan},
                "default-collections",
            ),
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

    assert response.kind == "plan"
    assert response.plan.sourceFingerprintSheets == []
    assert response.plan.assumptions == []
    assert response.plan.warnings == []
    assert client.call_count == 1
    assert submitted_plan["sourceFingerprintSheets"] is None


def test_invalid_free_text_is_compacted_before_repair_round(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    repair_plan = {
        "id": "content-repair",
        "title": "按地区拆分",
        "summary": "按地区建表并在表内按产品汇总",
        "actions": [
            {
                "type": "splitGroupAggregate",
                "sheet": "Sheet1",
                "splitBy": "地区",
                "groupBy": ["地区", "产品"],
                "metrics": [
                    {
                        "operation": "countRows",
                        "outputName": "记录数",
                        "ratioOutputName": "占比",
                    }
                ],
            }
        ],
    }

    def responder(call_count: int):
        if call_count == 1:
            return _content_payload("x" * 5000), _usage(15_000)
        return (
            _tool_payload("submit_plan", {"plan": repair_plan}, "fixed-content"),
            _usage(15_000),
        )

    client = _FakeClient(responder)
    _install_client(monkeypatch, client)

    response = asyncio.run(
        excel_agent.run_excel_agent(
            _request(),
            connection=_connection(),
            timeout=10,
            config=LoopConfig(max_wall_time_seconds=3600),
        )
    )

    assert response.kind == "plan"
    assert response.plan.id == "content-repair"
    assert response.plan.actions[0].type == "splitGroupAggregate"
    assert response.plan.actions[0].groupBy == ["产品"]
    assert client.call_kwargs[1]["max_tokens"] == excel_agent.REPAIR_MAX_OUTPUT_TOKENS
    serialized = json.dumps(client.message_snapshots[1], ensure_ascii=False)
    assert "x" * 100 not in serialized
    assert "未提交有效的结构化计划或回答" in serialized


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
