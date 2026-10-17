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
from server.app.models import PlanRequest, ResultContext


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


def _multi_tool_payload(
    calls: list[tuple[str, dict[str, Any], str]],
) -> dict[str, Any]:
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
                                "arguments": json.dumps(arguments, ensure_ascii=False),
                            },
                        }
                        for name, arguments, call_id in calls
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


@pytest.mark.parametrize("use_tool", [True, False])
def test_planner_returns_interactive_clarification_for_both_model_submission_paths(monkeypatch, use_tool):
    clarification = {
        "id": "statistics-unit", "summary": "需要统计单位", "question": "请选择统计单位",
        "reason": "单位会改变结果", "scopeLabel": "Sheet1", "options": [],
        "fields": [{"id": "unit", "label": "统计单位", "placeholder": "请说明单位"}],
    }
    payload = _tool_payload("submit_clarification", {"clarification": clarification}, "c1") if use_tool else _content_payload(
        json.dumps({"kind": "clarification", "clarification": clarification}, ensure_ascii=False)
    )
    client = _FakeClient(lambda _turn: (payload, _usage()))
    _install_client(monkeypatch, client)
    events = []

    async def on_event(event):
        events.append(event)

    result = asyncio.run(excel_agent.run_excel_agent(_request(), connection=_connection(), timeout=5, on_event=on_event))
    assert result.kind == "clarification"
    assert result.clarification.fields[0].label == "统计单位"
    assert client.call_count == 1
    assert any(event["title"] == "等待补充信息" for event in events)
    assert not hasattr(result, "plan")


def _install_client(
    monkeypatch: pytest.MonkeyPatch, client: _FakeClient
) -> None:
    monkeypatch.setattr(
        excel_agent,
        "OpenAICompatibleClient",
        lambda *_args, **_kwargs: client,
    )


def test_compact_mapping_does_not_carry_all_action_schemas(monkeypatch):
    request = _request().model_copy(update={"planningReadPolicy": "none"})
    request.workbook.worksheets[0].dataRows[0][0] = "LOCAL_PRIVATE_SENTINEL"
    request.lastResult = ResultContext(kind="table", title="上轮结果", headers=["标签"], rows=[["PRIVATE_PREVIOUS_SENTINEL"]], sourceSheets=["Sheet1"])
    assert request.workbook.worksheets[0].dataRows[0][0] == "LOCAL_PRIVATE_SENTINEL"
    assert request.lastResult.rows[0][0] == "PRIVATE_PREVIOUS_SENTINEL"
    args = {"sourceSheet": "Sheet1", "sourceRange": "A2:A3", "targetRange": "B2:B3",
            "mapping": [{"from": "甲", "to": "第一类"}, {"from": "乙", "to": "第二类"}]}
    client = _FakeClient(lambda _: (_tool_payload("map_column_values", args, "mapping"), _usage(100)))
    _install_client(monkeypatch, client)
    result = asyncio.run(excel_agent.run_excel_agent(request, connection=_connection(), timeout=5))
    assert result.kind == "local_plan_rule"
    assert client.call_count == 1
    tools = client.call_kwargs[0]["tools"]
    names = {tool["function"]["name"] for tool in tools}
    assert "map_column_values" in names
    assert "select_plan_actions" in names
    assert "submit_plan" not in names
    assert "read_range" not in names
    assert len(json.dumps(tools, ensure_ascii=False)) < len(json.dumps(excel_agent.AGENT_TOOLS_WITH_DATA, ensure_ascii=False)) / 2
    # Local sentinels exist, but never enter the model's initial context.
    assert "LOCAL_PRIVATE_SENTINEL" not in json.dumps(client.message_snapshots, ensure_ascii=False)
    assert "PRIVATE_PREVIOUS_SENTINEL" not in json.dumps(client.message_snapshots, ensure_ascii=False)


def test_schema_selection_loads_only_selected_types_and_survives_checkpoint(monkeypatch):
    plan = {"id": "compact-plan", "title": "清空内容", "summary": "清空已授权区域", "actions": [
        {"type": "clearRange", "sheet": "Sheet1", "range": "B2:B3", "applyTo": "contents"}
    ]}
    client = _FakeClient(lambda turn: (
        _tool_payload("select_plan_actions", {"actionTypes": ["clearRange"]}, "select") if turn == 1
        else _tool_payload("submit_plan", {"plan": plan}, "plan"), _usage()))
    _install_client(monkeypatch, client)
    checkpoints = []
    async def save(checkpoint):
        checkpoints.append(copy.deepcopy(checkpoint))
    result = asyncio.run(excel_agent.run_excel_agent(_request(), connection=_connection(), timeout=5, on_checkpoint=save))
    assert result.kind == "plan"
    tool = next(tool for tool in client.call_kwargs[1]["tools"] if tool["function"]["name"] == "submit_plan")
    variants = tool["function"]["parameters"]["properties"]["plan"]["properties"]["actions"]["items"]["anyOf"]
    assert [variant["properties"]["type"]["const"] for variant in variants] == ["clearRange"]
    assert "applyTo" in variants[0]["properties"]
    assert any(checkpoint["selectedActionTypes"] == ["clearRange"] for checkpoint in checkpoints)
    resumed_client = _FakeClient(lambda _: (_tool_payload("submit_plan", {"plan": plan}, "resumed"), _usage()))
    _install_client(monkeypatch, resumed_client)
    boundary = next(checkpoint for checkpoint in checkpoints if checkpoint["stage"] == "tool_batch_completed")
    resumed = asyncio.run(excel_agent.run_excel_agent(_request(), connection=_connection(), timeout=5, resume_checkpoint=boundary))
    assert resumed.kind == "plan"
    assert resumed_client.call_kwargs[0]["tools"] == client.call_kwargs[1]["tools"]


def test_schema_selection_rejects_unknown_or_unavailable_actions():
    with pytest.raises(ValueError):
        excel_agent._select_plan_actions(_request(), {"actionTypes": ["arbitraryScript"]})
    request = _request().model_copy(update={"hostId": "unrecognized-host"})
    with pytest.raises(ValueError):
        excel_agent._select_plan_actions(request, {"actionTypes": ["writeValues"]})


def test_no_data_policy_also_rejects_unadvertised_read_tool():
    request = _request().model_copy(update={"planningReadPolicy": "none"})
    with pytest.raises(ValueError, match="不允许"):
        excel_agent._execute_tool(request, "read_range", {"sheet": "Sheet1", "range": "A1:B3"})


def test_truncated_even_valid_tool_output_is_not_executed(monkeypatch):
    payload = _tool_payload("submit_answer", {"message": "不该被接受"}, "incomplete")
    payload["choices"][0]["finish_reason"] = "length"
    client = _FakeClient(lambda _: (payload, _usage()))
    _install_client(monkeypatch, client)
    result = asyncio.run(excel_agent.run_excel_agent(_request(), connection=_connection(), timeout=5))
    assert result.kind == "answer"
    assert result.outcome == "failed"
    assert result.stopReason == "output_length"
    assert "模型响应被截断" in result.message
    assert "格式要求" not in result.message
    assert "简化需求" not in result.message
    assert client.call_count == 1
    assert client.call_kwargs[0]["max_tokens"] is None


def test_production_output_budgets_use_provider_defaults():
    from server.app.capabilities import capability_optional_int
    assert excel_agent.MAX_OUTPUT_TOKENS is None
    assert excel_agent.REPAIR_MAX_OUTPUT_TOKENS is None
    assert capability_optional_int("intentContext", "maxOutputTokens") is None
    assert excel_agent._default_loop_config().max_total_tokens is None


@pytest.mark.parametrize("first_total", [40_000, 1_000_000])
def test_large_reasoning_response_can_select_actions_then_submit_plan(monkeypatch, first_total):
    plan = {"id": "long-reasoning", "title": "清空", "summary": "清空授权区域",
            "actions": [{"type": "clearRange", "sheet": "Sheet1", "range": "B2:B3", "applyTo": "contents"}]}
    def responder(call_count):
        if call_count == 1:
            return (_tool_payload("select_plan_actions", {"actionTypes": ["clearRange"]}, "select"),
                    {"prompt_tokens": first_total - 35000, "completion_tokens": 35000, "total_tokens": first_total, "reasoning_tokens": 34000})
        return _tool_payload("submit_plan", {"plan": plan}, "plan"), _usage(1000)
    client = _FakeClient(responder)
    _install_client(monkeypatch, client)
    result = asyncio.run(excel_agent.run_excel_agent(_request(), connection=_connection(), timeout=5))
    assert result.kind == "plan"
    assert result.plan.id == "long-reasoning"
    assert client.call_count == 2
    assert all(call["max_tokens"] is None for call in client.call_kwargs)


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


def test_budget_pause_resumes_complete_model_response_without_repeating_tool(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """总时限到达时在"启动下一次模型调用前"暂停，且已保存的响应不重复执行工具。

    2026-10-10 时间策略改造后，"短时预算暂停"已不存在：``approved_wall_time_seconds``
    只作 ``reminder_seconds`` 的兼容别名（只发非阻塞提醒）。真正的暂停由
    ``note_progress(..., enforce_total_wall_time=True)`` 在**累计活跃执行时间**达到
    ``LoopConfig.max_wall_time_seconds`` 时产生。所以这里把总时限设得极小来触发暂停，
    而不是靠旧的"批准预算"。
    """
    def responder(call_count: int):
        if call_count == 1:
            return (
                _tool_payload(
                    "read_range",
                    {"sheet": "Sheet1", "range": "A1:B3"},
                    "read-after-resume",
                ),
                _usage(1_000_000),
            )
        return (
            _tool_payload(
                "submit_answer",
                {"message": "已从保存的模型响应继续。"},
                "answer-after-resume",
            ),
            _usage(),
        )

    class SlowFirstClient(_FakeClient):
        async def chat_completions(self, **kwargs: Any):
            if self.call_count == 0:
                await asyncio.sleep(0.08)
            return await super().chat_completions(**kwargs)

    client = SlowFirstClient(responder)
    _install_client(monkeypatch, client)
    original_execute_tool = excel_agent._execute_tool
    executed_tools: list[str] = []

    def counted_execute_tool(request, name, arguments):
        executed_tools.append(name)
        return original_execute_tool(request, name, arguments)

    monkeypatch.setattr(excel_agent, "_execute_tool", counted_execute_tool)
    checkpoints: list[dict[str, Any]] = []

    async def save_checkpoint(checkpoint: dict[str, Any]) -> None:
        checkpoints.append(copy.deepcopy(checkpoint))

    with pytest.raises(excel_agent.AgentBudgetPause) as paused:
        asyncio.run(
            excel_agent.run_excel_agent(
                _request(),
                connection=_connection(),
                timeout=10,
                tool_cache={},
                on_checkpoint=save_checkpoint,
                # 总时限极小 ⇒ 第一次工具执行耗掉 0.08s 后，启动下一次模型调用前
                # 由 note_progress(enforce_total_wall_time=True) 收口。
                config=LoopConfig(max_wall_time_seconds=0.05),
            )
        )

    checkpoint = paused.value.checkpoint
    # 2026-10-10 时间策略改造后，总时限收口发生在**启动下一次模型调用之前**，
    # 所以暂停边界就是 ``before_model``（而不是旧的 ``model_response_saved``）。
    # 上一轮的完整模型响应早在 ``model_response_saved`` 边界落过盘，这里要
    # 断言的是"暂停时已保存的响应不会被丢掉、也不会被重复执行"。
    assert checkpoint["stage"] == "before_model"
    response_checkpoint = next(
        item for item in checkpoints if item["stage"] == "model_response_saved"
    )
    assert response_checkpoint["loopState"]["totalTokens"] == 1_000_000
    assert response_checkpoint["pendingAssistantMessage"]["tool_calls"][0]["function"][
        "name"
    ] == "read_range"
    # 从 ``model_response_saved`` 到收口为止，那个工具恰好执行了一次；恢复后
    # 不得再执行第二次（完整响应里的工具步骤已提交，不能重复）。
    assert executed_tools == ["read_range"]
    assert checkpoints[-1] == checkpoint

    response = asyncio.run(
        excel_agent.run_excel_agent(
            _request(),
            connection=_connection(),
            timeout=10,
            tool_cache={},
            resume_checkpoint=checkpoint,
            approved_wall_time_seconds=10,
            config=LoopConfig(max_wall_time_seconds=3600),
        )
    )

    assert response.kind == "answer"
    assert response.message == "已从保存的模型响应继续。"
    assert executed_tools.count("read_range") == 1
    assert client.call_count == 2


def test_complete_model_response_is_checkpointed_before_budget_pause(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    client = _FakeClient(
        lambda _call_count: (
            _content_payload(
                json.dumps(
                    {"kind": "answer", "message": "已经完成只读分析。"},
                    ensure_ascii=False,
                )
            ),
            _usage(),
        )
    )
    _install_client(monkeypatch, client)
    checkpoints: list[dict[str, Any]] = []

    async def save_checkpoint(checkpoint: dict[str, Any]) -> None:
        checkpoints.append(copy.deepcopy(checkpoint))

    response = asyncio.run(
        excel_agent.run_excel_agent(
            _request(),
            connection=_connection(),
            timeout=10,
            tool_cache={},
            on_checkpoint=save_checkpoint,
            approved_wall_time_seconds=60,
            config=LoopConfig(max_wall_time_seconds=3600),
        )
    )

    assert response.kind == "answer"
    stages = [checkpoint["stage"] for checkpoint in checkpoints]
    assert stages[:2] == ["before_model", "model_response_saved"]
    response_checkpoint = checkpoints[1]
    assert response_checkpoint["pendingAssistantMessage"]["content"]
    assert response_checkpoint["nextTurnIndex"] == 0


def test_resume_skips_tools_already_committed_in_same_model_response(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def responder(call_count: int):
        if call_count == 1:
            return (
                _multi_tool_payload(
                    [
                        (
                            "read_range",
                            {"sheet": "Sheet1", "range": "A1:A3"},
                            "read-first",
                        ),
                        (
                            "read_range",
                            {"sheet": "Sheet1", "range": "B1:B3"},
                            "read-second",
                        ),
                    ]
                ),
                _usage(),
            )
        return (
            _tool_payload(
                "submit_answer",
                {"message": "两个字段均已读取。"},
                "answer-after-tools",
            ),
            _usage(),
        )

    client = _FakeClient(responder)
    _install_client(monkeypatch, client)
    original_execute_tool = excel_agent._execute_tool
    executed_ranges: list[str] = []

    def counted_execute_tool(request, name, arguments):
        if name == "read_range":
            executed_ranges.append(str(arguments.get("range")))
        return original_execute_tool(request, name, arguments)

    monkeypatch.setattr(excel_agent, "_execute_tool", counted_execute_tool)
    committed_checkpoint: dict[str, Any] | None = None

    async def stop_after_first_tool(checkpoint: dict[str, Any]) -> None:
        nonlocal committed_checkpoint
        if (
            checkpoint.get("stage") == "tool_completed"
            and checkpoint.get("nextToolIndex") == 1
        ):
            committed_checkpoint = copy.deepcopy(checkpoint)
            raise RuntimeError("simulated process exit after first tool")

    with pytest.raises(RuntimeError, match="simulated process exit"):
        asyncio.run(
            excel_agent.run_excel_agent(
                _request(),
                connection=_connection(),
                timeout=10,
                tool_cache={},
                on_checkpoint=stop_after_first_tool,
                config=LoopConfig(max_wall_time_seconds=3600),
            )
        )

    assert committed_checkpoint is not None
    assert committed_checkpoint["schemaVersion"] == 2
    assert committed_checkpoint["nextToolIndex"] == 1
    assert [message["role"] for message in committed_checkpoint["messages"][-2:]] == [
        "assistant",
        "tool",
    ]
    assert executed_ranges == ["A1:A3"]

    response = asyncio.run(
        excel_agent.run_excel_agent(
            _request(),
            connection=_connection(),
            timeout=10,
            tool_cache={},
            resume_checkpoint=committed_checkpoint,
            approved_wall_time_seconds=60,
            config=LoopConfig(max_wall_time_seconds=3600),
        )
    )

    assert response.kind == "answer"
    assert response.message == "两个字段均已读取。"
    assert executed_ranges == ["A1:A3", "B1:B3"]
    assert client.call_count == 2
    resumed_messages = client.message_snapshots[-1]
    assistant_messages = [
        message for message in resumed_messages if message.get("role") == "assistant"
    ]
    assert len(assistant_messages) == 1
    assert [
        message.get("tool_call_id")
        for message in resumed_messages
        if message.get("role") == "tool"
    ] == ["read-first", "read-second"]


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
            config=LoopConfig(max_total_tokens=20_000, max_wall_time_seconds=3600),
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
