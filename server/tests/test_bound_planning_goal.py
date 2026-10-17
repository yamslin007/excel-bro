"""The goal is established before terminal submission, not by that submission."""
import asyncio

import pytest

from server.app import claude_sdk_agent, excel_agent
from server.app.models import AnswerResponse, PlanResponse
from server.app.result_submission import LocalResultIncomplete
from server.tests.test_planning_goal_completion import _register, _request, _sdk_request
from server.tests.test_claude_sdk_agent import _FakeSdk, _FakeResult, _install_fake, _connection


@pytest.fixture(autouse=True)
def isolated_results():
    from server.app.local_results import local_results
    local_results.reset()
    yield
    local_results.reset()


@pytest.mark.parametrize("declaration", [None, "read_only", "write_preview", "invalid"])
def test_local_result_cannot_downgrade_bound_goal(declaration):
    item = _register()
    request = _request(planningGoal="write_preview")
    with pytest.raises(LocalResultIncomplete):
        excel_agent._execute_tool(request, "submit_local_result", {
            "resultRef": item.handle, "planningGoal": declaration,
        })


def test_answer_cannot_bypass_bound_goal():
    with pytest.raises(LocalResultIncomplete):
        excel_agent._execute_tool(_request(planningGoal="write_preview"), "submit_answer", {
            "message": "分析已完成",
        })


def test_actual_read_only_request_can_finish():
    item = _register()
    response = excel_agent._execute_tool(
        _request(prompt="只展示汇总结果，不修改工作簿", planningGoal="read_only"),
        "submit_local_result", {"resultRef": item.handle, "planningGoal": "read_only"},
    )
    assert isinstance(response, AnswerResponse)


def test_sdk_rejects_early_answer_then_compiles_in_the_same_query(monkeypatch, tmp_path):
    item = _register()

    class RecoveringSdk(_FakeSdk):
        async def query(self, *, prompt, options):
            handlers = options.mcp_servers["excel_bro"]
            for name, args in [
                ("submit_local_result", {"resultRef": item.handle, "planningGoal": "read_only"}),
                ("submit_answer", {"message": "已完成"}),
            ]:
                response = await handlers[name].handler(args)
                assert response.get("is_error") is True
                assert "write_preview_incomplete" in str(response) or "规划目标尚未完成" in str(response)
            response = await handlers["submit_result_plan"].handler({
                "resultRef": item.handle, "target": {"sheet": "汇总表", "range": "A1:B4"},
            })
            assert not response.get("is_error")
            yield _FakeResult()

    _install_fake(monkeypatch, tmp_path, RecoveringSdk())
    request = _sdk_request().model_copy(update={"planningGoal": "write_preview"})
    result = asyncio.run(claude_sdk_agent.run_sdk_plan(request, _connection()))
    assert isinstance(result, PlanResponse)
    assert result.plan.localResultRef == item.handle


def test_sdk_text_answer_cannot_bypass_bound_goal(monkeypatch, tmp_path):
    _install_fake(monkeypatch, tmp_path, _FakeSdk(response='{"kind":"answer","message":"已完成"}'))
    request = _sdk_request().model_copy(update={"planningGoal": "write_preview"})
    from server.app.llm import LLMResponseError
    with pytest.raises(LLMResponseError):
        asyncio.run(claude_sdk_agent.run_sdk_plan(request, _connection()))


@pytest.mark.parametrize("use_tool", [True, False])
def test_normal_backend_repairs_early_answer_in_same_loop(monkeypatch, use_tool):
    import json
    from server.tests.test_agent_integration import (
        _FakeClient, _install_client, _tool_payload, _content_payload, _usage,
        _connection as normal_connection,
    )
    item = _register()
    early = (_tool_payload("submit_answer", {"message": "已完成"}, "early") if use_tool
             else _content_payload(json.dumps({"kind": "answer", "message": "已完成"})))
    finish = _tool_payload("submit_result_plan", {
        "resultRef": item.handle, "target": {"sheet": "汇总表", "range": "A1:B4"},
    }, "finish")
    client = _FakeClient(lambda n: (early if n == 1 else finish, _usage()))
    _install_client(monkeypatch, client)
    response = asyncio.run(excel_agent.run_excel_agent(
        _request(planningGoal="write_preview"), connection=normal_connection(), timeout=5,
    ))
    assert isinstance(response, PlanResponse)
    assert client.call_count == 2


def test_intent_goal_is_stored_separately_from_terminal_arguments():
    from server.app.turn_state import TurnRegistry
    from server.app.models import IntentProceedResponse
    registry = TurnRegistry()
    registry.record_decision("turn-goal-binding-test", IntentProceedResponse(
        summary="生成预览", confirmedPrompt="把结果写入汇总表", provider="model",
        planningGoal="write_preview",
    ))
    assert registry.touch("turn-goal-binding-test").planning_goal == "write_preview"


def test_query_first_does_not_mean_read_only():
    from server.app.models import IntentToolResponse
    from server.app.turn_state import TurnRegistry
    registry = TurnRegistry()
    registry.record_decision("turn-query-write", IntentToolResponse(
        summary="先查询再生成预览", confirmedPrompt="汇总并写入目标表", provider="model",
        planningGoal="write_preview", request={"id": "q1", "tool": "query_table",
            "arguments": {"mode": "aggregate", "metrics": [{"operation": "countRows", "outputName": "数量"}]}}
    ))
    assert registry.touch("turn-query-write").planning_goal == "write_preview"
