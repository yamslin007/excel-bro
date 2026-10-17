"""普通（非 SDK）后端补齐 ``submit_local_result`` / ``submit_result_plan``。

改动前的事实：``agent_tools.local_terminal_tool_definitions()`` 里定义了这两个
终局工具，但只有 SDK 适配器（``claude_sdk_agent``）注册并分派它们。普通后端的
``_execute_tool`` 没有对应分支，模型一旦调用就拿到
``{"error": "未知工具：submit_result_plan"}`` —— 于是"本地结果 → 白名单写入计划
→ 预览"这条在线链路在普通后端直接断掉，``docs/ARCHITECTURE.md`` 里"两后端共用"
的说法与实际实现不符。

本文件钉住三件事：

* **注册**：普通后端的工具目录里确实有这两个工具；``plan_only`` 模式仍然只给
  ``submit_plan``（该模式的定义就是"只允许提交计划"）。
* **分派**：按引用返回受控响应 —— 读任务用 ``AnswerResponse``（带结果卡指针），
  写任务用 ``PlanResponse``（**待确认**计划，提交不等于执行）。
* **归属校验**：无效 / 过期 / 属于别的任务 / 授权已被替换的引用一律拒绝，且
  拒绝文案里不含任何来自源数据的自由文本。
"""

from __future__ import annotations

import time

import pytest

from server.app import excel_agent
from server.app.excel_agent import (
    AGENT_TOOLS_WITH_DATA,
    PLAN_ONLY_TOOLS,
    _execute_tool,
    planning_tool_definitions,
)
from server.app.local_results import local_results
from server.app.models import (
    AnswerResponse,
    DataToolResult,
    PlanRequest,
    PlanResponse,
)
from server.app.result_submission import (
    RESULT_REF_INVALID_CODE,
    ResultSubmissionError,
    build_local_result_answer,
    local_result_card,
    resolve_result_ref,
)

SENTINEL = "SENTINEL_SECRET_9F3A"

TASK_ID = "turn-normal-result-0001"


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------


def _result(*, rows=None, headers=None, sources=None, title=None) -> DataToolResult:
    headers = headers or ["等级", "人数"]
    rows = rows if rows is not None else [["金卡", 4], ["银卡", 2]]
    return DataToolResult.model_validate({
        "requestId": "req-normal-1",
        "tool": "query_table",
        "title": title or f"会员等级汇总 {SENTINEL}",
        "headers": headers,
        "rows": rows,
        "sourceSheets": sources or ["订单"],
        "scannedRows": 8,
        "complete": True,
        "calculation": f"按等级分组计数 {SENTINEL}",
        "warnings": [f"注意 {SENTINEL}"],
    })


def _payload(**overrides) -> dict:
    payload = {
        "turnId": TASK_ID,
        "prompt": "把汇总结果写到汇总表",
        "planningReadPolicy": "none",
        "hostId": "excel-office",
        "hostVersion": "16.0",
        "workbook": {
            "name": "会员.xlsx",
            "capturedAt": "2026-10-09T00:00:00Z",
            "activeWorksheet": "订单",
            "sourceFingerprint": "fp-source-1",
            "worksheets": [
                {"name": "订单", "usedRange": "A1:C6", "rowCount": 6,
                 "columnCount": 3, "headers": ["订单号", "会员", "等级"]},
                {"name": "汇总", "usedRange": "A1:B4", "rowCount": 4,
                 "columnCount": 2, "headers": ["等级", "人数"]},
            ],
        },
    }
    payload.update(overrides)
    return payload


def _request(**overrides) -> PlanRequest:
    return PlanRequest.model_validate(_payload(**overrides))


def _register(**kwargs):
    return local_results.register(TASK_ID, _result(**kwargs), "fp-source-1")


@pytest.fixture(autouse=True)
def _clean_state():
    local_results.reset()
    yield
    local_results.reset()


# ---------------------------------------------------------------------------
# A. registration
# ---------------------------------------------------------------------------


def test_normal_backend_registers_the_local_result_tools():
    """普通后端的工具目录里必须有这两个终局工具，否则模型拿不到引用入口。"""
    names = [
        tool["function"]["name"]
        for tool in planning_tool_definitions(AGENT_TOOLS_WITH_DATA)
    ]
    assert "submit_local_result" in names
    assert "submit_result_plan" in names


def test_plan_only_mode_never_exposes_the_result_tools():
    """``plan_only`` 的定义就是只允许提交计划，不能被追加的工具破坏。"""
    tools = planning_tool_definitions(PLAN_ONLY_TOOLS)
    assert tools is PLAN_ONLY_TOOLS
    assert [tool["function"]["name"] for tool in tools] == ["submit_plan"]


# ---------------------------------------------------------------------------
# B. dispatch — local result display
# ---------------------------------------------------------------------------


def test_submit_local_result_returns_a_card_pointer_not_values():
    """读任务：按引用返回结果卡指针，响应本身不携带结果行。"""
    item = _register()
    response = _execute_tool(
        _request(), "submit_local_result", {"resultRef": item.handle}
    )
    assert isinstance(response, AnswerResponse)
    assert response.localResult is not None
    assert response.localResult.resultRef == item.handle
    assert response.localResult.headers == ["等级", "人数"]
    assert response.localResult.rowCount == 2
    assert not hasattr(response.localResult, "rows")
    # 整个响应会进聊天流 / 账本，不得携带来自源数据的自由文本。
    assert SENTINEL not in response.model_dump_json()


def test_submit_local_result_defaults_a_bounded_message():
    item = _register()
    response = _execute_tool(_request(), "submit_local_result", {"resultRef": item.handle})
    assert response.message
    assert len(response.message) <= 4000
    # 模型自报的超长 message 必须被截断，不能原样进账本。
    long_response = _execute_tool(
        _request(), "submit_local_result",
        {"resultRef": item.handle, "message": "x" * 5000},
    )
    assert len(long_response.message) <= 4000


# ---------------------------------------------------------------------------
# C. dispatch — result → plan (submit ≠ execute)
# ---------------------------------------------------------------------------


def test_submit_result_plan_compiles_a_pending_confirmation_plan():
    """写任务：按引用编译白名单计划并返回**待确认**预览，不执行任何写入。"""
    item = _register()
    response = _execute_tool(
        _request(),
        "submit_result_plan",
        {"resultRef": item.handle, "target": {"sheet": "汇总", "range": "G2:H4"}},
    )
    assert isinstance(response, PlanResponse)
    actions = response.plan.actions
    assert len(actions) == 1
    assert actions[0].type == "writeValues"
    assert actions[0].sheet == "汇总"
    assert actions[0].range == "G2:H4"
    # 表头 + 数据行（includeHeader 默认 true）。
    assert actions[0].values == [["等级", "人数"], ["金卡", 4], ["银卡", 2]]
    # 计划引用回本地结果，值只在本机内存。
    assert response.plan.localResultRef == item.handle
    # 验收条件由计划自动推断：写后按 rangeEquals 回读。
    assert any(
        criterion.type == "rangeEquals"
        for criterion in response.plan.acceptanceCriteria
    )


def test_submit_result_plan_can_omit_the_header_row():
    item = _register()
    response = _execute_tool(
        _request(),
        "submit_result_plan",
        {"resultRef": item.handle, "includeHeader": False,
         "target": {"sheet": "汇总", "range": "G2:H3"}},
    )
    assert response.plan.actions[0].values == [["金卡", 4], ["银卡", 2]]


# ---------------------------------------------------------------------------
# D. refusals — invalid / expired / foreign / re-authorized
# ---------------------------------------------------------------------------


def _refusal(request: PlanRequest, tool: str, arguments: dict) -> ResultSubmissionError:
    with pytest.raises(ResultSubmissionError) as caught:
        _execute_tool(request, tool, arguments)
    return caught.value


@pytest.mark.parametrize("tool", ["submit_local_result", "submit_result_plan"])
def test_unknown_reference_is_refused(tool):
    arguments: dict = {"resultRef": "res_" + "0" * 32}
    if tool == "submit_result_plan":
        arguments["target"] = {"sheet": "汇总", "range": "G2:H4"}
    error = _refusal(_request(), tool, arguments)
    assert error.code == RESULT_REF_INVALID_CODE
    assert SENTINEL not in str(error)


@pytest.mark.parametrize("tool", ["submit_local_result", "submit_result_plan"])
def test_a_reference_from_another_task_is_refused(tool):
    """知道句柄字符串不等于拥有它：归属按 task 校验。"""
    item = _register()
    arguments: dict = {"resultRef": item.handle}
    if tool == "submit_result_plan":
        arguments["target"] = {"sheet": "汇总", "range": "G2:H4"}
    error = _refusal(
        _request(turnId="turn-normal-result-other"), tool, arguments
    )
    assert error.code == RESULT_REF_INVALID_CODE


def test_an_expired_reference_is_refused():
    item = _register()
    stored = local_results._by_handle[item.handle]
    stored.created_at = time.monotonic() - 10_000
    error = _refusal(
        _request(), "submit_result_plan",
        {"resultRef": item.handle, "target": {"sheet": "汇总", "range": "G2:H4"}},
    )
    assert error.code == RESULT_REF_INVALID_CODE


def test_replacing_the_authorization_invalidates_retained_results():
    """重授权（同一任务、不同授权范围）后旧结果不得再被读出。"""
    local_results.bind(
        task_id=TASK_ID, host_id="excel-office", workbook_name="会员.xlsx",
        workbook_identity="url:https://example.com/会员.xlsx",
        authorized_key="a" * 64, source_fingerprint="fp-source-1",
    )
    item = _register()
    assert local_results.resolve(TASK_ID, item.handle) is item

    local_results.bind(
        task_id=TASK_ID, host_id="excel-office", workbook_name="会员.xlsx",
        workbook_identity="url:https://example.com/会员.xlsx",
        authorized_key="b" * 64, source_fingerprint="fp-source-1",
    )
    error = _refusal(
        _request(), "submit_result_plan",
        {"resultRef": item.handle, "target": {"sheet": "汇总", "range": "G2:H4"}},
    )
    assert error.code == RESULT_REF_INVALID_CODE


def test_a_plan_refusal_keeps_a_fixed_code_and_never_leaks_values():
    item = _register()
    error = _refusal(
        _request(), "submit_result_plan",
        {"resultRef": item.handle, "target": {"sheet": "汇总", "range": "G2:H2"}},
    )
    assert error.code == "size_mismatch"
    assert SENTINEL not in str(error)


def test_an_unauthorized_target_sheet_is_refused():
    item = _register()
    error = _refusal(
        _request(), "submit_result_plan",
        {"resultRef": item.handle, "target": {"sheet": "别的表", "range": "G2:H4"}},
    )
    assert error.code == "unauthorized_scope"


def test_a_stale_source_result_is_refused():
    item = _register()
    error = _refusal(
        _request(workbook={
            **_payload()["workbook"], "sourceFingerprint": "fp-source-2",
        }),
        "submit_result_plan",
        {"resultRef": item.handle, "target": {"sheet": "汇总", "range": "G2:H4"}},
    )
    assert error.code == "source_changed"


# ---------------------------------------------------------------------------
# E. shared implementation — both backends use one copy
# ---------------------------------------------------------------------------


def test_both_backends_share_one_result_submission_implementation():
    """``docs/ARCHITECTURE.md`` 的"两后端共用"必须是实现事实，不是文档措辞。"""
    from server.app import claude_sdk_agent

    assert claude_sdk_agent._local_result_card is local_result_card
    assert claude_sdk_agent.resolve_result_ref is resolve_result_ref
    assert claude_sdk_agent.build_local_result_answer is build_local_result_answer


def test_excel_agent_exposes_no_private_fork_of_the_card_builder():
    """普通后端不得自带一份结果卡构造，必须复用共享实现。"""
    assert excel_agent.resolve_result_ref is resolve_result_ref
