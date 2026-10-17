"""规划目标完成判据：把「分析完成」与「写入预览完成」分开。

改动前的事实（真实实机复现，turn-f493dba5ae61470d99a54dd7a933e53d）：用户明确
要求「把汇总与异常明细安排到汇总表，先展示写入预览」，模型调用了
``prepare_python_data`` 拿到结果引用后，**直接调用 submit_local_result 收尾**，
``result_json.kind == "answer"``。服务端把它当合法终局，面板因此没有计划卡、
没有「生成右侧副本并排审阅」按钮 —— 任务被标成完成，实际只做了一半。

根因不是"模型不肯调工具"（``submit_result_plan`` 确实已暴露且可调用），也没有
任何报错被吞掉：**全仓根本没有"本轮规划目标是否达成"的判据**，``AnswerResponse``
被无条件当作终局；SDK 侧还额外把 ``terminal`` 置位后**锁死后续所有工具**，模型
即使想补救也没有机会。

本文件钉住的通用语义（不按关键词硬编码写入意图）：

* 模型在提交本地结果时**必须声明**它读出的目标（``planningGoal``）。
* 声明 ``read_only`` → 照常返回结果卡（**只读任务不能被强迫生成写入计划**）。
* 声明 ``write_preview`` 却只交结果卡 → 固定未完成码 + 结论/下一步文案，
  **不产生终局**，模型可在剩余预算内改用 ``submit_result_plan``。
* 文案不含任何数据值，也不承诺"必然可以继续"。
"""

from __future__ import annotations

import asyncio

import pytest

from server.app import excel_agent
from server.app.excel_agent import _execute_tool
from server.app.local_results import local_results
from server.app.models import AnswerResponse, DataToolResult, PlanRequest, PlanResponse
from server.app.result_submission import (
    WRITE_PREVIEW_INCOMPLETE_CODE,
    LocalResultIncomplete,
    evaluate_local_result_submission,
    normalize_planning_goal,
)

SENTINEL = "SENTINEL_GOAL_5C1D"
TASK_ID = "turn-goal-completion-0001"


def _result() -> DataToolResult:
    return DataToolResult.model_validate({
        "requestId": "req-goal-1",
        "tool": "query_table",
        "title": f"等级汇总 {SENTINEL}",
        "headers": ["等级", "金额"],
        "rows": [["金卡", 400], ["银卡", 200], ["未匹配", 400]],
        "sourceSheets": ["订单表", "会员表"],
        "scannedRows": 4,
        "complete": True,
        "calculation": f"左连接汇总 {SENTINEL}",
        "warnings": [],
    })


def _request(**overrides) -> PlanRequest:
    payload = {
        "turnId": TASK_ID,
        "prompt": "把汇总和异常明细安排到汇总表，先展示写入预览",
        "planningReadPolicy": "none",
        "hostId": "wps-et",
        "hostVersion": "12.0",
        "workbook": {
            "name": "Python验收-WPS.xlsx",
            "capturedAt": "2026-10-11T00:00:00Z",
            "activeWorksheet": "汇总表",
            "sourceFingerprint": "fp-goal-1",
            "worksheets": [
                {"name": "汇总表", "usedRange": "A1:C12", "rowCount": 12,
                 "columnCount": 3, "headers": ["等级", "金额", "备注"]},
                {"name": "会员表", "usedRange": "A1:C6", "rowCount": 6,
                 "columnCount": 3, "headers": ["会员编号", "会员", "等级"]},
                {"name": "订单表", "usedRange": "A1:B5", "rowCount": 5,
                 "columnCount": 2, "headers": ["会员编号", "金额"]},
            ],
        },
    }
    payload.update(overrides)
    return PlanRequest.model_validate(payload)


def _register():
    return local_results.register(TASK_ID, _result(), "fp-goal-1")


@pytest.fixture(autouse=True)
def _clean_state():
    local_results.reset()
    yield
    local_results.reset()


# ---------------------------------------------------------------------------
# A. goal normalisation — read_only is the safe default
# ---------------------------------------------------------------------------


def test_missing_or_unknown_goal_defaults_to_read_only():
    """缺省/乱填一律按只读处理：绝不无中生有地强加写入要求。"""
    assert normalize_planning_goal(None) == "read_only"
    assert normalize_planning_goal("") == "read_only"
    assert normalize_planning_goal("whatever") == "read_only"
    assert normalize_planning_goal("read_only") == "read_only"
    assert normalize_planning_goal("write_preview") == "write_preview"


def test_read_only_goal_is_never_blocked():
    # 不抛异常 = 只读任务畅通。
    evaluate_local_result_submission("read_only")


def test_write_preview_goal_raises_a_fixed_incomplete_signal():
    with pytest.raises(LocalResultIncomplete) as caught:
        evaluate_local_result_submission("write_preview")
    error = caught.value
    assert error.code == WRITE_PREVIEW_INCOMPLETE_CODE
    # 结论 + 下一步，不含数据值，也不承诺必然成功。
    text = str(error)
    assert "未完成" in text
    assert "submit_result_plan" in text
    assert SENTINEL not in text


# ---------------------------------------------------------------------------
# B. normal backend dispatch — the judgement is actually wired in
# ---------------------------------------------------------------------------


def test_normal_backend_read_only_submission_still_returns_a_card():
    """只读：结果卡照常返回（回归保护，不能被新判据误伤）。"""
    item = _register()
    response = _execute_tool(
        _request(), "submit_local_result",
        {"resultRef": item.handle, "planningGoal": "read_only"},
    )
    assert isinstance(response, AnswerResponse)
    assert response.localResult is not None
    assert response.localResult.resultRef == item.handle


def test_normal_backend_omitting_the_goal_is_treated_as_read_only():
    """未声明目标时按只读放行 —— 默认值不制造写入要求。"""
    item = _register()
    response = _execute_tool(
        _request(), "submit_local_result", {"resultRef": item.handle}
    )
    assert isinstance(response, AnswerResponse)


def test_normal_backend_refuses_an_early_answer_for_a_write_preview_goal():
    """核心：声明要写入预览却只交结果卡 → 未完成，不是终局。"""
    item = _register()
    with pytest.raises(LocalResultIncomplete) as caught:
        _execute_tool(
            _request(), "submit_local_result",
            {"resultRef": item.handle, "planningGoal": "write_preview"},
        )
    assert caught.value.code == WRITE_PREVIEW_INCOMPLETE_CODE
    assert SENTINEL not in str(caught.value)


def test_the_same_reference_can_still_become_a_plan_after_the_refusal():
    """拒绝不消耗引用：模型随后用 submit_result_plan 仍能编译出待确认计划。"""
    item = _register()
    with pytest.raises(LocalResultIncomplete):
        _execute_tool(
            _request(), "submit_local_result",
            {"resultRef": item.handle, "planningGoal": "write_preview"},
        )
    response = _execute_tool(
        _request(), "submit_result_plan",
        {"resultRef": item.handle, "target": {"sheet": "汇总表", "range": "A1:B4"}},
    )
    assert isinstance(response, PlanResponse)
    assert response.plan.actions[0].type == "writeValues"
    assert response.plan.localResultRef == item.handle


# ---------------------------------------------------------------------------
# C. SDK backend — must NOT lock the terminal, or the model can never recover
# ---------------------------------------------------------------------------


def _sdk_request():
    """SDK 请求：复用既有范式，但把 turnId 与工作表对齐到本文件登记的任务。"""
    from server.tests.test_claude_sdk_agent import _plan_request

    payload = _plan_request().model_dump(mode="json")
    payload["turnId"] = TASK_ID
    payload["workbook"]["worksheets"] = [{
        "name": "汇总表", "usedRange": "A1:C12", "rowCount": 12,
        "columnCount": 3, "headers": ["等级", "金额", "备注"],
    }]
    payload["workbook"]["activeWorksheet"] = "汇总表"
    return PlanRequest.model_validate(payload)


def test_sdk_guard_does_not_set_terminal_on_an_incomplete_submission(monkeypatch, tmp_path):
    """SDK 侧一旦置 terminal 就会锁死后续工具 —— 未完成时必须只回错。

    这是**行为断言**（真正驱动 handler），不是源码文本匹配：变异把
    ``terminal = build_local_result_answer(...)`` 塞回未完成分支后，这里
    必须能观察到"后续工具被锁死"，用例才会变红。
    """
    from server.app import claude_sdk_agent
    from server.tests.test_claude_sdk_agent import _FakeSdk, _connection, _install_fake

    item = _register()
    fake = _FakeSdk(invoke=("submit_local_result", {
        "resultRef": item.handle, "planningGoal": "write_preview",
    }))
    _install_fake(monkeypatch, tmp_path, fake)
    with pytest.raises(Exception):
        # 模型只交结果卡、不再补救 ⇒ 本轮没有受控终局（不是被当成成功的 answer）。
        asyncio.run(claude_sdk_agent.run_sdk_plan(_sdk_request(), _connection()))

    first = fake.tool_result
    assert first is not None
    assert first.get("is_error") is True
    text = first["content"][0]["text"]
    assert WRITE_PREVIEW_INCOMPLETE_CODE in text
    assert "本轮结果已提交。" not in text  # 不得表现为已终结


def test_sdk_can_still_submit_a_plan_after_the_incomplete_refusal(monkeypatch, tmp_path):
    """判别力核心：未完成**不得**置 terminal ⇒ 后续工具仍可调用并成功终结。"""
    from server.app import claude_sdk_agent
    from server.tests.test_claude_sdk_agent import _FakeSdk, _connection, _install_fake

    item = _register()
    fake = _FakeSdk(invoke=("submit_result_plan", {
        "resultRef": item.handle,
        "target": {"sheet": "汇总表", "range": "A1:B4"},
    }))
    _install_fake(monkeypatch, tmp_path, fake)
    response = asyncio.run(claude_sdk_agent.run_sdk_plan(_sdk_request(), _connection()))
    assert isinstance(response, PlanResponse)
    assert response.plan.localResultRef == item.handle
    assert response.plan.actions[0].type == "writeValues"


def test_sdk_read_only_submission_still_completes(monkeypatch, tmp_path):
    """只读目标在 SDK 侧照常终结（回归保护）。"""
    from server.app import claude_sdk_agent
    from server.tests.test_claude_sdk_agent import _FakeSdk, _connection, _install_fake

    item = _register()
    fake = _FakeSdk(invoke=("submit_local_result", {
        "resultRef": item.handle, "planningGoal": "read_only",
    }))
    _install_fake(monkeypatch, tmp_path, fake)
    response = asyncio.run(claude_sdk_agent.run_sdk_plan(_sdk_request(), _connection()))
    assert isinstance(response, AnswerResponse)
    assert response.localResult is not None


def test_sdk_plan_compile_refusal_is_marked_as_an_error(monkeypatch, tmp_path):
    """编译被拒（尺寸不符等）必须回 is_error，否则模型会把失败当成功收尾。"""
    from server.app import claude_sdk_agent
    from server.tests.test_claude_sdk_agent import _FakeSdk, _connection, _install_fake

    item = _register()
    # 结果是 4 行 × 2 列，这里故意给 4 行 × 3 列 → size_mismatch。
    fake = _FakeSdk(invoke=("submit_result_plan", {
        "resultRef": item.handle,
        "target": {"sheet": "汇总表", "range": "A1:C4"},
    }))
    _install_fake(monkeypatch, tmp_path, fake)
    with pytest.raises(Exception):
        asyncio.run(claude_sdk_agent.run_sdk_plan(_sdk_request(), _connection()))
    assert fake.tool_result is not None
    assert fake.tool_result.get("is_error") is True, "被拒回执不得表现为成功"
    assert "size_mismatch" in fake.tool_result["content"][0]["text"]


def test_both_backends_call_the_same_completion_judgement():
    """共享判据必须是实现事实：两端都调同一个函数，不各留一份分叉。"""
    from server.app import claude_sdk_agent

    assert excel_agent.evaluate_local_result_submission is evaluate_local_result_submission
    assert claude_sdk_agent.evaluate_local_result_submission is evaluate_local_result_submission
    assert claude_sdk_agent.normalize_planning_goal is normalize_planning_goal
