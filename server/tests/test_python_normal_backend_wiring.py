"""普通（非 SDK）后端补齐 ``prepare_python_data`` 的注册 / 分派 / 可信绑定。

改动前的事实：``excel_agent`` 里既没有 ``prepare_python_data`` 的工具定义，
也没有分派分支，更没有桥会话 —— 但 ``run_python_analysis`` **是**注册给模型
的。于是模型在普通后端只能拿到一个永远失败关闭的分析工具（没有
``dataset://`` 句柄可引用）。本文件钉这条通道被真正接通。
"""
from __future__ import annotations

import asyncio

import pytest

from server.app import excel_agent
from server.app import python_analysis
from server.app.excel_agent import (
    _execute_tool_async,
    open_planning_bridge,
    python_tool_definitions,
)
from server.app.local_results import local_results
from server.app.models import PlanRequest
from server.app.python_coordinator import coordinator
from server.app.python_datasets import registry
from server.app.task_bridge import task_tool_bridge


def _payload(**overrides) -> dict:
    payload = {
        "turnId": "turn-normal",
        "prompt": "用 Python 汇总数据",
        "planningReadPolicy": "none",
        "hostId": "excel-office",
        "workbook": {
            "name": "数据.xlsx",
            "capturedAt": "2026-10-09T00:00:00Z",
            "activeWorksheet": "Sheet1",
            "sourceFingerprint": "fp-source-1",
            "workbookIdentity": "url:https://example.com/数据.xlsx",
            "worksheets": [
                {"name": "Sheet1", "usedRange": "A1:B2", "rowCount": 2,
                 "columnCount": 2, "headers": ["a", "b"]},
            ],
        },
    }
    payload.update(overrides)
    return payload


def _request(**overrides) -> PlanRequest:
    return PlanRequest.model_validate(_payload(**overrides))


class _FakeSession:
    def __init__(self, receipt):
        self.receipt = receipt
        self.calls: list[tuple[str, object]] = []

    async def request(self, name, arguments):
        self.calls.append((name, arguments))
        return self.receipt


@pytest.fixture(autouse=True)
def _clean(monkeypatch):
    task_tool_bridge.reset()
    coordinator.reset()
    registry._grants.clear()
    registry._budgets.clear()
    local_results.reset()
    # excel_agent 用的是 ``from .python_analysis import python_execution_enabled``
    # 的绑定，必须打在它自己的模块属性上（否则打在原模块上不生效）。
    monkeypatch.setattr(python_analysis, "python_execution_enabled", lambda: True)
    monkeypatch.setattr(excel_agent, "python_execution_enabled", lambda: True)
    yield
    task_tool_bridge.reset()
    coordinator.reset()
    registry._grants.clear()
    registry._budgets.clear()
    local_results.reset()


def test_python_tool_definitions_include_the_registration_channel():
    """登记工具与分析工具必须一起给 —— 只给分析工具等于给了个不可能成功的工具。"""
    names = [item["function"]["name"] for item in python_tool_definitions()]
    assert names == ["prepare_python_data", "run_python_analysis"]


def test_python_tool_definitions_empty_when_switch_off(monkeypatch):
    monkeypatch.setattr(excel_agent, "python_execution_enabled", lambda: False)
    assert python_tool_definitions() == []


def test_prepare_python_data_dispatched_through_bridge_session():
    """分派：交给易失桥，由宿主读表回执；模型自报的额外字段不带进去。"""
    session = _FakeSession({"status": "ok", "datasetHandles": ["dataset://" + "d" * 43]})
    result = asyncio.run(
        _execute_tool_async(
            _request(),
            "prepare_python_data",
            {"sources": [{"sheet": "Sheet1", "range": "A1:B2"}], "extra": "IGNORED"},
            bridge_session=session,
        )
    )
    assert result["status"] == "ok"
    # 只把白名单参数交给桥（模型自报的多余字段不带入可信通道）。
    assert session.calls == [
        ("prepare_python_data", {"sources": [{"sheet": "Sheet1", "range": "A1:B2"}]})
    ]


def test_prepare_python_data_fails_closed_without_bridge_session():
    """没有可信读表通道 → 不能登记任何数据（不在服务端凭自报区域假装读过表）。"""
    result = asyncio.run(
        _execute_tool_async(
            _request(), "prepare_python_data",
            {"sources": [{"sheet": "Sheet1", "range": "A1:B2"}]},
        )
    )
    assert result == {"status": "error", "errorCode": "bridge_unavailable"}


def test_prepare_python_data_rejects_empty_sources_before_bridge_call():
    session = _FakeSession({"status": "ok", "datasetHandles": []})
    result = asyncio.run(
        _execute_tool_async(
            _request(), "prepare_python_data", {"sources": []}, bridge_session=session
        )
    )
    assert result == {"status": "error", "errorCode": "invalid_arguments"}
    assert session.calls == []


def test_open_planning_bridge_binds_the_turn_and_emits_opaque_pointer():
    """普通后端也要有桥会话，且桥请求只发出不透明指针（参数不落盘）。"""
    events: list[dict] = []

    async def on_event(event):
        events.append(event)

    task_id, session = open_planning_bridge(_request(), on_event)
    assert task_id == "turn-normal"
    assert session is not None
    assert task_tool_bridge._sessions.get("turn-normal") is session

    async def scenario():
        pending = asyncio.ensure_future(
            session.request("prepare_python_data",
                            {"sources": [{"sheet": "Sheet1", "range": "A1:B2"}]})
        )
        for _ in range(50):
            await asyncio.sleep(0)
            if events:
                break
        pending.cancel()
        await asyncio.gather(pending, return_exceptions=True)

    asyncio.run(scenario())
    assert events, "桥请求没有发出指针事件，宿主无法知道要读表"
    assert events[0]["bridgeTool"]["tool"] == "prepare_python_data"
    assert "operationId" in events[0]["bridgeTool"]
    # 参数留在易失桥里，不写进进度事件/账本。
    assert "sources" not in events[0]

    task_tool_bridge.close(task_id)
    assert task_tool_bridge._sessions.get("turn-normal") is None


def test_open_planning_bridge_disabled_when_switch_off(monkeypatch):
    monkeypatch.setattr(excel_agent, "python_execution_enabled", lambda: False)
    assert open_planning_bridge(_request(), None) == (None, None)


def test_open_planning_bridge_requires_turn_id():
    payload = _payload()
    payload.pop("turnId")
    assert open_planning_bridge(PlanRequest.model_validate(payload), None) == (None, None)
