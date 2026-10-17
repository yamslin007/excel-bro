"""隔离 Python 的**同步容器运行**不得阻塞服务端/SDK 事件循环。

事实背景（改动前）：``run_python_isolated`` 是一次同步容器运行（上限
``maxWallTimeSeconds``）；FastAPI 端点与 SDK 的 MCP 工具处理器都是
``async def``、跑在同一个事件循环上，而全服没有任何 ``to_thread`` /
``run_in_executor`` —— 于是 ``_execute_tool`` 那次同步调用会把整台服务
（其它轮次的流式事件、工具桥 ``/poll``、取消、健康检查）一起冻住。

本文件钉两件事，各配反向变异：
1. 真实执行跑在**工作线程**上，执行期间事件循环仍在推进（心跳计数 > 0）；
2. 取消会**真实传导到执行器**（``cancel_check`` 被观察到），并且抛出
   ``CancelledError`` 之前真实执行已经终止 —— 不留脱离监管的容器。
"""
from __future__ import annotations

import asyncio
import threading
import time
from types import SimpleNamespace

import pytest

from server.app import python_analysis as pa
from server.app.local_results import local_results
from server.app.models import PythonDataScope, PythonDataSource
from server.app.python_coordinator import (
    coordinator,
    execute_python_analysis_async,
)
from server.app.python_datasets import registry

TASK = "turn-async"
ORIGIN = "bridge:excel-office"

_CELLS = [[1, 2], [3, 4]]


@pytest.fixture(autouse=True)
def _clean(monkeypatch):
    coordinator.reset()
    registry._grants.clear()
    registry._budgets.clear()
    local_results.reset()
    monkeypatch.setattr(pa, "python_execution_enabled", lambda: True)
    monkeypatch.setattr(pa, "_runner_status", lambda: "verified")
    yield
    coordinator.reset()
    registry._grants.clear()
    registry._budgets.clear()
    local_results.reset()


def _bind():
    """登记一块可信数据并建立任务绑定（与工具桥登记后的现场同形）。"""
    scope = PythonDataScope(
        version=1,
        taskId=TASK,
        hostId="excel-office",
        workbookIdentityHash="a" * 64,
        authorizationVersion="b" * 64,
        sources=[PythonDataSource(sheetId="S", sheetName="S", range="A1:B2")],
    )
    token = registry.create(scope, ORIGIN)
    receipt = registry.register(token, scope, ORIGIN, 0, _CELLS)
    coordinator.bind(
        task_id=TASK,
        host_id="excel-office",
        workbook_name="数据.xlsx",
        scope=scope,
        scope_token=token,
        origin=ORIGIN,
        receipts=[receipt],
        # 来源下标与源指纹都是登记侧记下来的（桥投递时按块记录）。
        source_indexes={receipt.datasetHandle: 0},
        source_fingerprint="fp-source-1",
    )
    return receipt


def _arguments(handle: str) -> dict:
    return {"code": "result = sum(1 for row in datasets[0])", "datasetHandles": [handle]}


def test_isolated_execution_runs_off_the_event_loop(monkeypatch):
    """执行期间事件循环必须还在推进，且真实执行不在循环线程上。"""
    started = threading.Event()
    release = threading.Event()
    observed: dict[str, int] = {}

    def fake(*, data_registry, scope_token, scope, origin, receipts, code,
             cancel_check=None):
        observed["thread"] = threading.get_ident()
        started.set()
        release.wait(2)  # 卡住真实执行，模拟容器运行
        return {"status": "completed",
                "result": {"kind": "matrix", "values": [[10]]}}

    monkeypatch.setattr(pa, "run_registered_analysis_internal", fake)
    receipt = _bind()
    request = SimpleNamespace(turnId=TASK)

    async def scenario():
        loop_thread = threading.get_ident()
        task = asyncio.create_task(
            execute_python_analysis_async(request, _arguments(receipt.datasetHandle))
        )
        for _ in range(200):
            if started.is_set():
                break
            await asyncio.sleep(0.01)
        ticks = 0
        for _ in range(50):
            if task.done():
                break
            await asyncio.sleep(0)
            ticks += 1
        release.set()
        result = await task
        return loop_thread, ticks, result

    loop_thread, ticks, result = asyncio.run(scenario())

    # 正向锚点：这次执行确实完成了，说明下面的判据不是"压根没执行"。
    assert result["status"] == "completed"
    assert result["resultHandle"].startswith("res_")
    # 判别点 1：真实执行不在事件循环线程上。
    assert observed.get("thread") is not None
    assert observed["thread"] != loop_thread
    # 判别点 2：执行被卡住的这段时间里，事件循环仍在推进。
    assert ticks >= 3


def test_cancel_reaches_the_executor_and_waits_for_termination(monkeypatch):
    """取消必须让真实执行停下来，且要等它停了才抛 CancelledError。"""
    state = {"observed": False, "terminated": False}

    def fake(*, data_registry, scope_token, scope, origin, receipts, code,
             cancel_check=None):
        for _ in range(1000):
            if cancel_check is not None and cancel_check():
                state["observed"] = True
                # 模拟"容器被终止"需要一点时间：卸载层必须等这一步走完。
                time.sleep(0.05)
                state["terminated"] = True
                return {"status": "cancelled",
                        "reasonCategory": "authorization_changed"}
            time.sleep(0.005)
        return {"status": "completed", "result": {"kind": "matrix", "values": [[1]]}}

    monkeypatch.setattr(pa, "run_registered_analysis_internal", fake)
    receipt = _bind()
    request = SimpleNamespace(turnId=TASK)

    async def scenario():
        task = asyncio.create_task(
            execute_python_analysis_async(request, _arguments(receipt.datasetHandle))
        )
        await asyncio.sleep(0.05)
        task.cancel()
        try:
            await task
        except asyncio.CancelledError:
            return "cancelled"
        return "returned"

    assert asyncio.run(scenario()) == "cancelled"
    # 判别点 1：取消信号真的传到了执行器（不是只在 asyncio 层打了个标记）。
    assert state["observed"] is True
    # 判别点 2：抛出取消时真实执行已经终止 —— 没有脱离监管的容器。
    assert state["terminated"] is True
    # 取消当次不发布任何结果。
    assert local_results.count_task(TASK) == 0


def test_async_path_stays_fail_closed_without_runner(monkeypatch):
    """没有已验证执行器时，offload 版本同样失败关闭（不因为换了调度就放行）。"""
    monkeypatch.setattr(pa, "_runner_status", lambda: None)
    receipt = _bind()
    request = SimpleNamespace(turnId=TASK)
    result = asyncio.run(
        execute_python_analysis_async(request, _arguments(receipt.datasetHandle))
    )
    assert result["status"] == "unavailable"
    assert result["reasonCategory"] == "runner_missing"


def test_excel_agent_routes_python_analysis_through_async_path(monkeypatch):
    """普通后端主循环必须走 offload 入口，不能同步调用容器。"""
    from server.app import excel_agent
    from server.app import python_coordinator

    seen: dict[str, object] = {}

    async def fake(request, arguments, *, cancel_check=None, on_event=None):
        seen["arguments"] = arguments
        return {"status": "completed", "resultHandle": "res_offloaded"}

    monkeypatch.setattr(python_coordinator, "execute_python_analysis_async", fake)
    result = asyncio.run(
        excel_agent._execute_tool_async(
            SimpleNamespace(turnId=TASK), "run_python_analysis", {"code": "x = 1"}
        )
    )
    assert result["resultHandle"] == "res_offloaded"
    assert seen.get("arguments") == {"code": "x = 1"}


def test_excel_agent_keeps_other_tools_synchronous(monkeypatch):
    """正向锚点：其它工具仍是同步确定性计算，不该被无谓地丢进线程池。"""
    from server.app import excel_agent

    monkeypatch.setattr(
        excel_agent, "_execute_tool",
        lambda request, name, arguments: {"via": "sync", "name": name},
    )
    result = asyncio.run(
        excel_agent._execute_tool_async(
            SimpleNamespace(turnId=TASK), "read_range", {}
        )
    )
    assert result == {"via": "sync", "name": "read_range"}
