"""取消 / 重授权 / 换簿 / 轮次结束 → Python 协调器与数据集登记的**失效链**。

为什么单独一份：此前桥的取消只清"在飞调用 + 本地结果"，Python 侧的
``coordinator`` 绑定与 ``registry`` 数据登记**没有任何一条失效路径** ——
取消之后跑完的那次隔离执行照样能把迟到结果登记出去，换簿之后旧的
``dataset://`` 句柄照样指向另一本工作簿的值。等 TTL 过期是**时间窗问题**，
不是失效。

本文件用的是合成执行器（真实 Docker 由 `test_python_docker_real.py` 在引擎
可用时单独验证）；失效链与执行器无关，是授权语义。
"""
from __future__ import annotations

import asyncio
from types import SimpleNamespace

import pytest

from server.app import python_analysis
from server.app import python_coordinator as coord
from server.app.local_results import local_results
from server.app.models import BridgeDatasetReceipt
from server.app.python_coordinator import coordinator
from server.app.python_datasets import DatasetRejected, registry
from server.app.task_bridge import BridgeBinding, TaskToolBridge, drop_python_state

TASK = "turn-py-cancel"


def _binding(**overrides) -> BridgeBinding:
    defaults = dict(
        task_id=TASK,
        host_id="excel-office",
        workbook_name="数据.xlsx",
        source_fingerprint="fp-source-1",
        authorized_key="c" * 64,
        workbook_identity="url:https://example.com/数据.xlsx",
    )
    defaults.update(overrides)
    return BridgeBinding(**defaults)


@pytest.fixture(autouse=True)
def _clean(monkeypatch):
    from server.app.task_bridge import task_tool_bridge

    task_tool_bridge.reset()
    coordinator.reset()
    registry._grants.clear()
    registry._budgets.clear()
    local_results.reset()
    monkeypatch.setattr(python_analysis, "python_execution_enabled", lambda: True)
    monkeypatch.setattr(python_analysis, "_runner_status", lambda: "verified")
    yield
    task_tool_bridge.reset()
    coordinator.reset()
    registry._grants.clear()
    registry._budgets.clear()
    local_results.reset()


def _register_dataset(session):
    """走真实桥登记一份数据，返回 ``dataset://`` 句柄。"""
    async def scenario():
        task = asyncio.create_task(
            session.request("prepare_python_data", {
                "sources": [{"sheet": "Sheet1", "range": "A1:B2"}],
            })
        )
        for _ in range(200):
            if session.pending_calls:
                break
            await asyncio.sleep(0)
        call = session.pending_calls[0]
        session.deliver(
            operation_id=call["operationId"],
            request_fingerprint=call["requestFingerprint"],
            host_id="excel-office",
            workbook_name="数据.xlsx",
            workbook_identity="url:https://example.com/数据.xlsx",
            source_fingerprint="fp-source-1",
            result=None,
            error_code=None,
            datasets=[BridgeDatasetReceipt(sourceIndex=0, values=[[1, 2], [3, 4]])],
        )
        return await task

    receipt = asyncio.run(scenario())
    assert receipt["status"] == "ok"
    return receipt["datasetHandles"][0]


def test_cancel_drops_python_binding_and_datasets():
    bridge = TaskToolBridge()
    session = bridge.open(_binding())
    handle = _register_dataset(session)
    assert coordinator.lookup(TASK) is not None

    bridge.cancel(TASK, "user_cancelled")

    # 绑定没了：run_python_analysis 连反查指纹的入口都没有。
    assert coordinator.lookup(TASK) is None
    # 数据没了：句柄即使被模型留着也解不出值。
    assert _grants_for_task() == 0

    result = coord.execute_python_analysis(
        SimpleNamespace(turnId=TASK),
        {"code": "result = 1", "datasetHandles": [handle]},
    )
    assert result["status"] == "rejected"


def test_reauthorization_invalidates_python_state():
    bridge = TaskToolBridge()
    bridge.open(_binding())
    session = bridge.open(_binding())  # 同一任务、同一授权
    _register_dataset(session)
    assert coordinator.lookup(TASK) is not None

    # 用户重新勾选范围 → 授权键变化 → 旧数据登记必须当场作废。
    bridge.open(_binding(authorized_key="d" * 64))

    assert coordinator.lookup(TASK) is None
    assert _grants_for_task() == 0


def test_workbook_switch_invalidates_python_state():
    bridge = TaskToolBridge()
    session = bridge.open(_binding())
    handle = _register_dataset(session)

    bridge.open(_binding(workbook_name="另一本.xlsx",
                         workbook_identity="url:https://example.com/另一本.xlsx"))

    assert coordinator.lookup(TASK) is None
    assert _grants_for_task() == 0
    # 旧句柄指向的是**另一本**工作簿的值，绝不能还能解出数据。
    assert registry.valid(_token_of(handle), _scope_of(handle), "bridge:excel-office",
                          []) is False


def test_close_at_turn_end_drops_python_state():
    bridge = TaskToolBridge()
    session = bridge.open(_binding())
    _register_dataset(session)

    bridge.close(TASK)  # 轮次正常收尾

    assert coordinator.lookup(TASK) is None
    assert _grants_for_task() == 0


def test_late_receipt_after_cancel_is_not_registered():
    """取消之后才到的宿主回执：不登记、不产生句柄。"""
    bridge = TaskToolBridge()
    session = bridge.open(_binding())

    async def scenario():
        task = asyncio.create_task(
            session.request("prepare_python_data", {
                "sources": [{"sheet": "Sheet1", "range": "A1:B2"}],
            })
        )
        for _ in range(200):
            if session.pending_calls:
                break
            await asyncio.sleep(0)
        call = session.pending_calls[0]
        bridge.cancel(TASK, "user_cancelled")
        with pytest.raises(Exception):
            session.deliver(
                operation_id=call["operationId"],
                request_fingerprint=call["requestFingerprint"],
                host_id="excel-office",
                workbook_name="数据.xlsx",
                workbook_identity="url:https://example.com/数据.xlsx",
                source_fingerprint="fp-source-1",
                result=None,
                error_code=None,
                datasets=[BridgeDatasetReceipt(sourceIndex=0, values=[[1, 2], [3, 4]])],
            )
        task.cancel()
        return call

    scenario_result = asyncio.run(scenario())
    assert _grants_for_task() == 0
    assert coordinator.lookup(TASK) is None


def test_cancel_during_execution_publishes_nothing(monkeypatch):
    """执行**期间**取消：runner 就算跑完并返回结果，也不得登记、不得给句柄。"""
    bridge = TaskToolBridge()
    session = bridge.open(_binding())
    handle = _register_dataset(session)

    def isolated_that_finishes_after_cancel(
        code, values, cancel_check=None, on_runner_start=None
    ):
        # 容器真的启动过：这条用例要证的是"跑完也不许登记"，不是"没开始"。
        if on_runner_start:
            on_runner_start()
        # 模拟"容器还在跑，用户取消了"：取消发生在结果返回之前。
        drop_python_state(TASK)
        return {"status": "completed", "result": {"kind": "matrix", "values": [[999]]}}

    monkeypatch.setattr(
        python_analysis, "run_isolated_analysis_internal", isolated_that_finishes_after_cancel
    )

    result = coord.execute_python_analysis(
        SimpleNamespace(turnId=TASK),
        {"code": "result = 999", "datasetHandles": [handle]},
    )
    assert result["status"] != "completed"
    assert "resultHandle" not in result
    # 没有任何本地结果被登记 → 也就不可能产生写计划。
    assert local_results.resolve(TASK, "res_anything") is None


def test_drop_python_state_is_task_scoped():
    """作废只针对该任务，别的任务的数据不受影响（否则会误伤并发轮次）。"""
    other = "turn-other"
    bridge = TaskToolBridge()
    session = bridge.open(_binding())
    _register_dataset(session)

    other_session = bridge.open(_binding(task_id=other))
    _register_dataset(other_session)  # 另一条任务也真的登记了数据
    coordinator.bind(
        task_id=other, host_id="excel-office", workbook_name="数据.xlsx",
        scope=_scope_of("x"), scope_token="t", origin="bridge:excel-office", receipts=[],
    )
    assert _grants_for(other) == 1

    bridge.cancel(TASK, "user_cancelled")
    assert coordinator.lookup(TASK) is None
    assert coordinator.lookup(other) is not None
    # 别的任务的**数据**也不能被顺手删掉（过度作废同样是缺陷）。
    assert _grants_for(other) == 1


def _grants_for_task() -> int:
    return _grants_for(TASK)


def _grants_for(task_id: str) -> int:
    return sum(1 for g in registry._grants.values() if g.scope.get("taskId") == task_id)


def _token_of(handle: str) -> str:
    """按句柄反查它所属的 grant token（只为断言"这个 grant 还在不在"）。"""
    for token, grant in registry._grants.items():
        if handle in grant.datasets:
            return token
    return ""


def _scope_of(_handle):
    from server.app.models import PythonDataScope, PythonDataSource

    return PythonDataScope(
        version=1,
        taskId=TASK,
        hostId="excel-office",
        workbookIdentityHash="0" * 64,
        authorizationVersion="c" * 64,
        sources=[PythonDataSource(sheetId="s1", sheetName="Sheet1", range="A1:B2")],
    )
