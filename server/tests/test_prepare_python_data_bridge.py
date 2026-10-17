"""``prepare_python_data`` 桥工具：宿主读表 → 值回执 → 数据集登记 → 协调器绑定。

覆盖正向闭环（句柄可被 run_python_analysis 消费）与失败关闭分支（源变化、非法
来源、folder 宿主、形状超限、投影不含值）。隔离执行用合成 mock runner（真实
Docker 由 test_python_docker_real.py 在引擎可用时单独验证）。
"""
from __future__ import annotations

import asyncio
import json

import pytest

from server.app import python_analysis
from server.app.agent_tools import project_prepare_python_data
from server.app.models import BridgeDatasetReceipt, PlanRequest
from server.app.task_bridge import BridgeBinding, TaskToolBridge

SENTINEL = "SENTINEL_SECRET_7B2D"


def _binding(task_id="turn-py-data", **overrides) -> BridgeBinding:
    defaults = dict(
        task_id=task_id,
        host_id="excel-office",
        workbook_name="数据.xlsx",
        source_fingerprint="fp-source-1",
        authorized_key="c" * 64,
        workbook_identity="url:https://example.com/数据.xlsx",
    )
    defaults.update(overrides)
    return BridgeBinding(**defaults)


def _plan_request(**overrides) -> PlanRequest:
    payload = {
        "turnId": "turn-py-data",
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
    return PlanRequest.model_validate(payload)


@pytest.fixture(autouse=True)
def _clean(monkeypatch):
    from server.app.python_coordinator import coordinator
    from server.app.python_datasets import registry
    from server.app.local_results import local_results
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


def _deliver_python(session, operation_id, fp, *, datasets=None, error_code=None,
                    source_fingerprint="fp-source-1", host_id="excel-office"):
    return session.deliver(
        operation_id=operation_id,
        request_fingerprint=fp,
        host_id=host_id,
        workbook_name="数据.xlsx",
        workbook_identity="url:https://example.com/数据.xlsx",
        source_fingerprint=source_fingerprint,
        result=None,
        error_code=error_code,
        datasets=datasets,
    )


def test_prepare_python_data_registers_dataset_and_binds_coordinator():
    from server.app.python_coordinator import coordinator

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
        fp = call["requestFingerprint"]
        _deliver_python(
            session, call["operationId"], fp,
            datasets=[BridgeDatasetReceipt(sourceIndex=0, values=[[1, 2], [3, 4]])],
        )
        return await task

    receipt = asyncio.run(scenario())
    assert receipt["status"] == "ok"
    handles = receipt["datasetHandles"]
    assert len(handles) == 1
    assert handles[0].startswith("dataset://")

    # 协调器已建立绑定，且句柄可从绑定反查。
    binding = coordinator.lookup("turn-py-data")
    assert binding is not None
    assert binding.find_receipt(handles[0]) is not None


def test_prepare_python_data_handle_consumed_by_run_python_analysis(monkeypatch):
    """正向闭环：登记 → run_python_analysis 用句柄 → 结果转本地结果。"""
    from server.app.python_coordinator import coordinator
    from server.app import python_coordinator as coord
    from types import SimpleNamespace

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
        _deliver_python(
            session, call["operationId"], call["requestFingerprint"],
            datasets=[BridgeDatasetReceipt(sourceIndex=0, values=[[1, 2], [3, 4]])],
        )
        return await task

    receipt = asyncio.run(scenario())
    handle = receipt["datasetHandles"][0]

    def fake_runner(data_registry, scope_token, scope, origin, receipts, code,
                    cancel_check=None):
        values = data_registry.resolve(scope_token, scope, origin, receipts)
        total = sum(cell for d in values["datasets"] for row in d["values"] for cell in row)
        return {"status": "completed", "result": {"kind": "matrix", "values": [[total]]}}

    monkeypatch.setattr(python_analysis, "run_registered_analysis_internal", fake_runner)
    result = coord.execute_python_analysis(
        SimpleNamespace(turnId="turn-py-data"),
        {"code": "result = sum(...)", "datasetHandles": [handle]},
    )
    assert result["status"] == "completed"
    assert result["resultHandle"].startswith("res_")

    from server.app.local_results import local_results
    item = local_results.resolve("turn-py-data", result["resultHandle"])
    assert item.rows == [[10]]


def test_prepare_python_data_rejects_source_changed():
    bridge = TaskToolBridge()
    session = bridge.open(_binding(source_fingerprint="fp-source-1"))

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
        _deliver_python(
            session, call["operationId"], call["requestFingerprint"],
            datasets=[BridgeDatasetReceipt(sourceIndex=0, values=[[1]])],
            source_fingerprint="fp-changed",
        )
        return await task

    receipt = asyncio.run(scenario())
    assert receipt == {"status": "error", "errorCode": "source_changed"}


def test_prepare_python_data_rejects_folder_host():
    bridge = TaskToolBridge()
    session = bridge.open(_binding(host_id="folder"))

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
        _deliver_python(
            session, call["operationId"], call["requestFingerprint"],
            datasets=[BridgeDatasetReceipt(sourceIndex=0, values=[[1]])],
            host_id="folder",
        )
        return await task

    receipt = asyncio.run(scenario())
    assert receipt["status"] == "error"
    assert receipt["errorCode"] == "invalid_arguments"


def test_prepare_python_data_rejects_empty_sources():
    bridge = TaskToolBridge()
    session = bridge.open(_binding())

    async def scenario():
        task = asyncio.create_task(
            session.request("prepare_python_data", {"sources": []})
        )
        for _ in range(200):
            if session.pending_calls:
                break
            await asyncio.sleep(0)
        call = session.pending_calls[0]
        _deliver_python(
            session, call["operationId"], call["requestFingerprint"],
            datasets=[BridgeDatasetReceipt(sourceIndex=0, values=[[1]])],
        )
        return await task

    receipt = asyncio.run(scenario())
    assert receipt["status"] == "error"
    assert receipt["errorCode"] == "invalid_arguments"


def test_prepare_python_data_rejects_missing_datasets():
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
        _deliver_python(session, call["operationId"], call["requestFingerprint"])
        return await task

    receipt = asyncio.run(scenario())
    assert receipt["status"] == "error"
    assert receipt["errorCode"] == "internal"


def test_prepare_python_data_rejects_shape_mismatch():
    """值矩阵与声明的 range 尺寸不符 → 登记器拒绝，映射 result_too_large。"""
    bridge = TaskToolBridge()
    session = bridge.open(_binding())

    async def scenario():
        task = asyncio.create_task(
            session.request("prepare_python_data", {
                "sources": [{"sheet": "Sheet1", "range": "A1:B2"}],  # 2x2
            })
        )
        for _ in range(200):
            if session.pending_calls:
                break
            await asyncio.sleep(0)
        call = session.pending_calls[0]
        # 只回 1 行，与 2 行不符。
        _deliver_python(
            session, call["operationId"], call["requestFingerprint"],
            datasets=[BridgeDatasetReceipt(sourceIndex=0, values=[[1, 2]])],
        )
        return await task

    receipt = asyncio.run(scenario())
    assert receipt["status"] == "error"
    assert receipt["errorCode"] == "result_too_large"


def _prepare_receipt(session, sources, datasets):
    """发起一次 ``prepare_python_data`` 并用给定 datasets 回执收口。"""
    async def scenario():
        task = asyncio.create_task(
            session.request("prepare_python_data", {"sources": sources})
        )
        for _ in range(200):
            if session.pending_calls:
                break
            await asyncio.sleep(0)
        call = session.pending_calls[0]
        _deliver_python(
            session, call["operationId"], call["requestFingerprint"],
            datasets=datasets,
        )
        return await task

    return asyncio.run(scenario())


def _block(index, values):
    return BridgeDatasetReceipt(sourceIndex=index, values=values)


TWO_SOURCES = [{"sheet": "Sheet1", "range": "A1:B2"},
               {"sheet": "Sheet2", "range": "A1:B2"}]
_CELLS_2X2 = [[1, 2], [3, 4]]


def test_prepare_python_data_accepts_complete_two_source_indexes():
    """正向锚点：声明两块并交回 0/1 两块 → 两个句柄、按声明顺序排。

    没有这条，下面那些 ``invalid_arguments`` 无法区分"闸门拦的"和"多来源本来就
    不可能成功"。
    """
    from server.app.python_coordinator import coordinator

    bridge = TaskToolBridge()
    session = bridge.open(_binding())
    receipt = _prepare_receipt(session, TWO_SOURCES,
                               [_block(0, _CELLS_2X2), _block(1, [[5, 6], [7, 8]])])
    assert receipt["status"] == "ok"
    handles = receipt["datasetHandles"]
    assert len(handles) == 2
    binding = coordinator.lookup("turn-py-data")
    assert binding.find_receipt(handles[0]) is not None
    assert binding.find_receipt(handles[1]) is not None


@pytest.mark.parametrize(
    "sources, datasets, label",
    [
        (TWO_SOURCES, [_block(0, _CELLS_2X2)], "少交一块"),
        (TWO_SOURCES, [_block(0, _CELLS_2X2), _block(0, _CELLS_2X2)], "sourceIndex 重复"),
        (TWO_SOURCES, [_block(0, _CELLS_2X2), _block(2, _CELLS_2X2)], "sourceIndex 越界"),
        (TWO_SOURCES,
         [_block(0, _CELLS_2X2), _block(1, _CELLS_2X2), _block(1, _CELLS_2X2)],
         "块数多于声明"),
        ([{"sheet": "Sheet1", "range": "A1:B2"}], [_block(1, _CELLS_2X2)],
         "单来源却交 index=1"),
    ],
)
def test_prepare_python_data_rejects_incomplete_source_indexes(sources, datasets, label):
    """声明 N 块来源，就必须交回 sourceIndex 完整且互不重复的一集合。

    少一块就放行 = 隔离 Python 在"少读了一张表"的数据上算出一份看起来成功的错
    答案；重复/越界会让同一份值登记两次或挂到错误的来源上。
    """
    from server.app.python_datasets import registry

    bridge = TaskToolBridge()
    session = bridge.open(_binding())
    receipt = _prepare_receipt(session, sources, datasets)
    assert receipt == {"status": "error", "errorCode": "invalid_arguments"}, label
    # 完整性闸在 create 之前 → 一格都没登记进去，不是"登记一半"。
    assert registry._grants == {}


def test_prepare_python_data_discards_partial_registration_on_failure():
    """第二块形状不符 → 整批拒绝，且第一块已经登记进去的授权必须当场清理。"""
    from server.app.python_datasets import registry

    bridge = TaskToolBridge()
    session = bridge.open(_binding())
    receipt = _prepare_receipt(
        session, TWO_SOURCES,
        [_block(0, _CELLS_2X2), _block(1, [[1, 2]])],  # 第二块只有 1 行
    )
    assert receipt == {"status": "error", "errorCode": "result_too_large"}
    assert registry._grants == {}


def test_prepare_python_data_projection_excludes_values():
    receipt = {
        "status": "ok",
        "datasetHandles": ["dataset://" + "d" * 43],
        # 这些字段即使同源回执里有也绝不能出站。
        "values": [[SENTINEL]],
        "fingerprint": SENTINEL,
    }
    projected = project_prepare_python_data(receipt)
    encoded = json.dumps(projected, ensure_ascii=False)
    assert SENTINEL not in encoded
    assert set(projected) == {"tool", "status", "datasetHandles"}
    assert projected["datasetHandles"] == ["dataset://" + "d" * 43]


def test_prepare_python_data_projection_maps_unknown_error():
    projected = project_prepare_python_data(
        {"status": "error", "errorCode": f"boom-{SENTINEL}"}
    )
    assert projected == {
        "tool": "prepare_python_data", "status": "error", "errorCode": "internal",
    }
