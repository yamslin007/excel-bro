"""可信协调器（python_coordinator）的正向接线与失败关闭测试。

覆盖：数据登记 → 任务绑定 → run_python_analysis 隔离执行 → 结果登记为本地
结果 → resultRef 可被 local_results / result_plan 消费；以及句柄不可信、
授权变化、非矩阵结果、取消竞态等失败关闭分支。

隔离执行用合成 mock runner（真实 Docker 执行由 test_python_docker_real.py
在引擎可用时单独验证，这里不 mock 掉协调器逻辑本身）。
"""

from __future__ import annotations

from types import SimpleNamespace

import pytest

from server.app import python_analysis
from server.app import python_coordinator as coord
from server.app.local_results import local_results
from server.app.models import PythonDataScope
from server.app.python_datasets import registry
from server.app.task_bridge import drop_python_state

ORIGIN = "https://localhost:3000"


def _scope(task_id="task-py", host_id="excel-office", sources=None):
    sources = sources or [dict(sheetId="s1", sheetName="Sheet1", range="A1:B2")]
    return PythonDataScope.model_validate(dict(
        version=1, taskId=task_id, hostId=host_id,
        workbookIdentityHash="a" * 64, authorizationVersion="b" * 64,
        sources=sources,
    ))


def _bind(scope=None, task_id="task-py", host_id="excel-office",
          workbook_name="w.xlsx", values=None, source_fingerprint=None):
    scope = scope or _scope(task_id=task_id, host_id=host_id)
    token = registry.create(scope, ORIGIN)
    values = values if values is not None else [[1, 2], [3, 4]]
    receipt = registry.register(token, scope, ORIGIN, 0, values)
    coord.coordinator.bind(
        task_id=task_id, host_id=host_id, workbook_name=workbook_name,
        scope=scope, scope_token=token, origin=ORIGIN, receipts=[receipt],
        # 来源下标由登记侧记录（桥投递时按块记录）；这里模拟同一现场。
        source_indexes={receipt.datasetHandle: 0},
        source_fingerprint=source_fingerprint,
    )
    return scope, receipt


def _bind_two_sources(source_fingerprint=None):
    """登记**两张**来源表，返回 (scope, receipt0, receipt1)。"""
    scope = _scope(sources=[
        dict(sheetId="s1", sheetName="Sheet1", range="A1:B2"),
        dict(sheetId="s2", sheetName="Sheet2", range="A1:B2"),
    ])
    token = registry.create(scope, ORIGIN)
    r0 = registry.register(token, scope, ORIGIN, 0, [[1, 2], [3, 4]])
    r1 = registry.register(token, scope, ORIGIN, 1, [[5, 6], [7, 8]])
    coord.coordinator.bind(
        task_id="task-py", host_id="excel-office", workbook_name="w.xlsx",
        scope=scope, scope_token=token, origin=ORIGIN, receipts=[r1, r0],  # 故意乱序
        source_indexes={r0.datasetHandle: 0, r1.datasetHandle: 1},
        source_fingerprint=source_fingerprint,
    )
    return scope, r0, r1


@pytest.fixture(autouse=True)
def _clean(monkeypatch):
    coord.coordinator.reset()
    registry._grants.clear()
    registry._budgets.clear()
    local_results.reset()
    monkeypatch.setattr(python_analysis, "python_execution_enabled", lambda: True)
    monkeypatch.setattr(python_analysis, "_runner_status", lambda: "verified")
    yield
    coord.coordinator.reset()
    registry._grants.clear()
    registry._budgets.clear()
    local_results.reset()


def _fake_runner_summing():
    def runner(data_registry, scope_token, scope, origin, receipts, code,
               cancel_check=None, on_runner_start=None):
        if on_runner_start:
            on_runner_start()
        # 模拟 run_registered_analysis_internal：解析数据 → 求和 → 返回矩阵结果。
        values = data_registry.resolve(scope_token, scope, origin, receipts)
        total = sum(
            cell
            for dataset in values["datasets"]
            for row in dataset["values"]
            for cell in row
        )
        return {"status": "completed", "result": {"kind": "matrix", "values": [[total]]}}
    return runner


def test_successful_run_registers_local_result(monkeypatch):
    scope, receipt = _bind()
    monkeypatch.setattr(python_analysis, "run_registered_analysis_internal", _fake_runner_summing())

    result = coord.execute_python_analysis(
        SimpleNamespace(turnId="task-py"),
        {"code": "result = ...", "datasetHandles": [receipt.datasetHandle]},
    )
    assert result["status"] == "completed"
    assert result["reasonCategory"] == "none"
    assert result["resultHandle"].startswith("res_")

    # 投影只暴露句柄，不含值。
    projected = python_analysis.project_python_analysis_receipt(result)
    assert projected["status"] == "completed"
    assert projected["resultHandle"] == result["resultHandle"]
    assert "result" not in projected and "values" not in projected

    # 结果能被本地结果登记器解析出真实值。
    item = local_results.resolve("task-py", result["resultHandle"])
    assert item is not None
    assert item.rows == [[10]]  # 1+2+3+4
    assert item.source_sheets == ["Sheet1"]


def test_result_handle_compiles_to_plan(monkeypatch):
    scope, receipt = _bind()
    monkeypatch.setattr(python_analysis, "run_registered_analysis_internal", _fake_runner_summing())
    result = coord.execute_python_analysis(
        SimpleNamespace(turnId="task-py"),
        {"code": "result = ...", "datasetHandles": [receipt.datasetHandle]},
    )

    from server.app.result_plan import compile_result_plan
    from server.app.models import PlanRequest

    item = local_results.resolve("task-py", result["resultHandle"])
    request = PlanRequest.model_validate({
        "prompt": "写入结果",
        "workbook": {
            "name": "w.xlsx",
            "capturedAt": "2026-10-09T00:00:00Z",
            "activeWorksheet": "Sheet1",
            "worksheets": [{"name": "Sheet1", "usedRange": "A1:B2",
                            "rowCount": 2, "columnCount": 2, "headers": ["a", "b"]}],
        },
    })
    plan = compile_result_plan(item, request, target_sheet="Sheet1", target_range="D1:D2")
    assert plan.actions[0].type == "writeValues"
    assert plan.actions[0].sheet == "Sheet1"
    # include_header=True 默认把表头「列1」写进第 1 行，数据写进第 2 行。
    assert plan.actions[0].values == [["列1"], [10]]
    assert plan.localResultRef == result["resultHandle"]


def test_unregistered_handle_is_rejected(monkeypatch):
    _bind()
    result = coord.execute_python_analysis(
        SimpleNamespace(turnId="task-py"),
        {"code": "result = 1", "datasetHandles": ["dataset://" + "z" * 43]},
    )
    assert result["status"] == "rejected"
    assert result["reasonCategory"] == "invalid_request"


def test_no_binding_is_rejected(monkeypatch):
    result = coord.execute_python_analysis(
        SimpleNamespace(turnId="task-unknown"),
        {"code": "result = 1", "datasetHandles": []},
    )
    assert result["status"] == "rejected"


def test_missing_task_id_is_rejected(monkeypatch):
    result = coord.execute_python_analysis(
        SimpleNamespace(turnId=None),
        {"code": "result = 1", "datasetHandles": []},
    )
    assert result["status"] == "rejected"


def test_non_matrix_result_fails_closed(monkeypatch):
    scope, receipt = _bind()
    monkeypatch.setattr(python_analysis, "run_registered_analysis_internal",
                        lambda **kw: {"status": "completed", "result": {"scalar": 42}})
    result = coord.execute_python_analysis(
        SimpleNamespace(turnId="task-py"),
        {"code": "result = 42", "datasetHandles": [receipt.datasetHandle]},
    )
    assert result["status"] == "unavailable"
    assert result["reasonCategory"] == "execution_error"


def test_runner_failure_maps_to_unavailable(monkeypatch):
    scope, receipt = _bind()
    monkeypatch.setattr(python_analysis, "run_registered_analysis_internal",
                        lambda **kw: {"status": "unavailable", "reasonCategory": "runner_missing"})
    result = coord.execute_python_analysis(
        SimpleNamespace(turnId="task-py"),
        {"code": "result = 1", "datasetHandles": [receipt.datasetHandle]},
    )
    assert result["status"] == "unavailable"
    assert result["reasonCategory"] == "runner_missing"


def test_cancelled_run_fails_closed_without_late_result(monkeypatch):
    scope, receipt = _bind()
    monkeypatch.setattr(python_analysis, "run_registered_analysis_internal",
                        lambda **kw: {"status": "cancelled", "reasonCategory": "authorization_changed"})
    result = coord.execute_python_analysis(
        SimpleNamespace(turnId="task-py"),
        {"code": "result = 1", "datasetHandles": [receipt.datasetHandle]},
    )
    assert result["status"] == "cancelled"
    assert result["reasonCategory"] == "authorization_changed"
    assert "resultHandle" not in result
    assert python_analysis.project_python_analysis_receipt(result)["status"] == "cancelled"
    assert local_results.count_task("task-py") == 0


def test_disabled_switch_never_runs(monkeypatch):
    monkeypatch.setattr(python_analysis, "python_execution_enabled", lambda: False)
    called = []
    monkeypatch.setattr(python_analysis, "run_registered_analysis_internal",
                        lambda **kw: called.append(1) or {"status": "completed", "result": {}})
    result = coord.execute_python_analysis(
        SimpleNamespace(turnId="task-py"), {"code": "result = 1", "datasetHandles": []},
    )
    assert result["status"] == "unavailable"
    assert result["reasonCategory"] == "disabled"
    assert not called


def test_runner_missing_fails_closed(monkeypatch):
    _bind()
    monkeypatch.setattr(python_analysis, "_runner_status", lambda: None)
    result = coord.execute_python_analysis(
        SimpleNamespace(turnId="task-py"), {"code": "result = 1", "datasetHandles": []},
    )
    assert result["status"] == "unavailable"
    assert result["reasonCategory"] == "runner_missing"


def test_result_source_sheets_only_cover_handles_actually_used(monkeypatch):
    """实际来源依赖：只算**真正交进来**的那批句柄，不虚报成登记的全部来源。"""
    _scope_obj, r0, r1 = _bind_two_sources()
    monkeypatch.setattr(python_analysis, "run_registered_analysis_internal",
                        _fake_runner_summing())

    result = coord.execute_python_analysis(
        SimpleNamespace(turnId="task-py"),
        {"code": "result = ...", "datasetHandles": [r1.datasetHandle]},
    )
    assert result["status"] == "completed"
    item = local_results.resolve("task-py", result["resultHandle"])
    assert item is not None
    # 登记了两张表，但这次只用了 Sheet2 —— 来源依赖就只是 Sheet2。
    assert item.source_sheets == ["Sheet2"]


def test_result_source_sheets_are_not_affected_by_delivery_order(monkeypatch):
    """回执块乱序交回时，来源依赖仍按**来源下标**认定（不用迭代顺序）。"""
    _scope_obj, r0, r1 = _bind_two_sources()
    monkeypatch.setattr(python_analysis, "run_registered_analysis_internal",
                        _fake_runner_summing())

    result = coord.execute_python_analysis(
        SimpleNamespace(turnId="task-py"),
        {"code": "result = ...", "datasetHandles": [r0.datasetHandle, r1.datasetHandle]},
    )
    assert result["status"] == "completed"
    item = local_results.resolve("task-py", result["resultHandle"])
    assert item.source_sheets == ["Sheet1", "Sheet2"]


def test_missing_source_index_fails_closed(monkeypatch):
    """查不到句柄对应哪张来源 → 失败关闭，不猜、不用全体来源顶替。"""
    scope = _scope(sources=[
        dict(sheetId="s1", sheetName="Sheet1", range="A1:B2"),
        dict(sheetId="s2", sheetName="Sheet2", range="A1:B2"),
    ])
    token = registry.create(scope, ORIGIN)
    receipt = registry.register(token, scope, ORIGIN, 0, [[1, 2], [3, 4]])
    coord.coordinator.bind(
        task_id="task-py", host_id="excel-office", workbook_name="w.xlsx",
        scope=scope, scope_token=token, origin=ORIGIN, receipts=[receipt],
        source_indexes={},  # 登记侧没记下标 —— 服务端自己漏了
    )
    monkeypatch.setattr(python_analysis, "run_registered_analysis_internal",
                        _fake_runner_summing())
    result = coord.execute_python_analysis(
        SimpleNamespace(turnId="task-py"),
        {"code": "result = ...", "datasetHandles": [receipt.datasetHandle]},
    )
    assert result["status"] == "unavailable"
    assert local_results.count_task("task-py") == 0


def test_result_carries_the_source_fingerprint_it_was_computed_from(monkeypatch):
    """结果要带上它据以算出的源快照指纹，否则源一变无法判断结果是否过期。"""
    _scope_obj, receipt = _bind(source_fingerprint="fp-source-1")
    monkeypatch.setattr(python_analysis, "run_registered_analysis_internal",
                        _fake_runner_summing())

    result = coord.execute_python_analysis(
        SimpleNamespace(turnId="task-py"),
        {"code": "result = ...", "datasetHandles": [receipt.datasetHandle]},
    )
    assert result["status"] == "completed"
    item = local_results.resolve("task-py", result["resultHandle"])
    assert item.source_fingerprint == "fp-source-1"
    # 源指纹一致 → 不该被判为过期源。
    assert local_results.invalidate_source("task-py", "fp-source-1") == 0
    assert local_results.resolve("task-py", result["resultHandle"]) is not None
    # 源换了 → 这份基于旧源的结果立刻失效。
    assert local_results.invalidate_source("task-py", "fp-source-2") == 1
    assert local_results.resolve("task-py", result["resultHandle"]) is None


def test_reauthorization_invalidates_binding(monkeypatch):
    """换授权（同任务重绑）后旧句柄不可再解析。"""
    scope, receipt = _bind()
    # 重绑同一 task：模拟换授权
    new_scope = _scope(task_id="task-py", sources=[dict(sheetId="s2", sheetName="Sheet2", range="A1:B2")])
    token = registry.create(new_scope, ORIGIN)
    new_receipt = registry.register(token, new_scope, ORIGIN, 0, [[9, 9], [9, 9]])
    coord.coordinator.bind(task_id="task-py", host_id="excel-office", workbook_name="w.xlsx",
                           scope=new_scope, scope_token=token, origin=ORIGIN, receipts=[new_receipt])
    # 旧句柄现在不在 binding 里
    result = coord.execute_python_analysis(
        SimpleNamespace(turnId="task-py"),
        {"code": "result = 1", "datasetHandles": [receipt.datasetHandle]},
    )
    assert result["status"] == "rejected"


def test_multiple_analyses_in_one_turn_each_carry_their_own_sources(monkeypatch):
    """同轮多次分析：每次结果只带它**那一次**实际用到的来源，不互相污染。"""
    _scope_obj, r0, r1 = _bind_two_sources()
    monkeypatch.setattr(python_analysis, "run_registered_analysis_internal",
                        _fake_runner_summing())

    first = coord.execute_python_analysis(
        SimpleNamespace(turnId="task-py"),
        {"code": "result = ...", "datasetHandles": [r0.datasetHandle]},
    )
    second = coord.execute_python_analysis(
        SimpleNamespace(turnId="task-py"),
        {"code": "result = ...", "datasetHandles": [r1.datasetHandle]},
    )
    assert first["status"] == "completed" and second["status"] == "completed"
    assert first["resultHandle"] != second["resultHandle"]
    first_item = local_results.resolve("task-py", first["resultHandle"])
    second_item = local_results.resolve("task-py", second["resultHandle"])
    # 两次结果各记各的来源，不因为同轮而合并成 ["Sheet1","Sheet2"]。
    assert first_item.source_sheets == ["Sheet1"]
    assert second_item.source_sheets == ["Sheet2"]
    assert local_results.count_task("task-py") == 2


def test_unregistered_handle_never_produces_a_result(monkeypatch):
    """旧结果引用：句柄不在本轮绑定里 → 拒绝，且不登记任何新结果。"""
    _scope_obj, r0, r1 = _bind_two_sources()
    monkeypatch.setattr(python_analysis, "run_registered_analysis_internal",
                        _fake_runner_summing())

    result = coord.execute_python_analysis(
        SimpleNamespace(turnId="task-py"),
        {"code": "result = ...", "datasetHandles": ["dataset://" + "z" * 43]},
    )
    assert result["status"] == "rejected"
    assert local_results.count_task("task-py") == 0


def test_source_change_makes_the_registered_result_unresolvable(monkeypatch):
    """源变化：基于旧源算出的结果不得再被读出（不是靠 TTL 兜底）。"""
    _scope_obj, receipt = _bind(source_fingerprint="fp-source-1")
    monkeypatch.setattr(python_analysis, "run_registered_analysis_internal",
                        _fake_runner_summing())
    result = coord.execute_python_analysis(
        SimpleNamespace(turnId="task-py"),
        {"code": "result = ...", "datasetHandles": [receipt.datasetHandle]},
    )
    assert result["status"] == "completed"
    handle = result["resultHandle"]
    assert local_results.resolve("task-py", handle) is not None

    # 源变化 → 与会话失效链同一处收口（drop_python_state）。
    drop_python_state("task-py")
    assert local_results.resolve("task-py", handle) is None
    assert local_results.count_task("task-py") == 0
