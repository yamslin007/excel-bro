"""A1-E1：结果引用过期的**精确故障入口**（开发验收专用）。

被测对象是生产代码：真实 `LocalResultRegistry`、真实 `task_tool_bridge` 绑定、
真实 `resolve_result_ref`（两个后端共用的那条漏斗）。只有"用户点了开发入口"这一步
由测试直接调用模块函数代替（等价于面板按钮背后的那次 arm 请求）。

覆盖：
* 单次触发；
* 任务隔离（同簿不同任务、不同簿）；
* 非开发模式不可用（默认关闭、接口 404、内部调用也无效）；
* 旧引用被拒且不可复活；
* 零宿主写入（没有任何计划被编译出来 —— 计划是宿主写入的唯一入口）；
* 普通路径不受影响（没启用时一切照旧）；
* 入口不接受任务/引用参数（模型不能指定别的任务或别的引用）。
"""

from __future__ import annotations

import pytest

from server.app import (
    dev_fault_target,
    dev_result_ref_fault,
    local_results as registry_module,
)
from server.app.models import DataToolResult, DevResultRefFaultArmRequest, PlanRequest
from server.app.result_submission import (
    RESULT_REF_INVALID_CODE,
    ResultSubmissionError,
    build_result_plan_response,
    resolve_result_ref,
)
from server.app.task_bridge import BridgeBinding, task_tool_bridge

WORKBOOK = "测试.xlsx"
#: 可核验身份（`name:` 回退值不再被接受：它证明不了是哪一本）。
IDENTITY = "url:https://example.com/测试.xlsx"


def _result(title: str = "Python 分析结果") -> DataToolResult:
    return DataToolResult.model_validate({
        "requestId": "req-1",
        "tool": "query_table",
        "title": title,
        "headers": ["类别", "数量"],
        "rows": [["甲", 2], ["乙", 1]],
        "sourceSheets": ["Sheet1"],
        "scannedRows": 12,
        "complete": True,
        "calculation": "按类别计数",
    })


def _binding(task_id: str, *, workbook: str = WORKBOOK, identity: str = IDENTITY) -> BridgeBinding:
    return BridgeBinding(
        task_id=task_id,
        host_id="excel-office",
        workbook_name=workbook,
        source_fingerprint="fp-source-1",
        authorized_key="a" * 64,
        workbook_identity=identity,
    )


def _plan_request(task_id: str) -> PlanRequest:
    return PlanRequest.model_validate({
        "turnId": task_id,
        "prompt": "把 Python 分析结果写入 Sheet1",
        "planningReadPolicy": "none",
        "hostId": "excel-office",
        "hostVersion": "16.0",
        "workbook": {
            "name": WORKBOOK,
            "capturedAt": "2026-10-25T00:00:00Z",
            "activeWorksheet": "Sheet1",
            "sourceFingerprint": "fp-source-1",
            "worksheets": [{
                "name": "Sheet1", "usedRange": "A1:B3", "rowCount": 3,
                "columnCount": 2, "headers": ["类别", "数量"],
            }],
        },
    })


def _open_task(task_id: str, *, workbook: str = WORKBOOK, identity: str = IDENTITY) -> None:
    """走真实的工具桥开桥路径：任务绑定 = 一次规划任务真正开始。"""
    task_tool_bridge.open(_binding(task_id, workbook=workbook, identity=identity))


def _python_result(task_id: str) -> str:
    """模拟 `python_coordinator` 的登记：先 register，再交给故障入口认领。"""
    registered = registry_module.local_results.register(task_id, _result(), "fp-source-1")
    dev_result_ref_fault.bind_python_result(task_id, registered.handle)
    return registered.handle


@pytest.fixture(autouse=True)
def _clean(monkeypatch):
    monkeypatch.setenv("EXCEL_BRO_RESULT_REF_FAULT", "1")
    registry_module.local_results.reset()
    task_tool_bridge.reset()
    dev_result_ref_fault.reset()
    dev_fault_target.reset()
    yield
    registry_module.local_results.reset()
    task_tool_bridge.reset()
    dev_result_ref_fault.reset()


# ---------------------------------------------------------------------------
# 普通路径：没启用时一切照旧
# ---------------------------------------------------------------------------


def test_plain_path_unaffected_when_nothing_is_armed():
    _open_task("turn-plain")
    handle = _python_result("turn-plain")
    assert resolve_result_ref("turn-plain", handle).rows == [["甲", 2], ["乙", 1]]
    plan_response = build_result_plan_response(
        resolve_result_ref("turn-plain", handle),
        _plan_request("turn-plain"),
        target={"sheet": "Sheet1", "range": "A1:B3"},
    )
    assert plan_response.plan.actions[0].type == "writeValues"
    # 正常路径可以被解析多次（引用没过期）。
    assert resolve_result_ref("turn-plain", handle).row_count == 2


# ---------------------------------------------------------------------------
# 单次触发：绑定的那一条失效，同一任务的其它引用不受影响
# ---------------------------------------------------------------------------


def test_armed_fault_expires_only_the_first_python_ref_of_the_bound_task():
    dev_result_ref_fault.arm(host_id="excel-office", workbook_name=WORKBOOK, workbook_identity=IDENTITY)
    _open_task("turn-armed")
    doomed = _python_result("turn-armed")

    with pytest.raises(ResultSubmissionError) as error:
        resolve_result_ref("turn-armed", doomed)
    assert error.value.code == RESULT_REF_INVALID_CODE
    # 引用不是"这次读不到"，而是真的没了：再读一次仍然无效。
    assert registry_module.local_results.count_task("turn-armed") == 0
    with pytest.raises(ResultSubmissionError):
        resolve_result_ref("turn-armed", doomed)


def test_single_shot_does_not_hit_later_python_results_of_the_same_task():
    dev_result_ref_fault.arm(host_id="excel-office", workbook_name=WORKBOOK, workbook_identity=IDENTITY)
    _open_task("turn-single")
    first = _python_result("turn-single")
    second = _python_result("turn-single")

    with pytest.raises(ResultSubmissionError):
        resolve_result_ref("turn-single", first)
    # 单次生效：同一任务的第二条 Python 结果照常可读。
    assert resolve_result_ref("turn-single", second).row_count == 2
    assert registry_module.local_results.count_task("turn-single") == 1


def test_query_result_is_not_the_target():
    """入口绑的是 **Python** 结果：同任务先登记的查询结果不被消耗、也不失效。"""
    dev_result_ref_fault.arm(host_id="excel-office", workbook_name=WORKBOOK, workbook_identity=IDENTITY)
    _open_task("turn-query")
    query = registry_module.local_results.register("turn-query", _result(), "fp-source-1")
    assert resolve_result_ref("turn-query", query.handle).row_count == 2

    doomed = _python_result("turn-query")
    with pytest.raises(ResultSubmissionError):
        resolve_result_ref("turn-query", doomed)
    # 查询结果仍在（失效只作用于绑定的那一个句柄）。
    assert resolve_result_ref("turn-query", query.handle).row_count == 2


# ---------------------------------------------------------------------------
# 任务隔离 / 工作簿隔离
# ---------------------------------------------------------------------------


def test_fault_stays_inside_the_bound_task():
    dev_result_ref_fault.arm(host_id="excel-office", workbook_name=WORKBOOK, workbook_identity=IDENTITY)
    _open_task("turn-a")
    _open_task("turn-b")
    handle_a = _python_result("turn-a")
    handle_b = _python_result("turn-b")

    with pytest.raises(ResultSubmissionError):
        resolve_result_ref("turn-a", handle_a)
    # 另一个任务的结果完全不受影响。
    assert resolve_result_ref("turn-b", handle_b).row_count == 2


def test_fault_does_not_cross_workbooks():
    dev_result_ref_fault.arm(host_id="excel-office", workbook_name=WORKBOOK, workbook_identity=IDENTITY)
    # 另一本簿的下一次任务：绑定不匹配，启用不会被消费。
    _open_task(
        "turn-other-book",
        workbook="别的.xlsx",
        identity="url:https://example.com/别的.xlsx",
    )
    handle = _python_result("turn-other-book")
    assert resolve_result_ref("turn-other-book", handle).row_count == 2
    # 启用还留着，等它自己那本簿的下一次任务。
    assert dev_result_ref_fault.status()["phase"] == "armed"
    _open_task("turn-own-book")
    doomed = _python_result("turn-own-book")
    with pytest.raises(ResultSubmissionError):
        resolve_result_ref("turn-own-book", doomed)


# ---------------------------------------------------------------------------
# 非开发模式不可用
# ---------------------------------------------------------------------------


def test_disabled_by_default_even_if_a_caller_tries_to_arm(monkeypatch):
    monkeypatch.delenv("EXCEL_BRO_RESULT_REF_FAULT", raising=False)
    dev_result_ref_fault.reset()
    assert dev_result_ref_fault.enabled() is False
    with pytest.raises(dev_result_ref_fault.FaultDisabled):
        dev_result_ref_fault.arm(host_id="excel-office", workbook_name=WORKBOOK, workbook_identity=IDENTITY)
    # 关掉之后即使内部状态被硬写，绑定与失效都不生效。
    _open_task("turn-off")
    handle = _python_result("turn-off")
    assert resolve_result_ref("turn-off", handle).row_count == 2


def test_arm_endpoint_is_hidden_when_disabled(monkeypatch):
    from fastapi.testclient import TestClient
    from server.app.main import app

    monkeypatch.delenv("EXCEL_BRO_RESULT_REF_FAULT", raising=False)
    dev_result_ref_fault.reset()
    client = TestClient(app, base_url="http://127.0.0.1", client=("127.0.0.1", 5000))
    response = client.post("/api/dev/result-ref-fault/arm",
        json={"hostId": "excel-office", "workbookName": WORKBOOK, "workbookIdentity": ""})
    assert response.status_code == 404
    assert response.json()["detail"]["code"] == "DEV_FAULT_UNAVAILABLE"


def test_arm_endpoint_cannot_name_another_task_or_reference():
    """入口请求里没有 taskId / resultRef 字段：多传一律 422，不给模型指定权。"""
    from fastapi.testclient import TestClient
    from server.app.main import app

    client = TestClient(app, base_url="http://127.0.0.1", client=("127.0.0.1", 5000))
    response = client.post("/api/dev/result-ref-fault/arm", json={
        "hostId": "excel-office", "workbookName": WORKBOOK, "workbookIdentity": "",
        "taskId": "turn-someone-else", "resultRef": "res_" + "0" * 32,
    })
    assert response.status_code == 422

    ok = client.post("/api/dev/result-ref-fault/arm",
        json={"hostId": "excel-office", "workbookName": WORKBOOK, "workbookIdentity": IDENTITY})
    assert ok.status_code == 200
    assert ok.json()["phase"] == "armed"
    assert ok.json()["enabled"] is True
    # 请求模型自身也没有这两个字段（双保险：即使路由改了也不会默默接受）。
    assert set(DevResultRefFaultArmRequest.model_fields) == {
        "hostId", "workbookName", "workbookIdentity"}


# ---------------------------------------------------------------------------
# 到期与取消
# ---------------------------------------------------------------------------


def test_arm_expires_without_firing(monkeypatch):
    now = [1000.0]
    monkeypatch.setattr(dev_result_ref_fault, "_now", lambda: now[0])
    dev_result_ref_fault.arm(host_id="excel-office", workbook_name=WORKBOOK, workbook_identity=IDENTITY)
    now[0] += dev_result_ref_fault.ttl_seconds() + 1
    _open_task("turn-late")
    handle = _python_result("turn-late")
    assert resolve_result_ref("turn-late", handle).row_count == 2
    assert dev_result_ref_fault.status()["phase"] == "expired"


def test_cancel_before_the_task_starts_leaves_everything_working():
    dev_result_ref_fault.arm(host_id="excel-office", workbook_name=WORKBOOK, workbook_identity=IDENTITY)
    dev_result_ref_fault.cancel()
    _open_task("turn-cancelled")
    handle = _python_result("turn-cancelled")
    assert resolve_result_ref("turn-cancelled", handle).row_count == 2


# ---------------------------------------------------------------------------
# 零宿主写入：失效发生在任何计划被编译出来之前
# ---------------------------------------------------------------------------


def test_expired_ref_never_reaches_a_plan_so_no_host_write_can_happen():
    dev_result_ref_fault.arm(host_id="excel-office", workbook_name=WORKBOOK, workbook_identity=IDENTITY)
    _open_task("turn-nowrite")
    doomed = _python_result("turn-nowrite")

    # 计划是宿主写入的唯一入口：拿不到 item 就编译不出计划。
    compiled = []
    with pytest.raises(ResultSubmissionError):
        item = resolve_result_ref("turn-nowrite", doomed)
        compiled.append(build_result_plan_response(
            item, _plan_request("turn-nowrite"),
            target={"sheet": "Sheet1", "range": "A1:B3"}))
    assert compiled == []
    # 事后也不会有"缓存正文"能顶上：句柄已经不在登记器里。
    assert registry_module.local_results.resolve("turn-nowrite", doomed) is None
