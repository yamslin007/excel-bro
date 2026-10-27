"""「没有观测」与「明确零写入」的分界（A1-E3 第三批）。

上一批把 `host-writes` 做成 best-effort 补报，等于承认**上报可能一条都到不了**。
那样 `reportedActions=0 / hostWriteCalls=0 / entries=[]` 就同时对应两件完全不同的事：

* 谁都没上报 ⇒ 我们对宿主一无所知（不能当成"没写"）；
* 上报到了、而且每一次都是 0 次调用 ⇒ 这才叫"明确零写调用"。

本文件钉住两者不可互换，并且"明确零"只能由**账本执行边界 + 完整覆盖**推出：

1. 一条上报都没有 ⇒ `not_collected` + `callEvidence=none`（**不是** zero）；
2. 边界内只覆盖了一部分 ⇒ `partial` + `unknown`；
3. 边界内全覆盖、每条都是 0 次 ⇒ `complete` + `zero`；
4. 至少一条非零 ⇒ `nonzero`（即使覆盖不全，`observation` 仍如实写 `partial`）；
5. 覆盖完整但有 `unknown` / `failed_after_write` ⇒ `unknown`，不许说零；
6. 上报不进边界（账本没认领过那个动作）⇒ `unexpectedReports>0` 且 `unknown`；
7. 执行记录不存在 / 跨命名空间 ⇒ 404，不会返回"看起来已验证的零"；
8. 报告丢了（只到一半）不会因为"没收到"就把已有的数字改成 0。
"""

from __future__ import annotations

import uuid

from fastapi.testclient import TestClient

from server.app.main import app


def _op_id(prefix: str = "office-host-observation") -> str:
    """每次运行用新的 operationId。

    账本是同一份 SQLite（跨这次运行持续存在），固定 ID 在重跑时会撞上上一轮已经
    推进过的 `next_action_index`，`actions/start` 于是 409 —— 那是测试自伤，
    不是产品缺陷。
    """
    return f"{prefix}-{uuid.uuid4().hex[:12]}"


def _plan(count: int = 2) -> dict[str, object]:
    return {
        "id": "office-plan-obs-1",
        "title": "观测口径",
        "summary": "宿主写调用观测覆盖",
        "assumptions": [],
        "warnings": [],
        "actions": [
            {
                "type": "writeValues",
                "sheet": "结果",
                "range": f"A{index + 1}",
                "values": [["ok"]],
            }
            for index in range(count)
        ],
    }


def _register(client: TestClient, operation_id: str, *, actions: int = 2) -> dict:
    response = client.post(
        "/api/office/executions/register",
        json={
            "operationId": operation_id,
            "plan": _plan(actions),
            "authorizedSheetNames": ["结果"],
            "allowedExternalFiles": [],
            "workbookIdentity": "url:c:/work/host-observation.xlsx",
            "confirmed": True,
        },
    )
    assert response.status_code == 200, response.text
    return response.json()


def _start(client: TestClient, operation_id: str, state: dict, index: int) -> dict:
    response = client.post(
        f"/api/office/executions/{operation_id}/actions/start",
        json={
            "expectedRevision": state["revision"],
            "ownerToken": state["ownerToken"],
            "actionIndex": index,
            "beforeSnapshot": {
                "sheetName": "结果",
                "range": f"A{index + 1}",
                "type": "writeValues",
                "beforeFingerprint": "f" * 64,
            },
        },
    )
    assert response.status_code == 200, response.text
    return response.json()


def _complete(client: TestClient, operation_id: str, state: dict, index: int) -> dict:
    response = client.post(
        f"/api/office/executions/{operation_id}/actions/complete",
        json={
            "expectedRevision": state["revision"],
            "ownerToken": state["ownerToken"],
            "actionIndex": index,
            "checkpoint": {
                "sheetName": "结果",
                "range": f"A{index + 1}",
                "type": "writeValues",
                "beforeFingerprint": "f" * 64,
                "afterFingerprint": "e" * 64,
            },
        },
    )
    assert response.status_code == 200, response.text
    return response.json()


def _entry(**overrides: object) -> dict[str, object]:
    entry: dict[str, object] = {
        "actionIndex": 0,
        "actionType": "writeValues",
        "hostWriteCalls": 1,
        "mutatingCalls": 1,
        "outcome": "written",
        "readback": "verified",
    }
    entry.update(overrides)
    return entry


def _report(client: TestClient, operation_id: str, owner: str, entries: list[dict]) -> dict:
    response = client.post(
        f"/api/office/executions/{operation_id}/host-writes",
        json={"ownerToken": owner, "entries": entries},
    )
    assert response.status_code == 200, response.text
    return response.json()


def _summary(client: TestClient, operation_id: str, namespace: str = "office") -> dict:
    response = client.get(
        f"/api/{namespace}/executions/{operation_id}/host-writes"
    )
    assert response.status_code == 200, response.text
    return response.json()


def test_no_report_is_not_zero_writes() -> None:
    """一条上报都没有：只能说"没有观测"，不能推出零写入。"""
    client = TestClient(app)
    op = _op_id()
    state = _register(client, op, actions=1)
    state = _start(client, op, state, 0)
    _complete(client, op, state, 0)

    body = _summary(client, op)
    assert body["entries"] == []
    assert body["totals"]["hostWriteCalls"] == 0
    observation = body["observation"]
    assert observation["observation"] == "not_collected"
    assert observation["callEvidence"] == "none"
    # 边界本身是真的：这个动作确实被认领并完成过（可能写过）。
    assert observation["attemptedActions"] == 1
    assert observation["reportedActions"] == 0
    assert observation["missingActions"] == 1
    assert observation["lastReportedAt"] is None
    assert observation["unit"] == "excel_sync_batches"


def test_no_action_claimed_is_a_ledger_fact_not_an_observation() -> None:
    """登记了但一个动作都没认领：账本说"没写过"，但仍不是"观测到零"。"""
    client = TestClient(app)
    op = _op_id()
    _register(client, op, actions=1)
    observation = _summary(client, op)["observation"]
    assert observation["attemptedActions"] == 0
    assert observation["observation"] == "not_collected"
    assert observation["callEvidence"] == "none"


def test_partial_coverage_can_never_claim_zero() -> None:
    """边界 2 个动作、只上报了 1 个：覆盖率不全 ⇒ 不许说零。"""
    client = TestClient(app)
    op = _op_id()
    state = _register(client, op, actions=2)
    state = _start(client, op, state, 0)
    state = _complete(client, op, state, 0)
    state = _start(client, op, state, 1)

    body = _report(
        client,
        op,
        state["ownerToken"],
        [_entry(actionIndex=0, hostWriteCalls=0, outcome="written")],
    )
    observation = body["observation"]
    assert observation["observation"] == "partial"
    assert observation["callEvidence"] == "unknown"
    assert observation["attemptedActions"] == 2
    assert observation["reportedActions"] == 1
    assert observation["missingActions"] == 1


def test_full_coverage_with_zero_calls_is_the_only_zero() -> None:
    """边界内全覆盖、每条都是 0 次调用：这才是"明确观测到零写调用"。"""
    client = TestClient(app)
    op = _op_id()
    state = _register(client, op, actions=1)
    state = _start(client, op, state, 0)
    _complete(client, op, state, 0)

    body = _report(
        client,
        op,
        state["ownerToken"],
        [_entry(actionIndex=0, hostWriteCalls=0, mutatingCalls=1)],
    )
    observation = body["observation"]
    assert observation["observation"] == "complete"
    assert observation["callEvidence"] == "zero"
    assert observation["missingActions"] == 0
    assert observation["unexpectedReports"] == 0
    assert observation["lastReportedAt"] is not None


def test_nonzero_calls_are_reported_even_when_coverage_is_incomplete() -> None:
    """已经看到非零调用：这一条与"覆盖不全"分别记录，不许互相抵消。"""
    client = TestClient(app)
    op = _op_id()
    state = _register(client, op, actions=2)
    state = _start(client, op, state, 0)
    state = _complete(client, op, state, 0)
    # 第 2 个动作也认领过（边界 = 2），但还没有它的观测。
    state = _start(client, op, state, 1)

    body = _report(
        client,
        op,
        state["ownerToken"],
        [_entry(actionIndex=0, hostWriteCalls=3)],
    )
    observation = body["observation"]
    assert observation["callEvidence"] == "nonzero"
    assert observation["observation"] == "partial"
    assert body["totals"]["hostWriteCalls"] == 3


def test_unknown_outcome_blocks_the_zero_verdict() -> None:
    """宿主抛错、断线只能记 `unknown`：覆盖再全也不能说零。"""
    client = TestClient(app)
    op = _op_id()
    state = _register(client, op, actions=1)
    state = _start(client, op, state, 0)

    body = _report(
        client,
        op,
        state["ownerToken"],
        [
            _entry(
                actionIndex=0,
                hostWriteCalls=0,
                outcome="unknown",
                readback="not_applicable",
            )
        ],
    )
    observation = body["observation"]
    assert observation["observation"] == "complete"
    assert observation["callEvidence"] == "unknown"
    assert body["totals"]["actionsUnknown"] == 1


def test_report_outside_the_ledger_boundary_is_flagged_not_counted_as_zero() -> None:
    """账本没认领过那个动作却收到它的观测：边界与观测不一致 ⇒ unknown。"""
    client = TestClient(app)
    op = _op_id()
    state = _register(client, op, actions=1)
    # 没有 actions/start：账本边界是空的。
    body = _report(
        client,
        op,
        state["ownerToken"],
        [_entry(actionIndex=0, hostWriteCalls=0, outcome="written")],
    )
    observation = body["observation"]
    assert observation["attemptedActions"] == 0
    assert observation["unexpectedReports"] == 1
    assert observation["callEvidence"] == "unknown"


def test_lost_half_of_the_reports_does_not_turn_into_zero() -> None:
    """先到的那一半照常算数；后一半丢了不会把已有数字抹成 0。"""
    client = TestClient(app)
    op = _op_id()
    state = _register(client, op, actions=2)
    state = _start(client, op, state, 0)
    state = _complete(client, op, state, 0)
    state = _start(client, op, state, 1)

    _report(
        client,
        op,
        state["ownerToken"],
        [_entry(actionIndex=0, hostWriteCalls=2)],
    )
    body = _summary(client, op)
    assert body["totals"]["hostWriteCalls"] == 2
    assert body["observation"]["callEvidence"] == "nonzero"
    assert body["observation"]["missingActions"] == 1


def test_unknown_operation_never_returns_a_verified_looking_zero() -> None:
    """执行记录不存在（含规划期就停下的场景）⇒ 404，没有"读到 0"这回事。"""
    client = TestClient(app)
    office_op = _op_id()
    _register(client, office_op, actions=1)
    for path in (
        "/api/office/executions/execution-never-created/host-writes",
        "/api/wps/executions/wps-execution-never-created/host-writes",
        # 另一端的执行 ID 在本端点上是"不存在"：不能跨命名空间探测。
        f"/api/wps/executions/{office_op}/host-writes",
    ):
        response = client.get(path)
        assert response.status_code == 404, path
        assert response.json()["detail"]["code"] == "EXECUTION_NOT_FOUND"


def test_wps_namespace_declares_its_own_unit() -> None:
    """两端计量单位不同：WPS 记的是宿主写 API 调用次数，不是刷新批次。"""
    client = TestClient(app)
    op = _op_id("wps-host-observation")
    response = client.post(
        "/api/wps/executions/register",
        json={
            "operationId": op,
            "taskId": "task-observation",
            "confirmed": True,
            "authorizedSheetNames": ["结果"],
            "workbookIdentity": "wps-path:c:/work/host-observation.xlsx",
            # WPS 的登记只带恢复描述，不带单元格值矩阵。
            "descriptor": {
                "actions": [
                    {
                        "actionIndex": 0,
                        "type": "writeValues",
                        "sheet": "结果",
                        "range": "A1:A1",
                        "rowCount": 1,
                        "columnCount": 1,
                        "cellCount": 1,
                        "previewFingerprint": "v2:fp-before",
                        "expectedFingerprint": "v2:fp-after",
                    }
                ]
            },
        },
    )
    assert response.status_code == 200, response.text
    body = _summary(client, op, "wps")
    assert body["observation"]["unit"] == "wps_api_calls"
