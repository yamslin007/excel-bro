"""「实际写入次数」的验收口径（A1-E3）。

过去验收说明拿 `/executions/{id}/actions/start` 的请求行数当写入次数 —— 那只是
**请求**次数：重复请求会被账本拒，认领成功也可能在碰宿主之前就失败。本文件钉住
新口径能证明什么、不能证明什么：

1. 上报只追加：一条 `failed_before_write`（宿主写调用 0 次）**不得**被算成已写入；
2. 同一动作先后两条时，以**最后一条**为准（补记不覆盖先到的现场）；
3. 别人的 `ownerToken` 不许往这次执行里写字；
4. 不存在的 operationId 一律 404，且不会因此创建任何记录；
5. 另一端（WPS）的 operationId 在 Office 端点上是"不存在"，不能跨命名空间探测；
6. 上报**不改**执行状态：revision 与 status 一字不动（观测不是恢复点）。
"""

from __future__ import annotations

from fastapi.testclient import TestClient

from server.app.main import app

OPERATION_ID = "office-host-writes-0001"


def _plan(action: dict[str, object]) -> dict[str, object]:
    return {
        "id": "office-plan-hw-1",
        "title": "写入测试",
        "summary": "宿主写调用审计",
        "assumptions": [],
        "warnings": [],
        "actions": [action],
    }


def _register(client: TestClient, operation_id: str = OPERATION_ID) -> dict:
    plan = _plan(
        {"type": "writeValues", "sheet": "结果", "range": "A1", "values": [["ok"]]}
    )
    response = client.post(
        "/api/office/executions/register",
        json={
            "operationId": operation_id,
            "plan": plan,
            "authorizedSheetNames": ["结果"],
            "allowedExternalFiles": [],
            "workbookIdentity": "url:c:/work/host-writes.xlsx",
            "confirmed": True,
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


def test_failed_before_write_is_never_counted_as_a_write() -> None:
    """认领后、写入前失败：只能记 `failed_before_write`，不能算已写入。"""
    client = TestClient(app)
    state = _register(client)

    response = client.post(
        f"/api/office/executions/{OPERATION_ID}/host-writes",
        json={
            "ownerToken": state["ownerToken"],
            "entries": [
                _entry(
                    hostWriteCalls=0,
                    mutatingCalls=1,
                    outcome="failed_before_write",
                    readback="not_applicable",
                    errorCode="EXCEL_ACTION_FAILED",
                )
            ],
        },
    )
    assert response.status_code == 200, response.text
    body = response.json()
    totals = body["totals"]
    assert totals["reportedActions"] == 1
    assert totals["actionsFailedBeforeWrite"] == 1
    assert totals["actionsWritten"] == 0
    # 关键：写前失败不把刷新次数算进"宿主写调用"
    assert totals["hostWriteCalls"] == 0

    # 上报不改执行状态：观测不是恢复点。
    after = client.get(f"/api/office/executions/{OPERATION_ID}").json()
    assert after["revision"] == state["revision"]
    assert after["status"] == state["status"]


def test_last_report_for_an_action_wins_and_counts_add_up() -> None:
    """同一次动作先后两条（先失败后补记成功）时取最后一条，计数只加一次。"""
    client = TestClient(app)
    state = _register(client, "office-host-writes-0002")
    url = "/api/office/executions/office-host-writes-0002/host-writes"

    first = client.post(
        url,
        json={
            "ownerToken": state["ownerToken"],
            "entries": [_entry(hostWriteCalls=0, outcome="unknown", readback="not_applicable")],
        },
    )
    assert first.status_code == 200, first.text
    assert first.json()["totals"]["actionsUnknown"] == 1

    second = client.post(
        url,
        json={"ownerToken": state["ownerToken"], "entries": [_entry(hostWriteCalls=2)]},
    )
    assert second.status_code == 200, second.text
    totals = second.json()["totals"]
    assert totals["reportedActions"] == 1
    assert totals["actionsWritten"] == 1
    assert totals["actionsUnknown"] == 0
    assert totals["hostWriteCalls"] == 2
    assert totals["readbackVerified"] == 1
    # 条目只有一个（按动作去重），不是两条相加。
    assert len(second.json()["entries"]) == 1


def test_another_executor_cannot_report_into_this_execution() -> None:
    """别人的 ownerToken 不许往这次执行里写字（哪怕是计数）。"""
    client = TestClient(app)
    _register(client, "office-host-writes-0003")
    response = client.post(
        "/api/office/executions/office-host-writes-0003/host-writes",
        json={"ownerToken": "x" * 40, "entries": [_entry()]},
    )
    assert response.status_code == 409
    assert response.json()["detail"]["code"] == "HOST_WRITES_REJECTED"
    assert client.get(
        "/api/office/executions/office-host-writes-0003/host-writes"
    ).json()["totals"]["reportedActions"] == 0


def test_unknown_or_other_namespace_operation_is_404() -> None:
    """不存在的 ID、或另一端的 ID，一律 404 —— 审计端点不能变成探测工具。"""
    client = TestClient(app)
    assert client.post(
        "/api/office/executions/nope-does-not-exist-0001/host-writes",
        json={"entries": [_entry()]},
    ).status_code == 404
    assert client.get(
        "/api/office/executions/nope-does-not-exist-0001/host-writes"
    ).status_code == 404
    # WPS 端点看不到 Office 的执行记录。
    assert client.get(
        f"/api/wps/executions/{OPERATION_ID}/host-writes"
    ).status_code == 404


def test_entries_reject_bad_counts_instead_of_rounding() -> None:
    """计数只接受非负整数：负值、小数、缺字段都被拒（不静默取整）。"""
    client = TestClient(app)
    state = _register(client, "office-host-writes-0004")
    url = "/api/office/executions/office-host-writes-0004/host-writes"
    for bad in (
        _entry(hostWriteCalls=-1),
        _entry(hostWriteCalls=1.5),
        _entry(outcome="done"),
        _entry(readback="maybe"),
    ):
        response = client.post(url, json={"ownerToken": state["ownerToken"], "entries": [bad]})
        assert response.status_code == 422, (bad, response.text)
