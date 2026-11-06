"""审阅生命周期必须是**账本里的持久事实**，不是界面里的一个内存状态。

真实现场（2026-11-03）：写入完成、回读通过，用户还没点「确认全部变更 / 放弃全部变更」
就把任务窗格刷新了 —— 待审阅界面消失，输入框回到普通状态，甚至显示「本次没有需要你
处理的后续操作」，而原表停在一个没人收尾的写入现场上。

根因（读代码得出，不是推测）：

1. `pending_office_executions` 的 SQL **明确排除 `completed`**，所以"完成但没审阅"
   从来不会被发现；
2. 账本只有 `status` / `abandoned_at` / `undone_at`，**没有"待审阅 / 已确认"** ——
   `abandoned_at` 记的是"撤销入口是不是用掉了"，不是"用户看过结果没有"；
3. 于是审阅状态只活在 `ResultReview` 的内存里，刷新即归零。

这组用例钉住补上的那一列与那一条端点：

1. 审阅初始状态与完成回执**同一条 UPDATE** —— 没有"已完成 + 没有审阅事实"的窗口；
2. 回执没带验证结论 ⇒ `NULL`（无法判断），**不是** `pending`，更不是 `confirmed`；
3. 确认 / 放弃 / 撤销四种结果 `applied` / `duplicate` / `conflict` / `missing` 分开表达：
   重复上报幂等且不改写"首次成立"的时刻，与已成立终态冲突必须 409；
4. 审阅决定**不碰** `status`、`revision`、动作游标、在途动作；
5. 发现只认显式审阅状态：`completed` 且 `review_state IN ('pending','blocked')` 才是
   `kind="review"` 候选；老记录（`review_state IS NULL`）一条都不许多出来；
6. 跨命名空间一律 404 —— 审阅端点不能成为探测另一宿主的工具。
"""

from __future__ import annotations

import pytest
from fastapi.testclient import TestClient

import server.app.main as main_module
import server.app.task_ledger as task_ledger_module
from server.app.main import app
from server.app.task_ledger import FolderOperationLedger

OFFICE_ID = "office-review-00000001"
WPS_ID = "wps-review-00000001"
OFFICE_IDENTITY = "url:c:/work/review.xlsx"
WPS_IDENTITY = "wps-path:c:/work/review.xlsx"


class _Clock:
    def __init__(self, start: float = 5000.0) -> None:
        self.now = start

    def time(self) -> float:
        self.now += 1.0
        return self.now


@pytest.fixture(autouse=True)
def isolate(monkeypatch: pytest.MonkeyPatch, tmp_path, clock: _Clock) -> None:
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path / "app-data"))
    main_module._folder_operations.clear()
    monkeypatch.setattr(task_ledger_module, "time", clock)


@pytest.fixture()
def clock() -> _Clock:
    return _Clock()


def _ledger() -> FolderOperationLedger:
    return main_module._folder_operation_ledger


def _office_plan() -> dict[str, object]:
    return {
        "id": "office-review-plan-1",
        "title": "写入测试",
        "summary": "审阅生命周期",
        "assumptions": [],
        "warnings": [],
        "actions": [
            {"type": "writeValues", "sheet": "结果", "range": "B2", "values": [["111"]]}
        ],
    }


def _register_office(client: TestClient, operation_id: str = OFFICE_ID) -> dict:
    response = client.post(
        "/api/office/executions/register",
        json={
            "operationId": operation_id,
            "plan": _office_plan(),
            "authorizedSheetNames": ["结果"],
            "allowedExternalFiles": [],
            "workbookIdentity": OFFICE_IDENTITY,
            "confirmed": True,
        },
    )
    assert response.status_code == 200, response.text
    return response.json()


def _finish_office(
    client: TestClient, state: dict, result: dict, operation_id: str = OFFICE_ID
) -> dict:
    # 先走完那一个动作：`complete` 要求"全部动作已完成"，跳过这一步拿到的是 409，
    # 那不是这次要测的东西（会掩盖审阅状态本身的结论）。
    started = client.post(
        f"/api/office/executions/{operation_id}/actions/start",
        json={
            "expectedRevision": state["revision"],
            "ownerToken": state["ownerToken"],
            "actionIndex": 0,
            "beforeSnapshot": {"range": "结果!B2", "values": [[""]]},
        },
    )
    assert started.status_code == 200, started.text
    completed = client.post(
        f"/api/office/executions/{operation_id}/actions/complete",
        json={
            "expectedRevision": started.json()["revision"],
            "ownerToken": state["ownerToken"],
            "actionIndex": 0,
            "checkpoint": {"range": "结果!B2", "values": [["111"]]},
        },
    )
    assert completed.status_code == 200, completed.text
    response = client.post(
        f"/api/office/executions/{operation_id}/complete",
        json={
            "expectedRevision": completed.json()["revision"],
            "ownerToken": state["ownerToken"],
            "result": result,
        },
    )
    assert response.status_code == 200, response.text
    return response.json()


def _verified_result() -> dict:
    return {
        "verification": {"status": "verified", "checks": [], "unverifiedActions": []},
        "reviews": [],
    }


def _register_wps(client: TestClient, operation_id: str = WPS_ID) -> dict:
    response = client.post(
        "/api/wps/executions/register",
        json={
            "operationId": operation_id,
            "taskId": "task-review",
            "confirmed": True,
            "workbookIdentity": WPS_IDENTITY,
            "authorizedSheetNames": ["Sheet1"],
            "descriptor": {
                "actions": [
                    {
                        "actionIndex": 0,
                        "type": "writeValues",
                        "sheet": "Sheet1",
                        "range": "B2:B2",
                        "rowCount": 1,
                        "columnCount": 1,
                        "cellCount": 1,
                        "previewFingerprint": "v2:before-0",
                        "expectedFingerprint": "v2:after-0",
                    }
                ]
            },
        },
    )
    assert response.status_code == 200, response.text
    return response.json()


def test_finish_writes_pending_review_in_the_same_row() -> None:
    """完成回执与审阅初始状态同一条 UPDATE：不存在"已完成但没有审阅事实"的窗口。"""
    client = TestClient(app)
    state = _register_office(client)
    finished = _finish_office(client, state, _verified_result())

    assert finished["status"] == "completed"
    assert finished["review"]["state"] == "pending"
    assert isinstance(finished["review"]["decidedAt"], float)

    record = _ledger().office_execution(OFFICE_ID)
    assert record is not None
    assert record.status == "completed"
    assert record.review_state == "pending"
    # 审阅初始状态不是"用户做过的处置"：两列留痕不许互相冒充。
    assert record.abandoned_at is None
    assert record.undone_at is None


def test_finish_without_verification_evidence_leaves_review_unknown() -> None:
    """回执没带验证结论 ⇒ `NULL`（无法判断）。把它当 `pending` 会凭空造出待审阅。"""
    client = TestClient(app)
    state = _register_office(client)
    finished = _finish_office(client, state, {"reviews": []})

    assert finished["status"] == "completed"
    assert finished["review"]["state"] is None
    record = _ledger().office_execution(OFFICE_ID)
    assert record is not None
    assert record.review_state is None


def test_unverified_result_is_blocked_not_pending() -> None:
    """验证没通过 = `blocked`（要人核对，但不能直接接受），不是 `pending`。"""
    client = TestClient(app)
    state = _register_office(client)
    finished = _finish_office(
        client,
        state,
        {"verification": {"status": "executed_unverified", "checks": [], "unverifiedActions": []}},
    )
    assert finished["review"]["state"] == "blocked"


def test_confirm_is_idempotent_and_keeps_first_decided_at() -> None:
    client = TestClient(app)
    state = _register_office(client)
    first = _finish_office(client, state, _verified_result())
    assert first["review"]["state"] == "pending"
    applied = client.post(f"/api/office/executions/{OFFICE_ID}/review", json={"decision": "confirmed"})
    assert applied.status_code == 200, applied.text
    assert applied.json()["review"]["state"] == "confirmed"
    decided_at = applied.json()["review"]["decidedAt"]

    # 重试 / 双窗格各报一次：幂等成功，且**不改写**"首次成立"的时刻。
    again = client.post(f"/api/office/executions/{OFFICE_ID}/review", json={"decision": "confirmed"})
    assert again.status_code == 200, again.text
    assert again.json()["review"]["decidedAt"] == decided_at
    assert again.json()["review"]["state"] == "confirmed"


def test_review_decision_does_not_touch_progress_or_status() -> None:
    """审阅决定是"用户处理过没有"，不是进度：status/revision/游标一概不动。"""
    client = TestClient(app)
    state = _register_office(client)
    _finish_office(client, state, _verified_result())
    before = _ledger().office_execution(OFFICE_ID)
    assert before is not None

    assert (
        client.post(
            f"/api/office/executions/{OFFICE_ID}/review", json={"decision": "confirmed"}
        ).status_code
        == 200
    )

    after = _ledger().office_execution(OFFICE_ID)
    assert after is not None
    assert after.status == before.status == "completed"
    assert after.revision == before.revision
    assert after.next_action_index == before.next_action_index
    assert after.in_flight_action_index == before.in_flight_action_index


def test_terminal_conflict_is_409_and_does_not_overwrite() -> None:
    """已经放弃的执行再收到"确认"：409 + 账本保持第一次的结论。"""
    client = TestClient(app)
    state = _register_office(client)
    _finish_office(client, state, _verified_result())

    assert (
        client.post(
            f"/api/office/executions/{OFFICE_ID}/review", json={"decision": "discarded"}
        ).status_code
        == 200
    )
    conflict = client.post(
        f"/api/office/executions/{OFFICE_ID}/review", json={"decision": "confirmed"}
    )
    assert conflict.status_code == 409, conflict.text
    assert conflict.json()["detail"]["code"] == "EXECUTION_REVIEW_CONFLICT"

    record = _ledger().office_execution(OFFICE_ID)
    assert record is not None
    assert record.review_state == "discarded"


def test_confirmed_can_become_undone_but_not_discarded() -> None:
    """确认之后又撤销是真实存在的一条路；确认之后"放弃"不是。"""
    client = TestClient(app)
    state = _register_office(client)
    _finish_office(client, state, _verified_result())
    client.post(f"/api/office/executions/{OFFICE_ID}/review", json={"decision": "confirmed"})

    assert (
        client.post(
            f"/api/office/executions/{OFFICE_ID}/review", json={"decision": "discarded"}
        ).status_code
        == 409
    )
    undone = client.post(
        f"/api/office/executions/{OFFICE_ID}/review", json={"decision": "undone"}
    )
    assert undone.status_code == 200, undone.text
    assert undone.json()["review"]["state"] == "undone"


def test_review_on_missing_or_foreign_record_is_404() -> None:
    client = TestClient(app)
    missing = client.post(
        "/api/office/executions/does-not-exist/review", json={"decision": "confirmed"}
    )
    assert missing.status_code == 404
    assert missing.json()["detail"]["code"] == "EXECUTION_REVIEW_NOT_FOUND"

    _register_wps(client)
    # WPS 的记录走 Office 的审阅端点 = 不存在（不跨命名空间作证）。
    foreign = client.post(
        f"/api/office/executions/{WPS_ID}/review", json={"decision": "confirmed"}
    )
    assert foreign.status_code == 404
    record = _ledger().office_execution(WPS_ID)
    assert record is not None and record.review_state is None


def test_review_requires_a_known_decision_value() -> None:
    client = TestClient(app)
    state = _register_office(client)
    _finish_office(client, state, _verified_result())
    invalid = client.post(
        f"/api/office/executions/{OFFICE_ID}/review", json={"decision": "accepted"}
    )
    assert invalid.status_code == 422


def test_pending_discovery_surfaces_unfinished_review_as_review_kind() -> None:
    """写完了、没人收尾 ⇒ 必须是 `kind="review"` 的候选，而不是"没有未完成任务"。"""
    client = TestClient(app)
    state = _register_office(client)
    _finish_office(client, state, _verified_result())

    response = client.get(
        "/api/office/executions/pending", params={"workbookIdentity": OFFICE_IDENTITY}
    )
    assert response.status_code == 200, response.text
    candidates = response.json()["candidates"]
    assert len(candidates) == 1
    candidate = candidates[0]
    assert candidate["kind"] == "review"
    assert candidate["status"] == "completed"
    assert candidate["reviewState"] == "pending"
    # 待审阅**不是**待续写：不给续跑资格。
    assert candidate["canResume"] is False


def test_completed_record_without_review_state_is_not_a_candidate() -> None:
    """老记录（本能力上线前）：既不推成待审阅，也不推成已确认 —— 一条都不许多出来。"""
    client = TestClient(app)
    state = _register_office(client)
    _finish_office(client, state, _verified_result())
    # 模拟老记录：账本里那一列没有值。
    import sqlite3

    connection = sqlite3.connect(_ledger()._path())
    try:
        connection.execute(
            "UPDATE office_executions SET review_state = NULL WHERE operation_id = ?",
            (OFFICE_ID,),
        )
        connection.commit()
    finally:
        connection.close()

    response = client.get(
        "/api/office/executions/pending", params={"workbookIdentity": OFFICE_IDENTITY}
    )
    assert response.status_code == 200, response.text
    assert response.json()["candidates"] == []


def test_terminal_review_stops_being_discovered() -> None:
    client = TestClient(app)
    state = _register_office(client)
    _finish_office(client, state, _verified_result())
    assert (
        client.post(
            f"/api/office/executions/{OFFICE_ID}/review", json={"decision": "confirmed"}
        ).status_code
        == 200
    )
    response = client.get(
        "/api/office/executions/pending", params={"workbookIdentity": OFFICE_IDENTITY}
    )
    assert response.json()["candidates"] == []


def test_wps_review_endpoint_uses_its_own_namespace() -> None:
    """WPS 也有自己的一条审阅端点：同一份判据、只认 `wps-path:` 的记录。"""
    client = TestClient(app)
    state = _register_wps(client)
    evidence = {
        "sheetName": "Sheet1",
        "range": "B2:B2",
        "type": "writeValues",
        "beforeFingerprint": "v2:before-0",
    }
    started = client.post(
        f"/api/wps/executions/{WPS_ID}/actions/start",
        json={
            "expectedRevision": state["revision"],
            "ownerToken": state["ownerToken"],
            "actionIndex": 0,
            "beforeSnapshot": evidence,
        },
    )
    assert started.status_code == 200, started.text
    completed = client.post(
        f"/api/wps/executions/{WPS_ID}/actions/complete",
        json={
            "expectedRevision": started.json()["revision"],
            "ownerToken": state["ownerToken"],
            "actionIndex": 0,
            "checkpoint": {**evidence, "afterFingerprint": "v2:after-0"},
        },
    )
    assert completed.status_code == 200, completed.text
    finished = client.post(
        f"/api/wps/executions/{WPS_ID}/complete",
        json={
            "expectedRevision": completed.json()["revision"],
            "ownerToken": state["ownerToken"],
            "result": {"actionCount": 1, "verified": True},
        },
    )
    assert finished.status_code == 200, finished.text
    assert finished.json()["review"]["state"] == "pending"

    applied = client.post(f"/api/wps/executions/{WPS_ID}/review", json={"decision": "discarded"})
    assert applied.status_code == 200, applied.text
    assert applied.json()["review"]["state"] == "discarded"

    record = _ledger().office_execution(WPS_ID)
    assert record is not None
    assert record.review_state == "discarded"
