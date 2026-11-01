"""未完成执行的**后端发现**：浏览器索引丢了也能重新找到这次执行。

真实现场（2026-10-27 WPS）：第二步写入并回读后精确中断，用户保存、完整退出 WPS、
再打开同一个工作簿。账本里 `status=needs_reconciliation / nextActionIndex=1 /
inFlightActionIndex=1 / ownerActive=false`，原表仍是 111/222/-303 —— 进度和现场都在，
但窗格再也找不到这次执行：浏览器 `localStorage` 的单条发现索引是**唯一**入口。

这组用例钉的是新入口本身：按当前工作簿稳定身份**精确**查找、只读、多候选不自动选、
截断必须明说、处置只写发现元数据。写入资格与恢复判据一条都没放宽 —— 这里查到的只是
"有这么一次未完成执行"，认领与续跑仍走既有接口。
"""

from __future__ import annotations

import json
import sqlite3
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

import server.app.llm.config as llm_config
import server.app.main as main_module
from server.app.main import app
from server.app.turn_state import turn_registry
from server.app.task_ledger import capability_int

IDENTITY = "wps-path:c:/work/recovery/report.xlsx"


@pytest.fixture(autouse=True)
def isolate(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    turn_registry.reset()
    main_module._folder_operations.clear()
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path / "app-data"))
    for name in ("AI_API_KEY", "AI_BASE_URL", "AI_MODEL", "AI_MODELS", "AI_VISION_MODELS"):
        monkeypatch.delenv(name, raising=False)
    monkeypatch.setattr(llm_config, "_config_path", lambda: tmp_path / "config.env")


def _descriptor_actions(count: int = 3) -> list[dict]:
    return [
        {
            "actionIndex": index,
            "type": "writeValues",
            "sheet": "Sheet1",
            "range": f"B{index + 2}:B{index + 2}",
            "rowCount": 1,
            "columnCount": 1,
            "cellCount": 1,
            "previewFingerprint": f"v2:before-{index}",
            "expectedFingerprint": f"v2:after-{index}",
        }
        for index in range(count)
    ]


def _register(client: TestClient, operation_id: str, *, identity: str = IDENTITY, actions: int = 3) -> dict:
    payload = {
        "operationId": operation_id,
        "taskId": "task-recovery",
        "confirmed": True,
        "workbookIdentity": identity,
        "authorizedSheetNames": ["Sheet1"],
        "descriptor": {"actions": _descriptor_actions(actions)},
    }
    response = client.post("/api/wps/executions/register", json=payload)
    assert response.status_code == 200, response.text
    return response.json()


def _evidence(index: int, *, after: bool = False) -> dict:
    evidence = {
        "sheetName": "Sheet1",
        "range": f"B{index + 2}:B{index + 2}",
        "type": "writeValues",
        "beforeFingerprint": f"v2:before-{index}",
    }
    if after:
        evidence["afterFingerprint"] = f"v2:after-{index}"
    return evidence


def _interrupt_at(client: TestClient, state: dict, action_index: int) -> dict:
    """走真实端点把一次执行推到"第 N 步在途、回执未到"。

    先完成前 `action_index` 个动作，再对第 `action_index` 个动作登记写前状态后挂起 ——
    与用户现场（`needs_reconciliation` + `inFlightActionIndex`）同形。
    """

    current = state
    for index in range(action_index):
        current = client.post(
            f"/api/wps/executions/{state['operationId']}/actions/start",
            json={
                "expectedRevision": current["revision"],
                "ownerToken": current["ownerToken"],
                "actionIndex": index,
                "beforeSnapshot": _evidence(index),
            },
        ).json()
        current = client.post(
            f"/api/wps/executions/{state['operationId']}/actions/complete",
            json={
                "expectedRevision": current["revision"],
                "ownerToken": current["ownerToken"],
                "actionIndex": index,
                "checkpoint": _evidence(index, after=True),
            },
        ).json()
    current = client.post(
        f"/api/wps/executions/{state['operationId']}/actions/start",
        json={
            "expectedRevision": current["revision"],
            "ownerToken": current["ownerToken"],
            "actionIndex": action_index,
            "beforeSnapshot": _evidence(action_index),
        },
    ).json()
    suspended = client.post(
        f"/api/wps/executions/{state['operationId']}/suspend",
        json={
            "expectedRevision": current["revision"],
            "ownerToken": current["ownerToken"],
            "code": "WPS_WRITE_INTERRUPTED",
            "message": "WPS 写入中断，需要核对现场。",
        },
    )
    assert suspended.status_code == 200, suspended.text
    body = suspended.json()
    assert body["status"] == "needs_reconciliation"
    assert body["inFlightActionIndex"] == action_index
    return body


def _pending(client: TestClient, identity: str = IDENTITY) -> dict:
    response = client.get("/api/wps/executions/pending", params={"workbookIdentity": identity})
    assert response.status_code == 200, response.text
    return response.json()


def test_discovery_finds_the_in_flight_execution_without_any_local_index() -> None:
    """核心：本地索引清空后，只凭工作簿身份就能查到这次未完成执行。"""

    client = TestClient(app)
    registered = _register(client, "wps-execution-1fd938fd")
    interrupted = _interrupt_at(client, registered, 1)

    body = _pending(client)
    assert body["schemaVersion"] == 1
    assert body["namespace"] == "wps"
    assert body["workbookIdentity"] == IDENTITY
    assert body["truncated"] is False
    assert body["limit"] == capability_int("taskRecovery", "pendingDiscoveryMaxCandidates")
    assert [item["operationId"] for item in body["candidates"]] == ["wps-execution-1fd938fd"]
    candidate = body["candidates"][0]
    assert candidate["status"] == "needs_reconciliation"
    assert candidate["nextActionIndex"] == 1
    assert candidate["inFlightActionIndex"] == 1
    assert candidate["totalActions"] == 3
    assert candidate["ownerActive"] is False
    assert candidate["requiresReconciliation"] is True
    assert candidate["canResume"] is True
    assert candidate["taskId"] == "task-recovery"
    # 账本时间换成毫秒，客户端只拿它排序与显示，不参与归属判断。
    assert candidate["createdAt"] > 0
    assert candidate["updatedAt"] >= candidate["createdAt"]
    assert interrupted["revision"] >= 1

    # 只给定位、时间与进度：执行权令牌、计划、恢复材料一律不下发。
    text = json.dumps(body)
    for forbidden in ("ownerToken", "descriptor", "plan", "values", "packageJson", "111", "222"):
        assert forbidden not in text
    assert set(candidate) == {
        "operationId",
        "taskId",
        "status",
        "createdAt",
        "updatedAt",
        "nextActionIndex",
        "totalActions",
        "inFlightActionIndex",
        "ownerActive",
        "requiresReconciliation",
        "canResume",
    }


def test_same_file_name_on_another_path_is_never_discovered() -> None:
    """同名不同路径：身份必须完整字符串相等，不看文件名、不看"像不像"。"""

    client = TestClient(app)
    _register(client, "wps-execution-same-name")

    assert _pending(client, "wps-path:d:/backup/report.xlsx")["candidates"] == []
    assert _pending(client, "wps-path:c:/work/recovery/REPORT.xlsx")["candidates"] == []
    assert [item["operationId"] for item in _pending(client)["candidates"]] == [
        "wps-execution-same-name"
    ]


def test_completed_is_not_a_pending_task_and_a_write_before_failure_cannot_shadow_the_in_flight_one() -> None:
    """已完成不算待续写；旧的写前失败记录不许盖住真正的在途执行（哪怕它更新）。"""

    client = TestClient(app)
    completed_state = _register(client, "wps-execution-completed")
    current = completed_state
    for index in range(3):
        current = client.post(
            f"/api/wps/executions/{completed_state['operationId']}/actions/start",
            json={
                "expectedRevision": current["revision"],
                "ownerToken": completed_state["ownerToken"],
                "actionIndex": index,
                "beforeSnapshot": _evidence(index),
            },
        ).json()
        current = client.post(
            f"/api/wps/executions/{completed_state['operationId']}/actions/complete",
            json={
                "expectedRevision": current["revision"],
                "ownerToken": completed_state["ownerToken"],
                "actionIndex": index,
                "checkpoint": _evidence(index, after=True),
            },
        ).json()
    finished = client.post(
        f"/api/wps/executions/{completed_state['operationId']}/complete",
        json={
            "expectedRevision": current["revision"],
            "ownerToken": completed_state["ownerToken"],
            "result": {"actionCount": 3},
        },
    )
    assert finished.status_code == 200, finished.text

    in_flight = _register(client, "wps-execution-in-flight")
    _interrupt_at(client, in_flight, 1)

    # 最后登记、最后失败的一条：一个动作都没推进。它更新，但排最后。
    write_before = _register(client, "wps-execution-write-before")
    failed = client.post(
        f"/api/wps/executions/{write_before['operationId']}/suspend",
        json={
            "expectedRevision": write_before["revision"],
            "ownerToken": write_before["ownerToken"],
            "code": "WPS_WRITE_INTERRUPTED",
            "message": "WPS 写入中断，需要核对现场。",
        },
    )
    assert failed.status_code == 200, failed.text
    assert failed.json()["status"] == "failed"

    body = _pending(client)
    assert [item["operationId"] for item in body["candidates"]] == [
        "wps-execution-in-flight",
        "wps-execution-write-before",
    ]
    assert body["candidates"][1]["nextActionIndex"] == 0


def test_discovery_is_read_only_and_repeatable() -> None:
    """发现全过程只读：不认领、不补记、不推进 revision、不改 updated_at。"""

    client = TestClient(app)
    registered = _register(client, "wps-execution-readonly")
    _interrupt_at(client, registered, 1)

    database = Path(main_module._folder_operation_ledger._path())  # noqa: SLF001 - 测试要读原始行
    with sqlite3.connect(database) as connection:
        connection.row_factory = sqlite3.Row
        before = dict(
            connection.execute(
                "SELECT * FROM office_executions WHERE operation_id = ?",
                ("wps-execution-readonly",),
            ).fetchone()
        )

    first = _pending(client)
    second = _pending(client)
    assert first == second

    with sqlite3.connect(database) as connection:
        connection.row_factory = sqlite3.Row
        after = dict(
            connection.execute(
                "SELECT * FROM office_executions WHERE operation_id = ?",
                ("wps-execution-readonly",),
            ).fetchone()
        )
    assert after == before
    # 尤其不许偷偷换执行权令牌或推进修订号。
    assert after["revision"] == before["revision"]
    assert after["owner_token"] == before["owner_token"]


def test_truncation_is_explicit_and_the_limit_comes_from_capabilities() -> None:
    client = TestClient(app)
    limit = capability_int("taskRecovery", "pendingDiscoveryMaxCandidates")
    for index in range(limit + 2):
        state = _register(client, f"wps-execution-{index:04d}")
        _interrupt_at(client, state, 0)

    body = _pending(client)
    assert body["limit"] == limit
    assert len(body["candidates"]) == limit
    assert body["truncated"] is True


@pytest.mark.parametrize(
    "identity",
    [
        "report.xlsx",
        "",
        "url:c:/work/recovery/report.xlsx",
        "wps-path:",
        "wps-path:c:/work/recovery/report.xls",
        "wps-path:\\\\server\\share\\report.xlsx",
        "wps-path:c:/work/" + "x" * 5000 + ".xlsx",
    ],
)
def test_unreadable_or_illegal_identity_is_refused_instead_of_guessed(identity: str) -> None:
    client = TestClient(app)
    response = client.get("/api/wps/executions/pending", params={"workbookIdentity": identity})
    assert response.status_code == 422, response.text


def test_missing_query_parameter_is_refused() -> None:
    client = TestClient(app)
    assert client.get("/api/wps/executions/pending").status_code == 422


def test_dismissal_only_records_discovery_metadata() -> None:
    """清除恢复提示：只写发现元数据。

    不许伪造账本完成、不许删恢复材料、不许改动作进度；但同一条提示也不能在重开后被
    无限复活 —— 处置时间记在账本上，界面不必把"清除了提示"伪装成"任务已结束"。
    """

    client = TestClient(app)
    registered = _register(client, "wps-execution-dismiss")
    interrupted = _interrupt_at(client, registered, 1)
    assert len(_pending(client)["candidates"]) == 1

    dismissed = client.post("/api/wps/executions/wps-execution-dismiss/dismiss-discovery")
    assert dismissed.status_code == 204, dismissed.text
    assert _pending(client)["candidates"] == []
    # 幂等：再点一次不报错、也不会改别的东西。
    assert client.post("/api/wps/executions/wps-execution-dismiss/dismiss-discovery").status_code == 204

    state = client.get("/api/wps/executions/wps-execution-dismiss").json()
    assert state["status"] == interrupted["status"]
    assert state["nextActionIndex"] == interrupted["nextActionIndex"]
    assert state["inFlightActionIndex"] == interrupted["inFlightActionIndex"]
    assert state["revision"] == interrupted["revision"]
    assert state["ownerToken"] is None
    # 恢复材料与现场判据一个都没被"清除提示"顺手删掉：状态端点仍要求核对在途动作。
    assert state["requiresReconciliation"] is True


def test_dismiss_metadata_does_not_hide_new_activity_after_reconciliation() -> None:
    client = TestClient(app)
    registered = _register(client, "wps-execution-resurface")
    interrupted = _interrupt_at(client, registered, 1)
    client.post("/api/wps/executions/wps-execution-resurface/dismiss-discovery")
    assert _pending(client)["candidates"] == []

    reconciled = client.post(
        "/api/wps/executions/wps-execution-resurface/actions/reconcile",
        json={
            "expectedRevision": interrupted["revision"],
            "actionIndex": 1,
            "outcome": "applied",
            "confirmed": True,
            "checkpoint": _evidence(1, after=True),
        },
    )
    assert reconciled.status_code == 200, reconciled.text
    body = _pending(client)
    assert [item["operationId"] for item in body["candidates"]] == ["wps-execution-resurface"]
    assert body["candidates"][0]["inFlightActionIndex"] is None


def test_office_namespace_has_the_same_discovery_entry_and_does_not_mix_identities() -> None:
    """Excel 端同类缺口：后端发现能力同批就位，命名空间互不冒充。

    本轮只接线 WPS 面板；Excel 客户端接线与实机验收另行安排（计划表如实记录），
    但"只有 WPS 能发现"不是本轮想要的结论 —— 服务端两侧都能查。
    """

    client = TestClient(app)
    office = client.get("/api/office/executions/pending", params={"workbookIdentity": "url:c:/work/report.xlsx"})
    assert office.status_code == 200, office.text
    assert office.json()["namespace"] == "office"
    assert office.json()["candidates"] == []
    assert client.get(
        "/api/office/executions/pending", params={"workbookIdentity": IDENTITY}
    ).status_code == 422
    assert client.get(
        "/api/wps/executions/pending", params={"workbookIdentity": "url:c:/work/report.xlsx"}
    ).status_code == 422
