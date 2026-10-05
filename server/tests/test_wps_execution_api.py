from __future__ import annotations

import json
import os
import sqlite3
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

import server.app.main as main_module
import server.app.llm.config as llm_config
from server.app.main import app
from server.app.turn_state import turn_registry


@pytest.fixture(autouse=True)
def isolate(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    turn_registry.reset()
    main_module._folder_operations.clear()
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path / "app-data"))
    for name in (
        "AI_API_KEY",
        "AI_BASE_URL",
        "AI_MODEL",
        "AI_MODELS",
        "AI_VISION_MODELS",
        "OPENAI_API_KEY",
        "OPENAI_BASE_URL",
        "OPENAI_MODEL",
    ):
        monkeypatch.delenv(name, raising=False)
    monkeypatch.setattr(llm_config, "_config_path", lambda: tmp_path / "config.env")


def _wps_register_payload(operation_id: str = "wps-execution-0001") -> dict:
    """登记载荷只带**恢复描述** —— 服务端账本不收单元格值矩阵。"""
    return {
        "operationId": operation_id,
        "taskId": "task-keep",
        "confirmed": True,
        "workbookIdentity": "wps-path:c:/work/report.xlsx",
        "authorizedSheetNames": ["Sheet1"],
        "descriptor": {
            "actions": [
                {
                    "actionIndex": 0,
                    "type": "writeValues",
                    "sheet": "Sheet1",
                    "range": "A1:A1",
                    "rowCount": 1,
                    "columnCount": 1,
                    "cellCount": 1,
                    "previewFingerprint": "v2:fp-before",
                    "expectedFingerprint": "v2:fp-after",
                }
            ]
        },
    }


def _evidence(*, after: bool = False) -> dict:
    result = {"sheetName": "Sheet1", "range": "A1:A1", "type": "writeValues", "beforeFingerprint": "v2:fp-before"}
    if after:
        result["afterFingerprint"] = "v2:fp-after"
    return result


def test_wps_execution_register_grants_owner_and_persists_task_id() -> None:
    client = TestClient(app)
    response = client.post("/api/wps/executions/register", json=_wps_register_payload())
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["ownerToken"]
    assert body["taskId"] == "task-keep"
    assert body["workbookIdentity"] == "wps-path:c:/work/report.xlsx"
    assert body["status"] == "running"

    # 下发状态只回描述，不回 `plan`，也不回任何矩阵字段。
    assert "plan" not in body
    assert "values" not in body["descriptor"]["actions"][0]
    assert body["totalActions"] == 1

    # 状态查询不会返回 ownerToken（不泄露写入权）。
    status = client.get("/api/wps/executions/wps-execution-0001").json()
    assert status["ownerToken"] is None
    assert status["taskId"] == "task-keep"
    assert "plan" not in status


def test_wps_execution_persists_no_cell_values(tmp_path: Path) -> None:
    """落盘的 `office_executions.plan_json` 里**没有**单元格内容。

    先证明哨兵确实存在于本地计划里（锚点），再证明它没进请求、也没进库。
    少了锚点，"库里没有"可能只是因为压根没传过。
    """
    client = TestClient(app)
    sentinel = "SENTINEL-CONFIDENTIAL-7f3a91c0"
    local_plan = {
        "id": "plan-wps",
        "title": "写入区域",
        "summary": f"写入 {sentinel}",
        "assumptions": [],
        "warnings": [],
        "actions": [
            {"type": "writeValues", "sheet": "Sheet1", "range": "A1:A1", "values": [[sentinel]]}
        ],
    }
    # 锚点：本机计划里确实有这个哨兵。
    assert sentinel in json.dumps(local_plan, ensure_ascii=False)

    payload = _wps_register_payload("wps-execution-privacy")
    assert sentinel not in json.dumps(payload, ensure_ascii=False)
    response = client.post("/api/wps/executions/register", json=payload)
    assert response.status_code == 200, response.text

    # 直接读持久账本，不看接口说了什么。
    database = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "tasks" / "state.sqlite3"
    assert database.exists(), "应持久化到受控 SQLite 账本"
    connection = sqlite3.connect(database)
    try:
        row = connection.execute(
            "SELECT plan_json FROM office_executions WHERE operation_id = ?",
            ("wps-execution-privacy",),
        ).fetchone()
    finally:
        connection.close()
    assert row is not None
    assert sentinel not in str(row[0])
    assert "values" not in str(row[0])


def test_wps_execution_refuses_cell_matrix_in_descriptor() -> None:
    """恢复描述里只要出现矩阵字段就 422 —— 不是"忽略后照常登记"。"""
    client = TestClient(app)
    payload = _wps_register_payload("wps-execution-0009")
    payload["descriptor"]["actions"][0]["values"] = [["机密"]]
    response = client.post("/api/wps/executions/register", json=payload)
    assert response.status_code == 422


def test_wps_execution_refuses_unsupported_action_variants() -> None:
    client = TestClient(app)
    payload = _wps_register_payload("wps-execution-0010")
    payload["descriptor"]["actions"][0]["type"] = "clearRange"
    payload["descriptor"]["actions"][0]["applyTo"] = "formats"
    response = client.post("/api/wps/executions/register", json=payload)
    assert response.status_code == 422


def test_wps_execution_refuses_office_url_identity() -> None:
    client = TestClient(app)
    payload = _wps_register_payload("wps-execution-0002")
    payload["workbookIdentity"] = "url:c:/work/report.xlsx"
    response = client.post("/api/wps/executions/register", json=payload)
    assert response.status_code == 422


def test_wps_execution_refuses_unprefixed_identity() -> None:
    client = TestClient(app)
    payload = _wps_register_payload("wps-execution-0003")
    payload["workbookIdentity"] = "c:/work/report.xlsx"
    response = client.post("/api/wps/executions/register", json=payload)
    assert response.status_code == 422


def test_wps_execution_rejects_stale_owner_token() -> None:
    client = TestClient(app)
    registered = client.post("/api/wps/executions/register", json=_wps_register_payload()).json()
    owner = registered["ownerToken"]

    start = client.post(
        "/api/wps/executions/wps-execution-0001/actions/start",
        json={
            "expectedRevision": registered["revision"],
            "ownerToken": owner,
            "actionIndex": 0,
            "beforeSnapshot": _evidence(),
        },
    )
    assert start.status_code == 200, start.text
    next_revision = start.json()["revision"]

    # 旧 revision + 旧 ownerToken 的动作回执 → 冲突，不能推进第二次写入。
    complete = client.post(
        "/api/wps/executions/wps-execution-0001/actions/complete",
        json={
            "expectedRevision": registered["revision"],
            "ownerToken": owner,
            "actionIndex": 0,
            "checkpoint": _evidence(after=True),
        },
    )
    assert complete.status_code == 409

    # 用新 revision 才能推进。
    complete_ok = client.post(
        "/api/wps/executions/wps-execution-0001/actions/complete",
        json={
            "expectedRevision": next_revision,
            "ownerToken": start.json()["ownerToken"],
            "actionIndex": 0,
            "checkpoint": _evidence(after=True),
        },
    )
    assert complete_ok.status_code == 200, complete_ok.text
    assert complete_ok.json()["nextActionIndex"] == 1


def _expire_owner() -> None:
    database = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "tasks" / "state.sqlite3"
    with sqlite3.connect(database) as connection:
        connection.execute("UPDATE office_executions SET updated_at = 0")


def test_wps_active_owner_cannot_be_stolen_or_replayed() -> None:
    client = TestClient(app)
    registered = client.post("/api/wps/executions/register", json=_wps_register_payload()).json()
    assert registered["ownerActive"] is True
    assert registered["canResume"] is False
    replay = client.post("/api/wps/executions/register", json=_wps_register_payload()).json()
    assert replay["ownerToken"] is None
    response = client.post("/api/wps/executions/wps-execution-0001/resume",
        json={"expectedRevision": registered["revision"], "confirmed": True})
    assert response.status_code == 409
    assert client.get("/api/wps/executions/wps-execution-0001").json()["revision"] == registered["revision"]


def test_office_routes_cannot_bypass_wps_owner_guard() -> None:
    client = TestClient(app)
    registered = client.post("/api/wps/executions/register", json=_wps_register_payload()).json()
    assert client.get("/api/office/executions/wps-execution-0001").status_code == 404
    assert client.post("/api/office/executions/wps-execution-0001/resume", json={
        "expectedRevision": registered["revision"], "confirmed": True}).status_code == 409
    assert client.get("/api/wps/executions/wps-execution-0001").json()["revision"] == registered["revision"]


def test_wps_cannot_finish_with_remaining_actions() -> None:
    client = TestClient(app)
    registered = client.post("/api/wps/executions/register", json=_wps_register_payload()).json()
    response = client.post("/api/wps/executions/wps-execution-0001/complete", json={
        "expectedRevision": registered["revision"], "ownerToken": registered["ownerToken"], "result": {"actionCount": 1}})
    assert response.status_code == 409
    assert client.get("/api/wps/executions/wps-execution-0001").json()["nextActionIndex"] == 0


def test_wps_unknown_query_creates_no_execution() -> None:
    client = TestClient(app)
    assert client.get("/api/wps/executions/wps-execution-unknown").status_code == 404
    database = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "tasks" / "state.sqlite3"
    with sqlite3.connect(database) as connection:
        assert connection.execute("SELECT COUNT(*) FROM office_executions").fetchone()[0] == 0


@pytest.mark.parametrize("column", ["plan_json", "before_snapshot_json"])
def test_legacy_sensitive_record_fails_closed_without_echo(column: str) -> None:
    client = TestClient(app)
    registered = client.post("/api/wps/executions/register", json=_wps_register_payload()).json()
    sentinel = "LEGACY-CELL-SECRET-193dd"
    legacy = {"actions": [{"values": [[sentinel]]}]} if column == "plan_json" else {"values": [[sentinel]]}
    assert sentinel in json.dumps(legacy)
    database = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "tasks" / "state.sqlite3"
    with sqlite3.connect(database) as connection:
        # The column name is parametrized from a fixed test allowlist, not input.
        connection.execute(f"UPDATE office_executions SET {column} = ?", (json.dumps(legacy),))
    for response in [client.get("/api/wps/executions/wps-execution-0001"),
        client.post("/api/wps/executions/wps-execution-0001/resume", json={
            "expectedRevision": registered["revision"], "confirmed": True})]:
        assert response.status_code != 200
        assert sentinel not in response.text


def test_wps_lost_final_receipt_only_finishes_after_idle_claim() -> None:
    client = TestClient(app)
    registered = client.post("/api/wps/executions/register", json=_wps_register_payload()).json()
    started = client.post("/api/wps/executions/wps-execution-0001/actions/start", json={
        "expectedRevision": registered["revision"], "ownerToken": registered["ownerToken"],
        "actionIndex": 0, "beforeSnapshot": _evidence()}).json()
    completed = client.post("/api/wps/executions/wps-execution-0001/actions/complete", json={
        "expectedRevision": started["revision"], "ownerToken": started["ownerToken"],
        "actionIndex": 0, "checkpoint": _evidence(after=True)}).json()
    assert completed["nextActionIndex"] == completed["totalActions"] == 1
    _expire_owner()
    status = client.get("/api/wps/executions/wps-execution-0001").json()
    assert status["ownerActive"] is False and status["canResume"] is True
    claim = client.post("/api/wps/executions/wps-execution-0001/resume", json={
        "expectedRevision": status["revision"], "confirmed": True}).json()
    assert claim["ownerToken"] != registered["ownerToken"]
    finished = client.post("/api/wps/executions/wps-execution-0001/complete", json={
        "expectedRevision": claim["revision"], "ownerToken": claim["ownerToken"],
        "result": {"actionCount": 1, "backfilled": True}})
    assert finished.status_code == 200, finished.text
    assert finished.json()["status"] == "completed"
    assert finished.json()["canResume"] is False


def test_wps_inflight_reconcile_requires_idle_and_exact_checkpoint() -> None:
    client = TestClient(app)
    registered = client.post("/api/wps/executions/register", json=_wps_register_payload()).json()
    started = client.post("/api/wps/executions/wps-execution-0001/actions/start", json={
        "expectedRevision": registered["revision"], "ownerToken": registered["ownerToken"],
        "actionIndex": 0, "beforeSnapshot": _evidence()}).json()
    payload = {"expectedRevision": started["revision"], "actionIndex": 0, "outcome": "applied",
        "confirmed": True, "checkpoint": _evidence(after=True)}
    assert client.post("/api/wps/executions/wps-execution-0001/actions/reconcile", json=payload).status_code == 409
    _expire_owner()
    payload["checkpoint"]["afterFingerprint"] = "wrong"
    assert client.post("/api/wps/executions/wps-execution-0001/actions/reconcile", json=payload).status_code == 422
    payload["checkpoint"] = _evidence(after=True)
    response = client.post("/api/wps/executions/wps-execution-0001/actions/reconcile", json=payload)
    assert response.status_code == 200, response.text
    assert response.json()["nextActionIndex"] == 1
    assert client.post("/api/wps/executions/wps-execution-0001/actions/reconcile", json=payload).status_code == 409


@pytest.mark.parametrize("route,field", [("actions/start", "beforeSnapshot"),
    ("actions/complete", "checkpoint"), ("actions/reconcile", "checkpoint"), ("complete", "result")])
def test_wps_no_matrix_in_any_persisted_receipt(route: str, field: str) -> None:
    client = TestClient(app)
    local_plan = {"actions": [{"values": [["PRIVATE-RECEIPT-SENTINEL"]]}]}
    assert "PRIVATE-RECEIPT-SENTINEL" in json.dumps(local_plan)
    registered = client.post("/api/wps/executions/register", json=_wps_register_payload()).json()
    payload = {"expectedRevision": registered["revision"], "ownerToken": registered["ownerToken"],
        "actionIndex": 0, field: {**_evidence(after=True), "values": local_plan["actions"][0]["values"]}}
    if route == "actions/reconcile":
        payload.pop("ownerToken")
        payload.update(outcome="applied", confirmed=True)
    if route == "complete":
        payload.pop("actionIndex")
        payload[field] = {"actionCount": 1, "values": local_plan["actions"][0]["values"]}
    response = client.post(f"/api/wps/executions/wps-execution-0001/{route}", json=payload)
    assert response.status_code == 422, response.text
    database = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "tasks" / "state.sqlite3"
    with sqlite3.connect(database) as connection:
        row = connection.execute("SELECT plan_json, before_snapshot_json, checkpoint_json, result_json FROM office_executions").fetchone()
    assert row is not None
    assert "PRIVATE-RECEIPT-SENTINEL" not in str(row)
    assert client.get("/api/wps/executions/wps-execution-0001").json()["revision"] == registered["revision"]
