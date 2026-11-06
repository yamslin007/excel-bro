from __future__ import annotations

import json
import os
import re
import sqlite3
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

import server.app.main as main_module
import server.app.llm.config as llm_config
from server.app.main import app
from server.app.models import WPS_FORMAT_FIELD_NAMES, WPS_RECOVERY_ACTION_TYPES
from server.app.turn_state import turn_registry

REPO_ROOT = Path(__file__).resolve().parents[2]


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


def test_wps_suspend_persists_caller_supplied_error() -> None:
    """挂起落账的是调用方给的真实原因，不是服务端写死的通用文案。

    2026-11-01 真机（WPS 公式 B 首轮）账本里只剩"WPS 写入中断，需要核对现场。"，
    事后既查不出最初失败点，也分不清零写入与写到一半。请求体本就要求 `code` /
    `message` 必填，服务端不该覆盖它。
    """
    client = TestClient(app)
    registered = client.post("/api/wps/executions/register", json=_wps_register_payload()).json()
    suspended = client.post(
        "/api/wps/executions/wps-execution-0001/suspend",
        json={
            "expectedRevision": registered["revision"],
            "ownerToken": registered["ownerToken"],
            "code": "WPS_RECOVERY_PROBE_CANCELLED",
            "message": "恢复验收已取消：本次计划未登记可恢复的执行。",
        },
    )
    assert suspended.status_code == 200, suspended.text
    body = suspended.json()
    assert body["status"] == "failed"
    # 挂起后不泄露写入权。
    assert body["ownerToken"] is None
    # 真实原因原样进入下发的 error。
    assert body["error"]["code"] == "WPS_RECOVERY_PROBE_CANCELLED"
    assert body["error"]["message"] == "恢复验收已取消：本次计划未登记可恢复的执行。"

    # 落库的 error_json 同样是调用方给的那份（不是通用文案）。
    database = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "tasks" / "state.sqlite3"
    with sqlite3.connect(database) as connection:
        stored = connection.execute(
            "select error_json from office_executions where operation_id = ?",
            ("wps-execution-0001",),
        ).fetchone()[0]
    assert json.loads(stored)["code"] == "WPS_RECOVERY_PROBE_CANCELLED"
    assert json.loads(stored)["message"] == "恢复验收已取消：本次计划未登记可恢复的执行。"


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


# ---------------------------------------------------------------------------
# 格式动作（setFont / setAlignment / resizeRange）的恢复账本链路
# ---------------------------------------------------------------------------
#
# 这一组用例复现的是真实 WPS 验收里失败的那条链路：格式动作在**登记**这一步就被拒，
# 界面提示「无法持久登记本次写入现场，已取消写入；原表未被修改。WPS 恢复描述的动作
# 定位或形状无效」。直接原因是三处动作白名单只认三个值动作，格式动作在 fetch 之前
# （`contracts.ts`）或登记之前（`main.py`）就没了。
#
# 指纹用 `f1:` 前缀的**格式证据指纹**：格式动作不改值，值矩阵指纹在写前写后完全
# 一样，对"到底改没改到"没有区分力。这里传的是与执行器同一口径的字符串形态；服务端
# 只做命名空间与属性级校验 —— 它算不出也不需要算出真实指纹，那要求上传格式快照，
# 而快照含逐格属性值，正是协议不许出站的东西。

_WPS_FORMAT_ACTIONS: list[dict] = [
    {
        "actionIndex": 0,
        "type": "setFont",
        "sheet": "Sheet1",
        "range": "A2:B3",
        "rowCount": 2,
        "columnCount": 2,
        "cellCount": 4,
        "previewFingerprint": "f1:88:0a0a0a0a",
        "expectedFingerprint": "f1:88:1b1b1b1b",
        "formatFields": ["bold", "size"],
    },
    {
        "actionIndex": 1,
        "type": "setAlignment",
        "sheet": "Sheet1",
        "range": "A2:B3",
        "rowCount": 2,
        "columnCount": 2,
        "cellCount": 4,
        "previewFingerprint": "f1:72:2c2c2c2c",
        "expectedFingerprint": "f1:72:3d3d3d3d",
        "formatFields": ["horizontalAlignment"],
    },
    {
        # 行高是**行级**动作：两行两列，写入只落在整行上。
        "actionIndex": 2,
        "type": "resizeRange",
        "sheet": "Sheet1",
        "range": "A3:B4",
        "rowCount": 2,
        "columnCount": 2,
        "cellCount": 4,
        "previewFingerprint": "f1:60:4e4e4e4e",
        "expectedFingerprint": "f1:60:5f5f5f5f",
        "formatFields": ["rowHeight"],
    },
]


def _wps_format_register_payload(operation_id: str = "wps-format-0001") -> dict:
    """格式动作的登记载荷：三个动作共 12 格，`f1:` 格式证据指纹 + 属性名。

    载荷里**没有**任何快照字段 —— 只有"改了哪些属性"的名字，没有属性取值。
    """
    return {
        "operationId": operation_id,
        "taskId": "task-format",
        "confirmed": True,
        "workbookIdentity": "wps-path:c:/work/report.xlsx",
        "authorizedSheetNames": ["Sheet1"],
        "descriptor": {"actions": [dict(action) for action in _WPS_FORMAT_ACTIONS]},
    }


def _format_evidence(action: dict, *, after: bool = False, fingerprint: str | None = None) -> dict:
    evidence = {
        "sheetName": action["sheet"],
        "range": action["range"],
        "type": action["type"],
        "beforeFingerprint": action["previewFingerprint"],
    }
    if after:
        evidence["afterFingerprint"] = fingerprint or action["expectedFingerprint"]
    return evidence


def _wps_single_format_payload(action: dict, operation_id: str) -> dict:
    """只含一个格式动作的登记载荷（动作必须落在 0 号位：在途动作只能是下一个待办）。

    用于"这个动作自己的证据判据"这类用例 —— 要让 2 号动作成为在途动作，得先把 0、
    1 号走完；那属于另一条用例（`test_wps_format_actions_register_...`）的职责。
    """
    payload = _wps_format_register_payload(operation_id)
    payload["descriptor"]["actions"] = [{**action, "actionIndex": 0}]
    return payload


def test_wps_format_actions_register_start_complete_and_finish() -> None:
    """格式动作在**真实服务端账本**上跑通 登记 → 逐动作写前检查点 → 写后回执 → 完成。"""
    client = TestClient(app)
    registered = client.post("/api/wps/executions/register", json=_wps_format_register_payload())
    assert registered.status_code == 200, registered.text
    state = registered.json()
    assert state["totalActions"] == 3
    owner = state["ownerToken"]
    revision = state["revision"]

    for index, action in enumerate(_WPS_FORMAT_ACTIONS):
        started = client.post(
            "/api/wps/executions/wps-format-0001/actions/start",
            json={
                "expectedRevision": revision,
                "ownerToken": owner,
                "actionIndex": index,
                "beforeSnapshot": _format_evidence(action),
            },
        )
        assert started.status_code == 200, f"第 {index + 1} 步写前检查点被拒：{started.text}"
        owner = started.json()["ownerToken"]
        revision = started.json()["revision"]
        completed = client.post(
            "/api/wps/executions/wps-format-0001/actions/complete",
            json={
                "expectedRevision": revision,
                "ownerToken": owner,
                "actionIndex": index,
                "checkpoint": _format_evidence(action, after=True),
            },
        )
        assert completed.status_code == 200, f"第 {index + 1} 步写后回执被拒：{completed.text}"
        owner = completed.json()["ownerToken"]
        revision = completed.json()["revision"]
        assert completed.json()["nextActionIndex"] == index + 1

    finished = client.post(
        "/api/wps/executions/wps-format-0001/complete",
        json={
            "expectedRevision": revision,
            "ownerToken": owner,
            "result": {"actionCount": 3, "cellCount": 12, "backfilled": True},
        },
    )
    assert finished.status_code == 200, finished.text
    assert finished.json()["status"] == "completed"


def test_wps_format_status_discloses_only_property_names() -> None:
    """状态读取认识格式动作，但下发的只有**属性名**，没有任何快照或取值。"""
    client = TestClient(app)
    payload = _wps_format_register_payload("wps-format-status")
    assert client.post("/api/wps/executions/register", json=payload).status_code == 200

    status = client.get("/api/wps/executions/wps-format-status")
    assert status.status_code == 200, status.text
    body = status.json()
    actions = body["descriptor"]["actions"]
    assert [action["formatFields"] for action in actions] == [
        ["bold", "size"],
        ["horizontalAlignment"],
        ["rowHeight"],
    ]
    assert "plan" not in body
    # 下发的载荷里不许出现快照 / 取值载体。
    serialized = json.dumps(body, ensure_ascii=False)
    for forbidden in ("cells", "rowHeights", "values", "formulas", "snapshot"):
        assert forbidden not in serialized, f"状态下发里出现了不该出站的字段：{forbidden}"
    # 直接读持久账本，不看接口说了什么。
    database = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "tasks" / "state.sqlite3"
    with sqlite3.connect(database) as connection:
        row = connection.execute(
            "SELECT plan_json FROM office_executions WHERE operation_id = ?",
            ("wps-format-status",),
        ).fetchone()
    assert row is not None
    persisted = str(row[0])
    for forbidden in ("cells", "rowHeights", "values", "formulas"):
        assert forbidden not in persisted, f"落盘账本里出现了不该出站的字段：{forbidden}"
    assert "horizontalAlignment" in persisted  # 属性名本身是登记必须的信息。


@pytest.mark.parametrize("index", [0, 1, 2])
def test_wps_format_checkpoint_after_fingerprint_must_match_target(index: int) -> None:
    """写后回执必须精确等于登记的目标指纹 —— 多一格、少一格、换一个现场都拒。

    三种现场（写前 / 写后 / 都不是）在指纹上是三个不同的值：写前指纹能通过写前
    检查点，写后目标能通过写后回执，而"都不是"（= 期间被外部改过，或部分写入）
    两个都过不了。
    """
    client = TestClient(app)
    action = _WPS_FORMAT_ACTIONS[index]
    registered = client.post(
        "/api/wps/executions/register",
        json=_wps_single_format_payload(action, "wps-format-drift"),
    ).json()
    started = client.post(
        "/api/wps/executions/wps-format-drift/actions/start",
        json={
            "expectedRevision": registered["revision"],
            "ownerToken": registered["ownerToken"],
            "actionIndex": 0,
            "beforeSnapshot": _format_evidence(action),
        },
    )
    assert started.status_code == 200, started.text

    def complete(fingerprint: str) -> int:
        response = client.post(
            "/api/wps/executions/wps-format-drift/actions/complete",
            json={
                "expectedRevision": started.json()["revision"],
                "ownerToken": started.json()["ownerToken"],
                "actionIndex": 0,
                "checkpoint": _format_evidence(action, after=True, fingerprint=fingerprint),
            },
        )
        return response.status_code

    # 外部改过 / 部分写入：既不是写前现场，也不是登记的目标现场。
    assert complete("f1:88:deadbeef") == 422
    # 写前现场冒充写后：格式动作不改值，值矩阵指纹恰好会一直"自洽"，这条正是它挡不住的。
    assert complete(action["previewFingerprint"]) == 422
    # 精确的目标指纹才通过。
    assert complete(action["expectedFingerprint"]) == 200


def test_wps_format_action_reconcile_and_resume_recognize_formats() -> None:
    """失败恢复与对账也走同一套格式证据，不是只认值动作。"""
    client = TestClient(app)
    action = _WPS_FORMAT_ACTIONS[2]  # 行高动作：对账/恢复同样按整行语义走
    registered = client.post(
        "/api/wps/executions/register",
        json=_wps_single_format_payload(action, "wps-format-recover"),
    ).json()
    started = client.post(
        "/api/wps/executions/wps-format-recover/actions/start",
        json={
            "expectedRevision": registered["revision"],
            "ownerToken": registered["ownerToken"],
            "actionIndex": 0,
            "beforeSnapshot": _format_evidence(action),
        },
    ).json()
    assert started["inFlightActionIndex"] == 0

    def reconcile(fingerprint: str) -> int:
        return client.post(
            "/api/wps/executions/wps-format-recover/actions/reconcile",
            json={
                "expectedRevision": started["revision"],
                "actionIndex": 0,
                "outcome": "applied",
                "confirmed": True,
                "checkpoint": _format_evidence(action, after=True, fingerprint=fingerprint),
            },
        ).status_code

    # 主人的租约还在 ⇒ 先要空闲；不能凭"我记得写了"就把在途动作记成已应用。
    assert reconcile(action["expectedFingerprint"]) == 409
    _expire_owner()
    # 空下来之后，错误的写后格式指纹照样不许把在途动作记成已应用。
    assert reconcile("f1:60:deadbeef") == 422
    assert reconcile(action["previewFingerprint"]) == 422
    reconciled = client.post(
        "/api/wps/executions/wps-format-recover/actions/reconcile",
        json={
            "expectedRevision": started["revision"],
            "actionIndex": 0,
            "outcome": "applied",
            "confirmed": True,
            "checkpoint": _format_evidence(action, after=True),
        },
    )
    assert reconciled.status_code == 200, reconciled.text
    assert reconciled.json()["nextActionIndex"] == 1

    # 中断 → 状态可续跑 → 重新接管，全程都认得格式动作。
    suspended = client.post(
        "/api/wps/executions/wps-format-recover/suspend",
        json={
            "expectedRevision": reconciled.json()["revision"],
            "ownerToken": reconciled.json()["ownerToken"],
            "code": "WPS_WRITE_INTERRUPTED",
            "message": "宿主在写回过程中被关闭。",
        },
    )
    assert suspended.status_code == 200, suspended.text
    _expire_owner()
    status = client.get("/api/wps/executions/wps-format-recover").json()
    assert status["canResume"] is True
    assert status["descriptor"]["actions"][0]["formatFields"] == ["rowHeight"]
    claimed = client.post(
        "/api/wps/executions/wps-format-recover/resume",
        json={"expectedRevision": status["revision"], "confirmed": True},
    )
    assert claimed.status_code == 200, claimed.text
    assert claimed.json()["descriptor"]["actions"][0]["formatFields"] == ["rowHeight"]


@pytest.mark.parametrize(
    ("action_index", "mutation", "expected_status"),
    [
        # 格式动作没声明属性：指纹是按什么算的变得不可判定。
        (0, {"formatFields": []}, 422),
        (0, {"formatFields": None}, 422),
        # 值动作声明格式属性：两个命名空间混在一起。
        (0, {"type": "writeValues", "formatFields": ["bold"]}, 422),
    ],
)
def test_wps_format_property_declaration_is_required_and_exclusive(
    action_index: int, mutation: dict, expected_status: int
) -> None:
    client = TestClient(app)
    payload = _wps_format_register_payload("wps-format-decl")
    payload["descriptor"]["actions"][action_index].update(mutation)
    response = client.post("/api/wps/executions/register", json=payload)
    assert response.status_code == expected_status, response.text


@pytest.mark.parametrize(
    ("action_type", "declared"),
    [
        # 未通过实机验证的属性：哨兵没过，不许"忽略后照常登记"。
        ("setFont", ["bold", "underline"]),
        ("setAlignment", ["vertical"]),
        ("resizeRange", ["rowHeight", "columnWidth"]),
        # 属性名合法，但不属于这个动作。
        ("setFont", ["rowHeight"]),
        ("setAlignment", ["bold"]),
        ("resizeRange", ["horizontalAlignment"]),
    ],
)
def test_wps_format_action_rejects_properties_outside_its_whitelist(
    action_type: str, declared: list[str]
) -> None:
    client = TestClient(app)
    payload = _wps_format_register_payload("wps-format-props")
    payload["descriptor"]["actions"][0].update({"type": action_type, "formatFields": declared})
    response = client.post("/api/wps/executions/register", json=payload)
    assert response.status_code == 422, response.text


@pytest.mark.parametrize(
    "mutation",
    [
        # 忘了 expectedFingerprint：格式动作的"写后目标"缺了，写后回执无从核对。
        {"expectedFingerprint": None},
        # 用值矩阵指纹登记格式动作：值矩阵指纹在写前写后一样，没有区分力。
        {"previewFingerprint": "v2:20:0a0a0a0a", "expectedFingerprint": "v2:20:1b1b1b1b"},
        # 一半格式、一半值，同样是不可判定的混合体。
        {"expectedFingerprint": "v2:20:1b1b1b1b"},
    ],
)
def test_wps_format_action_requires_format_evidence_fingerprints(mutation: dict) -> None:
    client = TestClient(app)
    payload = _wps_format_register_payload("wps-format-fp")
    payload["descriptor"]["actions"][0].update(mutation)
    response = client.post("/api/wps/executions/register", json=payload)
    assert response.status_code == 422, response.text


def test_wps_value_action_requires_value_fingerprints() -> None:
    """反向：值动作不许拿格式证据指纹登记（命名空间不能互相冒充）。"""
    client = TestClient(app)
    payload = _wps_register_payload("wps-value-fp")
    payload["descriptor"]["actions"][0]["previewFingerprint"] = "f1:88:0a0a0a0a"
    assert client.post("/api/wps/executions/register", json=payload).status_code == 422


@pytest.mark.parametrize("field", ["cells", "rowHeights", "snapshot", "formatSnapshot"])
def test_wps_format_descriptor_refuses_snapshot_smuggling(field: str) -> None:
    """格式快照不许出站 —— 它不是"用格式指纹绕过协议"的许可。

    判据落在**落盘**上，不落在回显上：422 的响应体本来就会把请求原样带回去（数据
    来自调用方自己，没有跨信任边界），而会话记录的权威副本是 SQLite 账本 ——
    看不进账本才算真的没被接收。
    """
    client = TestClient(app)
    sentinel = "SNEAKED-FORMAT-SNAPSHOT-2b7c14"
    payload = _wps_format_register_payload("wps-format-snap")
    payload["descriptor"]["actions"][0][field] = [[sentinel]]
    assert sentinel in json.dumps(payload, ensure_ascii=False)  # 锚点：确实传过
    response = client.post("/api/wps/executions/register", json=payload)
    assert response.status_code == 422, response.text
    database = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "tasks" / "state.sqlite3"
    if database.exists():
        with sqlite3.connect(database) as connection:
            count = connection.execute("SELECT COUNT(*) FROM office_executions").fetchone()
        assert count is not None and count[0] == 0, "被拒的登记不得在账本里留下任何记录"


@pytest.mark.parametrize("action_type", ["setFill", "setNumberFormat", "setBorders", "autofit"])
def test_wps_formats_refuse_invalid_field_claims(action_type: str) -> None:
    """端上没接通的动作（fail-closed 拒绝）连登记都不该放行。

    它们曾经被错误地列在服务端白名单里 —— "能登记"读起来像"能执行"。
    """
    client = TestClient(app)
    payload = _wps_format_register_payload("wps-format-unopened")
    payload["descriptor"]["actions"][0].update({"type": action_type, "formatFields": ["bold"]})
    response = client.post("/api/wps/executions/register", json=payload)
    assert response.status_code == 422, response.text


_TS_CONTRACTS = REPO_ROOT / "apps" / "excel-addin" / "src" / "contracts.ts"


def _ts_recoverable_action_types() -> list[str]:
    source = _TS_CONTRACTS.read_text(encoding="utf-8")
    match = re.search(
        r"export const WPS_RECOVERABLE_ACTION_TYPES\s*=\s*\[(?P<body>[^\]]*)\]\s*as const;",
        source,
        re.S,
    )
    assert match is not None, "contracts.ts 里找不到 WPS_RECOVERABLE_ACTION_TYPES"
    return re.findall(r'"([^"]+)"', match.group("body"))


def _ts_format_action_fields() -> dict[str, list[str]]:
    source = _TS_CONTRACTS.read_text(encoding="utf-8")
    match = re.search(
        r"export const WPS_FORMAT_ACTION_FIELDS\s*:\s*Record<string,\s*readonly string\[\]>\s*=\s*\{(?P<body>.*?)\n\};",
        source,
        re.S,
    )
    assert match is not None, "contracts.ts 里找不到 WPS_FORMAT_ACTION_FIELDS"
    return {
        name: re.findall(r'"([^"]+)"', body)
        for name, body in re.findall(r"(\w+)\s*:\s*\[([^\]]*)\]", match.group("body"))
    }


def test_wps_action_whitelist_matches_capabilities() -> None:
    """四份动作/属性白名单必须逐字一致（任一处漂移都是真实故障形态）。

    `config/capabilities.json` 是**客户端打包**用的那一份（`wpsPlanExecutor` 直接
    `import` 它）；`main._WPS_RECOVERABLE_ACTION_TYPES` 是服务端闸；`contracts.ts`
    是出站契约闸；`models.WPS_RECOVERY_ACTION_TYPES` 是模型层闸。少同步一处就会出现
    「端上放行、服务端 422」（本轮的故障）或反过来「服务端放行、端上根本执行不了」。
    """
    capabilities = json.loads(
        (REPO_ROOT / "config" / "capabilities.json").read_text(encoding="utf-8")
    )
    # `clearRange:contents` 的 `:contents` 是变体细分，不在动作类型命名空间里。
    from_capabilities = {
        name.split(":", 1)[0] for name in capabilities["wpsWrite"]["allowedActions"]
    }
    expansions = capabilities["wpsWrite"].get("actionExpansion", {})
    assert set(expansions) <= from_capabilities
    for primitives in expansions.values():
        assert set(primitives) <= set(WPS_RECOVERY_ACTION_TYPES)
    assert from_capabilities - set(expansions) == set(WPS_RECOVERY_ACTION_TYPES)
    assert set(main_module._WPS_RECOVERABLE_ACTION_TYPES) == set(WPS_RECOVERY_ACTION_TYPES)
    assert set(_ts_recoverable_action_types()) == set(WPS_RECOVERY_ACTION_TYPES)

    ts_fields = _ts_format_action_fields()
    assert set(ts_fields) == set(main_module._WPS_FORMAT_ACTION_PROPERTIES)
    for action_type, names in ts_fields.items():
        assert len(names) == len(set(names)), f"{action_type} 的属性名有重复"
        assert set(names) == set(
            main_module._WPS_FORMAT_ACTION_PROPERTIES[action_type]
        ), f"{action_type} 的属性白名单两端不一致"
        assert set(names) <= WPS_FORMAT_FIELD_NAMES, f"{action_type} 声明了未验证的属性"

    # 未实现的属性及旧的非规范字段名不能登记；verticalAlignment 已有实现，实机证据另列。
    assert not (WPS_FORMAT_FIELD_NAMES & {"underline", "vertical", "columnWidth"})
    assert not (set(ts_fields["setFont"]) & {"underline"})
    assert not (set(ts_fields["setAlignment"]) & {"vertical"})
    assert not (set(ts_fields["resizeRange"]) & {"columnWidth"})


@pytest.mark.parametrize(("action_type", "fields"), [
    ("setFill", ["fillColor", "fillPattern"]),
    ("setNumberFormat", ["numberFormat"]),
    ("setBorders", ["borderTop", "borderBottom", "borderLeft", "borderRight"]),
    ("setAlignment", ["verticalAlignment"]),
])
def test_wps_added_formats_register_without_snapshots(action_type, fields):
    client = TestClient(app)
    payload = _wps_format_register_payload("wps-added-formats")
    action = payload["descriptor"]["actions"][0]
    action.update(type=action_type, formatFields=fields)
    response = client.post("/api/wps/executions/register", json=payload)
    assert response.status_code == 200, response.text
    assert response.json()["descriptor"]["actions"][0]["formatFields"] == fields
    assert "cells" not in response.json()["descriptor"]["actions"][0]
    assert "values" not in response.json()["descriptor"]["actions"][0]


def test_sort_recovery_registers_value_evidence_and_rejects_format_evidence():
    client = TestClient(app)
    payload = _wps_register_payload("wps-sort-checkpoint")
    payload["descriptor"]["actions"][0]["type"] = "sortRange"
    response = client.post("/api/wps/executions/register", json=payload)
    assert response.status_code == 200, response.text
    assert response.json()["descriptor"]["actions"][0]["type"] == "sortRange"
    payload["operationId"] = "wps-sort-invalid"
    payload["descriptor"]["actions"][0]["previewFingerprint"] = "f1:wrong"
    assert client.post("/api/wps/executions/register", json=payload).status_code == 422


def _resume_payload() -> dict:
    payload = _wps_register_payload("wps-resume-context-1")
    payload["authorizedSheetNames"].append("Source")
    payload["descriptor"]["resumeContext"] = {
        "version": 1, "hostVersion": "12.0", "sources": [{
            "sheetName": "Source", "range": "A1:A2", "rowCount": 2, "columnCount": 1,
            "previewFingerprint": "v2:42:12345678", "valueFingerprint": "v2:40:abcdef12",
        }],
    }
    return payload


@pytest.mark.parametrize("version", [1, 2])
def test_wps_resume_context_roundtrip_is_value_free(version: int) -> None:
    client = TestClient(app)
    payload = _resume_payload()
    if version == 2:
        payload["descriptor"]["resumeContext"] = {"version": 2, "hostVersion": "12.0", "sources": []}
    response = client.post("/api/wps/executions/register", json=payload)
    assert response.status_code == 200, response.text
    fetched = client.get("/api/wps/executions/" + payload["operationId"]).json()
    assert fetched["descriptor"]["resumeContext"] == payload["descriptor"]["resumeContext"]
    assert fetched["ownerToken"] is None
    assert "values" not in json.dumps(fetched["descriptor"])


@pytest.mark.parametrize("issue", ["raw", "unauthorized", "overlap", "shape", "format", "fingerprint", "limit"])
def test_wps_resume_context_rejects_unsafe_metadata(issue: str) -> None:
    """来源证据带"不该出站"或"说不通"的内容一律拒绝。

    `formula` 这一档在 2026-11-01 改成 `format`：`writeFormulas` **可以**带来源证据
    （依赖表证据与值写入同构，端上 `isResumableValueAction` 的判据），把它当成不安全
    元数据拒掉正是"WPS 公式计划一律登记不了"的真机故障。这里要守的是另一半——
    格式动作的恢复点是属性快照，语义上没有来源值，出现这种组合就是伪造或协议漂移。
    正面锚点见 `test_wps_formula_resume.py`。
    """
    payload = _resume_payload()
    source = payload["descriptor"]["resumeContext"]["sources"][0]
    if issue == "raw":
        source["values"] = [["PRIVATE-NEVER-SEND"]]
    elif issue == "unauthorized":
        source["sheetName"] = "Secret"
    elif issue == "overlap":
        source["sheetName"] = "Sheet1"
    elif issue == "shape":
        source["rowCount"] = 3
    elif issue == "format":
        payload["descriptor"]["actions"][0]["type"] = "setFont"
        payload["descriptor"]["actions"][0]["formatFields"] = ["bold"]
    elif issue == "fingerprint":
        source["valueFingerprint"] = "raw-data"
    elif issue == "limit":
        source["range"] = "A1:A1048576"
        source["rowCount"] = 1048576
    response = TestClient(app).post("/api/wps/executions/register", json=payload)
    # 授权类越界沿用 403（与"目标表不在授权内"同一口径），其余是 422。
    expected = 403 if issue == "unauthorized" else 422
    assert response.status_code == expected, response.text
