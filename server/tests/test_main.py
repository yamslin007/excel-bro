from __future__ import annotations

import asyncio
import json
import os
from pathlib import Path
import sqlite3

import httpx
import pytest
from fastapi.testclient import TestClient
from openpyxl import Workbook

import server.app.main as main_module
import server.app.llm.config as llm_config
import server.app.planner as planner_module
import server.app.folder_workbooks as folder_workbooks_module
from server.app.main import app
from server.app.llm import ModelConnection, selected_model_config
from server.app.models import (
    AnswerResponse,
    IntentConversationResponse,
    IntentClarificationResponse,
    IntentProceedResponse,
    LocalPlanRuleResponse,
    PlanRequest,
)
from server.app.turn_state import turn_registry


@pytest.fixture(autouse=True)
def clear_model_environment(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
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
    monkeypatch.setattr(
        llm_config,
        "_config_path",
        lambda: tmp_path / "config.env",
    )


def test_health_reports_unconfigured_model_without_basic_mode() -> None:
    """未配置任何外部模型时，健康状态必须如实表达「没有模型」。

    历史上这里返回 mode="local"，等价于把「基础模式」当成一种可用模型，
    会让前台认为已经选好了模型。现在只保留唯一模式 "model"，
    用 configured=False 表达尚未配置。
    """

    response = TestClient(app).get("/health")

    assert response.status_code == 200
    assert response.json() == {
        "status": "ok",
        "configured": False,
        "mode": "model",
        "model": None,
    }


def _intent_payload(prompt: str, sheet_count: int = 2) -> dict[str, object]:
    """构造一个「已经没有可用模型」的 /api/intent 请求。

    modelId 留空是唯一诚实的表达：目录里已经不存在 local 占位项，
    旧客户端缓存下来的 "local" 会被服务端等同为「没有选择模型」，
    两条路径都必须走到同一个「尚未配置模型」拒绝语义。
    """

    return {
        "prompt": prompt,
        "scope": {
            "workbookName": "电影报表.xlsx",
            "sourceMode": "workbook",
            "selectionMode": "manual",
            "activeWorksheet": "恐怖游轮",
            "selectedRange": "恐怖游轮!A1:E19",
            "totalWorksheetCount": sheet_count,
            "worksheetNames": [f"影片{index + 1}" for index in range(sheet_count)],
            "sheets": [
                {
                    "name": f"影片{index + 1}",
                    "usedRange": f"影片{index + 1}!A1:E20",
                    "rowCount": 20,
                    "columnCount": 5,
                    "headers": [
                        "影院名称",
                        "影院编码",
                        "影片名称",
                        "计数项:影厅",
                        "占比",
                    ],
                }
                for index in range(sheet_count)
            ],
        },
        "imageCount": 0,
        "modelId": "",
    }


def _turn_completion_payload(turn_id: str, **overrides: object) -> dict[str, object]:
    payload: dict[str, object] = {
        "turnId": turn_id,
        "prompt": "查看当前数据",
        "workbook": {
            "name": "通用数据.xlsx",
            "capturedAt": "2026-07-26T00:00:00.000Z",
            "activeWorksheet": "数据",
            "selectedRange": "数据!A1:B2",
            "worksheets": [
                {
                    "name": "数据",
                    "usedRange": "数据!A1:B2",
                    "rowCount": 2,
                    "columnCount": 2,
                    "headers": ["分类", "数值"],
                    "dataRows": [["甲", 10]],
                }
            ],
        },
        "modelId": "",
    }
    payload.update(overrides)
    return payload


def _office_plan(action: dict[str, object]) -> dict[str, object]:
    return {
        "id": "office-plan-1",
        "title": "写入测试",
        "summary": "测试 Office 动作级断点",
        "assumptions": [],
        "warnings": [],
        "actions": [action],
    }


def test_office_recovery_api_requires_supported_saved_workbook_action() -> None:
    client = TestClient(app)
    unsupported = client.post(
        "/api/office/executions/register",
        json={
            "operationId": "office-api-unsupported",
            "plan": _office_plan(
                {"type": "deleteWorksheet", "sheet": "Sheet1"}
            ),
            "authorizedSheetNames": ["Sheet1"],
            "allowedExternalFiles": [],
            "workbookIdentity": "url:c:/work/report.xlsx",
            "confirmed": True,
        },
    )
    assert unsupported.status_code == 422
    assert unsupported.json()["detail"]["code"] == "OFFICE_RECOVERY_UNSUPPORTED_ACTION"

    unsaved = client.post(
        "/api/office/executions/register",
        json={
            "operationId": "office-api-unsaved",
            "plan": _office_plan(
                {
                    "type": "writeValues",
                    "sheet": "Sheet1",
                    "range": "A1",
                    "values": [["ok"]],
                }
            ),
            "authorizedSheetNames": ["Sheet1"],
            "allowedExternalFiles": [],
            "workbookIdentity": "unsaved",
            "confirmed": True,
        },
    )
    assert unsaved.status_code == 422


def test_office_recovery_api_persists_action_receipt_and_completion() -> None:
    client = TestClient(app)
    operation_id = "office-api-execution-0001"
    plan = _office_plan(
        {
            "type": "writeValues",
            "sheet": "Sheet1",
            "range": "A1",
            "values": [["ok"]],
        }
    )
    registered = client.post(
        "/api/office/executions/register",
        json={
            "operationId": operation_id,
            "plan": plan,
            "authorizedSheetNames": ["Sheet1"],
            "allowedExternalFiles": [],
            "workbookIdentity": "url:c:/work/report.xlsx",
            "confirmed": True,
        },
    )
    assert registered.status_code == 200
    state = registered.json()
    assert state["ownerToken"]
    before = {
        "actionIndex": 0,
        "sheet": "Sheet1",
        "range": "A1",
        "formulas": [[None]],
        "numberFormat": [["General"]],
        "fillColor": "#FFFFFF",
        "fontBold": False,
        "fontColor": "#000000",
    }
    started = client.post(
        f"/api/office/executions/{operation_id}/actions/start",
        json={
            "expectedRevision": state["revision"],
            "ownerToken": state["ownerToken"],
            "actionIndex": 0,
            "beforeSnapshot": before,
        },
    )
    assert started.status_code == 200
    state = started.json()
    assert state["requiresReconciliation"] is True
    checkpoint = {
        "planId": "office-plan-1",
        "executionId": operation_id,
        "workbookIdentity": "url:c:/work/report.xlsx",
        "capturedAt": "2026-09-09T00:00:00Z",
        "coverage": "full",
        "completedActionCount": 1,
        "uncertainActionIndexes": [],
        "postStateCoverage": "complete",
        "ranges": [{**before, "postFingerprint": "fnv1a32:12345678"}],
        "createdWorksheets": [],
        "createdWorksheetStates": [],
        "deletedWorksheets": [],
        "skippedWorksheets": [],
    }
    action_completed = client.post(
        f"/api/office/executions/{operation_id}/actions/complete",
        json={
            "expectedRevision": state["revision"],
            "ownerToken": state["ownerToken"],
            "actionIndex": 0,
            "checkpoint": checkpoint,
        },
    )
    assert action_completed.status_code == 200
    state = action_completed.json()
    assert state["nextActionIndex"] == 1
    finished = client.post(
        f"/api/office/executions/{operation_id}/complete",
        json={
            "expectedRevision": state["revision"],
            "ownerToken": state["ownerToken"],
            "result": {
                "actionResults": [],
                "verification": {"status": "verified"},
            },
        },
    )
    assert finished.status_code == 200
    assert finished.json()["status"] == "completed"

    status = client.get(f"/api/office/executions/{operation_id}")
    assert status.status_code == 200
    assert status.json()["ownerToken"] is None


def test_diagnostics_records_unhandled_exceptions() -> None:
    """500 请求必须进诊断日志。

    以前只在 call_next 正常返回后 append，未捕获异常的请求对 /api/diagnostics
    完全隐形；而这类响应又不带 CORS 头，浏览器只报「No Access-Control-Allow-Origin」。
    两者叠加导致真正的错误在前后端都看不见，只能去翻 uvicorn 终端。
    """

    @app.get("/api/_test_unhandled_error")
    async def _boom() -> None:
        raise RuntimeError("测试用未捕获异常")

    main_module._diagnostic_events.clear()
    response = TestClient(app, raise_server_exceptions=False).get(
        "/api/_test_unhandled_error"
    )

    assert response.status_code == 500
    events = [
        event
        for event in main_module._diagnostic_events
        if event.get("path") == "/api/_test_unhandled_error"
    ]
    assert len(events) == 1
    assert events[0]["status"] == 500
    assert "RuntimeError" in str(events[0]["error"])
    assert "测试用未捕获异常" in str(events[0]["error"])


def test_capabilities_are_loaded_from_shared_project_config() -> None:
    response = TestClient(app).get("/api/capabilities")

    assert response.status_code == 200
    value = response.json()
    assert value["version"] == 1
    assert value["snapshot"]["dataRows"] == 200
    assert value["queryTable"]["maxRows"] == 250000
    assert value["images"]["maxAttachments"] == 3
    assert value["intentContext"] == {
        "maxFieldsPerSheet": 30,
        "maxPriorResultRows": 20,
        "maxOutputTokens": None,
    }
    assert value["savedTools"]["maxItems"] == 50
    assert value["agent"]["maxConsecutiveErrors"] == 2
    assert value["agent"]["maxOutputTokens"] is None
    assert value["agent"]["repairMaxOutputTokens"] is None
    assert value["agent"]["maxTotalTokens"] is None


def test_health_reports_model_without_exposing_key(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("AI_BASE_URL", "https://api.example.test/v1")
    monkeypatch.setenv("AI_MODEL", "example-model")
    monkeypatch.setenv("AI_API_KEY", "secret-value")

    response = TestClient(app).get("/health")

    assert response.status_code == 200
    assert response.json() == {
        "status": "ok",
        "configured": True,
        "mode": "model",
        "model": "example-model",
    }
    assert "secret-value" not in response.text


def test_model_catalog_lists_allowlisted_models_without_a_local_placeholder(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """目录里只能出现真实连接，不能再有 "local" 这个产品入口占位项。"""

    monkeypatch.setenv("AI_BASE_URL", "https://api.example.test/v1")
    monkeypatch.setenv("AI_MODEL", "default-model")
    monkeypatch.setenv("AI_MODELS", "fast-model, default-model, reasoning-model")

    response = TestClient(app).get("/api/models")

    assert response.status_code == 200
    assert response.json() == {
        "defaultModelId": "default-model",
        "models": [
            {
                "id": "default-model",
                "label": "default-model",
                "provider": "model",
                "available": True,
                "supportsVision": False,
            },
            {
                "id": "fast-model",
                "label": "fast-model",
                "provider": "model",
                "available": True,
                "supportsVision": False,
            },
            {
                "id": "reasoning-model",
                "label": "reasoning-model",
                "provider": "model",
                "available": True,
                "supportsVision": False,
            },
        ],
    }
    assert "local" not in response.text
    assert "基础模式" not in response.text


def test_model_catalog_is_empty_and_has_no_default_without_any_model() -> None:
    """一条真实连接都没有时，目录必须是空的、默认值必须是 null。

    这对应「成功但为空」这一态：不能凭空补一个默认选择，
    否则界面会显示一个并不存在的模型。
    """

    response = TestClient(app).get("/api/models")

    assert response.status_code == 200
    assert response.json() == {"defaultModelId": None, "models": []}


def test_model_api_key_can_be_updated_without_exposing_it(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    config_path = tmp_path / "excel-bro.env"
    config_path.write_text(
        "AI_BASE_URL=https://api.example.test/v1\n"
        "AI_MODEL=example-model\n"
        "AI_API_KEY=old-secret\n",
        encoding="utf-8",
    )
    monkeypatch.setattr(llm_config, "_config_path", lambda: config_path)
    monkeypatch.setenv("AI_BASE_URL", "https://api.example.test/v1")
    monkeypatch.setenv("AI_MODEL", "example-model")
    monkeypatch.setenv("AI_API_KEY", "old-secret")

    response = TestClient(app).put(
        "/api/settings/model",
        json={"apiKey": "replacement-secret"},
    )

    assert response.status_code == 200
    assert response.json() == {
        "baseUrl": "https://api.example.test/v1",
        "defaultModel": "example-model",
        "apiKeyConfigured": True,
        "apiKeyHint": "••••cret",
        "formulaModelId": "",
        "connections": [],
    }
    assert "replacement-secret" not in response.text
    assert "replacement-secret" in config_path.read_text(encoding="utf-8")
    assert os.environ["AI_API_KEY"] == "replacement-secret"

    current = TestClient(app).get("/api/settings/model")
    assert current.status_code == 200
    assert "replacement-secret" not in current.text
    assert current.json()["apiKeyHint"] == "••••cret"


def test_model_api_key_update_rejects_blank_values(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    config_path = tmp_path / "excel-bro.env"
    config_path.write_text("AI_API_KEY=existing\n", encoding="utf-8")
    monkeypatch.setattr(llm_config, "_config_path", lambda: config_path)

    response = TestClient(app).put(
        "/api/settings/model",
        json={"apiKey": "   "},
    )

    assert response.status_code == 422
    assert "existing" in config_path.read_text(encoding="utf-8")


def test_managed_model_connections_have_independent_credentials(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    config_path = tmp_path / "excel-bro.env"
    config_path.write_text("", encoding="utf-8")
    monkeypatch.setattr(llm_config, "_config_path", lambda: config_path)

    created = TestClient(app).post(
        "/api/settings/model/connections",
        json={
            "label": "DeepSeek",
            "baseUrl": "https://api.deepseek.com/v1/",
            "modelId": "deepseek-chat",
            "apiKey": "deepseek-private-key",
            "supportsVision": False,
        },
    )

    assert created.status_code == 200
    value = created.json()
    assert len(value["connections"]) == 1
    connection = value["connections"][0]
    assert connection["label"] == "DeepSeek"
    assert connection["baseUrl"] == "https://api.deepseek.com/v1"
    assert connection["apiKeyHint"] == "••••-key"
    assert "deepseek-private-key" not in created.text

    catalog = TestClient(app).get("/api/models").json()
    catalog_option = next(
        option
        for option in catalog["models"]
        if option["id"] == connection["catalogModelId"]
    )
    assert catalog_option["label"] == "DeepSeek · deepseek-chat"

    selected = selected_model_config(connection["catalogModelId"])
    assert selected is not None
    assert selected.base_url == "https://api.deepseek.com/v1"
    assert selected.model == "deepseek-chat"
    assert selected.api_key == "deepseek-private-key"

    store_path = tmp_path / "model-connections.json"
    assert "deepseek-private-key" in store_path.read_text(encoding="utf-8")

    deleted = TestClient(app).delete(
        f"/api/settings/model/connections/{connection['id']}"
    )
    assert deleted.status_code == 200
    assert deleted.json()["connections"] == []

    # 删除最后一个连接后，目录回到「成功但为空」，默认值为 null：
    # 不能因为删光了就退回一个并不存在的默认模型。
    empty_catalog = TestClient(app).get("/api/models").json()
    assert empty_catalog == {"defaultModelId": None, "models": []}

    # 已删除连接的 id 不再可用：查询路径返回 None；
    # 写入路径抛稳定错误（此时一个连接都没有，命中的是「尚未配置」）。
    # 两条路径都不静默回落到某个别的模型。
    assert llm_config.resolve_model_config(connection["catalogModelId"]) is None
    with pytest.raises(llm_config.ModelNotConfiguredError):
        llm_config.selected_model_config(connection["catalogModelId"])

    # 留一条别的连接后，目录非空，此时「已删除的 id」才是 ModelNotAvailableError，
    # 与「一条都没有」区分开——前台据此提示「该连接已失效」而不是「去添加模型」。
    surviving = TestClient(app).post(
        "/api/settings/model/connections",
        json={
            "label": "Kimi",
            "baseUrl": "https://api.moonshot.cn/v1",
            "modelId": "moonshot-v1-8k",
            "apiKey": "surviving-key",
            "supportsVision": False,
        },
    ).json()["connections"][0]
    assert surviving["catalogModelId"] != connection["catalogModelId"]

    still_gone = connection["catalogModelId"]
    with pytest.raises(llm_config.ModelNotAvailableError):
        llm_config.selected_model_config(still_gone)
    assert llm_config.resolve_model_config(still_gone) is None
    assert llm_config.resolve_model_config(surviving["catalogModelId"]) is not None


def test_model_connection_can_be_tested_without_saving(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    async def fake_test_model_connection(connection: object) -> None:
        captured["connection"] = connection

    monkeypatch.setattr(
        main_module,
        "test_model_connection",
        fake_test_model_connection,
    )
    response = TestClient(app).post(
        "/api/settings/model/connections/test",
        json={
            "label": "Test",
            "baseUrl": "https://api.example.test/v1",
            "modelId": "test-model",
            "apiKey": "temporary-secret",
            "supportsVision": False,
        },
    )

    assert response.status_code == 200
    assert response.json()["ok"] is True
    connection = captured["connection"]
    assert getattr(connection, "api_key") == "temporary-secret"
    assert not (llm_config._connection_store_path()).exists()


def test_same_model_connections_keep_distinct_ids_and_keys() -> None:
    client = TestClient(app)
    first = client.post(
        "/api/settings/model/connections",
        json={
            "label": "Primary",
            "baseUrl": "https://api.example.test/v1",
            "modelId": "test-model",
            "apiKey": "private-key",
            "supportsVision": False,
        },
    ).json()["connections"][0]

    cleared = client.post(
        "/api/settings/model/connections",
        json={
            "id": first["id"],
            "label": "Primary",
            "baseUrl": "https://api.example.test/v1",
            "modelId": "test-model",
            "apiKey": None,
            "clearApiKey": True,
            "supportsVision": False,
        },
    )
    assert cleared.status_code == 200
    assert cleared.json()["connections"][0]["apiKeyConfigured"] is False

    duplicate = client.post(
        "/api/settings/model/connections",
        json={
            "label": "Duplicate",
            "baseUrl": "https://api.example.test/v1/",
            "modelId": "TEST-MODEL",
            "apiKey": "another-key",
            "supportsVision": False,
        },
    )
    assert duplicate.status_code == 200
    second = duplicate.json()["connections"][1]
    assert second["id"] != first["id"]
    assert second["catalogModelId"] != first["catalogModelId"]
    assert selected_model_config(first["catalogModelId"]).api_key == ""
    assert selected_model_config(second["catalogModelId"]).api_key == "another-key"
    labels = [item["label"] for item in client.get("/api/models").json()["models"]]
    assert labels == ["Primary · test-model (1)", "Duplicate · TEST-MODEL (2)"]

    updated = client.post(
        "/api/settings/model/connections",
        json={
            "id": first["id"],
            "label": "Primary",
            "baseUrl": "https://api.example.test/v1",
            "modelId": "test-model",
            "apiKey": "first-new-key",
            "supportsVision": False,
        },
    )
    assert updated.status_code == 200
    assert selected_model_config(first["catalogModelId"]).api_key == "first-new-key"
    assert selected_model_config(second["catalogModelId"]).api_key == "another-key"


def test_corrupt_model_connection_store_is_reported(
    tmp_path: Path,
) -> None:
    (tmp_path / "model-connections.json").write_text(
        "{not-json",
        encoding="utf-8",
    )

    response = TestClient(app).get("/api/settings/model")

    assert response.status_code == 500
    assert response.json()["detail"]["code"] == (
        "MODEL_CONNECTION_STORE_INVALID"
    )


def test_model_vision_capability_requires_explicit_configuration(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("AI_BASE_URL", "https://api.example.test/v1")
    monkeypatch.setenv("AI_MODEL", "vision-looking-name")

    # 目录里只有真实连接，第一项就是它本身（不再有 local 占位项顶在前面）。
    unconfigured = TestClient(app).get("/api/models").json()
    assert unconfigured["models"][0]["supportsVision"] is False

    monkeypatch.setenv("AI_VISION_MODELS", "vision-looking-name")
    configured = TestClient(app).get("/api/models").json()
    assert configured["models"][0]["supportsVision"] is True


def test_folder_refresh_endpoint_keeps_session_id(tmp_path: Path) -> None:
    source = tmp_path / "scores.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "得分"
    sheet.append(["人员", "得分"])
    sheet.append(["嘟嘟嘟", 33])
    workbook.save(source)
    workbook.close()

    catalog = folder_workbooks_module.scan_folder(tmp_path)
    source.rename(tmp_path / "renamed.xlsx")

    response = TestClient(app).post(
        "/api/folders/refresh",
        json={"sessionId": catalog.sessionId},
    )

    assert response.status_code == 200
    value = response.json()
    assert value["sessionId"] == catalog.sessionId
    assert [file["name"] for file in value["files"]] == ["renamed.xlsx"]


def test_folder_refresh_endpoint_rejects_unknown_session() -> None:
    response = TestClient(app).post(
        "/api/folders/refresh",
        json={"sessionId": "missing-session"},
    )

    assert response.status_code == 422
    assert response.json()["detail"]["code"] == "FOLDER_DATA_ERROR"


def test_folder_backup_management_endpoints_use_the_selected_work_directory(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "scores.xlsx"
    workbook = Workbook()
    workbook.active.append(["人员", "得分"])
    workbook.save(source)
    workbook.close()
    (tmp_path / ".excel-bro-backups").mkdir()
    catalog = folder_workbooks_module.scan_folder(tmp_path)
    client = TestClient(app)

    listed = client.get(
        "/api/folders/backups", params={"sessionId": catalog.sessionId}
    )
    cleaned = client.post(
        "/api/folders/backups/cleanup", json={"sessionId": catalog.sessionId}
    )
    opened: list[str] = []
    monkeypatch.setattr(
        folder_workbooks_module.os,
        "startfile",
        lambda path: opened.append(path),
        raising=False,
    )
    monkeypatch.setattr(
        folder_workbooks_module,
        "_explorer_window_handles",
        lambda: [],
    )
    monkeypatch.setattr(
        folder_workbooks_module,
        "_focus_opened_explorer_window",
        lambda directory, handles: True,
    )
    opened_response = client.post(
        "/api/folders/backups/open",
        json={"directoryToken": listed.json()["openDirectoryToken"]},
    )

    assert listed.status_code == 200
    assert listed.json()["folderPath"] == str(tmp_path)
    assert listed.json()["backupDirectory"] == str(
        tmp_path / ".excel-bro-backups"
    )
    assert cleaned.status_code == 200
    assert cleaned.json()["deletedExecutions"] == 0
    assert opened_response.status_code == 200
    assert opened_response.json()["directory"] == str(
        tmp_path / ".excel-bro-backups"
    )
    assert opened == [str(tmp_path / ".excel-bro-backups")]


def test_folder_backup_management_rejects_unknown_session() -> None:
    response = TestClient(app).get(
        "/api/folders/backups", params={"sessionId": "missing-session"}
    )

    assert response.status_code == 422
    assert response.json()["detail"]["code"] == "FOLDER_BACKUP_ERROR"

    open_response = TestClient(app).post(
        "/api/folders/backups/open",
        json={"directoryToken": "missing-token"},
    )
    assert open_response.status_code == 422
    assert open_response.json()["detail"]["code"] == "FOLDER_BACKUP_OPEN_ERROR"

    delete_response = TestClient(app).post(
        "/api/folders/backups/delete",
        json={
            "backupManagementToken": "missing-token",
            "undoId": "a" * 32,
            "confirmed": True,
        },
    )
    assert delete_response.status_code == 422
    assert delete_response.json()["detail"]["code"] == "FOLDER_BACKUP_DELETE_ERROR"


def _expect_model_not_configured(response: httpx.Response) -> None:
    """断言一次「没有可用模型」的调用被明确拒绝，而不是静默降级。

    旧行为是悄悄回落到基础模式并返回 200；现在必须是稳定错误码 409。
    文案必须给出可执行的下一步（服务地址 / 模型 ID / API Key），
    而不是含糊的「请求无效」。
    """

    assert response.status_code == 409
    detail = response.json()["detail"]
    assert detail["code"] == "MODEL_NOT_CONFIGURED"
    assert "服务地址" in detail["message"]
    assert "API Key" in detail["message"]


# 测试里反复使用的「已配置模型」目录 id。
_CONFIGURED_MODEL_ID = "connection:test-primary"


def _configure_single_model(monkeypatch: pytest.MonkeyPatch) -> None:
    """给当前测试装一条可用的真实模型连接。

    需要同时换掉三处，因为它们各自持有不同的引用：

    - ``planner_module.selected_model_config``：planner 是
      ``from .llm import selected_model_config`` 的模块级引用，且现在要求
      "不可用时抛错"（返回 None 会让它当成解析失败）。
    - ``main_module.resolve_model_config``：``/api/turn`` 的意图分支用它查
      "当前有没有可用模型"。
    - ``main_module.selected_model_config``：``/api/intent`` 端点直接用。
    """

    connection = ModelConnection(
        base_url="https://api.example.test/v1",
        model="test-primary",
        api_key="test-secret",
        supports_vision=False,
    )

    def resolve(model_id: str | None) -> ModelConnection:
        if model_id == _CONFIGURED_MODEL_ID:
            return connection
        raise llm_config.ModelNotAvailableError(
            f"指定的模型「{model_id}」不存在或已被删除。"
        )

    monkeypatch.setattr(
        main_module,
        "model_settings_view",
        lambda: {
            "connections": [
                {
                    "catalogModelId": _CONFIGURED_MODEL_ID,
                    "label": "测试模型",
                    "modelId": "test-primary",
                    "supportsVision": False,
                }
            ]
        },
    )
    monkeypatch.setattr(planner_module, "selected_model_config", resolve)
    monkeypatch.setattr(main_module, "selected_model_config", resolve)
    monkeypatch.setattr(main_module, "resolve_model_config", resolve)


def _expect_model_not_configured(response: httpx.Response) -> None:
    """断言一次「没有可用模型」的调用被明确拒绝，而不是静默降级。

    旧行为是悄悄回落到基础模式并返回 200；现在必须是稳定错误码 409。
    文案必须给出可执行的下一步（服务地址 / 模型 ID / API Key），
    而不是含糊的「请求无效」。
    """

    assert response.status_code == 409
    detail = response.json()["detail"]
    assert detail["code"] == "MODEL_NOT_CONFIGURED"
    assert "服务地址" in detail["message"]
    assert "API Key" in detail["message"]


def test_semantic_intent_is_rejected_without_a_configured_model() -> None:
    """语义请求在无模型时必须显式失败，不得静默降级为基础模式。"""

    response = TestClient(app).post(
        "/api/intent",
        json=_intent_payload("找出占比最高的一个电影"),
    )

    _expect_model_not_configured(response)


def test_turn_endpoint_rejects_semantic_prompt_without_model() -> None:
    assert TestClient(app).post(
        "/api/intent",
        json=_intent_payload("查看当前数据"),
    ).status_code == 409


def _stub_planner_answer(monkeypatch: pytest.MonkeyPatch) -> None:
    """把真正的模型调用换成固定应答。

    这些用例校验的是 turnId / 幂等等协议语义，不该真的发网络请求；
    但必须保留「有模型」的前提，否则会先被 MODEL_NOT_CONFIGURED 拦住。
    """

    async def fake_create_plan(
        _request: object,
        *,
        tool_cache: object = None,
        on_event=None,
        **_options: object,
    ) -> AnswerResponse:
        return AnswerResponse(
            kind="answer",
            message="查看完成",
            provider="model",
        )

    async def fake_check_intent(
        _request: object, **_kwargs: object
    ) -> IntentProceedResponse:
        return IntentProceedResponse(
            kind="proceed",
            summary="查看当前数据",
            confirmedPrompt="查看当前数据",
            provider="model",
        )

    monkeypatch.setattr(planner_module, "create_plan", fake_create_plan)
    monkeypatch.setattr(main_module, "create_plan", fake_create_plan)
    monkeypatch.setattr(main_module, "check_intent", fake_check_intent)


def test_turn_endpoint_completes_direct_conversation(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _configure_single_model(monkeypatch)

    async def fake_check_intent(_request: object, **_kwargs: object) -> IntentConversationResponse:
        return IntentConversationResponse(message="你好！", provider="model")

    monkeypatch.setattr(main_module, "check_intent", fake_check_intent)
    response = TestClient(app).post(
        "/api/turn",
        json={**_intent_payload("你好"), "modelId": _CONFIGURED_MODEL_ID},
    )

    assert response.status_code == 200
    assert response.json()["kind"] == "conversation"
    assert response.json()["message"] == "你好！"
    assert turn_registry.touch(response.json()["turnId"]).stage == "completed"


def test_turn_endpoint_preserves_turn_id_across_decisions(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """turnId 跨「决策 → 完成」两阶段的语义不变，与模型选择无关。

    这里给一个已配置的模型，让请求能走到真正的规划路径；
    「没有模型」的拒绝语义由 test_turn_endpoint_rejects_semantic_prompt_without_model 覆盖。
    """

    _configure_single_model(monkeypatch)
    _stub_planner_answer(monkeypatch)
    first = TestClient(app).post(
        "/api/turn",
        json={**_intent_payload("查看当前数据"), "modelId": _CONFIGURED_MODEL_ID},
    )

    assert first.status_code == 200
    turn_id = first.json()["turnId"]
    assert turn_id.startswith("turn-")

    payload = _intent_payload("继续查看")
    payload["turnId"] = turn_id
    payload["modelId"] = _CONFIGURED_MODEL_ID
    second = TestClient(app).post("/api/turn", json=payload)

    assert second.status_code == 200
    assert second.json()["turnId"] == turn_id


def test_turn_completion_is_idempotent(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _configure_single_model(monkeypatch)
    _stub_planner_answer(monkeypatch)
    completion = _turn_completion_payload("turn-idempotent-test")
    completion["modelId"] = _CONFIGURED_MODEL_ID

    first = TestClient(app).post("/api/turn", json=completion)
    repeated = TestClient(app).post("/api/turn", json=completion)

    assert first.status_code == 200
    assert repeated.status_code == 200
    assert repeated.json() == first.json()

    changed = {**completion, "prompt": "改用另一种口径"}
    conflict = TestClient(app).post("/api/turn", json=changed)
    assert conflict.status_code == 409
    assert conflict.json() == {
        "detail": {
            "code": "TURN_ALREADY_COMPLETED",
            "message": "该轮次已经用另一组参数完成，请开始新的请求。",
            "retryable": False,
        }
    }


def test_planning_clarification_is_cached_but_keeps_the_turn_open_for_answers(monkeypatch):
    _configure_single_model(monkeypatch)
    calls = []

    async def planner(request, **_kwargs):
        calls.append(request)
        if len(calls) == 1:
            return IntentClarificationResponse.model_validate({
                "provider": "model", "clarification": {
                    "id": "unit", "summary": "补充口径", "question": "请填写单位", "reason": "单位未知",
                    "scopeLabel": "Sheet1", "options": [], "fields": [{"id": "unit", "label": "统计单位"}],
                },
            })
        return AnswerResponse(provider="model", message="已收到单位，尚未执行写入。")

    monkeypatch.setattr(main_module, "create_plan", planner)
    client = TestClient(app)
    payload = _turn_completion_payload("turn-form-test", modelId=_CONFIGURED_MODEL_ID)
    first = client.post("/api/turn", json=payload)
    repeated = client.post("/api/turn", json=payload)
    assert first.status_code == repeated.status_code == 200
    assert first.json() == repeated.json()
    assert first.json()["kind"] == "clarification"
    assert first.json()["clarification"]["fields"][0]["id"] == "unit"
    assert len(calls) == 1
    assert turn_registry.touch("turn-form-test").stage == "awaiting_clarification"
    turn_registry.ensure_decision_allowed("turn-form-test")
    answered = client.post("/api/turn", json={**payload, "prompt": "统计已选表，补充确认：单位为件"})
    assert answered.status_code == 200
    assert answered.json()["kind"] == "answer"
    assert len(calls) == 2
    assert turn_registry.touch("turn-form-test").stage == "completed"


def test_turn_endpoint_returns_local_mapping_rule_without_cell_values(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """WPS 无行数据请求必须经真实 HTTP 响应校验返回规则，重试仍使用缓存。"""

    calls = 0

    async def fake_create_plan(request: PlanRequest, **_kwargs: object) -> LocalPlanRuleResponse:
        nonlocal calls
        calls += 1
        assert request.planningReadPolicy == "none"
        assert request.workbook.worksheets[0].dataRows == []
        return LocalPlanRuleResponse.model_validate(
            {
                "rule": {
                    "sourceSheet": "数据",
                    "targetSheet": "数据",
                    "sourceRange": "B2:B2",
                    "targetRange": "C2:C2",
                    "mapping": [{"source": "甲", "target": "一类"}],
                    "blankLabel": "未分类",
                },
                "summary": "按给定映射填写目标列",
                "warnings": [],
                "assumptions": [],
                "provider": "model",
            }
        )

    monkeypatch.setattr(main_module, "create_plan", fake_create_plan)
    payload = _turn_completion_payload(
        "turn-local-rule-response-test", planningReadPolicy="none"
    )
    workbook = payload["workbook"]
    assert isinstance(workbook, dict)
    worksheets = workbook["worksheets"]
    assert isinstance(worksheets, list)
    worksheets[0]["dataRows"] = []

    client = TestClient(app, raise_server_exceptions=False)
    wps_origin = {"Origin": "http://127.0.0.1:3889"}
    first = client.post("/api/turn", json=payload, headers=wps_origin)
    repeated = client.post("/api/turn", json=payload, headers=wps_origin)

    assert first.status_code == 200
    assert first.headers["access-control-allow-origin"] == wps_origin["Origin"]
    assert repeated.status_code == 200
    assert repeated.json() == first.json()
    assert calls == 1
    body = first.json()
    assert body["kind"] == "local_plan_rule"
    assert body["turnId"] == "turn-local-rule-response-test"
    assert body["rule"]["mapping"] == [{"source": "甲", "target": "一类"}]
    assert "values" not in body["rule"]
    assert "actions" not in body


def test_concurrent_turn_completion_only_runs_planner_once(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls = 0

    async def fake_create_plan(
        _request: object, **_kwargs: object
    ) -> AnswerResponse:
        nonlocal calls
        calls += 1
        await asyncio.sleep(0.02)
        return AnswerResponse(
            kind="answer",
            message="完成",
            provider="model",
        )

    monkeypatch.setattr(main_module, "create_plan", fake_create_plan)
    completion = _turn_completion_payload("turn-concurrent-test")

    async def run_requests() -> list[httpx.Response]:
        transport = httpx.ASGITransport(app=app)
        async with httpx.AsyncClient(
            transport=transport,
            base_url="http://testserver",
        ) as client:
            return await asyncio.gather(
                client.post("/api/turn", json=completion),
                client.post("/api/turn", json=completion),
            )

    responses = asyncio.run(run_requests())

    assert [response.status_code for response in responses] == [200, 200]
    assert responses[0].json() == responses[1].json()
    assert calls == 1


def _parse_sse(text: str) -> list[tuple[str, dict]]:
    events: list[tuple[str, dict]] = []
    for block in text.strip().split("\n\n"):
        if not block.strip():
            continue
        event_name = ""
        data = ""
        for line in block.splitlines():
            if line.startswith("event:"):
                event_name = line[len("event:") :].strip()
            elif line.startswith("data:"):
                data = line[len("data:") :].strip()
        events.append((event_name, json.loads(data)))
    return events


_STREAM_PAYLOAD = {
    "turnId": "turn-stream-test",
    "prompt": "查看当前数据",
    "workbook": {
        "name": "通用数据.xlsx",
        "capturedAt": "2026-07-26T00:00:00.000Z",
        "activeWorksheet": "数据",
        "worksheets": [
            {
                "name": "数据",
                "rowCount": 2,
                "columnCount": 2,
                "headers": ["分类", "数值"],
                "dataRows": [["甲", 10]],
            }
        ],
    },
    "modelId": "",
}


def test_turn_stream_emits_steps_then_result(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    async def fake_create_plan(
        _request: object,
        *,
        tool_cache: object = None,
        on_event=None,
        **_checkpoint_options: object,
    ) -> AnswerResponse:
        if on_event is not None:
            await on_event({"phase": "planning", "title": "正在查看工作簿结构"})
            await on_event(
                {
                    "phase": "planning",
                    "title": "正在查找字段「分类」",
                    "completedStep": "已查看工作簿结构",
                }
            )
        return AnswerResponse(kind="answer", message="完成", provider="model")

    monkeypatch.setattr(main_module, "create_plan", fake_create_plan)

    response = TestClient(app).post("/api/turn/stream", json=_STREAM_PAYLOAD)

    assert response.status_code == 200
    assert response.headers["content-type"].startswith("text/event-stream")
    events = _parse_sse(response.text)
    step_events = [payload for name, payload in events if name == "step"]
    result_events = [payload for name, payload in events if name == "result"]
    assert len(step_events) == 2
    assert step_events[0]["title"] == "正在查看工作簿结构"
    assert len(result_events) == 1
    assert result_events[0]["kind"] == "answer"
    assert result_events[0]["turnId"] == "turn-stream-test"


def test_turn_stream_persists_budget_stop_as_failed_not_completed(monkeypatch):
    async def stopped(_request, **_kwargs):
        return AnswerResponse(message="规划预算已用完", provider="system", outcome="failed", stopReason="max_tokens")
    monkeypatch.setattr(main_module, "create_plan", stopped)
    client = TestClient(app)
    payload = {**_STREAM_PAYLOAD, "turnId": "turn-stopped-test"}
    response = client.post("/api/turn/stream", json=payload)
    result = next(value for name, value in _parse_sse(response.text) if name == "result")
    assert result["outcome"] == "failed"
    assert result["stopReason"] == "max_tokens"
    record = main_module._folder_operation_ledger.planning_task("turn-stopped-test")
    assert record.status == "failed"
    assert record.result["outcome"] == "failed"
    assert main_module.turn_registry.touch("turn-stopped-test").stage == "failed"
    status = client.get("/api/turn/turn-stopped-test/status").json()
    assert status["status"] == "failed"
    assert status["canContinue"] is False
    assert status["error"]["code"] == "PLANNING_STOPPED"
    replay = client.post("/api/turn/stream", json=payload)
    assert all(name != "result" or value.get("outcome") == "failed" for name, value in _parse_sse(replay.text))


def test_folder_execute_stream_reuses_operation_without_writing_twice(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    calls = 0

    def fake_execute(request, on_progress=None):
        nonlocal calls
        calls += 1
        if on_progress is not None:
            on_progress(
                {
                    "executionId": "0123456789abcdef0123456789abcdef",
                    "phase": "executing",
                    "stepId": "folder.action.1",
                    "status": "completed",
                    "title": "第 1 步已完成",
                    "summary": "已创建结果表",
                    "current": 1,
                    "total": 1,
                    "unit": "actions",
                    "toolName": "submit_plan",
                    "facts": [
                        {"label": "动作", "value": "createWorksheet", "style": "code"}
                    ],
                }
            )
        return main_module.FolderExecuteResponse.model_validate(
            {
                "filesModified": ["数据.xlsx"],
                "backups": [],
                "actionResults": [
                    {
                        "index": 0,
                        "type": "createWorksheet",
                        "sheet": "结果",
                        "status": "succeeded",
                    }
                ],
                "verification": {
                    "status": "verified",
                    "passed": True,
                    "checks": [],
                    "unverifiedActions": [],
                },
                "executionMs": 1,
                "verificationMs": 1,
            }
        )

    monkeypatch.setattr(main_module, "execute_folder_plan", fake_execute)
    monkeypatch.setattr(main_module, "folder_session_root", lambda _session_id: tmp_path)
    payload = {
        "sessionId": "session-stream-test",
        "operationId": "folder-operation-test",
        "plan": {
            "id": "folder-plan-test",
            "title": "创建结果表",
            "summary": "创建一张结果表",
            "assumptions": [],
            "warnings": [],
            "actions": [{"type": "createWorksheet", "sheet": "结果"}],
        },
    }
    client = TestClient(app)

    first = client.post("/api/folders/execute/stream", json=payload)
    main_module._folder_operations.clear()
    monkeypatch.setattr(
        main_module,
        "folder_session_root",
        lambda _session_id: (_ for _ in ()).throw(ValueError("会话已失效")),
    )
    second = client.post(
        "/api/folders/execute/stream",
        json={**payload, "afterSequence": 1},
    )
    status = client.get(
        "/api/folders/execute/status",
        params={"operationId": "folder-operation-test"},
    )

    assert first.status_code == 200
    assert second.status_code == 200
    assert calls == 1
    progress_events = [
        payload for name, payload in _parse_sse(first.text) if name == "progress"
    ]
    assert progress_events
    assert progress_events[0]["toolName"] == "submit_plan"
    assert progress_events[0]["facts"][0]["value"] == "createWorksheet"
    second_events = _parse_sse(second.text)
    assert [name for name, _ in second_events] == ["result"]
    assert status.json()["status"] == "completed"
    assert status.json()["lastSequence"] == 1


def test_folder_execute_stream_reports_unknown_effect_as_needs_reconciliation(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    def fake_execute(request, on_progress=None):
        raise RuntimeError("模拟保存窗口中断")

    monkeypatch.setattr(main_module, "execute_folder_plan", fake_execute)
    monkeypatch.setattr(main_module, "folder_session_root", lambda _session_id: tmp_path)
    payload = {
        "sessionId": "session-reconciliation-test",
        "operationId": "folder-reconciliation-operation",
        "plan": {
            "id": "folder-plan-reconciliation",
            "title": "创建结果表",
            "summary": "创建一张结果表",
            "assumptions": [],
            "warnings": [],
            "actions": [{"type": "createWorksheet", "sheet": "结果"}],
        },
    }
    client = TestClient(app)

    response = client.post("/api/folders/execute/stream", json=payload)
    status = client.get(
        "/api/folders/execute/status", params={"operationId": payload["operationId"]}
    )

    assert response.status_code == 200
    assert [name for name, _payload in _parse_sse(response.text)][-1] == "error"
    assert status.status_code == 200
    assert status.json()["status"] == "needs_reconciliation"
    assert status.json()["error"]["code"] == (
        "FOLDER_EXECUTION_REQUIRES_RECONCILIATION"
    )


def test_folder_resume_stream_claims_checkpoint_and_returns_original_result(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    operation_id = "folder-resume-operation"
    ledger = main_module._folder_operation_ledger
    ledger.register(operation_id, "resume-fingerprint")
    ledger.save_execution_context(
        operation_id,
        request={"operationId": operation_id, "plan": {"id": "split-plan"}},
        authorization={"schemaVersion": 1, "root": str(tmp_path), "targets": []},
    )
    ledger.commit_checkpoint(
        operation_id,
        kind="splitGroupAggregate",
        payload={
            "schemaVersion": 1,
            "kind": "splitGroupAggregate",
            "stage": "publish_prepared",
            "executionId": "a" * 32,
        },
    )
    record = ledger.get(operation_id)
    assert record is not None
    captured: dict[str, object] = {}

    def fake_resume(
        received_operation_id,
        authorization,
        checkpoint,
        on_progress=None,
        on_checkpoint=None,
    ):
        captured["operationId"] = received_operation_id
        captured["authorization"] = authorization
        captured["checkpoint"] = checkpoint
        return main_module.FolderExecuteResponse.model_validate(
            {
                "executionId": "a" * 32,
                "filesModified": ["cinema.xlsx"],
                "backups": [".excel-bro-backups/executions/backup/cinema.xlsx"],
                "backupWarnings": [],
                "actionResults": [],
                "verification": {
                    "status": "verified",
                    "passed": True,
                    "checks": [],
                    "unverifiedActions": [],
                },
                "executionMs": 0,
                "verificationMs": 1,
            }
        )

    monkeypatch.setattr(main_module, "resume_folder_plan", fake_resume)
    monkeypatch.setattr(
        main_module,
        "folder_resume_checkpoint_summary",
        lambda *_args: {"canResume": True, "filesApplied": 0, "filesPending": 1},
    )
    response = TestClient(app).post(
        "/api/folders/execute/resume/stream",
        json={
            "operationId": operation_id,
            "expectedRevision": record.revision,
            "confirmed": True,
        },
    )

    assert response.status_code == 200
    events = _parse_sse(response.text)
    assert events[-1][0] == "result"
    assert events[-1][1]["filesModified"] == ["cinema.xlsx"]
    assert captured["operationId"] == operation_id
    assert captured["authorization"] == record.authorization
    assert captured["checkpoint"] == ledger.checkpoint(operation_id).payload
    assert ledger.get(operation_id).status == "completed"


def test_folder_resume_stream_rejects_stale_revision(tmp_path: Path) -> None:
    operation_id = "folder-resume-stale"
    ledger = main_module._folder_operation_ledger
    ledger.register(operation_id, "resume-fingerprint")
    ledger.save_execution_context(
        operation_id,
        request={"operationId": operation_id, "plan": {"id": "split-plan"}},
        authorization={"schemaVersion": 1, "root": str(tmp_path), "targets": []},
    )
    ledger.commit_checkpoint(
        operation_id,
        kind="splitGroupAggregate",
        payload={
            "schemaVersion": 1,
            "kind": "splitGroupAggregate",
            "stage": "publish_prepared",
            "executionId": "b" * 32,
        },
    )

    response = TestClient(app).post(
        "/api/folders/execute/resume/stream",
        json={
            "operationId": operation_id,
            "expectedRevision": 1,
            "confirmed": True,
        },
    )

    assert response.status_code == 409
    assert response.json()["detail"]["code"] == "RESUME_CLAIM_CONFLICT"


def test_folder_execute_stream_refuses_unfinished_persisted_operation(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    calls = 0

    def fake_execute(request, on_progress=None):
        nonlocal calls
        calls += 1
        raise AssertionError("遗留 running 任务不能被重新执行")

    monkeypatch.setattr(main_module, "execute_folder_plan", fake_execute)
    monkeypatch.setattr(main_module, "folder_session_root", lambda _session_id: tmp_path)
    payload = {
        "sessionId": "session-stream-test",
        "operationId": "folder-interrupted-operation",
        "plan": {
            "id": "folder-plan-test",
            "title": "创建结果表",
            "summary": "创建一张结果表",
            "assumptions": [],
            "warnings": [],
            "actions": [{"type": "createWorksheet", "sheet": "结果"}],
        },
    }
    validated_request = main_module.FolderExecuteRequest.model_validate(payload)
    request_fingerprint = main_module._folder_request_fingerprint(validated_request)
    main_module._folder_operation_ledger.register(
        payload["operationId"], request_fingerprint
    )
    main_module._folder_operations.clear()
    response = TestClient(app).post("/api/folders/execute/stream", json=payload)

    assert response.status_code == 409
    assert response.json()["detail"]["code"] == "OPERATION_REQUIRES_RECONCILIATION"
    assert calls == 0


def test_folder_execute_stream_refuses_expired_operation_without_reexecution(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    calls = 0

    def fake_execute(request, on_progress=None):
        nonlocal calls
        calls += 1
        raise AssertionError("过期任务不能被重新执行")

    monkeypatch.setattr(main_module, "execute_folder_plan", fake_execute)
    monkeypatch.setattr(main_module, "folder_session_root", lambda _session_id: tmp_path)
    payload = {
        "sessionId": "session-expired-test",
        "operationId": "folder-expired-operation",
        "plan": {
            "id": "folder-plan-expired",
            "title": "创建结果表",
            "summary": "创建一张结果表",
            "assumptions": [],
            "warnings": [],
            "actions": [{"type": "createWorksheet", "sheet": "结果"}],
        },
    }
    validated_request = main_module.FolderExecuteRequest.model_validate(payload)
    request_fingerprint = main_module._folder_request_fingerprint(validated_request)
    main_module._folder_operation_ledger.register(
        payload["operationId"], request_fingerprint
    )
    main_module._folder_operation_ledger.finish(
        payload["operationId"],
        status="completed",
        result={"filesModified": ["结果.xlsx"]},
    )
    with sqlite3.connect(main_module._folder_operation_ledger._path()) as connection:
        connection.execute(
            "UPDATE folder_operations SET updated_at = 0 WHERE operation_id = ?",
            (payload["operationId"],),
        )
    assert main_module._folder_operation_ledger.compact_terminal_operations(30) == 1

    client = TestClient(app)
    response = client.post("/api/folders/execute/stream", json=payload)
    status = client.get(
        "/api/folders/execute/status",
        params={"operationId": payload["operationId"]},
    )

    assert response.status_code == 410
    assert response.json()["detail"]["code"] == "OPERATION_EXPIRED"
    assert status.status_code == 410
    assert status.json()["detail"]["code"] == "OPERATION_EXPIRED"
    assert calls == 0


def test_legacy_folder_execute_endpoint_cannot_bypass_idempotency() -> None:
    response = TestClient(app).post(
        "/api/folders/execute",
        json={
            "sessionId": "legacy-session",
            "plan": {
                "id": "folder-plan-test",
                "title": "创建结果表",
                "summary": "创建一张结果表",
                "assumptions": [],
                "warnings": [],
                "actions": [{"type": "createWorksheet", "sheet": "结果"}],
            },
        },
    )

    assert response.status_code == 409
    assert response.json()["detail"]["code"] == "IDEMPOTENT_EXECUTION_REQUIRED"


def test_opening_backup_manager_registers_directory_for_restart_discovery(
    tmp_path: Path,
) -> None:
    root = tmp_path / "work"
    root.mkdir()
    client = TestClient(app)

    backups = client.get(
        "/api/folders/backups", params={"folderPath": str(root)}
    )
    directories = client.get("/api/folders/recovery/directories")

    assert backups.status_code == 200
    assert directories.status_code == 200
    registered = directories.json()["directories"]
    assert len(registered) == 1
    assert Path(registered[0]["folderPath"]) == root
    assert registered[0]["status"] == "active"
    assert not (root / ".excel-bro-backups").exists()

    removed = client.delete(
        f"/api/folders/recovery/directories/{registered[0]['directoryId']}"
    )
    assert removed.status_code == 200
    assert client.get("/api/folders/recovery/directories").json()["directories"] == []


def test_turn_stream_serializes_model_error_as_event(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from server.app.llm import LLMTimeoutError

    async def failing_create_plan(
        _request: object,
        *,
        tool_cache: object = None,
        on_event=None,
        **_checkpoint_options: object,
    ) -> AnswerResponse:
        raise LLMTimeoutError("模型响应超时")

    monkeypatch.setattr(main_module, "create_plan", failing_create_plan)

    response = TestClient(app).post("/api/turn/stream", json=_STREAM_PAYLOAD)

    assert response.status_code == 200  # SSE 已 200，错误在流内
    events = _parse_sse(response.text)
    error_events = [payload for name, payload in events if name == "error"]
    assert len(error_events) == 1
    assert error_events[0]["code"] == "MODEL_TIMEOUT"
    assert error_events[0]["status"] == 504
    assert error_events[0]["retryable"] is True


def test_turn_stream_persists_unexpected_failure_and_returns_safe_error(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    async def failing_create_plan(
        _request: object,
        *,
        tool_cache: object = None,
        on_event=None,
        **_checkpoint_options: object,
    ) -> AnswerResponse:
        raise RuntimeError("不应返回给用户的内部错误")

    monkeypatch.setattr(main_module, "create_plan", failing_create_plan)
    payload = {**_STREAM_PAYLOAD, "turnId": "turn-unexpected-failure-test"}

    response = TestClient(app).post("/api/turn/stream", json=payload)

    assert response.status_code == 200
    events = _parse_sse(response.text)
    error = next(value for name, value in events if name == "error")
    assert error == {
        "status": 500,
        "code": "PLANNING_TASK_FAILED",
        "message": "规划任务未能安全保存当前状态，请重新发起任务。",
        "retryable": False,
    }
    assert "内部错误" not in response.text
    record = main_module._folder_operation_ledger.planning_task(
        "turn-unexpected-failure-test"
    )
    assert record is not None
    assert record.status == "failed"
    assert record.error == error


def test_closing_planning_stream_stops_worker_before_releasing_active_fence(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    worker_cancelled = asyncio.Event()

    async def hanging_create_plan(
        _request: object,
        *,
        tool_cache: object = None,
        on_event=None,
        on_checkpoint=None,
        **_checkpoint_options: object,
    ) -> AnswerResponse:
        checkpoint = {
            "schemaVersion": 1,
            "kind": "planning",
            "stage": "tool_batch_completed",
            "messages": [{"role": "user", "content": "继续前的安全现场"}],
            "toolCache": {},
            "loopState": {"activeElapsedSeconds": 1},
            "toolMode": "all",
            "repairingSubmission": False,
            "nextTurnIndex": 1,
            "pendingAssistantMessage": None,
        }
        if on_checkpoint is not None:
            await on_checkpoint(checkpoint)
        if on_event is not None:
            await on_event({"phase": "planning", "title": "已保存安全现场"})
        try:
            await asyncio.Event().wait()
        except asyncio.CancelledError:
            worker_cancelled.set()
            raise
        return AnswerResponse(kind="answer", message="不会到达", provider="local")

    monkeypatch.setattr(main_module, "create_plan", hanging_create_plan)
    payload = {**_STREAM_PAYLOAD, "turnId": "turn-stream-close-test"}
    request = main_module.PlanRequest.model_validate(payload)
    record, created = main_module._folder_operation_ledger.register_planning_task(
        request.turnId or "",
        main_module._planning_request_fingerprint(request),
        request.model_dump(mode="json", exclude_none=True),
        approved_wall_time_seconds=60,
    )
    assert created is True

    async def close_after_first_event() -> None:
        stream = main_module._planning_task_stream(request, record)
        first = await anext(stream)
        assert "event: step" in first
        assert request.turnId in main_module._active_planning_tasks
        await stream.aclose()

    asyncio.run(close_after_first_event())

    assert worker_cancelled.is_set()
    assert request.turnId not in main_module._active_planning_tasks
    persisted = main_module._folder_operation_ledger.planning_task(request.turnId or "")
    assert persisted is not None
    assert persisted.status == "running"
    assert persisted.checkpoint is not None


def test_turn_stream_maps_overloaded_429_to_friendly_message(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from server.app.llm import LLMHTTPStatusError

    async def failing_create_plan(
        _request: object,
        *,
        tool_cache: object = None,
        on_event=None,
        **_checkpoint_options: object,
    ) -> AnswerResponse:
        raise LLMHTTPStatusError(
            429,
            '{"error":{"message":"The engine is currently overloaded",'
            '"type":"engine_overloaded_error"}}',
        )

    monkeypatch.setattr(main_module, "create_plan", failing_create_plan)

    response = TestClient(app).post("/api/turn/stream", json=_STREAM_PAYLOAD)

    assert response.status_code == 200
    events = _parse_sse(response.text)
    error_events = [payload for name, payload in events if name == "error"]
    assert len(error_events) == 1
    assert error_events[0]["code"] == "MODEL_OVERLOADED"
    assert error_events[0]["status"] == 503
    assert error_events[0]["retryable"] is True
    # 不透传上游英文原始 body
    assert "engine_overloaded_error" not in error_events[0]["message"]
    assert "overloaded" not in error_events[0]["message"]
    assert "繁忙" in error_events[0]["message"]


def test_turn_stream_returns_cached_completion_without_steps(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls = 0

    async def fake_create_plan(
        _request: object,
        *,
        tool_cache: object = None,
        on_event=None,
        **_checkpoint_options: object,
    ) -> AnswerResponse:
        nonlocal calls
        calls += 1
        if on_event is not None:
            await on_event({"phase": "planning", "title": "正在规划操作"})
        return AnswerResponse(kind="answer", message="完成", provider="local")

    monkeypatch.setattr(main_module, "create_plan", fake_create_plan)

    client = TestClient(app)
    first = client.post("/api/turn/stream", json=_STREAM_PAYLOAD)
    second = client.post("/api/turn/stream", json=_STREAM_PAYLOAD)

    assert first.status_code == 200
    assert second.status_code == 200
    second_events = _parse_sse(second.text)
    assert [name for name, _ in second_events] == ["result"]
    assert calls == 1  # 第二次命中缓存，不再规划


def test_turn_stream_pauses_persists_and_continues_from_checkpoint(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls = 0
    resumed_checkpoints: list[dict | None] = []

    async def pausing_create_plan(
        _request: object,
        *,
        tool_cache: object = None,
        on_event=None,
        on_checkpoint=None,
        resume_checkpoint=None,
        approved_wall_time_seconds=None,
    ) -> AnswerResponse:
        nonlocal calls
        calls += 1
        resumed_checkpoints.append(resume_checkpoint)
        if resume_checkpoint is None:
            checkpoint = {
                "schemaVersion": 1,
                "kind": "planning",
                "stage": "tool_batch_completed",
                "messages": [{"role": "user", "content": "查看当前数据"}],
                "toolCache": {"context": "saved"},
                "loopState": {"activeElapsedSeconds": 60},
                "toolMode": "terminal",
                "repairingSubmission": False,
                "nextTurnIndex": 1,
                "pendingAssistantMessage": None,
            }
            if on_checkpoint is not None:
                await on_checkpoint(checkpoint)
            raise main_module.AgentBudgetPause(
                checkpoint, float(approved_wall_time_seconds or 60)
            )
        return AnswerResponse(kind="answer", message="续跑完成", provider="local")

    monkeypatch.setattr(main_module, "create_plan", pausing_create_plan)
    payload = {
        **_STREAM_PAYLOAD,
        "turnId": "turn-planning-pause-test",
        "workbook": {
            **_STREAM_PAYLOAD["workbook"],
            "sourceFingerprint": "sha256:original-source",
            "sourceFingerprintSheets": ["数据"],
        },
    }
    client = TestClient(app)

    first = client.post("/api/turn/stream", json=payload)
    first_events = _parse_sse(first.text)
    pause = next(value for name, value in first_events if name == "paused")
    assert pause["kind"] == "planning_pause"
    assert pause["canContinue"] is True
    assert pause["sourceFingerprint"] == "sha256:original-source"

    status = client.get("/api/turn/turn-planning-pause-test/status")
    assert status.status_code == 200
    assert status.json()["status"] == "waiting_user"
    assert status.json()["revision"] == pause["revision"]

    continued = client.post(
        "/api/turn/continue/stream",
        json={
            "turnId": "turn-planning-pause-test",
            "expectedRevision": pause["revision"],
            "requestId": "continue-request-0001",
            "confirmed": True,
            "afterSequence": pause["lastSequence"],
        },
    )
    continued_events = _parse_sse(continued.text)
    result = next(value for name, value in continued_events if name == "result")
    assert result["message"] == "续跑完成"
    assert calls == 2
    assert resumed_checkpoints[0] is None
    assert resumed_checkpoints[1] is not None
    assert resumed_checkpoints[1]["toolCache"] == {"context": "saved"}
    record = main_module._folder_operation_ledger.planning_task(
        "turn-planning-pause-test"
    )
    assert record is not None
    assert record.status == "completed"
    assert record.extension_count == 1
    assert record.approved_wall_time_seconds == 120


def test_model_timeout_pauses_and_user_can_switch_configured_fallback(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls: list[str | None] = []

    async def model_fails_then_resumes(
        request: object,
        *,
        tool_cache: object = None,
        on_event=None,
        on_checkpoint=None,
        resume_checkpoint=None,
        **_options: object,
    ) -> AnswerResponse:
        model_id = getattr(request, "modelId", None)
        calls.append(model_id)
        if resume_checkpoint is None:
            checkpoint = {
                "schemaVersion": 2,
                "kind": "planning",
                "stage": "before_model",
                "messages": [{"role": "user", "content": "分析数据"}],
                "toolCache": {"context": "saved"},
                "loopState": {"activeElapsedSeconds": 12},
                "toolMode": "full",
                "repairingSubmission": False,
                "nextTurnIndex": 0,
                "pendingAssistantMessage": None,
                "nextToolIndex": 0,
                "toolBatchState": {},
            }
            assert on_checkpoint is not None
            await on_checkpoint(checkpoint)
            raise main_module.LLMTimeoutError("timeout")
        assert model_id == "connection:fallback"
        assert resume_checkpoint["toolCache"] == {"context": "saved"}
        return AnswerResponse(kind="answer", message="备用模型接续完成", provider="model")

    monkeypatch.setattr(main_module, "create_plan", model_fails_then_resumes)
    monkeypatch.setattr(
        main_module,
        "model_settings_view",
        lambda: {
            "connections": [
                {
                    "catalogModelId": "connection:primary",
                    "label": "主模型",
                    "modelId": "primary-model",
                    "supportsVision": True,
                },
                {
                    "catalogModelId": "connection:fallback",
                    "label": "备用模型",
                    "modelId": "fallback-model",
                    "supportsVision": True,
                },
            ]
        },
    )
    monkeypatch.setattr(
        main_module,
        "selected_model_config",
        lambda model_id: ModelConnection(
            base_url="https://fallback.example.test/v1",
            model="fallback-model",
            api_key="secret-not-persisted",
            supports_vision=True,
        )
        if model_id == "connection:fallback"
        else None,
    )
    payload = {
        **_STREAM_PAYLOAD,
        "turnId": "turn-model-fallback-test",
        "modelId": "connection:primary",
    }
    client = TestClient(app)

    first = client.post("/api/turn/stream", json=payload)
    pause = next(
        value for name, value in _parse_sse(first.text) if name == "paused"
    )
    assert pause["pauseType"] == "model_error"
    assert pause["requiresExtension"] is False
    assert pause["canSwitchModel"] is True
    assert pause["currentModelId"] == "connection:primary"
    assert [item["catalogModelId"] for item in pause["availableModelConnections"]] == [
        "connection:fallback"
    ]
    assert "secret-not-persisted" not in first.text

    continued = client.post(
        "/api/turn/continue/stream",
        json={
            "turnId": "turn-model-fallback-test",
            "expectedRevision": pause["revision"],
            "requestId": "fallback-request-0001",
            "confirmed": True,
            "afterSequence": pause["lastSequence"],
            "modelId": "connection:fallback",
            "confirmModelSwitch": True,
        },
    )
    continued_events = _parse_sse(continued.text)
    assert any(
        name == "step" and value["title"] == "已切换到用户选择的备用模型"
        for name, value in continued_events
    )
    result = next(value for name, value in continued_events if name == "result")
    assert result["message"] == "备用模型接续完成"
    assert calls == ["connection:primary", "connection:fallback"]
    record = main_module._folder_operation_ledger.planning_task(
        "turn-model-fallback-test"
    )
    assert record is not None
    assert record.current_model_id == "connection:fallback"
    assert record.extension_count == 0
    assert record.approved_wall_time_seconds == 60
    assert len(record.model_switches) == 1
    assert "secret-not-persisted" not in json.dumps(record.model_switches)


def test_interrupted_planning_task_can_switch_model_and_resume_checkpoint(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    resumed: list[tuple[str | None, object]] = []

    async def resume_from_checkpoint(
        request: object,
        *,
        resume_checkpoint=None,
        **_options: object,
    ) -> AnswerResponse:
        resumed.append((getattr(request, "modelId", None), resume_checkpoint))
        return AnswerResponse(
            kind="answer",
            message="中断任务已由备用模型接续",
            provider="model",
        )

    monkeypatch.setattr(main_module, "create_plan", resume_from_checkpoint)
    monkeypatch.setattr(
        main_module,
        "model_settings_view",
        lambda: {
            "connections": [
                {
                    "catalogModelId": "connection:primary-interrupted",
                    "label": "原模型",
                    "modelId": "primary-model",
                    "supportsVision": True,
                },
                {
                    "catalogModelId": "connection:fallback-interrupted",
                    "label": "备用模型",
                    "modelId": "fallback-model",
                    "supportsVision": True,
                },
            ]
        },
    )
    monkeypatch.setattr(
        main_module,
        "selected_model_config",
        lambda model_id: ModelConnection(
            base_url="https://fallback.example.test/v1",
            model="fallback-model",
            api_key="secret-not-persisted",
            supports_vision=True,
        )
        if model_id == "connection:fallback-interrupted"
        else None,
    )
    payload = {
        **_STREAM_PAYLOAD,
        "turnId": "turn-interrupted-model-switch-test",
        "modelId": "connection:primary-interrupted",
    }
    request = main_module.PlanRequest.model_validate(payload)
    record, created = main_module._folder_operation_ledger.register_planning_task(
        request.turnId or "",
        main_module._planning_request_fingerprint(request),
        request.model_dump(mode="json", exclude_none=True),
        approved_wall_time_seconds=60,
    )
    assert created is True
    checkpoint = {
        "schemaVersion": 2,
        "kind": "planning",
        "stage": "tool_batch_completed",
        "messages": [{"role": "user", "content": "保留的任务上下文"}],
        "toolCache": {"context": "saved-before-disconnect"},
        "loopState": {"activeElapsedSeconds": 8},
        "toolMode": "full",
        "repairingSubmission": False,
        "nextTurnIndex": 1,
        "pendingAssistantMessage": None,
        "nextToolIndex": 0,
        "toolBatchState": {},
    }
    main_module._folder_operation_ledger.save_planning_checkpoint(
        record.turn_id,
        checkpoint,
        active_elapsed_seconds=8,
    )
    client = TestClient(app)

    status_response = client.get(f"/api/turn/{record.turn_id}/status")
    assert status_response.status_code == 200
    status = status_response.json()
    assert status["status"] == "interrupted"
    assert status["pauseType"] == "interrupted"
    assert status["canSwitchModel"] is True

    continued = client.post(
        "/api/turn/continue/stream",
        json={
            "turnId": record.turn_id,
            "expectedRevision": status["revision"],
            "requestId": "interrupted-fallback-request-0001",
            "confirmed": True,
            "afterSequence": 0,
            "modelId": "connection:fallback-interrupted",
            "confirmModelSwitch": True,
        },
    )
    events = _parse_sse(continued.text)
    result = next(value for name, value in events if name == "result")
    assert result["message"] == "中断任务已由备用模型接续"
    assert resumed == [("connection:fallback-interrupted", checkpoint)]
    persisted = main_module._folder_operation_ledger.planning_task(record.turn_id)
    assert persisted is not None
    assert persisted.current_model_id == "connection:fallback-interrupted"
    assert len(persisted.model_switches) == 1


def test_turn_stream_rejects_intent_request() -> None:
    payload = _intent_payload("帮我看看这些数据")
    payload["turnId"] = "turn-stream-intent"
    response = TestClient(app).post("/api/turn/stream", json=payload)
    assert response.status_code == 422
    assert response.json()["detail"]["code"] == "INVALID_REQUEST"


def test_intent_check_is_rejected_without_a_model_even_when_prompt_is_explicit() -> None:
    """即使提示词已经足够明确（不需要模型澄清），无模型时也必须拒绝。

    这里曾经返回 200 + provider="local"，等于用规则路径冒充一次模型判断。
    现在统一走 MODEL_NOT_CONFIGURED，前端据此提示「添加模型」。
    """

    response = TestClient(app).post(
        "/api/intent",
        json=_intent_payload(
            "按电影汇总后重新计算整体占比，找出占比最高的电影"
        ),
    )

    _expect_model_not_configured(response)


def test_confirmed_intent_still_requires_a_configured_model() -> None:
    """用户已确认意图也不能绕过模型检查。"""

    payload = _intent_payload("找出占比最高的一个电影")
    payload["intentConfirmed"] = True

    response = TestClient(app).post("/api/intent", json=payload)

    _expect_model_not_configured(response)


def test_intent_rejection_does_not_special_case_a_single_sheet_workbook() -> None:
    """单工作表（自动选区）同样按「没有模型」拒绝，不因表数而走特殊路径。"""

    payload = _intent_payload("找出占比最高的一个电影", sheet_count=1)
    scope = payload["scope"]
    assert isinstance(scope, dict)
    scope["selectionMode"] = "auto"
    scope["totalWorksheetCount"] = 23
    scope["activeWorksheet"] = "功夫女足"
    sheets = scope["sheets"]
    assert isinstance(sheets, list)
    sheets[0]["name"] = "功夫女足"

    response = TestClient(app).post("/api/intent", json=payload)

    _expect_model_not_configured(response)


def test_model_intent_retries_invalid_structure_and_keeps_prior_intent(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("AI_BASE_URL", "https://api.example.test/v1")
    monkeypatch.setenv("AI_MODEL", "example-model")
    calls: list[dict[str, object]] = []

    async def handler(request: httpx.Request) -> httpx.Response:
        calls.append(json.loads(request.content))
        if len(calls) == 1:
            return httpx.Response(
                200,
                json={"choices": [{"message": {"content": "我需要想一下"}}]},
            )
        return httpx.Response(
            200,
            json={
                "choices": [
                    {
                        "message": {
                            "content": json.dumps(
                                {
                                    "kind": "proceed",
                                    "summary": "继承上一轮并改为最小值",
                                    "confirmedPrompt": "在相同范围和口径下找出占比最小的一项",
                                },
                                ensure_ascii=False,
                            )
                        }
                    }
                ]
            },
        )

    transport = httpx.MockTransport(handler)
    original_client = httpx.AsyncClient
    monkeypatch.setattr(
        httpx,
        "AsyncClient",
        lambda **kwargs: original_client(transport=transport, **kwargs),
    )
    payload = _intent_payload("那最小的呢")
    payload["modelId"] = "example-model"
    payload["priorIntent"] = {
        "confirmedPrompt": "找出占比最大的一项",
        "toolRequest": {
            "id": "maximum",
            "tool": "query_table",
            "arguments": {
                "mode": "rows",
                "fields": ["占比"],
                "sortBy": "占比",
                "sortDirection": "desc",
                "limit": 1,
            },
        },
    }
    payload["priorResult"] = {
        "kind": "table",
        "title": "上一轮极值结果",
        "headers": ["名称", "占比"],
        "rows": [["甲", 0.42]],
        "primaryValueColumn": 1,
        "sourceSheets": ["影片1", "影片2"],
        "warnings": [],
    }
    scope = payload["scope"]
    assert isinstance(scope, dict)
    sheets = scope["sheets"]
    assert isinstance(sheets, list)
    sheets[0]["previewRows"] = [["不应发送的原始数据"]]

    response = TestClient(app).post("/api/intent", json=payload)

    assert response.status_code == 200
    assert response.json()["confirmedPrompt"].endswith("占比最小的一项")
    assert len(calls) == 2
    second_prompt = calls[1]["messages"][1]["content"]  # type: ignore[index]
    assert "上一轮结构化意图" in second_prompt
    assert "上一轮紧凑结果" in second_prompt
    assert "上一轮极值结果" in second_prompt
    assert "不应发送的原始数据" not in second_prompt
    assert "未通过结构化协议校验" in second_prompt
