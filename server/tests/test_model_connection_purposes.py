from __future__ import annotations

import asyncio
import json
from pathlib import Path
from types import SimpleNamespace

import pytest
from fastapi.testclient import TestClient

from server.app import claude_sdk_agent as sdk_agent
from server.app import main as main_module
from server.app.llm import config
from server.app.llm.errors import LLMResponseError
from server.app.models import UpsertModelConnectionRequest


@pytest.fixture(autouse=True)
def isolated_connections(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "_config_path", lambda: tmp_path / ".env")
    monkeypatch.setattr(config, "load_model_settings", lambda: None)
    monkeypatch.setattr(config, "_claude_sdk_connection", lambda: None)
    monkeypatch.setattr(config, "_claude_sdk_state_ready", lambda: True)


def save(**changes):
    return config.upsert_managed_connection(**{
        "connection_id": None, "label": "Gateway", "base_url": "https://gateway.invalid/anthropic",
        "model": "chosen-model", "api_key": "secret-test-key", "clear_api_key": False,
        "supports_vision": True, "protocol": "anthropic", **changes,
    })["connections"][0]


def test_explicit_protocol_roundtrip_preserves_identity_credentials_and_formula_setting():
    item = save()
    connection = config.selected_model_config(item["catalogModelId"])
    assert connection == config.ModelConnection("https://gateway.invalid/anthropic", "chosen-model", "secret-test-key", True, "claude_agent_sdk")
    updated = save(connection_id=item["id"], api_key=None, protocol=None, label="Renamed")
    assert updated["id"] == item["id"] and updated["protocol"] == "anthropic"
    assert config.selected_model_config(item["catalogModelId"]).api_key == "secret-test-key"
    catalog = config.model_catalog()
    assert len(catalog["models"]) == 1
    assert catalog["models"][0]["modelId"] == "chosen-model"
    assert catalog["models"][0]["id"] == item["catalogModelId"]
    response = TestClient(main_module.app).get("/api/models")
    assert response.status_code == 200
    assert response.json()["models"][0]["modelId"] == "chosen-model"
    assert catalog["models"][0]["agentAvailable"] is True
    assert catalog["models"][0]["formulaAvailable"] is True
    assert "secret-test-key" not in json.dumps(catalog)
    config.set_formula_model_id(item["catalogModelId"])
    assert config.formula_model_config() == connection


def test_old_json_defaults_to_openai_and_does_not_guess_gateway_endpoint():
    config._connection_store_path().write_text(json.dumps({"connections": [{
        "id": "old-entry", "label": "Legacy", "baseUrl": "https://gateway.invalid/v1",
        "modelId": "chosen-model", "apiKey": "key", "supportsVision": False,
    }]}), encoding="utf-8")
    item = config.model_settings_view()["connections"][0]
    assert item["protocol"] == "openai_compatible"
    assert item["agentAvailable"] is True and item["formulaAvailable"] is True
    assert item["agentUnavailableReason"] is None
    config.set_formula_model_id(item["catalogModelId"])
    assert config.load_formula_model_id() == item["catalogModelId"]
    assert config.selected_model_config(item["catalogModelId"]).base_url == "https://gateway.invalid/v1"


def test_deepseek_has_one_catalog_entry_but_old_sdk_id_still_resolves():
    item = save(base_url="https://api.deepseek.com/v1", protocol="openai_compatible")
    options = config.model_catalog()["models"]
    assert len(options) == 1 and options[0]["id"] == item["catalogModelId"]
    assert options[0]["agentAvailable"] and options[0]["formulaAvailable"]
    old = config.selected_model_config("claude-agent-sdk:" + item["id"])
    assert old.base_url == config.DEEPSEEK_ANTHROPIC_URL
    assert old.api_key == "secret-test-key"


def test_missing_sdk_runtime_does_not_disable_a_text_connection(monkeypatch):
    item = save()
    monkeypatch.setattr(config, "_claude_sdk_state_ready", lambda: False)
    option = config.model_catalog()["models"][0]
    assert option["id"] == item["catalogModelId"] and option["agentAvailable"] is True
    assert option["agentUnavailableReason"] is None


def test_invalid_protocol_rejected_before_persistence():
    with pytest.raises(ValueError):
        UpsertModelConnectionRequest(label="x", baseUrl="https://gateway.invalid", modelId="m", protocol="guess")
    assert not config._connection_store_path().exists()


@pytest.mark.parametrize("purpose", ["connection", "agent", "formula"])
@pytest.mark.parametrize("protocol", ["anthropic", "openai_compatible"])
def test_connection_test_is_unified_even_for_old_clients_and_does_not_save(monkeypatch, purpose, protocol):
    calls = []
    async def agent(connection):
        pytest.fail("unified connection test must not require the SDK runtime or tools")
    async def formula(connection):
        calls.append(("formula", connection))
    monkeypatch.setattr(sdk_agent, "check_sdk_connection", agent)
    monkeypatch.setattr(main_module, "test_model_connection", formula)
    response = TestClient(main_module.app).post("/api/settings/model/connections/test", json={
        "label": "test", "baseUrl": "https://gateway.invalid", "modelId": "chosen-model",
        "apiKey": "secret", "protocol": protocol, "purpose": purpose,
    })
    assert response.status_code == 200 and response.json()["testedFor"] == "connection"
    assert [call[0] for call in calls] == ["formula"]
    assert calls[0][1].model == "chosen-model" and calls[0][1].api_key == "secret"
    assert not config._connection_store_path().exists()


def test_anthropic_test_preserves_protocol(monkeypatch):
    async def check(connection):
        assert connection.adapter == "claude_agent_sdk"
    monkeypatch.setattr(main_module, "test_model_connection", check)
    response = TestClient(main_module.app).post("/api/settings/model/connections/test", json={
        "label": "test", "baseUrl": "https://gateway.invalid", "modelId": "m",
        "protocol": "anthropic", "purpose": "formula",
    })
    assert response.status_code == 200
    assert response.json()["testedFor"] == "connection"


@pytest.mark.parametrize("tool_call", ["valid", "wrong_token", "text_only"])
def test_sdk_probe_requires_real_tool_receipt_and_exposes_no_workbook_tools(monkeypatch, tmp_path, tool_call):
    from server.tests.test_claude_sdk_agent import _FakeSdk, _FakeResult

    class ProbeSdk(_FakeSdk):
        async def query(self, *, prompt, options):
            assert options.tools == [] and options.setting_sources == []
            assert options.allowed_tools == ["mcp__excel_bro__connection_probe"]
            assert options.extra_args == {"no-session-persistence": None}
            assert list(options.mcp_servers["excel_bro"]) == ["connection_probe"]
            if tool_call != "text_only":
                token = prompt.rsplit(": ", 1)[1] if tool_call == "valid" else "invalid"
                await options.mcp_servers["excel_bro"]["connection_probe"].handler({"token": token})
            yield _FakeResult("OK")

    monkeypatch.setattr(sdk_agent, "_sdk", lambda: ProbeSdk())
    monkeypatch.setattr(sdk_agent, "_state_dir", lambda: tmp_path)
    connection = config.ModelConnection("https://gateway.invalid", "m", "secret", False, "claude_agent_sdk")
    if tool_call == "valid":
        asyncio.run(sdk_agent.check_sdk_connection(connection))
    else:
        with pytest.raises(LLMResponseError, match="未完成 SDK 工具调用"):
            asyncio.run(sdk_agent.check_sdk_connection(connection))


def test_resume_candidates_ignore_legacy_agent_capability_gate(monkeypatch):
    monkeypatch.setattr(main_module, "model_settings_view", lambda: {"connections": [
        {"catalogModelId": "connection:formula", "label": "formula", "modelId": "m", "agentAvailable": False},
        {"catalogModelId": "connection:agent", "label": "agent", "modelId": "m", "agentAvailable": True},
    ]})
    record = SimpleNamespace(active_elapsed_seconds=1, error={"code": "MODEL_ERROR", "retryable": True},
        status="waiting_user", request={}, current_model_id="connection:old", checkpoint={"kind": "planning"},
        pause_reason="", approved_wall_time_seconds=600, turn_id="t", revision=1,
        last_sequence=0, extension_count=0, model_switches=[])
    result = main_module._planning_pause_payload(record)
    assert [item["catalogModelId"] for item in result["availableModelConnections"]] == ["connection:formula", "connection:agent"]


def test_probe_with_no_key_cannot_borrow_cached_sdk_login(monkeypatch):
    def forbidden():
        pytest.fail("must reject before constructing SDK or reading cached login")
    monkeypatch.setattr(sdk_agent, "_sdk", forbidden)
    with pytest.raises(config.ModelNotAvailableError, match="API Key"):
        asyncio.run(sdk_agent.check_sdk_connection(config.ModelConnection(
            "https://gateway.invalid", "m", "", False, "claude_agent_sdk")))
