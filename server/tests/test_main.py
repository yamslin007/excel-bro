from __future__ import annotations

import asyncio
import json
import os
from pathlib import Path

import httpx
import pytest
from fastapi.testclient import TestClient

import server.app.main as main_module
import server.app.llm.config as llm_config
from server.app.main import app
from server.app.llm import selected_model_config
from server.app.models import AnswerResponse
from server.app.turn_state import turn_registry


@pytest.fixture(autouse=True)
def clear_model_environment(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    turn_registry.reset()
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


def test_health_reports_local_mode_without_model_config() -> None:
    response = TestClient(app).get("/health")

    assert response.status_code == 200
    assert response.json() == {
        "status": "ok",
        "configured": False,
        "mode": "local",
        "model": None,
    }


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
    }
    assert value["savedTools"]["maxItems"] == 50


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


def test_model_catalog_includes_local_and_allowlisted_models(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("AI_BASE_URL", "https://api.example.test/v1")
    monkeypatch.setenv("AI_MODEL", "default-model")
    monkeypatch.setenv("AI_MODELS", "fast-model, default-model, reasoning-model")

    response = TestClient(app).get("/api/models")

    assert response.status_code == 200
    assert response.json() == {
        "defaultModelId": "default-model",
        "models": [
            {
                "id": "local",
                "label": "基础模式",
                "provider": "local",
                "available": True,
                "supportsVision": False,
            },
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
    assert TestClient(app).get("/api/models").json()["models"] == [
        {
            "id": "local",
            "label": "基础模式",
            "provider": "local",
            "available": True,
            "supportsVision": False,
        }
    ]


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


def test_model_connection_edit_can_clear_key_and_reject_duplicates() -> None:
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
    assert duplicate.status_code == 409
    assert duplicate.json()["detail"]["code"] == "MODEL_CONNECTION_DUPLICATE"


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

    unconfigured = TestClient(app).get("/api/models").json()
    assert unconfigured["models"][1]["supportsVision"] is False

    monkeypatch.setenv("AI_VISION_MODELS", "vision-looking-name")
    configured = TestClient(app).get("/api/models").json()
    assert configured["models"][1]["supportsVision"] is True


def _intent_payload(prompt: str, sheet_count: int = 2) -> dict[str, object]:
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
        "modelId": "local",
    }


def test_local_intent_fallback_does_not_apply_business_specific_rules() -> None:
    response = TestClient(app).post(
        "/api/intent",
        json=_intent_payload("找出占比最高的一个电影"),
    )

    assert response.status_code == 200
    value = response.json()
    assert value["kind"] == "proceed"
    assert value["provider"] == "local"


def test_turn_endpoint_preserves_turn_id_across_decisions() -> None:
    first = TestClient(app).post(
        "/api/turn",
        json=_intent_payload("查看当前数据"),
    )

    assert first.status_code == 200
    turn_id = first.json()["turnId"]
    assert turn_id.startswith("turn-")

    payload = _intent_payload("继续查看")
    payload["turnId"] = turn_id
    second = TestClient(app).post("/api/turn", json=payload)

    assert second.status_code == 200
    assert second.json()["turnId"] == turn_id


def test_turn_completion_is_idempotent() -> None:
    decision = TestClient(app).post(
        "/api/turn",
        json=_intent_payload("查看当前数据"),
    )
    turn_id = decision.json()["turnId"]
    completion = {
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
        "modelId": "local",
    }

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
            provider="local",
        )

    monkeypatch.setattr(main_module, "create_plan", fake_create_plan)
    completion = {
        "turnId": "turn-concurrent-test",
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
        "modelId": "local",
    }

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
    "modelId": "local",
}


def test_turn_stream_emits_steps_then_result(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    async def fake_create_plan(
        _request: object,
        *,
        tool_cache: object = None,
        on_event=None,
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
        return AnswerResponse(kind="answer", message="完成", provider="local")

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


def test_turn_stream_serializes_model_error_as_event(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from server.app.llm import LLMTimeoutError

    async def failing_create_plan(
        _request: object,
        *,
        tool_cache: object = None,
        on_event=None,
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


def test_turn_stream_maps_overloaded_429_to_friendly_message(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from server.app.llm import LLMHTTPStatusError

    async def failing_create_plan(
        _request: object,
        *,
        tool_cache: object = None,
        on_event=None,
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


def test_turn_stream_rejects_intent_request() -> None:
    payload = _intent_payload("帮我看看这些数据")
    payload["turnId"] = "turn-stream-intent"
    response = TestClient(app).post("/api/turn/stream", json=payload)
    assert response.status_code == 422
    assert response.json()["detail"]["code"] == "INVALID_REQUEST"


def test_intent_check_allows_explicit_multi_sheet_ratio_request() -> None:
    response = TestClient(app).post(
        "/api/intent",
        json=_intent_payload(
            "按电影汇总后重新计算整体占比，找出占比最高的电影"
        ),
    )

    assert response.status_code == 200
    assert response.json() == {
        "kind": "proceed",
        "summary": "按电影汇总后重新计算整体占比，找出占比最高的电影",
        "confirmedPrompt": "按电影汇总后重新计算整体占比，找出占比最高的电影",
        "provider": "local",
    }


def test_confirmed_intent_is_not_clarified_again() -> None:
    payload = _intent_payload("找出占比最高的一个电影")
    payload["intentConfirmed"] = True

    response = TestClient(app).post("/api/intent", json=payload)

    assert response.status_code == 200
    assert response.json()["kind"] == "proceed"


def test_local_intent_fallback_does_not_special_case_sheet_names() -> None:
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

    assert response.status_code == 200
    value = response.json()
    assert value["kind"] == "proceed"
    assert value["provider"] == "local"


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
