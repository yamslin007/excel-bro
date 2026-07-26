from __future__ import annotations

import asyncio
import json

import httpx
import pytest
from fastapi.testclient import TestClient

import server.app.main as main_module
from server.app.main import app
from server.app.models import AnswerResponse
from server.app.turn_state import turn_registry


@pytest.fixture(autouse=True)
def clear_model_environment(monkeypatch: pytest.MonkeyPatch) -> None:
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

    async def fake_create_plan(_request: object) -> AnswerResponse:
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
