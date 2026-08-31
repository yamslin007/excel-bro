from __future__ import annotations

from fastapi.testclient import TestClient

from server.app.main import app

client = TestClient(app)


def base_plan() -> dict:
    return {
        "id": "edited-plan",
        "title": "手改的计划",
        "summary": "从数据表按分类汇总金额",
        "assumptions": [],
        "warnings": [],
        "actions": [
            {
                "type": "computeValues",
                "sheet": "结果",
                "sourceSheet": "数据",
                "destinationCell": "A1",
                "hasHeaders": True,
                "includeHeaders": True,
                "query": {
                    "mode": "aggregate",
                    "groupBy": ["分类"],
                    "metrics": [
                        {
                            "operation": "sum",
                            "field": "金额",
                            "outputName": "合计",
                        }
                    ],
                },
            }
        ],
    }


def validate_request(
    plan: dict,
    *,
    authorized_sheets: list[str] | None = None,
    existing_sheets: list[str] | None = None,
) -> dict:
    authorized = authorized_sheets if authorized_sheets is not None else ["数据"]
    return {
        "plan": plan,
        "authorized_sheets": authorized,
        "existing_sheets": (
            existing_sheets if existing_sheets is not None else authorized
        ),
    }


def test_returns_normalized_plan_with_server_computed_criteria() -> None:
    """前端只送内容，验收条件必须是服务端加的 —— 这是编辑功能的核心约定。"""
    payload = base_plan()
    assert "acceptanceCriteria" not in payload

    response = client.post("/api/plans/validate", json=validate_request(payload))

    assert response.status_code == 200
    normalized = response.json()
    criteria = normalized["acceptanceCriteria"]
    assert criteria, "服务端应补上验收条件"
    # computeValues 的行数在执行时才确定，只能补 worksheetExists。
    assert {criterion["type"] for criterion in criteria} == {"worksheetExists"}


def test_rejects_unknown_action_type() -> None:
    payload = base_plan()
    payload["actions"] = [{"type": "runShellCommand", "sheet": "结果"}]

    response = client.post("/api/plans/validate", json=validate_request(payload))

    assert response.status_code == 422


def test_rejects_multi_cell_destination() -> None:
    payload = base_plan()
    payload["actions"][0]["destinationCell"] = "A1:C9"

    response = client.post("/api/plans/validate", json=validate_request(payload))

    assert response.status_code == 422


def test_rejects_criteria_targeting_untouched_sheets() -> None:
    """编辑后的计划不能比模型生成的宽松：验收条件只能查计划真的动过的表。"""
    payload = base_plan()
    payload["acceptanceCriteria"] = [
        {"type": "worksheetExists", "sheet": "没动过的表"}
    ]

    response = client.post("/api/plans/validate", json=validate_request(payload))

    assert response.status_code == 422


def test_rejects_an_unauthorized_source_fingerprint_sheet() -> None:
    plan = base_plan()
    plan["sourceFingerprintSheets"] = ["秘密"]
    response = client.post(
        "/api/plans/validate",
        json=validate_request(
            plan,
            authorized_sheets=["数据"],
            existing_sheets=["数据", "秘密"],
        ),
    )

    assert response.status_code == 422
    assert "数据来源指纹尝试读取未授权的工作表「秘密」" in response.text


def test_rejects_an_unauthorized_existing_source_sheet() -> None:
    response = client.post(
        "/api/plans/validate",
        json=validate_request(
            base_plan(),
            authorized_sheets=["其他"],
            existing_sheets=["数据", "其他"],
        ),
    )

    assert response.status_code == 422
    assert "尝试读取未授权的工作表「数据」" in response.text


def test_rejects_writes_and_deletes_on_unauthorized_existing_sheets() -> None:
    for action in (
        {
            "type": "writeValues",
            "sheet": "秘密",
            "range": "A1",
            "values": [["不允许"]],
        },
        {"type": "deleteWorksheet", "sheet": "秘密"},
        {"type": "createWorksheet", "sheet": "秘密"},
    ):
        plan = base_plan()
        plan["actions"] = [action]
        response = client.post(
            "/api/plans/validate",
            json=validate_request(
                plan,
                authorized_sheets=["数据"],
                existing_sheets=["数据", "秘密"],
            ),
        )

        assert response.status_code == 422
        assert f"动作 {action['type']} 尝试写入未授权的工作表「秘密」" in response.text


def test_allows_a_new_output_sheet_outside_the_authorized_existing_set() -> None:
    response = client.post(
        "/api/plans/validate",
        json=validate_request(
            base_plan(),
            authorized_sheets=["数据"],
            existing_sheets=["数据", "秘密"],
        ),
    )

    assert response.status_code == 200
    assert response.json()["actions"][0]["sheet"] == "结果"


def test_blocks_cross_site_origins() -> None:
    """状态变更端点必须落在 Origin 白名单后面。"""
    response = client.post(
        "/api/plans/validate",
        json=validate_request(base_plan()),
        headers={"Origin": "https://evil.example"},
    )

    assert response.status_code == 403
    assert response.json()["detail"]["code"] == "FORBIDDEN_ORIGIN"
