"""账本恢复描述：持久侧只存**版本化恢复描述 + 结果引用**，绝不存值矩阵。

缺陷复现：``ledger_safe_response`` 把计划里的必填矩阵换成形状之后，账本里那条记录
**长得仍然像一个完整的 ``PlanResponse``**。状态查询 / 重放通道于是照原样把它回放，
调用方拿到一个"矩阵全空"的假计划 —— 看着完整，其实值被剥掉了。

修好之后的判据：

* 持久侧存的是带 ``schemaVersion`` 的**恢复描述**，明确写出"值被省略了"、
  "要取值请用这个 ``resultRef``"；
* 状态端点在**没有值被省略**时才按 ``result`` 回放完整响应（普通后端的完成结果
  取回走这条），一旦有值被省略就只给 ``recovery``，**不再**冒充完整响应。
"""

from __future__ import annotations

import json
import uuid

import pytest
from fastapi.testclient import TestClient

from server.app import main as main_module
from server.app.ledger_privacy import (
    RECOVERY_SCHEMA_VERSION,
    ledger_recovery_descriptor,
)
from server.app.models import (
    AnalysisPlan,
    AnswerResponse,
    PlanResponse,
    ResultContext,
)
from server.app.main import app

SENTINEL = "SENTINEL_CELL_7K2Q"
RESULT_REF = "res_" + "1" * 32


@pytest.fixture(autouse=True)
def isolated_ledger(monkeypatch, tmp_path):
    from server.app.task_ledger import FolderOperationLedger

    monkeypatch.setattr(main_module, "_folder_operation_ledger",
                        FolderOperationLedger(path_provider=lambda: tmp_path / "tasks.sqlite3"))


# ---------------------------------------------------------------------------
# A. the projection itself
# ---------------------------------------------------------------------------


def _plan_response() -> PlanResponse:
    plan = AnalysisPlan.model_validate(
        {
            "id": "result-write-abc",
            "title": "写入结果",
            "summary": "把本地结果写入 Sheet1!A1:B2",
            "localResultRef": RESULT_REF,
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": "Sheet1",
                    "range": "A1:B2",
                    "values": [[SENTINEL, 1], [SENTINEL, 2]],
                }
            ],
            "acceptanceCriteria": [
                {
                    "type": "rangeEquals",
                    "sheet": "Sheet1",
                    "range": "A1:B2",
                    "expected": [[SENTINEL, 1], [SENTINEL, 2]],
                }
            ],
        }
    )
    return PlanResponse.model_validate(
        {"kind": "plan", "provider": "model", "plan": plan.model_dump(mode="json")}
    )


def test_plan_recovery_descriptor_is_versioned_and_carries_a_reference():
    descriptor = ledger_recovery_descriptor(_plan_response())
    assert descriptor["schemaVersion"] == RECOVERY_SCHEMA_VERSION
    assert descriptor["kind"] == "plan"
    # 明确写出"值被省略"，而不是留一个看起来完整的 plan。
    assert descriptor["valuesOmitted"] is True
    assert descriptor["resultRef"] == RESULT_REF
    assert SENTINEL not in json.dumps(descriptor, ensure_ascii=False)


def test_answer_result_context_values_are_omitted_too():
    response = AnswerResponse.model_validate(
        {
            "kind": "answer",
            "provider": "local",
            "message": "已统计",
            "resultContext": {
                "title": "汇总",
                "headers": ["类别", "数量"],
                "rows": [[SENTINEL, 2]],
                "sourceSheets": ["Sheet1"],
                "warnings": [],
            },
        }
    )
    descriptor = ledger_recovery_descriptor(response)
    assert descriptor["valuesOmitted"] is True
    assert SENTINEL not in json.dumps(descriptor, ensure_ascii=False)


def test_lossless_response_is_not_flagged_and_keeps_its_content():
    response = AnswerResponse.model_validate(
        {"kind": "answer", "provider": "model", "message": "已收到完整规则"}
    )
    descriptor = ledger_recovery_descriptor(response)
    assert descriptor["valuesOmitted"] is False
    assert descriptor["resultRef"] is None
    assert descriptor["response"]["message"] == "已收到完整规则"


# ---------------------------------------------------------------------------
# B. status / replay must not present a projection as a complete response
# ---------------------------------------------------------------------------


def _seed_completed(turn_id: str, response) -> None:
    request = main_module.PlanRequest.model_validate(
        {
            "turnId": turn_id,
            "prompt": "把结果写进表",
            "planningReadPolicy": "none",
            "hostId": "excel-office",
            "hostVersion": "16.0",
            "workbook": {
                "name": "测试.xlsx",
                "capturedAt": "2026-10-09T00:00:00Z",
                "activeWorksheet": "Sheet1",
                "worksheets": [
                    {
                        "name": "Sheet1",
                        "usedRange": "A1:B2",
                        "rowCount": 2,
                        "columnCount": 2,
                        "headers": ["类别", "数量"],
                    }
                ],
            },
        }
    )
    main_module._folder_operation_ledger.register_planning_task(
        turn_id,
        main_module._planning_request_fingerprint(request),
        request.model_dump(mode="json", exclude_none=True),
        approved_wall_time_seconds=60,
    )
    main_module._folder_operation_ledger.finish_planning_task(
        turn_id,
        status="completed",
        result=ledger_recovery_descriptor(response),
    )


def test_status_does_not_replay_a_value_stripped_plan_as_a_plan_response():
    turn_id = f"turn-ledger-{uuid.uuid4().hex[:12]}"
    _seed_completed(turn_id, _plan_response())
    client = TestClient(app)
    body = client.get(f"/api/turn/{turn_id}/status").json()
    assert body["status"] == "completed"
    # 关键：**不能**给一个 `result` 让它看起来像完整 PlanResponse。
    assert "result" not in body
    recovery = body["recovery"]
    assert recovery["schemaVersion"] == RECOVERY_SCHEMA_VERSION
    assert recovery["valuesOmitted"] is True
    assert recovery["resultRef"] == RESULT_REF
    assert SENTINEL not in json.dumps(body, ensure_ascii=False)


def test_status_still_returns_a_lossless_completed_result():
    """普通后端的完成结果取回：没有值被省略时仍按 `result` 回放。"""
    turn_id = f"turn-ledger-{uuid.uuid4().hex[:12]}"
    _seed_completed(
        turn_id,
        AnswerResponse.model_validate(
            {"kind": "answer", "provider": "model", "message": "已找回原任务结果"}
        ),
    )
    client = TestClient(app)
    body = client.get(f"/api/turn/{turn_id}/status").json()
    assert body["status"] == "completed"
    assert body["result"]["message"] == "已找回原任务结果"
    assert "recovery" not in body
