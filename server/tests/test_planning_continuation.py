"""A0「继续分析」的服务端判别力测试。

全部经**真实入口**：

- 规划暂停与受控现场：`POST /api/turn/stream`（真实 `planning_task` 账本 + 真实盖章）
- 查询：`GET /api/turn/{turnId}/status`
- 继续：`POST /api/turn/continue/stream`

不直接调用 `planning_context.validate` 断言"能不能继续"——纯函数已由
`test_planning_context.py`（共享夹具）覆盖；这里关心的是**接线之后**的行为：
损坏/别处/旧版本的现场会不会真的被拒、`canContinue` 会不会漏过一个已经写过
原表的任务、双击会不会跑两遍模型。
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
from server.app.models import AnswerResponse
from server.app.turn_state import turn_registry

SHEET = "数据"
IDENTITY = "url:https://example.invalid/a0-continuation.xlsx"
SOURCE_FINGERPRINT = "sha256:a0-source"


@pytest.fixture(autouse=True)
def isolate_app_state(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """每个用例一份账本：任务身份、检查点、执行记录都不跨用例泄漏。"""
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


def _payload(turn_id: str, *, identity: str = IDENTITY) -> dict:
    return {
        "turnId": turn_id,
        "prompt": "查看当前数据",
        "workbook": {
            "name": "A0 数据.xlsx",
            "capturedAt": "2026-10-19T00:00:00.000Z",
            "activeWorksheet": SHEET,
            "workbookIdentity": identity,
            "sourceFingerprint": SOURCE_FINGERPRINT,
            "sourceFingerprintSheets": [SHEET],
            "worksheets": [
                {
                    "name": SHEET,
                    "rowCount": 2,
                    "columnCount": 2,
                    "headers": ["分类", "数值"],
                    "dataRows": [["甲", 10]],
                }
            ],
        },
        "modelId": "",
    }


SDK_PROGRESS = {
    "stage": "after_tool",
    "confirmations": [
        {"questionId": "q1", "label": "统计口径", "answerLabel": "按月"}
    ],
    "completedTools": ["query_table"],
    "resultRefs": [],
    "failedTool": None,
}


def _install_pausing_planner(monkeypatch: pytest.MonkeyPatch) -> list[dict | None]:
    """让本轮规划在保存现场后暂停；续跑时返回一条普通回答。"""
    resumed: list[dict | None] = []

    async def pausing_create_plan(
        _request: object,
        *,
        tool_cache: object = None,
        on_event=None,
        on_checkpoint=None,
        resume_checkpoint=None,
        resume_revision=None,
        approved_wall_time_seconds=None,
        reminder_seconds=None,
    ) -> AnswerResponse:
        resumed.append(resume_checkpoint)
        if resume_checkpoint is None:
            if on_checkpoint is not None:
                await on_checkpoint(dict(SDK_PROGRESS))
            raise main_module.AgentBudgetPause(
                dict(SDK_PROGRESS), float(approved_wall_time_seconds or 60)
            )
        return AnswerResponse(kind="answer", message="已按当前授权重新分析", provider="local")

    monkeypatch.setattr(main_module, "create_plan", pausing_create_plan)
    return resumed


def _pause(client: TestClient, turn_id: str, *, identity: str = IDENTITY) -> dict:
    response = client.post("/api/turn/stream", json=_payload(turn_id, identity=identity))
    assert response.status_code == 200
    paused = [value for name, value in _parse_sse(response.text) if name == "paused"]
    assert paused, response.text
    return paused[0]


def _continue_body(pause: dict, request_id: str, **overrides: object) -> dict:
    """一份**与检查点一致**的继续请求（含协议必填的现场证据）。

    证据不是可选项：缺了它服务端直接 422。所以默认值必须"对得上"，
    各个用例再用 `overrides` 把某一项弄坏，才能证明是**那一项**在起作用。
    """
    body: dict[str, object] = {
        "turnId": pause["turnId"],
        "expectedRevision": pause["revision"],
        "requestId": request_id,
        "confirmed": True,
        "afterSequence": pause["lastSequence"],
        "workbookIdentity": IDENTITY,
        "authorizedSheetNames": [SHEET],
        "sourceVerification": "verified",
        "sourceFingerprint": SOURCE_FINGERPRINT,
    }
    body.update(overrides)
    return body


def _force_checkpoint(turn_id: str, checkpoint: dict) -> None:
    """把一份**坏现场**直接写进账本。

    这是唯一能构造"库里躺着一份损坏/别处/旧版本的现场"的方式：
    `save_planning_checkpoint` / `pause_planning_task` 都要求任务正处于
    `running`，而真实事故（进程被杀、旧版本写下的记录、文件损坏）留下的
    就是这种状态。用例要验证的正是：**读取侧**会不会无条件相信它。
    """
    ledger = main_module._folder_operation_ledger
    with sqlite3.connect(ledger._path()) as connection:
        connection.execute(
            "UPDATE planning_tasks SET checkpoint_json = ? WHERE turn_id = ?",
            (
                json.dumps(checkpoint, ensure_ascii=False, sort_keys=True),
                turn_id,
            ),
        )
        connection.commit()


def _valid_checkpoint(pause: dict) -> dict:
    record = main_module._folder_operation_ledger.planning_task(pause["turnId"])
    assert record is not None and record.checkpoint is not None
    return record.checkpoint


def test_pause_reports_continuation_from_the_shared_contract(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """合法现场：状态与暂停载荷都给出"可继续"，且带阶段。"""
    _install_pausing_planner(monkeypatch)
    client = TestClient(app)
    pause = _pause(client, "turn-a0-ok")

    assert pause["canContinue"] is True
    assert pause["continuation"] == {
        "available": True,
        "code": None,
        "message": "",
        "stage": "after_tool",
        # 这份现场记录了来源指纹与来源表 ⇒ 继续前必须重读比对。
        "sourceVerificationRequired": True,
    }
    status = client.get("/api/turn/turn-a0-ok/status").json()
    assert status["canContinue"] is True
    assert status["continuation"]["available"] is True


@pytest.mark.parametrize(
    ("mutate", "expected_code"),
    [
        # 旧循环检查点：没有 kind，绝不能当 planning_context 用。
        (
            lambda base: {
                "schemaVersion": 2,
                "kind": "planning",
                "messages": [{"role": "user", "content": "查看当前数据"}],
                "toolCache": {},
                "loopState": {"step": 3},
            },
            "PLANNING_CONTEXT_KIND_MISMATCH",
        ),
        # 版本不受支持：字段对了也不认。
        (
            lambda base: {**base, "schemaVersion": 99},
            "PLANNING_CONTEXT_UNSUPPORTED_VERSION",
        ),
        # 损坏：缺必填字段。
        (
            lambda base: {
                key: value
                for key, value in base.items()
                if key not in {"target", "authorization"}
            },
            "PLANNING_CONTEXT_INVALID",
        ),
        # 含值矩阵：一次都不许落盘，更不许据此继续。
        (
            lambda base: {
                **base,
                "confirmations": [
                    {
                        "questionId": "q1",
                        "label": "统计口径",
                        "answerLabel": "按月",
                        "values": [["甲", 10]],
                    }
                ],
            },
            "PLANNING_CONTEXT_SENSITIVE",
        ),
        # 别的任务：现场不能串用。
        (
            lambda base: {**base, "taskId": "turn-somebody-else"},
            "PLANNING_CONTEXT_TASK_MISMATCH",
        ),
        # 阶段不支持：只有"尚未进入宿主写入"的两个阶段可以继续。
        (
            lambda base: {**base, "stage": "after_execution"},
            "PLANNING_CONTEXT_STAGE_UNSUPPORTED",
        ),
        # 继续资格是 replan：说明结果不确定，不能直接重跑。
        (
            lambda base: {**base, "resumability": "replan"},
            "PLANNING_CONTEXT_NOT_CONTINUABLE",
        ),
        # 版本比记录新：拿未来的现场续跑等于跳过用户已看到的进度。
        (
            lambda base: {**base, "taskRevision": 999},
            "PLANNING_CONTEXT_STALE_REVISION",
        ),
    ],
)
def test_bad_checkpoint_is_refused_by_the_continue_endpoint(
    monkeypatch: pytest.MonkeyPatch,
    mutate,
    expected_code: str,
) -> None:
    """损坏 / 伪造 / 旧版本 / 别处来的现场：状态如实说不能继续，继续接口也拒。"""
    resumed = _install_pausing_planner(monkeypatch)
    client = TestClient(app)
    pause = _pause(client, f"turn-a0-bad-{expected_code.lower()}")
    base = _valid_checkpoint(pause)

    _force_checkpoint(pause["turnId"], mutate(base))

    status = client.get(f"/api/turn/{pause['turnId']}/status").json()
    assert status["canContinue"] is False
    assert status["continuation"]["available"] is False
    assert status["continuation"]["code"] == expected_code

    refused = client.post(
        "/api/turn/continue/stream",
        json=_continue_body(pause, "continue-bad-0001"),
    )
    assert refused.status_code in {404, 409}
    assert refused.json()["detail"]["code"] == expected_code
    # 一次都没续跑：拒绝发生在 claim 之前。
    assert resumed == [None]


def test_checkpoint_with_foreign_workbook_is_refused_before_planning(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """换工作簿：现场本身合法，但客户端送上来的身份对不上就拒。"""
    resumed = _install_pausing_planner(monkeypatch)
    client = TestClient(app)
    pause = _pause(client, "turn-a0-identity")

    refused = client.post(
        "/api/turn/continue/stream",
        json=_continue_body(
            pause,
            "continue-identity-0001",
            workbookIdentity="url:https://example.invalid/another.xlsx",
        ),
    )
    assert refused.status_code == 409
    assert refused.json()["detail"]["code"] == "PLANNING_CONTEXT_TARGET_CHANGED"
    assert resumed == [None]


@pytest.mark.parametrize(
    ("field", "value", "expected_code"),
    [
        # 扩大：绝不允许用"继续"把授权推宽到原任务没授权的表。
        (
            "authorizedSheetNames",
            [SHEET, "另一张表"],
            "PLANNING_CONTEXT_SCOPE_CHANGED",
        ),
        # 来源变了：旧结果引用不再可信。
        ("sourceFingerprint", "sha256:changed", "PLANNING_CONTEXT_SOURCE_CHANGED"),
    ],
)
def test_scope_or_source_change_is_refused_by_the_server(
    monkeypatch: pytest.MonkeyPatch,
    field: str,
    value: object,
    expected_code: str,
) -> None:
    resumed = _install_pausing_planner(monkeypatch)
    client = TestClient(app)
    pause = _pause(client, f"turn-a0-{field.lower()}-{len(str(value))}")

    refused = client.post(
        "/api/turn/continue/stream",
        json=_continue_body(pause, "continue-scope-0001", **{field: value}),
    )
    assert refused.status_code == 409
    assert refused.json()["detail"]["code"] == expected_code
    assert resumed == [None]


def test_narrowed_authorization_does_not_reuse_the_old_checkpoint(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """收窄授权同样不沿用旧现场：旧现场是按**更宽**范围读取的。

    这里用两张授权表的任务 + 只提交一张表的核实结果。放任它继续，服务端就会
    按旧请求（两张表）重新读取，等于把用户刚撤销的授权又用了一次。
    """
    resumed = _install_pausing_planner(monkeypatch)
    client = TestClient(app)
    turn_id = "turn-a0-narrowed"
    payload = _payload(turn_id)
    payload["workbook"]["worksheets"].append(
        {"name": "另一张表", "rowCount": 1, "columnCount": 1, "headers": ["x"]}
    )
    payload["workbook"]["sourceFingerprintSheets"] = [SHEET, "另一张表"]
    response = client.post("/api/turn/stream", json=payload)
    assert response.status_code == 200
    pause = [value for name, value in _parse_sse(response.text) if name == "paused"][0]

    refused = client.post(
        "/api/turn/continue/stream",
        json=_continue_body(
            pause, "continue-narrowed-0001", authorizedSheetNames=[SHEET]
        ),
    )
    assert refused.status_code == 409
    assert refused.json()["detail"]["code"] == "PLANNING_CONTEXT_SCOPE_CHANGED"
    assert resumed == [None]


def test_verified_target_still_continues(monkeypatch: pytest.MonkeyPatch) -> None:
    """正向锚点：核实结果一致时照常继续，不会被上面那些闸门误伤。"""
    resumed = _install_pausing_planner(monkeypatch)
    client = TestClient(app)
    pause = _pause(client, "turn-a0-verified")

    continued = client.post(
        "/api/turn/continue/stream",
        json=_continue_body(
            pause,
            "continue-verified-0001",
            workbookIdentity=IDENTITY,
            authorizedSheetNames=[SHEET],
            sourceFingerprint=SOURCE_FINGERPRINT,
        ),
    )
    assert continued.status_code == 200
    result = [value for name, value in _parse_sse(continued.text) if name == "result"]
    assert result and result[0]["message"] == "已按当前授权重新分析"
    assert resumed[1] is not None


def _register_execution(turn_id: str, operation_id: str) -> dict:
    client = TestClient(app)
    registered = client.post(
        "/api/office/executions/register",
        json={
            "operationId": operation_id,
            "plan": {
                "id": "a0-plan",
                "title": "写入测试",
                "summary": "A0 执行账本",
                "assumptions": [],
                "warnings": [],
                "actions": [
                    {
                        "type": "writeValues",
                        "sheet": SHEET,
                        "range": "A1",
                        "values": [["ok"]],
                    }
                ],
            },
            "authorizedSheetNames": [SHEET],
            "allowedExternalFiles": [],
            "workbookIdentity": IDENTITY,
            "taskId": turn_id,
            "confirmed": True,
        },
    )
    assert registered.status_code == 200, registered.text
    return registered.json()


def _complete_execution(state: dict) -> None:
    """把这次执行走到"已完成"，用真实端点而不是改库。"""
    client = TestClient(app)
    operation_id = state["operationId"]
    before = {
        "actionIndex": 0,
        "sheet": SHEET,
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
    assert started.status_code == 200, started.text
    state = started.json()
    completed = client.post(
        f"/api/office/executions/{operation_id}/actions/complete",
        json={
            "expectedRevision": state["revision"],
            "ownerToken": state["ownerToken"],
            "actionIndex": 0,
            "checkpoint": {
                "planId": "a0-plan",
                "executionId": operation_id,
                "workbookIdentity": IDENTITY,
                "capturedAt": "2026-10-19T00:00:00Z",
                "coverage": "full",
                "completedActionCount": 1,
                "uncertainActionIndexes": [],
                "postStateCoverage": "complete",
                "ranges": [{**before, "postFingerprint": "fnv1a32:deadbeef"}],
                "createdWorksheets": [],
                "createdWorksheetStates": [],
                "deletedWorksheets": [],
                "skippedWorksheets": [],
            },
        },
    )
    assert completed.status_code == 200, completed.text
    state = completed.json()
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
    assert finished.status_code == 200, finished.text
    assert finished.json()["status"] == "completed"


@pytest.mark.parametrize("settle", ["running", "completed", "suspended"])
def test_registered_execution_blocks_continuation_in_every_state(
    monkeypatch: pytest.MonkeyPatch,
    settle: str,
) -> None:
    """executing / completed / partial-unknown 都**不**等于"尚未执行"。

    只要这个任务登记过写入，就不许用重新分析绕过去 —— 无论那次写入是正在跑、
    已经写完，还是结果未知被挂起。
    """
    resumed = _install_pausing_planner(monkeypatch)
    client = TestClient(app)
    pause = _pause(client, f"turn-a0-exec-{settle}")
    state = _register_execution(pause["turnId"], f"a0-execution-{settle}-0001")
    if settle == "completed":
        _complete_execution(state)
    elif settle == "suspended":
        main_module._folder_operation_ledger.suspend_office_execution(
            state["operationId"],
            expected_revision=state["revision"],
            owner_token=state["ownerToken"],
            error={"code": "EXECUTION_UNKNOWN", "message": "结果未知"},
        )

    status = client.get(f"/api/turn/{pause['turnId']}/status").json()
    assert status["canContinue"] is False
    assert status["continuation"]["code"] == "PLANNING_CONTEXT_HAS_EXECUTION"
    assert "核对" in status["continuation"]["message"]

    refused = client.post(
        "/api/turn/continue/stream",
        json=_continue_body(pause, f"continue-exec-{settle}-0001"),
    )
    assert refused.status_code == 409
    assert refused.json()["detail"]["code"] == "PLANNING_CONTEXT_HAS_EXECUTION"
    # 没有绕过执行核对：模型一步都没跑。
    assert resumed == [None]
    # 执行账本与规划账本仍然是两件事：规划任务没有被这次拒绝改写。
    record = main_module._folder_operation_ledger.planning_task(pause["turnId"])
    assert record is not None and record.status == "waiting_user"


def test_continuation_never_registers_an_execution(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """规划恢复 ≠ 执行恢复：继续分析不产生任何写入登记。"""
    _install_pausing_planner(monkeypatch)
    client = TestClient(app)
    pause = _pause(client, "turn-a0-no-write")

    continued = client.post(
        "/api/turn/continue/stream",
        json=_continue_body(pause, "continue-no-write-0001"),
    )
    assert continued.status_code == 200

    ledger = main_module._folder_operation_ledger
    assert ledger.has_execution_for_task(pause["turnId"]) is False
    record = ledger.planning_task(pause["turnId"])
    assert record is not None and record.status == "completed"


def test_double_click_claims_the_continuation_only_once(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """双击 / 双客户端：只有一次取得执行权，模型不会被跑两遍。"""
    resumed = _install_pausing_planner(monkeypatch)
    client = TestClient(app)
    pause = _pause(client, "turn-a0-double")

    first = client.post(
        "/api/turn/continue/stream",
        json=_continue_body(pause, "continue-double-0001"),
    )
    assert first.status_code == 200
    second = client.post(
        "/api/turn/continue/stream",
        json=_continue_body(pause, "continue-double-0002"),
    )
    assert second.status_code == 409
    assert second.json()["detail"]["code"] in {
        "PLANNING_CONTINUE_CONFLICT",
        "PLANNING_TASK_STOPPED",
        "PLANNING_CONTEXT_MISSING",
        "PLANNING_TASK_STILL_RUNNING",
    }
    # 初始一轮 + 恰好一次续跑。
    assert len(resumed) == 2


def test_task_without_checkpoint_reports_missing_not_replan(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """没有现场：如实说"没有可继续的现场"，不假装可以重新分析。"""
    _install_pausing_planner(monkeypatch)
    client = TestClient(app)
    pause = _pause(client, "turn-a0-missing")

    ledger = main_module._folder_operation_ledger
    with sqlite3.connect(ledger._path()) as connection:
        connection.execute(
            "UPDATE planning_tasks SET checkpoint_json = NULL WHERE turn_id = ?",
            (pause["turnId"],),
        )
        connection.commit()

    status = client.get(f"/api/turn/{pause['turnId']}/status").json()
    assert status["continuation"]["available"] is False
    assert status["continuation"]["code"] == "PLANNING_CONTEXT_MISSING"
    assert status["canContinue"] is False

    refused = client.post(
        "/api/turn/continue/stream",
        json=_continue_body(pause, "continue-missing-0001"),
    )
    assert refused.status_code == 409
    assert refused.json()["detail"]["code"] == "PLANNING_CONTEXT_MISSING"


def test_unknown_turn_is_not_found_without_creating_a_task() -> None:
    """查询别人的 / 不存在的任务：404，且不产生任何新任务记录。"""
    client = TestClient(app)
    # 证据齐全：这里要证明的是"未知任务 404"，不是"缺证据 422"。
    response = client.post(
        "/api/turn/continue/stream",
        json={
            "turnId": "turn-never-existed",
            "expectedRevision": 1,
            "requestId": "continue-unknown-0001",
            "confirmed": True,
            "afterSequence": 0,
            "workbookIdentity": IDENTITY,
            "authorizedSheetNames": [SHEET],
            "sourceVerification": "not_collected",
        },
    )
    assert response.status_code == 404
    assert response.json()["detail"]["code"] == "PLANNING_TASK_NOT_FOUND"
    assert main_module._folder_operation_ledger.planning_task("turn-never-existed") is None


# ---------------------------------------------------------------------------
# 现场证据闸（2026-10-20）：缺证据不许继续
#
# 这一组用例证明的是**"省略字段不能通过"**：上一版的 `_reverify_continuation_target`
# 只在字段非 None 时比对，三项全省略就返回 None —— 等于不核实也能继续。
# 现在证据是协议必填（缺了 422），并且"读取失败"与"尚未采集"是两种状态。
# ---------------------------------------------------------------------------


def _dropped(body: dict, key: str) -> dict:
    """删掉一个**必填**字段：模拟漏传，而不是故意填空值。"""
    payload = dict(body)
    payload.pop(key, None)
    return payload


@pytest.mark.parametrize(
    ("missing", "label"),
    [
        ("workbookIdentity", "工作簿身份"),
        ("authorizedSheetNames", "授权成员"),
        ("sourceVerification", "来源核实状态"),
    ],
)
def test_missing_evidence_field_is_rejected_before_anything_runs(
    monkeypatch: pytest.MonkeyPatch, missing: str, label: str
) -> None:
    """漏传任一必填证据 ⇒ 422，且**没有续跑**（模型一次都没被调用）。"""
    resumed = _install_pausing_planner(monkeypatch)
    client = TestClient(app)
    pause = _pause(client, f"turn-evidence-missing-{missing.lower()}")

    refused = client.post(
        "/api/turn/continue/stream",
        json=_dropped(_continue_body(pause, "continue-evidence-miss"), missing),
    )
    assert refused.status_code == 422, refused.text
    assert resumed == [None]

    # 后续一次**合法**的继续必须仍然成功：说明上一次没有领取续跑权、没有扣预算。
    ok = client.post(
        "/api/turn/continue/stream",
        json=_continue_body(pause, "continue-evidence-ok-1"),
    )
    assert ok.status_code == 200, ok.text
    assert resumed[1] is not None


@pytest.mark.parametrize(
    ("field", "value"),
    [
        ("workbookIdentity", ""),
        ("authorizedSheetNames", []),
        ("authorizedSheetNames", ["  "]),
        ("sourceVerification", "verified"),  # 声明已核实却不给指纹
        ("sourceVerification", "maybe"),  # 未知状态
    ],
)
def test_blank_or_inconsistent_evidence_is_rejected(
    monkeypatch: pytest.MonkeyPatch, field: str, value: object
) -> None:
    """空值 / 未知状态 / 自相矛盾：都在协议层被拒，进不到业务判断。"""
    _install_pausing_planner(monkeypatch)
    client = TestClient(app)
    pause = _pause(client, f"turn-evidence-blank-{field.lower()}-{len(str(value))}")

    body = _continue_body(pause, "continue-evidence-blank")
    if field == "sourceVerification" and value == "verified":
        body.pop("sourceFingerprint", None)
    body[field] = value
    refused = client.post("/api/turn/continue/stream", json=body)
    assert refused.status_code == 422, refused.text


def test_source_read_failure_is_not_treated_as_not_collected(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """`failed` 是**读取失败**，必须拒绝，不能退化成"尚未采集"。

    这条是缺口一的核心：以前"读不到宿主"会走到 `verified = {}`，字段全部省略，
    服务端 `None` 直接放行 —— 读不到 = 不需要读。
    """
    resumed = _install_pausing_planner(monkeypatch)
    client = TestClient(app)
    pause = _pause(client, "turn-evidence-source-failed")

    refused = client.post(
        "/api/turn/continue/stream",
        json=_continue_body(
            pause,
            "continue-evidence-failed",
            sourceVerification="failed",
            sourceFingerprint=None,
        ),
    )
    assert refused.status_code == 409
    assert refused.json()["detail"]["code"] == "PLANNING_CONTEXT_EVIDENCE_UNREADABLE"
    assert resumed == [None]


def test_not_collected_is_refused_when_the_task_depends_on_source(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """原任务读过数据（记录了来源指纹/来源表）⇒ "尚未采集"不成立。"""
    resumed = _install_pausing_planner(monkeypatch)
    client = TestClient(app)
    pause = _pause(client, "turn-evidence-source-required")

    refused = client.post(
        "/api/turn/continue/stream",
        json=_continue_body(
            pause,
            "continue-evidence-notcollected",
            sourceVerification="not_collected",
            sourceFingerprint=None,
        ),
    )
    assert refused.status_code == 409
    assert refused.json()["detail"]["code"] == "PLANNING_CONTEXT_SOURCE_UNVERIFIED"
    assert resumed == [None]


def test_structural_task_may_continue_without_reading_source(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """正向锚点：**确实没读过来源**的结构任务，报"尚未采集"必须能继续。

    没有这条，上面那道闸就可以靠"一律要求重读来源"通过全部用例 —— 而那会把
    "整理表头 / 规划结构"这类根本不碰数据的任务全部堵死。
    """
    resumed = _install_pausing_planner(monkeypatch)
    client = TestClient(app)
    turn_id = "turn-evidence-structural"
    payload = _payload(turn_id)
    # 去掉来源指纹与来源表：这份现场从未读过数据。
    payload["workbook"].pop("sourceFingerprint", None)
    payload["workbook"].pop("sourceFingerprintSheets", None)
    response = client.post("/api/turn/stream", json=payload)
    assert response.status_code == 200, response.text
    pause = [value for name, value in _parse_sse(response.text) if name == "paused"][0]
    assert pause["continuation"]["sourceVerificationRequired"] is False

    continued = client.post(
        "/api/turn/continue/stream",
        json=_continue_body(
            pause,
            "continue-evidence-structural",
            sourceVerification="not_collected",
            sourceFingerprint=None,
        ),
    )
    assert continued.status_code == 200, continued.text
    assert resumed[1] is not None


def test_evidence_rejection_leaves_no_execution_and_no_state_change(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """被证据闸拒绝的任务：状态、版本、执行登记都不许被动过。"""
    _install_pausing_planner(monkeypatch)
    client = TestClient(app)
    pause = _pause(client, "turn-evidence-untouched")

    before = client.get(f"/api/turn/{pause['turnId']}/status").json()
    refused = client.post(
        "/api/turn/continue/stream",
        json=_continue_body(
            pause,
            "continue-evidence-untouched",
            workbookIdentity="url:https://example.invalid/other.xlsx",
        ),
    )
    assert refused.status_code == 409
    after = client.get(f"/api/turn/{pause['turnId']}/status").json()
    assert after["status"] == before["status"] == "waiting_user"
    assert after["revision"] == before["revision"]
    assert after["canContinue"] is True
    assert main_module._folder_operation_ledger.has_execution_for_task(
        pause["turnId"]
    ) is False
