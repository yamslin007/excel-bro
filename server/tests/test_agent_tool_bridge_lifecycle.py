"""缺陷复现：结果 / 会话 / 授权三条生命周期，以及规划账本的隐私边界。

本文件覆盖三个必须先修的服务端阻断：

1. ``run_sdk_plan`` 的 ``finally`` 关闭会话时把结果一起删掉，导致终局
   ``localResult.resultRef`` 在本轮返回后取不到值。
2. ``submit_result_plan`` 产出的含矩阵 ``PlanResponse`` 被规划账本完整持久化，
   值矩阵落进 ``planning_tasks.result_json``。
3. 桥归属只比 ``hostId`` / 工作簿名，同名工作簿互相冒充；``register`` 端点
   不校验活跃会话归属。

这里的断言全部针对**行为**与**真实请求体**：桥、登记器、编译器、账本与回环
端点的都是生产对象。只有模型传输是测试替身。
"""
from __future__ import annotations

import asyncio
import json
from pathlib import Path

import httpx
import pytest

import server.app.main as main_module
from server.app import claude_sdk_agent
from server.app.local_results import local_results
from server.app.models import (
    AnswerResponse,
    DataToolResult,
    PlanRequest,
    PlanResponse,
)
from server.app.task_bridge import (
    BridgeBinding,
    BridgeError,
    TaskToolBridge,
    register_bridge_result,
    task_tool_bridge,
)
from server.app.turn_state import turn_registry

SENTINEL = "SENTINEL_SECRET_4B71"
WORKBOOK = "报表.xlsx"
# 本地结果句柄：账本与恢复副本里只留它，真值留在本机易失登记器。
RESULT_REF = "res_" + "a1" * 16


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------


def _result(rows=None, headers=None, *, complete=True) -> DataToolResult:
    headers = headers or ["类别", "数量"]
    rows = rows if rows is not None else [["甲", 2], ["乙", 1]]
    return DataToolResult.model_validate({
        "requestId": "req-1",
        "tool": "query_table",
        "title": f"汇总结果 {SENTINEL}",
        "headers": headers,
        "rows": rows,
        "sourceSheets": ["Sheet1"],
        "scannedRows": 12,
        "complete": complete,
        "calculation": f"按类别分组计数 {SENTINEL}",
        "warnings": [f"注意 {SENTINEL}"],
    })


def _binding(task_id: str = "turn-life-0001", **overrides) -> BridgeBinding:
    defaults = dict(
        task_id=task_id,
        host_id="excel-office",
        workbook_name=WORKBOOK,
        workbook_identity=f"url:https://contoso.example/{WORKBOOK}",
        source_fingerprint="fp-source-1",
        authorized_key="a" * 64,
    )
    defaults.update(overrides)
    return BridgeBinding(**defaults)


_TASK = "turn-life-0001"


@pytest.fixture(autouse=True)
def _clean_state(monkeypatch: pytest.MonkeyPatch, tmp_path: Path):
    turn_registry.reset()
    local_results.reset()
    task_tool_bridge.reset()
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path / "app-data"))
    monkeypatch.setattr(
        main_module._folder_operation_ledger,
        "_path_provider",
        lambda: tmp_path / "app-data" / "tasks.sqlite3",
    )
    yield
    turn_registry.reset()
    local_results.reset()
    task_tool_bridge.reset()


def _localhost_client() -> httpx.AsyncClient:
    from server.app.main import app

    return httpx.AsyncClient(
        transport=httpx.ASGITransport(app=app), base_url="http://localhost"
    )


def _body(task_id: str = _TASK, **overrides) -> dict:
    payload = {
        "taskId": task_id,
        "hostId": "excel-office",
        "workbookName": WORKBOOK,
        "workbookIdentity": f"url:https://contoso.example/{WORKBOOK}",
        "sourceFingerprint": "fp-source-1",
    }
    payload.update(overrides)
    return payload


async def _deliver_into(bridge: TaskToolBridge, binding: BridgeBinding, *, result=None):
    """Run one real query_table round trip and return the typed receipt."""
    session = bridge.open(binding)

    async def scenario():
        task = asyncio.create_task(session.request("query_table", {"mode": "rows"}))
        for _ in range(2000):
            if session.pending_calls:
                break
            await asyncio.sleep(0)
        call = session.pending_calls[0]
        session.deliver(
            operation_id=call["operationId"],
            request_fingerprint=call["requestFingerprint"],
            host_id=binding.host_id,
            workbook_name=binding.workbook_name,
            workbook_identity=binding.workbook_identity,
            source_fingerprint=binding.source_fingerprint,
            result=result or _result(),
            error_code=None,
        )
        return await task

    return await scenario()


# ---------------------------------------------------------------------------
# 缺陷 1：结果生命周期必须长于会话生命周期
# ---------------------------------------------------------------------------


def test_registered_result_survives_session_close():
    """会话关闭只结束会话，不结束结果：本轮返回后终局引用仍可取回。"""
    bridge = TaskToolBridge()
    binding = _binding()
    bridge.open(binding)
    handle = register_bridge_result(_TASK, _result(), binding.source_fingerprint)

    bridge.close(_TASK)

    assert local_results.resolve(_TASK, handle) is not None


def test_sdk_turn_result_ref_is_still_resolvable_after_run_sdk_plan(monkeypatch, tmp_path):
    """真实 SDK 闭环：``run_sdk_plan`` 返回后终局引用仍能解析出值。"""
    from server.app.llm.config import ModelConnection

    holder: dict = {}

    class _FakeResult:
        def __init__(self, result: str = "") -> None:
            self.result = result
            self.is_error = False
            self.subtype = "success"
            self.api_error_status = None

    class _Sdk:
        ResultMessage = _FakeResult
        CLINotFoundError = type("CLINotFoundError", (Exception,), {})
        CLIConnectionError = type("CLIConnectionError", (Exception,), {})
        ProcessError = type("ProcessError", (Exception,), {})
        ResultError = type("ResultError", (Exception,), {})
        CLIJSONDecodeError = type("CLIJSONDecodeError", (Exception,), {})

        def ClaudeAgentOptions(self, **kwargs):
            from types import SimpleNamespace

            return SimpleNamespace(**kwargs)

        def tool(self, name, _description, _schema):
            return lambda handler: type("T", (), {"name": name, "handler": handler})

        def create_sdk_mcp_server(self, *, name, version, tools):
            return {tool.name: tool for tool in tools}

        async def query(self, *, prompt, options):
            tool = options.mcp_servers["excel_bro"]["query_table"]
            task = asyncio.create_task(tool.handler({"mode": "rows"}))
            session = task_tool_bridge.session(_TASK)
            for _ in range(2000):
                if session is not None and session.pending_calls:
                    break
                await asyncio.sleep(0)
            call = session.pending_calls[0]
            holder["ref"] = session.deliver(
                operation_id=call["operationId"],
                request_fingerprint=call["requestFingerprint"],
                host_id="excel-office",
                workbook_name=WORKBOOK,
                workbook_identity=f"url:https://contoso.example/{WORKBOOK}",
                source_fingerprint="fp-source-1",
                result=_result(),
                error_code=None,
            )
            first = await task
            ref = json.loads(first["content"][0]["text"])["resultRef"]
            submit = options.mcp_servers["excel_bro"]["submit_local_result"]
            await submit.handler({"resultRef": ref, "message": "查询完成"})
            yield _FakeResult("")

    monkeypatch.setattr(claude_sdk_agent, "_sdk", lambda: _Sdk())
    monkeypatch.setattr(claude_sdk_agent, "_state_dir", lambda: tmp_path)

    request = PlanRequest.model_validate({
        "turnId": _TASK,
        "prompt": "把两个表按类别关联并汇总",
        "planningReadPolicy": "none",
        "hostId": "excel-office",
        "hostVersion": "16.0",
        "workbook": {
            "name": WORKBOOK,
            "workbookIdentity": f"url:https://contoso.example/{WORKBOOK}",
            "capturedAt": "2026-10-09T00:00:00Z",
            "activeWorksheet": "Sheet1",
            "sourceFingerprint": "fp-source-1",
            "worksheets": [
                {
                    "name": "Sheet1", "usedRange": "A1:B3", "rowCount": 3,
                    "columnCount": 2, "headers": ["类别", "数量"],
                },
            ],
        },
    })
    connection = ModelConnection(
        base_url="https://api.anthropic.com",
        model="test-claude",
        api_key="test-key",
        supports_vision=False,
        adapter="claude_agent_sdk",
    )
    response = asyncio.run(claude_sdk_agent.run_sdk_plan(request, connection))

    assert isinstance(response, AnswerResponse)
    assert response.localResult is not None
    # 本轮已返回、会话已关闭 —— 引用仍必须可解析。
    resolved = local_results.resolve(_TASK, response.localResult.resultRef)
    assert resolved is not None
    assert resolved.rows == [["甲", 2], ["乙", 1]]


def test_result_endpoint_still_serves_values_after_the_session_closes():
    """终局引用要能通过回环端点取回真实值（结果卡渲染路径）。"""
    bridge = TaskToolBridge()
    binding = _binding()
    handle = asyncio.run(_deliver_into(bridge, binding)).get("resultRef")
    assert isinstance(handle, str) and handle.startswith("res_")

    async def scenario():
        async with _localhost_client() as client:
            fetched = await client.post(
                "/api/sdk/tool-bridge/result",
                json={**_body(), "resultRef": handle},
            )
            assert fetched.status_code == 200
            return fetched.json()

    payload = asyncio.run(scenario())
    assert payload["rows"] == [["甲", 2], ["乙", 1]]


# ---------------------------------------------------------------------------
# 缺陷 2：含值计划不得进入持久规划账本
# ---------------------------------------------------------------------------


def _stream_payload(turn_id: str) -> dict:
    return {
        "turnId": turn_id,
        "prompt": "把结果写到 D1",
        "modelId": "",
        "workbook": {
            "name": WORKBOOK,
            "capturedAt": "2026-10-09T00:00:00.000Z",
            "activeWorksheet": "Sheet1",
            "worksheets": [
                {
                    "name": "Sheet1", "rowCount": 3, "columnCount": 2,
                    "headers": ["类别", "数量"], "dataRows": [["甲", 2]],
                }
            ],
        },
    }


def _plan_response_with_values() -> PlanResponse:
    from server.app.models import AnalysisPlan

    plan = AnalysisPlan.model_validate({
        "id": "result-write-deadbeef",
        "title": "写入 Sheet1!D1:E3",
        "summary": "把本地结果写入 Sheet1 D1:E3",
        "actions": [{
            "type": "writeValues",
            "sheet": "Sheet1",
            "range": "D1:E3",
            "values": [["类别", "数量"], [SENTINEL, 2], ["乙", 1]],
        }],
        # `submit_result_plan` 编译出来的计划一定带本地结果引用 —— 账本要留的就是它。
        "localResultRef": RESULT_REF,
    })
    return PlanResponse(plan=plan, provider="model")


def test_planning_ledger_never_stores_the_value_matrix(monkeypatch):
    """含值计划可以在本机响应里交付，但绝不能落进持久账本。

    账本里存的**不是**"一个看起来完整的 `PlanResponse`"，而是一份**版本化的恢复
    描述**：明确写出"值被省略过"，并给出取真值该走的本地结果引用。这样状态端点
    就不会把一个缺了必填矩阵的形状当成完整响应回放给客户端。
    """
    from fastapi.testclient import TestClient

    from server.app.main import app

    turn_id = "turn-ledger-privacy"

    async def fake_create_plan(_request, **_kwargs):
        return _plan_response_with_values()

    monkeypatch.setattr(main_module, "create_plan", fake_create_plan)

    response = TestClient(app).post("/api/turn/stream", json=_stream_payload(turn_id))
    assert response.status_code == 200

    record = main_module._folder_operation_ledger.planning_task(turn_id)
    assert record is not None and record.result is not None
    stored = json.dumps(record.result, ensure_ascii=False)
    assert SENTINEL not in stored, "结果值不得落进规划账本"

    descriptor = record.result
    assert descriptor["schemaVersion"] == 1
    assert descriptor["kind"] == "plan"
    # 有值被省略 → 必须明确标记；状态端点据此给 `recovery` 而不是 `result`。
    assert descriptor["valuesOmitted"] is True
    # 取真值走**本地结果引用**，不是从账本里还原。
    assert descriptor["resultRef"] == RESULT_REF
    # 正向锚点：描述里仍留着可复算的形状（动作类型与区域），只是没有矩阵。
    action = descriptor["response"]["plan"]["actions"][0]
    assert "values" not in action
    assert action["type"] == "writeValues"
    assert action["range"] == "D1:E3"


def test_stream_still_hands_the_full_plan_to_the_local_client(monkeypatch):
    """本机客户端仍要拿到完整计划（预览/确认需要值），只是不落盘。"""
    from fastapi.testclient import TestClient

    from server.app.main import app

    turn_id = "turn-ledger-delivery"

    async def fake_create_plan(_request, **_kwargs):
        return _plan_response_with_values()

    monkeypatch.setattr(main_module, "create_plan", fake_create_plan)

    response = TestClient(app).post("/api/turn/stream", json=_stream_payload(turn_id))
    result_frame = next(
        payload
        for name, payload in _parse_sse(response.text)
        if name == "result"
    )
    values = result_frame["plan"]["actions"][0]["values"]
    assert values[1][0] == SENTINEL


def _parse_sse(text: str) -> list[tuple[str, dict]]:
    events: list[tuple[str, dict]] = []
    for frame in text.split("\n\n"):
        name = "message"
        data = None
        for line in frame.split("\n"):
            if line.startswith("event:"):
                name = line[6:].strip()
            elif line.startswith("data:"):
                data = line[5:].lstrip()
        if data:
            events.append((name, json.loads(data)))
    return events


def test_sse_parser_self_check():
    """自检：SSE 解析器本身可用（避免测试替身自身出错）。"""
    assert _parse_sse("event: result\ndata: {\"kind\": \"plan\"}") == [
        ("result", {"kind": "plan"})
    ]


# ---------------------------------------------------------------------------
# 缺陷 3：桥归属与 register 会话校验
# ---------------------------------------------------------------------------


def test_same_named_workbook_cannot_deliver_into_another_binding():
    """同名但不同工作簿（workbookIdentity 不同）的回执必须被拒。"""
    bridge = TaskToolBridge()
    binding = _binding(workbook_name=WORKBOOK)
    session = bridge.open(binding)
    session._pending["op_" + "1" * 32] = _pending_entry()

    with pytest.raises(BridgeError) as caught:
        session.deliver(
            operation_id="op_" + "1" * 32,
            request_fingerprint="0" * 64,
            host_id="excel-office",
            workbook_name=WORKBOOK,
            workbook_identity="url:https://evil.example/报表.xlsx",
            source_fingerprint="fp-source-1",
            result=_result(),
            error_code=None,
        )
    assert caught.value.code == "stale_binding"


def test_poll_does_not_leak_calls_to_a_same_named_workbook():
    binding = _binding()
    task_tool_bridge.open(binding)

    async def scenario():
        session = task_tool_bridge.session(_TASK)
        task = asyncio.create_task(session.request("query_table", {"mode": "rows"}))
        for _ in range(2000):
            if session.pending_calls:
                break
            await asyncio.sleep(0)
        async with _localhost_client() as client:
            impostor = await client.post(
                "/api/sdk/tool-bridge/poll",
                json=_body(workbookIdentity="url:https://evil.example/报表.xlsx"),
            )
            owner = await client.post(
                "/api/sdk/tool-bridge/poll", json=_body()
            )
        session.cancel("test_done")
        await task
        assert impostor.status_code == 200, impostor.text
        assert owner.status_code == 200, owner.text
        return impostor.json(), owner.json()

    impostor, owner = asyncio.run(scenario())
    assert impostor["calls"] == []
    assert len(owner["calls"]) == 1


def test_register_endpoint_requires_a_live_owned_session():
    """register 不得凭可猜 taskId 登记结果：必须有活跃且归属一致的会话。"""

    async def scenario():
        async with _localhost_client() as client:
            # (a) 完全没有会话
            no_session = await client.post(
                "/api/sdk/tool-bridge/register",
                json={**_body("turn-none"), "result": _result().model_dump(mode="json")},
            )
            # (b) 有会话但归属不符（同名不同工作簿）
            task_tool_bridge.open(_binding())
            mismatched = await client.post(
                "/api/sdk/tool-bridge/register",
                json={
                    **_body(workbookIdentity="url:https://evil.example/报表.xlsx"),
                    "result": _result().model_dump(mode="json"),
                },
            )
            # (c) 归属正确
            ok = await client.post(
                "/api/sdk/tool-bridge/register",
                json={**_body(), "result": _result().model_dump(mode="json")},
            )
            return no_session.status_code, mismatched.status_code, ok.status_code

    no_session, mismatched, ok = asyncio.run(scenario())
    assert no_session != 200
    assert mismatched != 200
    assert ok == 200


# ---------------------------------------------------------------------------
# 授权生命周期：重授权 / 换簿 / 换会话必须作废已保留的结果
# ---------------------------------------------------------------------------


def test_reauthorization_invalidates_retained_results():
    bridge = TaskToolBridge()
    bridge.open(_binding())
    handle = register_bridge_result(_TASK, _result(), "fp-source-1")

    bridge.open(_binding(authorized_key="b" * 64))

    assert local_results.resolve(_TASK, handle) is None


def test_workbook_change_invalidates_retained_results():
    bridge = TaskToolBridge()
    bridge.open(_binding())
    handle = register_bridge_result(_TASK, _result(), "fp-source-1")

    bridge.open(
        _binding(
            workbook_name="别的.xlsx",
            workbook_identity="url:https://contoso.example/别的.xlsx",
        )
    )

    assert local_results.resolve(_TASK, handle) is None


def test_result_ref_is_rejected_from_another_session():
    bridge = TaskToolBridge()
    bridge.open(_binding())
    handle = register_bridge_result(_TASK, _result(), "fp-source-1")

    other = TaskToolBridge()
    other.open(_binding(task_id="turn-other"))

    assert local_results.resolve("turn-other", handle) is None
    assert local_results.resolve(_TASK, handle) is not None


def _pending_entry():
    from types import SimpleNamespace

    return SimpleNamespace(
        operation_id="op_" + "1" * 32,
        tool="query_table",
        arguments={"mode": "rows"},
        fingerprint="0" * 64,
        future=asyncio.new_event_loop().create_future(),
        created_at=0.0,
    )
