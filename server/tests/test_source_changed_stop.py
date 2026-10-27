"""来源变化 = 本次规划的终态：桥停止派发、SDK 循环结束、轮次如实失败。

为什么单独一份：2026-10-25 的真实 SDK 实验里，模型在 ``source_changed`` 之后
**反复重试**（5 次 ``prepare_python_data`` + 1 次 ``query_table`` 全被拒），最终
以 ``MODEL_TURN_LIMIT`` 收场 —— 用户看到的是"调用轮数上限"，与"来源变了"毫无
关系，而每一次都真的派给了宿主、真的读了一次表。提示词里的"不要重试"解决不了：
模型看不到"再读也不会变"。

本文件钉的是**生产路径**上的六个事实：

1. 守卫确认 ``source_changed`` 之后，桥不再把调用派给宿主（宿主读表次数不增加）；
2. 再次调用得到的是同一个固定码，不是 ``cancelled``（原因不能被说成"用户取消"）；
3. SDK 的流式循环在拿到下一条消息时结束，轮次以 ``SOURCE_CHANGED`` 终止，
   不是 ``MODEL_TURN_LIMIT``、不是普通暂停；
4. 已登记结果当场作废，旧引用提交不了计划；
5. ``timeout`` 这类暂时错误**不**触发停止（重试仍然正当）；
6. 迟到回执被拒；独立暂停的另一个任务保留继续入口；重新发起后可正常完成。

除模型传输（``query``）外，桥、登记器、编译器、账本与 SSE 端点都是生产对象。
"""
from __future__ import annotations

import asyncio
import json
from types import SimpleNamespace

import pytest
from fastapi.testclient import TestClient

from server.app import claude_sdk_agent
from server.app.models import DataToolResult, PlanRequest
from server.app.planning_source_guard import (
    SOURCE_CHANGED_RECEIPT_CODE,
    SOURCE_CHANGED_STOP_MESSAGE,
    SOURCE_CHANGED_TURN_CODE,
    SourceChangedDuringPlanning,
)
from server.app.task_bridge import BridgeBinding, BridgeError, TaskToolBridge

TASK_ID = "turn-source-stop"
WORKBOOK_IDENTITY = "url:https://example.com/测试.xlsx"


# ---------------------------------------------------------------------------
# fixtures / helpers
# ---------------------------------------------------------------------------


def _plan_request(**overrides) -> PlanRequest:
    payload = {
        "turnId": TASK_ID,
        "prompt": "按类别汇总数量",
        "planningReadPolicy": "none",
        "hostId": "excel-office",
        "workbook": {
            "name": "测试.xlsx",
            "capturedAt": "2026-10-09T00:00:00Z",
            "activeWorksheet": "Sheet1",
            "sourceFingerprint": "fp-source-1",
            "workbookIdentity": WORKBOOK_IDENTITY,
            "worksheets": [
                {
                    "name": "Sheet1", "usedRange": "A1:B3", "rowCount": 3,
                    "columnCount": 2, "headers": ["类别", "数量"],
                },
            ],
        },
    }
    payload.update(overrides)
    return PlanRequest.model_validate(payload)


def _binding(task_id: str = TASK_ID, **overrides) -> BridgeBinding:
    defaults = dict(
        task_id=task_id,
        host_id="excel-office",
        workbook_name="测试.xlsx",
        source_fingerprint="fp-source-1",
        authorized_key="a" * 64,
        workbook_identity=WORKBOOK_IDENTITY,
    )
    defaults.update(overrides)
    return BridgeBinding(**defaults)


def _result() -> DataToolResult:
    return DataToolResult.model_validate({
        "requestId": "req-1",
        "tool": "query_table",
        "title": "汇总结果",
        "headers": ["类别", "数量"],
        "rows": [["甲", 2], ["乙", 1]],
        "sourceSheets": ["Sheet1"],
        "scannedRows": 3,
        "complete": True,
        "calculation": "按类别分组计数",
        "warnings": [],
    })


@pytest.fixture(autouse=True)
def _clean_state():
    from server.app import dev_fault_target, dev_source_changed_fault
    from server.app.local_results import local_results
    from server.app.python_coordinator import coordinator
    from server.app.python_datasets import registry
    from server.app.task_bridge import task_tool_bridge

    dev_source_changed_fault.reset()
    dev_fault_target.reset()
    task_tool_bridge.reset()
    coordinator.reset()
    registry._grants.clear()
    registry._budgets.clear()
    local_results.reset()
    yield
    dev_source_changed_fault.reset()
    task_tool_bridge.reset()
    coordinator.reset()
    registry._grants.clear()
    registry._budgets.clear()
    local_results.reset()


def _connection():
    from server.app.llm.config import ModelConnection

    return ModelConnection(
        base_url="https://api.anthropic.com",
        model="test-claude",
        api_key="test-key",
        supports_vision=False,
        adapter="claude_agent_sdk",
    )


def _deliver(
    session,
    *,
    fingerprint: str,
    error_code: str | None = None,
    workbook_name: str = "测试.xlsx",
    workbook_identity: str = WORKBOOK_IDENTITY,
):
    call = session.pending_calls[0]
    return session.deliver(
        operation_id=call["operationId"],
        request_fingerprint=call["requestFingerprint"],
        host_id="excel-office",
        workbook_name=workbook_name,
        workbook_identity=workbook_identity,
        source_fingerprint=fingerprint,
        result=_result(),
        error_code=error_code,
    )


def _one_call(bridge, binding) -> tuple[dict, list[str]]:
    """开一条桥、发一次受控读表；有派发就按 bridge 回收执。

    返回 ``(回执, 派给宿主的操作 ID 列表)``。列表为空 = **宿主一次都没被要求读表**，
    这是"停止"与"只是回了个错误码"的区别。
    """
    session = bridge.open(binding)
    dispatched: list[str] = []

    async def on_request(operation_id: str, tool: str, args: dict, fingerprint: str) -> None:
        dispatched.append(operation_id)

    session.on_request = on_request

    async def scenario():
        task = asyncio.create_task(session.request("query_table", {"mode": "rows"}))
        for _ in range(64):
            if task.done():
                break
            if session.pending_calls:
                _deliver(
                    session,
                    fingerprint="fp-source-1",
                    workbook_name=binding.workbook_name,
                    workbook_identity=binding.identity,
                )
                break
            await asyncio.sleep(0)
        return await asyncio.wait_for(task, timeout=5)

    return asyncio.run(scenario()), dispatched


class _FakeResult:
    def __init__(self, result: str = "") -> None:
        self.result = result
        self.is_error = False
        self.subtype = "success"
        self.api_error_status = None


class _RetrySdk:
    """模拟"被拒之后继续重试"的模型，并按真实 SDK 的样子**逐条**流式输出。

    真实 SDK 把每次工具回执作为一条消息推给 ``_result``；这里每次工具调用后
    ``yield`` 一条，正是为了验证"停止发生在拿到下一条消息时"。若只在生成器末尾
    yield，停止会看起来生效，而真实流式链路上却晚了一步。
    """

    ResultMessage = _FakeResult
    CLINotFoundError = type("CLINotFoundError", (Exception,), {})
    CLIConnectionError = type("CLIConnectionError", (Exception,), {})
    ProcessError = type("ProcessError", (Exception,), {})
    ResultError = type("ResultError", (Exception,), {})
    CLIJSONDecodeError = type("CLIJSONDecodeError", (Exception,), {})

    def __init__(
        self,
        *,
        attempts: int = 4,
        first_error: str | None = SOURCE_CHANGED_RECEIPT_CODE,
        later_error: str | None = SOURCE_CHANGED_RECEIPT_CODE,
        fingerprint: str = "fp-changed",
        task_id: str = TASK_ID,
    ) -> None:
        self.attempts = attempts
        self.first_error = first_error
        self.later_error = later_error
        self.fingerprint = fingerprint
        self.task_id = task_id
        self.schemas: dict[str, dict] = {}
        self.tool_results: list[dict] = []
        self.steps: list[str] = []
        self.host_polls = 0

    # -- SDK surface -------------------------------------------------------
    def ClaudeAgentOptions(self, **kwargs):
        return SimpleNamespace(**kwargs)

    def tool(self, name, _description, _schema):
        self.schemas[name] = _schema
        return lambda handler: SimpleNamespace(name=name, handler=handler)

    def create_sdk_mcp_server(self, *, name, version, tools):
        assert name == "excel_bro"
        return {tool.name: tool for tool in tools}

    # -- host side ---------------------------------------------------------
    async def _await_pending(self, task) -> bool:
        """等宿主侧被要求读表；handler 提前结束（桥已停）就立刻返回 False。"""
        from server.app.task_bridge import task_tool_bridge

        session = task_tool_bridge.session(self.task_id)
        for _ in range(2000):
            if session is not None and session.pending_calls:
                return True
            if task.done():
                return False
            await asyncio.sleep(0)
        return False

    async def _serve(self, task, error_code: str | None):
        if not await self._await_pending(task):
            # 桥停了：宿主**没有被要求**读表。这是最关键的观测点。
            self.steps.append("host_idle")
            return
        from server.app.task_bridge import task_tool_bridge

        self.host_polls += 1
        _deliver(
            task_tool_bridge.session(self.task_id),
            fingerprint=self.fingerprint,
            error_code=error_code,
        )

    async def query(self, *, prompt, options):
        tool = options.mcp_servers["excel_bro"]["query_table"]
        for index in range(self.attempts):
            error_code = self.first_error if index == 0 else self.later_error
            task = asyncio.create_task(tool.handler({
                "mode": "aggregate", "groupBy": ["类别"],
                "metrics": [{"operation": "countRows", "outputName": "数量"}],
            }))
            await self._serve(task, error_code)
            outcome = await task
            self.tool_results.append(outcome)
            yield _FakeResult("")
        yield _FakeResult("")


def _install(monkeypatch, tmp_path, fake):
    monkeypatch.setattr(claude_sdk_agent, "_sdk", lambda: fake)
    monkeypatch.setattr(claude_sdk_agent, "_state_dir", lambda: tmp_path)


def _receipt_text(outcome: dict) -> dict:
    return json.loads(outcome["content"][0]["text"])


# ---------------------------------------------------------------------------
# 1. 桥：停止派发、回同一个固定码、作废旧结果、拒收迟到回执
# ---------------------------------------------------------------------------


def test_bridge_stops_dispatching_after_source_changed():
    bridge = TaskToolBridge()
    session = bridge.open(_binding())
    dispatched: list[str] = []

    async def on_request(operation_id: str, tool: str, args: dict, fingerprint: str) -> None:
        dispatched.append(operation_id)

    session.on_request = on_request

    async def scenario():
        task = asyncio.create_task(session.request("query_table", {"mode": "rows"}))
        while not session.pending_calls:
            await asyncio.sleep(0)
        _deliver(session, fingerprint="fp-changed")
        return await task

    first = asyncio.run(scenario())
    assert first == {"status": "error", "errorCode": SOURCE_CHANGED_RECEIPT_CODE}
    assert session.source_stopped is True
    assert len(dispatched) == 1

    # 模型再试一次：回同一个固定码，**且不再派给宿主**。
    second = asyncio.run(session.request("query_table", {"mode": "rows"}))
    assert second == {"status": "error", "errorCode": SOURCE_CHANGED_RECEIPT_CODE}
    assert len(dispatched) == 1, "停止之后不许再派发一次宿主读取"
    assert session.pending_calls == []
    # 原因不能被说成"用户取消"。
    assert second["errorCode"] != "cancelled"


def test_late_host_receipt_after_source_stop_is_refused():
    bridge = TaskToolBridge()
    session = bridge.open(_binding())
    seen: dict[str, dict] = {}

    async def scenario():
        task = asyncio.create_task(session.request("query_table", {"mode": "rows"}))
        while not session.pending_calls:
            await asyncio.sleep(0)
        seen["call"] = dict(session.pending_calls[0])
        _deliver(session, fingerprint="fp-changed")
        await task
        return await session.request("query_table", {"mode": "rows"})

    assert asyncio.run(scenario()) == {
        "status": "error", "errorCode": SOURCE_CHANGED_RECEIPT_CODE,
    }
    # 同一条调用的回执晚到（宿主慢半拍）：会话已停，回执被拒。
    with pytest.raises(BridgeError) as caught:
        session.deliver(
            operation_id=seen["call"]["operationId"],
            request_fingerprint=seen["call"]["requestFingerprint"],
            host_id="excel-office",
            workbook_name="测试.xlsx",
            workbook_identity=WORKBOOK_IDENTITY,
            source_fingerprint="fp-source-1",
            result=_result(),
            error_code=None,
        )
    assert caught.value.code == "cancelled"


def test_registered_result_is_released_after_source_stop():
    """旧结果必须当场作废：否则模型还能把旧那一版提交成白名单计划。"""
    from server.app.local_results import local_results

    bridge = TaskToolBridge()
    session = bridge.open(_binding())
    handle = local_results.register(TASK_ID, _result(), "fp-source-1")
    assert local_results.resolve(TASK_ID, handle.handle) is not None

    bridge.stop_for_source_change(TASK_ID)

    assert session.source_stopped is True
    assert local_results.claim(TASK_ID) is None
    # 旧引用当场取不到值：模型再拿它提交计划，编译器会按"引用无效"拒绝。
    assert local_results.resolve(TASK_ID, handle.handle) is None


# ---------------------------------------------------------------------------
# 2. SDK 循环：停下来，且不是 MODEL_TURN_LIMIT
# ---------------------------------------------------------------------------


def test_sdk_planning_stops_on_first_source_changed(monkeypatch, tmp_path):
    fake = _RetrySdk(attempts=4)
    _install(monkeypatch, tmp_path, fake)

    with pytest.raises(SourceChangedDuringPlanning) as caught:
        asyncio.run(claude_sdk_agent.run_sdk_plan(_plan_request(), _connection()))

    assert caught.value.code == SOURCE_CHANGED_TURN_CODE
    assert SOURCE_CHANGED_STOP_MESSAGE in str(caught.value)
    # 只读了一次：第一次调用到宿主，之后桥停了，宿主再没被要求读表。
    assert fake.host_polls == 1
    assert _receipt_text(fake.tool_results[0])["errorCode"] == SOURCE_CHANGED_RECEIPT_CODE
    # 停止发生在拿到"下一条消息"时：模型后面的重试根本没有被消费。
    assert len(fake.tool_results) == 1


def test_in_flight_tool_calls_after_the_stop_are_refused_locally(monkeypatch, tmp_path):
    """模型同一批里已经发出的调用也不会到宿主，更不会算出旧结果。

    真实 SDK 里工具调用与消息是交错到达的；停止信号发出的那一刻，可能还有一条
    调用在本机排队。这一条专门覆盖"来不及拦"的那个窗口。
    """

    class _BurstSdk(_RetrySdk):
        async def query(self, *, prompt, options):
            tool = options.mcp_servers["excel_bro"]["query_table"]
            args = {
                "mode": "aggregate", "groupBy": ["类别"],
                "metrics": [{"operation": "countRows", "outputName": "数量"}],
            }
            first = asyncio.create_task(tool.handler(args))
            await self._serve(first, self.first_error)
            self.tool_results.append(await first)
            # 停止已经发生，但模型在同一批里又发了两条调用。
            for _ in range(2):
                self.tool_results.append(await tool.handler(args))
            yield _FakeResult("")

    fake = _BurstSdk(attempts=1)
    _install(monkeypatch, tmp_path, fake)

    with pytest.raises(SourceChangedDuringPlanning):
        asyncio.run(claude_sdk_agent.run_sdk_plan(_plan_request(), _connection()))

    assert fake.host_polls == 1, "停止之后不许再派发宿主读取"
    assert len(fake.tool_results) == 3
    assert _receipt_text(fake.tool_results[0])["errorCode"] == SOURCE_CHANGED_RECEIPT_CODE
    for outcome in fake.tool_results[1:]:
        assert outcome.get("is_error") is True
        assert SOURCE_CHANGED_STOP_MESSAGE in outcome["content"][0]["text"]


def test_sdk_turn_reports_stopped_step_and_never_claims_success(monkeypatch, tmp_path):
    fake = _RetrySdk(attempts=3)
    _install(monkeypatch, tmp_path, fake)
    events: list[dict] = []

    async def on_event(event):
        events.append(dict(event))

    with pytest.raises(SourceChangedDuringPlanning):
        asyncio.run(claude_sdk_agent.run_sdk_plan(
            _plan_request(), _connection(), on_event=on_event
        ))

    by_step = {(event.get("stepId"), event.get("status")) for event in events}
    # 明确的一步：不是"正常完成"，也不是"用户暂停"。
    assert ("planning.source_changed", "failed") in by_step
    assert ("planning.model", "failed") in by_step
    assert ("planning.model", "started") in by_step
    assert ("planning.model", "completed") not in by_step
    assert ("planning.submit", "completed") not in by_step
    # 被拒的那一步如实记成"未完成"（warning），而不是 completed。
    warnings = [
        event for event in events
        if str(event.get("stepId", "")).startswith("bridge.query_table:")
    ]
    assert any(event.get("status") == "warning" for event in warnings)
    assert all(event.get("status") != "completed" for event in warnings)


def test_retryable_timeout_does_not_stop_planning(monkeypatch, tmp_path):
    """正向锚点：暂时错误不是终态。把 timeout 也做成停止会废掉"重算一次就好"。"""
    fake = _RetrySdk(
        attempts=2, first_error="timeout", later_error=None, fingerprint="fp-source-1"
    )
    _install(monkeypatch, tmp_path, fake)
    events: list[dict] = []

    async def on_event(event):
        events.append(dict(event))

    with pytest.raises(Exception) as caught:
        asyncio.run(claude_sdk_agent.run_sdk_plan(
            _plan_request(), _connection(), on_event=on_event
        ))

    assert not isinstance(caught.value, SourceChangedDuringPlanning)
    # 两次都真的到了宿主：现场没变时重读是正当的。
    assert fake.host_polls == 2
    assert not any(event.get("stepId") == "planning.source_changed" for event in events)


def test_successful_reread_does_not_stop_planning(monkeypatch, tmp_path):
    """同一批来源读到的是同一版本值时，重读不是缺陷，不该被打断。"""
    fake = _RetrySdk(
        attempts=3, first_error=None, later_error=None, fingerprint="fp-source-1"
    )
    _install(monkeypatch, tmp_path, fake)
    events: list[dict] = []

    async def on_event(event):
        events.append(dict(event))

    with pytest.raises(Exception) as caught:
        asyncio.run(claude_sdk_agent.run_sdk_plan(
            _plan_request(), _connection(), on_event=on_event
        ))

    assert not isinstance(caught.value, SourceChangedDuringPlanning)
    assert fake.host_polls == 3
    assert not any(event.get("stepId") == "planning.source_changed" for event in events)


# ---------------------------------------------------------------------------
# 3. 轮次终态：SSE 错误帧、账本 failed、无旧继续入口、不影响别的暂停任务
# ---------------------------------------------------------------------------


_STREAM_PAYLOAD = {
    "turnId": TASK_ID,
    "prompt": "按类别汇总数量并写入目标表",
    "workbook": {
        "name": "通用数据.xlsx",
        "capturedAt": "2026-10-09T00:00:00.000Z",
        "activeWorksheet": "数据",
        "worksheets": [
            {"name": "数据", "rowCount": 2, "columnCount": 2, "headers": ["分类", "数值"]},
        ],
    },
    "modelId": "",
}


def _parse_sse(text: str) -> list[tuple[str, dict]]:
    events: list[tuple[str, dict]] = []
    for block in text.strip().split("\n\n"):
        if not block.strip():
            continue
        name = ""
        data = ""
        for line in block.splitlines():
            if line.startswith("event:"):
                name = line[len("event:"):].strip()
            elif line.startswith("data:"):
                data = line[len("data:"):].strip()
        events.append((name, json.loads(data)))
    return events


def test_turn_stream_reports_source_changed_as_stopped_not_limit(monkeypatch):
    import server.app.main as main_module
    from server.app.main import app

    async def stopped(_request, **_kwargs):
        raise SourceChangedDuringPlanning()

    monkeypatch.setattr(main_module, "create_plan", stopped)
    client = TestClient(app)
    response = client.post("/api/turn/stream", json=_STREAM_PAYLOAD)
    events = _parse_sse(response.text)

    errors = [payload for name, payload in events if name == "error"]
    assert len(errors) == 1
    assert errors[0]["code"] == SOURCE_CHANGED_TURN_CODE
    assert errors[0]["status"] == 409
    assert errors[0]["retryable"] is False
    assert SOURCE_CHANGED_STOP_MESSAGE in errors[0]["message"]
    # 不是普通暂停：没有 paused 帧，也就不会长出"继续"入口。
    assert all(name != "paused" for name, _ in events)
    assert all(name != "result" for name, _ in events)

    record = main_module._folder_operation_ledger.planning_task(TASK_ID)
    assert record.status == "failed"
    assert record.error["code"] == SOURCE_CHANGED_TURN_CODE

    status = client.get(f"/api/turn/{TASK_ID}/status").json()
    assert status["status"] == "failed"
    assert status["canContinue"] is False


def test_another_paused_turn_keeps_its_continue_entry(monkeypatch):
    """撤下的是**这一条**失效任务；另一个独立暂停的任务不许被一起作废。"""
    import server.app.main as main_module
    from server.app.main import app

    other = "turn-independent-pause"
    paused_payload = {
        **_STREAM_PAYLOAD,
        "turnId": other,
        "workbook": {
            **_STREAM_PAYLOAD["workbook"],
            "workbookIdentity": "url:https://example.invalid/other.xlsx",
            "sourceFingerprint": "sha256:other-source",
            "sourceFingerprintSheets": ["数据"],
        },
    }

    async def pausing(_request, *, on_checkpoint=None, **_kwargs):
        if on_checkpoint is not None:
            await on_checkpoint({
                "stage": "after_tool",
                "confirmations": [],
                "completedTools": ["query_table"],
                "resultRefs": [],
                "failedTool": None,
            })
        raise main_module.AgentBudgetPause(
            {
                "stage": "after_tool",
                "confirmations": [],
                "completedTools": ["query_table"],
                "resultRefs": [],
                "failedTool": None,
            },
            60.0,
        )

    monkeypatch.setattr(main_module, "create_plan", pausing)
    client = TestClient(app)
    client.post("/api/turn/stream", json=paused_payload)
    before = client.get(f"/api/turn/{other}/status").json()
    assert before["status"] == "waiting_user"
    assert before["canContinue"] is True

    async def stopped(_request, **_kwargs):
        raise SourceChangedDuringPlanning()

    monkeypatch.setattr(main_module, "create_plan", stopped)
    client.post("/api/turn/stream", json=_STREAM_PAYLOAD)

    after = client.get(f"/api/turn/{other}/status").json()
    assert after["status"] == "waiting_user"
    assert after["canContinue"] is True
    assert client.get(f"/api/turn/{TASK_ID}/status").json()["canContinue"] is False


def test_new_turn_after_stop_reads_the_current_source_again(monkeypatch):
    """用户重新发起是一次**新任务**：新绑定、新来源指纹，正常完成一次。"""
    import server.app.main as main_module
    from server.app.main import app

    async def stopped(_request, **_kwargs):
        raise SourceChangedDuringPlanning()

    monkeypatch.setattr(main_module, "create_plan", stopped)
    client = TestClient(app)
    client.post("/api/turn/stream", json=_STREAM_PAYLOAD)

    from server.app.models import AnswerResponse

    async def fine(_request, **_kwargs):
        return AnswerResponse(kind="answer", message="已按最新数据重新计算", provider="model")

    monkeypatch.setattr(main_module, "create_plan", fine)
    restarted = {**_STREAM_PAYLOAD, "turnId": "turn-restarted-after-stop"}
    response = client.post("/api/turn/stream", json=restarted)
    events = _parse_sse(response.text)
    results = [payload for name, payload in events if name == "result"]
    assert len(results) == 1
    body = results[0].get("response", results[0])
    assert body["message"] == "已按最新数据重新计算"
    assert body["outcome"] == "completed"
    # 原任务的终态没有被新任务改回"可继续"。
    assert client.get(f"/api/turn/{TASK_ID}/status").json()["canContinue"] is False


def test_dev_fault_is_off_by_default(monkeypatch):
    """验收开关默认关闭：状态是 `off`，桥的行为与生产逐字节一致。

    这条必须存在：注入点是**生产** `request()`。如果开关没关严，每个真实用户都会
    在下一次工具调用上莫名收到 `source_changed`。
    """
    from server.app import dev_source_changed_fault

    monkeypatch.delenv(dev_source_changed_fault.ENV_FLAG, raising=False)
    assert dev_source_changed_fault.enabled() is False
    assert dev_source_changed_fault.status()["phase"] == "off"
    assert dev_source_changed_fault.take(TASK_ID) is False
    # 未开启时连 arm 都不该被允许。
    with pytest.raises(dev_source_changed_fault.FaultDisabled):
        dev_source_changed_fault.arm(
            host_id="excel-office", workbook_name="测试.xlsx"
        )

    bridge = TaskToolBridge()
    session = bridge.open(_binding())
    dispatched: list[str] = []

    async def on_request(operation_id: str, tool: str, args: dict, fingerprint: str) -> None:
        dispatched.append(operation_id)

    session.on_request = on_request

    async def scenario():
        task = asyncio.create_task(session.request("query_table", {"mode": "rows"}))
        while not session.pending_calls:
            await asyncio.sleep(0)
        _deliver(session, fingerprint="fp-source-1")
        return await task

    receipt = asyncio.run(scenario())
    assert receipt["status"] == "ok"
    assert len(dispatched) == 1
    assert session.source_stopped is False


def test_dev_fault_turns_the_next_call_into_a_real_stop(monkeypatch):
    """开启后：**下一次**调用直接拿到生产那份 `source_changed`，并且真的停。

    注入只是让"来源变化"在真实面板上成为一次真实发生的终止 —— 回执、停止、
    `drop_python_state` 全部走生产链路，不另写一套。单次：后面那本簿不受影响。
    """
    from server.app import dev_source_changed_fault

    monkeypatch.setenv(dev_source_changed_fault.ENV_FLAG, "1")
    assert dev_source_changed_fault.status()["phase"] == "idle"
    armed = dev_source_changed_fault.arm(
        host_id="excel-office",
        workbook_name="测试.xlsx",
        workbook_identity=WORKBOOK_IDENTITY,
    )
    assert armed["phase"] == "armed"

    bridge = TaskToolBridge()
    session = bridge.open(_binding())
    dispatched: list[str] = []

    async def on_request(operation_id: str, tool: str, args: dict, fingerprint: str) -> None:
        dispatched.append(operation_id)

    session.on_request = on_request
    receipt = asyncio.run(session.request("query_table", {"mode": "rows"}))
    assert receipt == {"status": "error", "errorCode": SOURCE_CHANGED_RECEIPT_CODE}
    assert session.source_stopped is True
    # 关键：这一次**根本没有派给宿主**（不读表）。
    assert dispatched == []
    assert session.pending_calls == []
    assert dev_source_changed_fault.status()["phase"] == "consumed"

    # 单次：换一个任务（同一本簿也换了 owner）不会再被注入。
    other = TaskToolBridge().open(_binding(task_id="turn-source-stop-2"))
    other_dispatched: list[str] = []

    async def on_other(operation_id: str, tool: str, args: dict, fingerprint: str) -> None:
        other_dispatched.append(operation_id)

    other.on_request = on_other

    async def other_scenario():
        task = asyncio.create_task(other.request("query_table", {"mode": "rows"}))
        while not other.pending_calls:
            await asyncio.sleep(0)
        _deliver(other, fingerprint="fp-source-1")
        return await task

    assert asyncio.run(other_scenario())["status"] == "ok"
    assert other.source_stopped is False
    assert len(other_dispatched) == 1


def test_dev_fault_status_is_readable_and_writes_still_need_json(monkeypatch):
    """验收入口的状态必须**真读得到**；写方法仍然要求 JSON 声明。

    本轮踩到的坑：桥的本机守卫原先对**所有**方法都要求 `content-type:
    application/json`，而 GET 根本没有体 —— 于是"查一下入口状态"永远 403。
    放松的只有这一条（无体方法），回环与来源判据一字未动；写方法照旧。
    """
    from server.app import dev_source_changed_fault

    monkeypatch.setenv(dev_source_changed_fault.ENV_FLAG, "1")

    async def scenario():
        import httpx

        from server.app.main import app

        transport = httpx.ASGITransport(app=app)
        async with httpx.AsyncClient(
            transport=transport, base_url="http://127.0.0.1:8765"
        ) as client:
            return await client.get("/api/dev/source-changed-fault")

    response = asyncio.run(scenario())
    assert response.status_code == 200
    assert response.json()["phase"] == "idle"

    from types import SimpleNamespace

    from fastapi import HTTPException

    from server.app.main import _bridge_local_guard

    def fake_request(method: str, content_type: str):
        return SimpleNamespace(
            headers={"content-type": content_type},
            client=SimpleNamespace(host="127.0.0.1"),
            url=SimpleNamespace(hostname="127.0.0.1"),
            method=method,
        )

    # 无体方法放行（这正是本轮修的那一条）。
    _bridge_local_guard(fake_request("GET", ""))
    # 写方法仍然必须声明 JSON —— 放松不许蔓延到写入面。
    with pytest.raises(HTTPException) as caught:
        _bridge_local_guard(fake_request("POST", "text/plain"))
    assert caught.value.status_code == 403


# ---------------------------------------------------------------------------
# D3–D5：验收入口的**作用域**（这一批补的失败回归）
#
# 旧实现只按"下一次工具调用"命中：`arm()` 不带任何绑定时，任何一本簿、任何一个
# 任务的第一次读表都会把它消费掉。真实验收里这意味着注入可能落到别人（并发任务、
# 另一本打开的簿）的调用上 —— 看到"没停"就以为生产停止链路坏了，或者反过来把
# 一次无关任务停在半路。下面三条钉的是"注入属于哪一次任务"。
# ---------------------------------------------------------------------------


def test_dev_fault_requires_a_host_and_workbook_binding(monkeypatch):
    """启用时必须说清是哪本簿：留空的"下一次调用不论哪本簿"不再被接受。"""
    from server.app import dev_source_changed_fault

    monkeypatch.setenv(dev_source_changed_fault.ENV_FLAG, "1")
    with pytest.raises(dev_source_changed_fault.FaultDisabled):
        dev_source_changed_fault.arm()
    with pytest.raises(dev_source_changed_fault.FaultDisabled):
        dev_source_changed_fault.arm(host_id="excel-office")
    with pytest.raises(dev_source_changed_fault.FaultDisabled):
        dev_source_changed_fault.arm(workbook_name="测试.xlsx")
    assert dev_source_changed_fault.status()["phase"] == "idle"


def test_dev_fault_cannot_be_stolen_by_another_workbook(monkeypatch):
    """另一本簿的任务先开桥、先读表 —— 不许把这次启用抢走。

    抢走的后果不是"多注入一次"这么轻：被抢的那次调用会**真的停在半路**，
    而用户本来要验的那本簿一次都看不到停止。
    """
    from server.app import dev_source_changed_fault

    monkeypatch.setenv(dev_source_changed_fault.ENV_FLAG, "1")
    dev_source_changed_fault.arm(
        host_id="excel-office",
        workbook_name="测试.xlsx",
        workbook_identity=WORKBOOK_IDENTITY,
    )

    other_binding = _binding(
        task_id="turn-other-workbook",
        workbook_name="别的.xlsx",
        workbook_identity="url:https://example.com/别的.xlsx",
    )
    other_receipt, other_dispatched = _one_call(TaskToolBridge(), other_binding)
    # 另一本簿拿到的是**正常**读表回执，且真的派给了宿主。
    assert other_receipt["status"] == "ok"
    assert len(other_dispatched) == 1
    # 启用没有被消费掉：它还等着目标那本簿。
    assert dev_source_changed_fault.status()["phase"] == "armed"

    target_receipt, target_dispatched = _one_call(TaskToolBridge(), _binding())
    assert target_receipt == {"status": "error", "errorCode": SOURCE_CHANGED_RECEIPT_CODE}
    assert target_dispatched == []
    assert dev_source_changed_fault.status()["phase"] == "consumed"


def test_dev_fault_does_not_match_a_same_name_different_workbook(monkeypatch):
    """同名但不同工作簿：只比名字会张冠李戴，必须连身份一起比。"""
    from server.app import dev_source_changed_fault

    monkeypatch.setenv(dev_source_changed_fault.ENV_FLAG, "1")
    dev_source_changed_fault.arm(
        host_id="excel-office",
        workbook_name="测试.xlsx",
        workbook_identity=WORKBOOK_IDENTITY,
    )

    same_name = _binding(
        task_id="turn-same-name",
        workbook_identity="url:https://example.com/另一个/测试.xlsx",
    )
    receipt, dispatched = _one_call(TaskToolBridge(), same_name)
    assert receipt["status"] == "ok"
    assert len(dispatched) == 1
    assert dev_source_changed_fault.status()["phase"] == "armed"


def test_dev_fault_binds_the_task_when_its_bridge_opens(monkeypatch):
    """任务归属在**开桥那一刻**由服务端绑定，此后只有那一次任务能命中。

    绑定时机是这条判据的关键：放到"下一次调用"就等于把归属交给谁先说话；
    放到开桥（规划请求已经确定 taskId、宿主、工作簿）才是服务端在可信路径上定。
    """
    from server.app import dev_source_changed_fault

    monkeypatch.setenv(dev_source_changed_fault.ENV_FLAG, "1")
    dev_source_changed_fault.arm(
        host_id="excel-office",
        workbook_name="测试.xlsx",
        workbook_identity=WORKBOOK_IDENTITY,
    )

    bridge = TaskToolBridge()
    bridge.open(_binding())
    status = dev_source_changed_fault.status()
    assert status["phase"] == "bound"
    assert status["taskBound"] is True

    # 第二本同样匹配绑定的簿（同一个工作簿、另一个任务）不能再抢这次启用。
    second_receipt, second_dispatched = _one_call(
        TaskToolBridge(), _binding(task_id="turn-second-same-book")
    )
    assert second_receipt["status"] == "ok"
    assert len(second_dispatched) == 1
    assert dev_source_changed_fault.status()["phase"] == "bound"

    # 真正被绑定的那一次任务：一次真实发生的停止，且不派给宿主。
    receipt, dispatched = _one_call(TaskToolBridge(), _binding())
    assert receipt == {"status": "error", "errorCode": SOURCE_CHANGED_RECEIPT_CODE}
    assert dispatched == []
    consumed = dev_source_changed_fault.status()
    assert consumed["phase"] == "consumed"
    assert consumed["taskBound"] is False


# ---------------------------------------------------------------------------
# D6–D10：启用的**工作簿身份**（第三批）
#
# 上一批只把"哪一次任务"钉住了，启用的目标仍然可以是"任意一本叫这个名字的簿"：
# `arm` 允许不给身份，`bind_task` 于是只比 host + name。同名不同簿（两个文件夹里
# 各一份 `测试.xlsx`）时，先开桥的那本把单发抢走 —— 被抢的那一轮真的停在半路，
# 而要验的那本一次都看不到停止。探针实证：
# `outputs/verify-a1e3-round3/fault-identity-before.txt`。
#
# 判据：启用与绑定都必须落在**可核验的工作簿身份**上；无法唯一确定就拒绝启用，
# 并把原因说清楚，不靠"谁先说话"猜。
# ---------------------------------------------------------------------------

IDENTITY_A = "url:https://example.com/甲/测试.xlsx"
IDENTITY_B = "url:https://example.com/乙/测试.xlsx"


def test_dev_fault_refuses_to_arm_without_a_workbook_identity(monkeypatch):
    """只给宿主 + 名字不算"唯一确定目标"：同名两本簿都会匹配。"""
    from server.app import dev_source_changed_fault

    monkeypatch.setenv(dev_source_changed_fault.ENV_FLAG, "1")
    with pytest.raises(dev_source_changed_fault.FaultDisabled) as caught:
        dev_source_changed_fault.arm(host_id="excel-office", workbook_name="测试.xlsx")
    # 拒绝理由必须指出缺的是"身份"，而不是含糊的"参数不对"。
    assert "身份" in str(caught.value)
    assert dev_source_changed_fault.status()["phase"] == "idle"


def test_dev_fault_uses_the_only_identity_the_server_has_seen(monkeypatch):
    """服务端只见过这本簿一个身份时，可以据此补全 —— 这是证据，不是猜测。"""
    from server.app import dev_source_changed_fault

    monkeypatch.setenv(dev_source_changed_fault.ENV_FLAG, "1")
    # 先让服务端在真实开桥路径上见过甲的这一次身份。
    TaskToolBridge().open(_binding(task_id="turn-warmup-a", workbook_identity=IDENTITY_A))

    armed = dev_source_changed_fault.arm(
        host_id="excel-office", workbook_name="测试.xlsx"
    )
    assert armed["phase"] == "armed"
    assert armed["identityBound"] is True
    assert armed["identitySource"] == "observed"

    # 之后出现的同名簿（乙）不能被绑上：启用认的是甲那个身份。
    other_receipt, other_dispatched = _one_call(
        TaskToolBridge(),
        _binding(task_id="turn-b", workbook_identity=IDENTITY_B),
    )
    assert other_receipt["status"] == "ok"
    assert len(other_dispatched) == 1
    assert dev_source_changed_fault.status()["phase"] == "armed"

    # 甲仍然命中，并且真的停在半路（没派给宿主）。
    target_receipt, target_dispatched = _one_call(
        TaskToolBridge(),
        _binding(task_id="turn-a", workbook_identity=IDENTITY_A),
    )
    assert target_receipt == {"status": "error", "errorCode": SOURCE_CHANGED_RECEIPT_CODE}
    assert target_dispatched == []


def test_dev_fault_refuses_when_the_name_carries_two_identities(monkeypatch):
    """同一个名字下见过两个身份：拒绝启用，让人显式指定，而不是先到先得。"""
    from server.app import dev_source_changed_fault

    monkeypatch.setenv(dev_source_changed_fault.ENV_FLAG, "1")
    TaskToolBridge().open(_binding(task_id="turn-seen-a", workbook_identity=IDENTITY_A))
    TaskToolBridge().open(_binding(task_id="turn-seen-b", workbook_identity=IDENTITY_B))

    with pytest.raises(dev_source_changed_fault.FaultDisabled) as caught:
        dev_source_changed_fault.arm(host_id="excel-office", workbook_name="测试.xlsx")
    assert "身份" in str(caught.value)
    assert dev_source_changed_fault.status()["phase"] == "idle"

    # 显式指定后就能启用（两条路互不替代）。
    armed = dev_source_changed_fault.arm(
        host_id="excel-office",
        workbook_name="测试.xlsx",
        workbook_identity=IDENTITY_B,
    )
    assert armed["identitySource"] == "declared"
    assert armed["identityBound"] is True


def test_dev_fault_rejects_an_unverifiable_identity(monkeypatch):
    """`name:...` 之类不是可核验身份：它本身就是"只认名字"，不能拿来做绑定依据。"""
    from server.app import dev_source_changed_fault

    monkeypatch.setenv(dev_source_changed_fault.ENV_FLAG, "1")
    with pytest.raises(dev_source_changed_fault.FaultDisabled) as caught:
        dev_source_changed_fault.arm(
            host_id="excel-office",
            workbook_name="测试.xlsx",
            workbook_identity="name:测试.xlsx",
        )
    assert "身份" in str(caught.value)


def test_dev_fault_does_not_bind_a_task_that_only_carries_a_name(monkeypatch):
    """任务侧只报了名字（旧客户端）时**不许**绑定：无法证明是同一本簿。

    这条是"有绑定"≠"绑定可信"的分界：真的绑上就等于按名字张冠李戴。
    """
    from server.app import dev_source_changed_fault

    monkeypatch.setenv(dev_source_changed_fault.ENV_FLAG, "1")
    dev_source_changed_fault.arm(
        host_id="excel-office",
        workbook_name="测试.xlsx",
        workbook_identity=IDENTITY_A,
    )
    TaskToolBridge().open(
        _binding(task_id="turn-name-only", workbook_identity="")
    )
    assert dev_source_changed_fault.status()["phase"] == "armed"
    assert dev_source_changed_fault.status()["taskBound"] is False

    # 报名字的那一次调用照常派给宿主（没有被当成目标）。
    receipt, dispatched = _one_call(
        TaskToolBridge(), _binding(task_id="turn-name-only-2", workbook_identity="")
    )
    assert receipt["status"] == "ok"
    assert len(dispatched) == 1


def test_dev_fault_from_a_reopened_unsaved_workbook_session_never_binds(monkeypatch):
    """未保存簿用的是会话内身份：窗格重开就是新会话，旧启用不得绑到新会话上。"""
    from server.app import dev_source_changed_fault

    monkeypatch.setenv(dev_source_changed_fault.ENV_FLAG, "1")
    dev_source_changed_fault.arm(
        host_id="excel-office",
        workbook_name="工作簿1",
        workbook_identity="unsaved:1-aaaa1111",
    )
    # 重开后的窗格：同一个显示名，另一个会话令牌。
    receipt, dispatched = _one_call(
        TaskToolBridge(),
        _binding(
            task_id="turn-reopened",
            workbook_name="工作簿1",
            workbook_identity="unsaved:2-bbbb2222",
        ),
    )
    assert receipt["status"] == "ok"
    assert len(dispatched) == 1
    assert dev_source_changed_fault.status()["phase"] == "armed"


def test_dev_fault_cancel_and_expiry_never_leave_a_half_armed_entry(monkeypatch):
    """取消 / 过期之后：状态可见、调用不受影响、也不会落到下一次任务上。"""
    from server.app import dev_source_changed_fault

    monkeypatch.setenv(dev_source_changed_fault.ENV_FLAG, "1")
    dev_source_changed_fault.arm(
        host_id="excel-office",
        workbook_name="测试.xlsx",
        workbook_identity=IDENTITY_A,
    )
    cancelled = dev_source_changed_fault.cancel()
    assert cancelled["phase"] == "idle"
    assert cancelled["identityBound"] is False
    receipt, dispatched = _one_call(TaskToolBridge(), _binding(task_id="turn-after-cancel"))
    assert receipt["status"] == "ok"
    assert len(dispatched) == 1

    # 过期：把单调时钟推到 TTL 之后。
    import time as time_module

    dev_source_changed_fault.arm(
        host_id="excel-office",
        workbook_name="测试.xlsx",
        workbook_identity=IDENTITY_A,
        ttl_seconds=1,
    )
    real = time_module.monotonic
    monkeypatch.setattr(
        dev_source_changed_fault.time, "monotonic", lambda: real() + 10_000
    )
    assert dev_source_changed_fault.status()["phase"] == "expired"
    receipt, dispatched = _one_call(TaskToolBridge(), _binding(task_id="turn-after-expiry"))
    assert receipt["status"] == "ok"
    assert len(dispatched) == 1
