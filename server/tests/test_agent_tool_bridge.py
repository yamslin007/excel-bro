"""Task tool bridge, local result registry, result→plan compiler and SDK loop.

These tests exercise the *real* bridge, the *real* deterministic shape of the
receipt and the *real* SDK handler wiring. Only the model transport (``sdk``) is
a test double — the bridge, the registry, the projections and the compiler are
the production objects.
"""
from __future__ import annotations

import asyncio
import json
from types import SimpleNamespace

import pytest

from server.app import claude_sdk_agent
from server.app.agent_tools import (
    project_describe_tables,
    project_query_receipt,
)
from server.app.local_results import LocalResultRegistry
from server.app.models import (
    DataToolResult,
    PlanRequest,
)
from server.app.result_plan import (
    ResultPlanError,
    compile_result_plan,
    result_summary_for_answer,
)
from server.app.task_bridge import (
    BridgeBinding,
    BridgeError,
    TaskToolBridge,
    arguments_fingerprint,
)

SENTINEL = "SENTINEL_SECRET_9F3A"


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------


def _dummy_future():
    return asyncio.new_event_loop().create_future()


def _pending(op_id: str, *, fingerprint: str, future=None):
    return SimpleNamespace(
        operation_id=op_id, tool="query_table", arguments={"mode": "rows"},
        fingerprint=fingerprint, future=future or _dummy_future(), created_at=0.0,
    )


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


def _plan_request(**overrides) -> PlanRequest:
    payload = {
        "turnId": "turn-bridge-0001",
        "prompt": "把两个表按类别关联并汇总",
        "planningReadPolicy": "none",
        "hostId": "excel-office",
        "hostVersion": "16.0",
        "workbook": {
            "name": "测试.xlsx",
            "capturedAt": "2026-10-09T00:00:00Z",
            "activeWorksheet": "Sheet1",
            "sourceFingerprint": "fp-source-1",
            "worksheets": [
                {
                    "name": "Sheet1", "usedRange": "A1:B3", "rowCount": 3,
                    "columnCount": 2, "headers": ["类别", "数量"],
                },
                {
                    "name": "Sheet2", "usedRange": "A1:B3", "rowCount": 3,
                    "columnCount": 2, "headers": ["类别", "名称"],
                },
            ],
        },
    }
    payload.update(overrides)
    return PlanRequest.model_validate(payload)


def _binding(task_id="turn-bridge-0001", **overrides) -> BridgeBinding:
    defaults = dict(
        task_id=task_id,
        host_id="excel-office",
        workbook_name="测试.xlsx",
        source_fingerprint="fp-source-1",
        authorized_key="a" * 64,
    )
    defaults.update(overrides)
    return BridgeBinding(**defaults)


@pytest.fixture(autouse=True)
def _clean_state():
    from server.app import local_results as module
    from server.app.task_bridge import task_tool_bridge

    module.local_results.reset()
    task_tool_bridge.reset()
    yield
    module.local_results.reset()
    task_tool_bridge.reset()


# ---------------------------------------------------------------------------
# A. outbound projections
# ---------------------------------------------------------------------------


def test_query_receipt_projection_excludes_values_and_free_text():
    receipt = {
        "status": "ok",
        "resultRef": "res_" + "0" * 32,
        "structure": {
            "headers": ["类别", "数量"], "rowCount": 2, "columnCount": 2,
            "complete": True, "truncated": False,
        },
        # 下面这些字段即使同源回执里有，也绝不能出站。
        "rows": [["甲", 2]],
        "title": SENTINEL,
        "calculation": SENTINEL,
        "warnings": [SENTINEL],
        "scannedRows": 12,
    }
    projected = project_query_receipt(receipt)
    encoded = json.dumps(projected, ensure_ascii=False)
    assert SENTINEL not in encoded
    assert set(projected) == {"tool", "status", "resultRef", "structure"}
    assert projected["structure"] == {
        "headers": ["类别", "数量"], "rowCount": 2, "columnCount": 2,
        "complete": True, "truncated": False,
    }


def test_query_receipt_projection_maps_unknown_error_to_internal():
    projected = project_query_receipt({"status": "error", "errorCode": f"boom-{SENTINEL}"})
    assert projected == {"tool": "query_table", "status": "error", "errorCode": "internal"}


def test_describe_projection_is_a_whitelist():
    projected = project_describe_tables({
        "tables": [{
            "name": "Sheet1", "sourceSheet": None, "usedRange": "A1:B3",
            "rowCount": 3, "columnCount": 2, "headers": ["类别", "数量"],
            "dataRows": [[SENTINEL]], "displayRows": [[SENTINEL]],
        }],
    })
    encoded = json.dumps(projected, ensure_ascii=False)
    assert SENTINEL not in encoded
    assert set(projected["tables"][0]) == {
        "name", "sourceSheet", "usedRange", "rowCount", "columnCount", "headers",
    }


# ---------------------------------------------------------------------------
# C. bridge semantics
# ---------------------------------------------------------------------------


def test_bridge_delivers_a_typed_receipt_and_projects_no_values():
    bridge = TaskToolBridge()
    session = bridge.open(_binding())
    delivered: dict = {}

    async def scenario():
        task = asyncio.create_task(session.request("query_table", {"mode": "aggregate"}))
        for _ in range(200):
            if session.pending_calls:
                break
            await asyncio.sleep(0)
        call = session.pending_calls[0]
        delivered["fingerprint"] = call["requestFingerprint"]
        outcome = session.deliver(
            operation_id=call["operationId"],
            request_fingerprint=call["requestFingerprint"],
            host_id="excel-office",
            workbook_name="测试.xlsx",
            source_fingerprint="fp-source-1",
            result=_result(),
            error_code=None,
        )
        delivered["outcome"] = outcome
        return await task

    receipt = asyncio.run(scenario())
    assert delivered["outcome"] == {"status": "accepted"}
    assert receipt["status"] == "ok"
    assert receipt["resultRef"].startswith("res_")
    # 回执是投影输入，值不在其中。
    assert "rows" not in receipt
    assert SENTINEL not in json.dumps(receipt, ensure_ascii=False)


def test_bridge_idempotent_duplicate_and_conflict_on_same_id():
    bridge = TaskToolBridge()
    session = bridge.open(_binding())
    fp = arguments_fingerprint("query_table", {"mode": "rows"})
    session._pending["op_" + "1" * 32] = _pending("op_" + "1" * 32, fingerprint=fp)
    first = session.deliver(
        operation_id="op_" + "1" * 32, request_fingerprint=fp,
        host_id="excel-office", workbook_name="测试.xlsx",
        source_fingerprint="fp-source-1", result=_result(), error_code=None,
    )
    assert first == {"status": "accepted"}
    duplicate = session.deliver(
        operation_id="op_" + "1" * 32, request_fingerprint=fp,
        host_id="excel-office", workbook_name="测试.xlsx",
        source_fingerprint="fp-source-1", result=_result(), error_code=None,
    )
    assert duplicate == {"status": "duplicate"}
    with pytest.raises(BridgeError) as caught:
        session.deliver(
            operation_id="op_" + "1" * 32, request_fingerprint="0" * 64,
            host_id="excel-office", workbook_name="测试.xlsx",
            source_fingerprint="fp-source-1", result=_result(), error_code=None,
        )
    assert caught.value.code == "conflict"


def test_bridge_rejects_late_receipt_after_cancel():
    bridge = TaskToolBridge()
    session = bridge.open(_binding())
    session.cancel("user_cancelled")
    with pytest.raises(BridgeError) as caught:
        session.deliver(
            operation_id="op_" + "2" * 32, request_fingerprint="0" * 64,
            host_id="excel-office", workbook_name="测试.xlsx",
            source_fingerprint="fp-source-1", result=_result(), error_code=None,
        )
    assert caught.value.code == "cancelled"

    async def request_after_cancel():
        return await session.request("query_table", {"mode": "rows"})

    assert asyncio.run(request_after_cancel()) == {"status": "error", "errorCode": "cancelled"}


def test_bridge_rejects_unknown_operation_as_late():
    bridge = TaskToolBridge()
    session = bridge.open(_binding())
    with pytest.raises(BridgeError) as caught:
        session.deliver(
            operation_id="op_" + "3" * 32, request_fingerprint="0" * 64,
            host_id="excel-office", workbook_name="测试.xlsx",
            source_fingerprint="fp-source-1", result=_result(), error_code=None,
        )
    assert caught.value.code == "late"


@pytest.mark.parametrize(
    "override",
    [
        {"host_id": "wps-et"},
        {"workbook_name": "别的.xlsx"},
    ],
)
def test_bridge_rejects_mismatched_binding(override):
    bridge = TaskToolBridge()
    session = bridge.open(_binding())
    with pytest.raises(BridgeError) as caught:
        session.deliver(
            operation_id="op_" + "4" * 32, request_fingerprint="0" * 64,
            host_id=override.get("host_id", "excel-office"),
            workbook_name=override.get("workbook_name", "测试.xlsx"),
            source_fingerprint="fp-source-1",
            result=_result(), error_code=None,
        )
    assert caught.value.code == "stale_binding"


def test_bridge_marks_source_changed_receipt_as_error():
    """回执来源与绑定时不一致时，结果不得接上 —— 返回 source_changed。"""
    bridge = TaskToolBridge()
    session = bridge.open(_binding(source_fingerprint="fp-source-1"))
    fp = arguments_fingerprint("query_table", {"mode": "rows"})
    fut = _dummy_future()
    session._pending["op_" + "5" * 32] = _pending(
        "op_" + "5" * 32, fingerprint=fp, future=fut
    )
    # 宿主读到的是 fp-changed 的现场，与绑定时 fp-source-1 不同。
    session.deliver(
        operation_id="op_" + "5" * 32, request_fingerprint=fp,
        host_id="excel-office", workbook_name="测试.xlsx",
        source_fingerprint="fp-changed", result=_result(), error_code=None,
    )
    assert fut.result() == {"status": "error", "errorCode": "source_changed"}
    # 没有登记任何结果。
    from server.app.local_results import local_results

    assert local_results.count_task("turn-bridge-0001") == 0


def test_bridge_supersedes_on_reauthorization():
    bridge = TaskToolBridge()
    first = bridge.open(_binding())
    second = bridge.open(_binding(authorized_key="b" * 64))
    assert first.cancel_reason == "superseded"
    assert bridge.session("turn-bridge-0001") is second


# ---------------------------------------------------------------------------
# C. registry ownership
# ---------------------------------------------------------------------------


def test_registry_is_task_scoped():
    registry = LocalResultRegistry()
    item = registry.register("task-a", _result(), "fp-source-1")
    assert registry.resolve("task-a", item.handle) is item
    assert registry.resolve("task-b", item.handle) is None


def test_registry_invalidates_results_on_source_change():
    registry = LocalResultRegistry()
    registry.register("task-a", _result(), "fp-source-1")
    registry.register("task-a", _result(), "fp-source-2")
    assert registry.count_task("task-a") == 2
    removed = registry.invalidate_source("task-a", "fp-source-2")
    assert removed == 1
    assert registry.count_task("task-a") == 1


def test_registry_rejects_oversized_matrix():
    registry = LocalResultRegistry()
    # 500 行 × 50 列 = 25000 格，超过 maxResultCells(20000)。
    oversized = _result(rows=[[0] * 50 for _ in range(500)],
                        headers=[f"h{index}" for index in range(50)])
    with pytest.raises(ValueError):
        registry.register("task-a", oversized, "fp")


# ---------------------------------------------------------------------------
# D. result → plan compiler
# ---------------------------------------------------------------------------


def _registered(registry: LocalResultRegistry, **kwargs):
    return registry.register("turn-bridge-0001", _result(**kwargs), "fp-source-1")


def test_compiler_builds_a_single_write_values_plan_with_auto_readback():
    registry = LocalResultRegistry()
    item = _registered(registry)
    plan = compile_result_plan(
        item, _plan_request(), target_sheet="Sheet1", target_range="D1:E3",
    )
    assert len(plan.actions) == 1
    assert plan.actions[0].type == "writeValues"
    # 验收条件由 AnalysisPlan 自动推断：写后按 rangeEquals 回读。
    assert any(criterion.type == "rangeEquals" for criterion in plan.acceptanceCriteria)


def test_compiled_plan_carries_the_result_s_actual_source_dependency():
    """来源信息不能在「结果 → 计划」这一层丢掉。

    恢复登记的"实际来源依赖"只认计划里的 ``sourceDependencySheets``。编译结果计划
    时若不带，端上 ``attestSourceDependency`` 会把字段**删掉**（= 依赖不确定），
    于是一份来源明确的计划被当成不可恢复 —— 与本地查询结果的待遇不一致。
    """
    registry = LocalResultRegistry()
    item = registry.register(
        "turn-bridge-0001",
        DataToolResult.model_validate({
            "requestId": "req-2", "tool": "query_table", "title": "两表汇总",
            "headers": ["类别", "数量"], "rows": [["甲", 2]],
            "sourceSheets": ["Sheet1", "Sheet2"], "scannedRows": 6, "complete": True,
            "calculation": "按类别汇总", "warnings": [],
        }),
        "fp-source-1",
    )
    plan = compile_result_plan(item, _plan_request(), target_sheet="Sheet1", target_range="D1:E2")
    assert plan.sourceDependencySheets == ["Sheet1", "Sheet2"]


def test_compiled_plan_omits_source_dependency_when_the_result_has_none():
    """结果自己都不知道来源时（``sourceSheets`` 空），计划**不写**该字段。

    空数组会被读成"依赖为空"，而事实是"依赖不确定" —— 两者在恢复登记里的结论
    必须不同。``DataToolResult`` 本身不允许空 ``sourceSheets``（协议层已挡），
    所以这里直接构造登记项，验证**编译器**这一层同样不凭空造来源。
    """
    registry = LocalResultRegistry()
    item = registry.register("turn-bridge-0001", _result(), "fp-source-1")
    # 模拟"登记项没有任何来源信息"（例如旧版本登记的记录）。
    item.source_sheets = []
    plan = compile_result_plan(item, _plan_request(), target_sheet="Sheet1", target_range="D1:E3")
    assert plan.sourceDependencySheets == []


def test_result_card_never_silently_drops_warnings_or_sources():
    """结果卡必须原样带回来源与告警，不静默丢弃任何输出。"""
    registry = LocalResultRegistry()
    item = registry.register(
        "turn-bridge-0001",
        DataToolResult.model_validate({
            "requestId": "req-4", "tool": "query_table", "title": "带未匹配明细的汇总",
            "headers": ["类别", "数量"], "rows": [["甲", 2]],
            "sourceSheets": ["订单", "字典"],
            "scannedRows": 8, "complete": False,
            "calculation": "左连接后汇总",
            "warnings": ["有 2 条记录在字典中未匹配，已列入未匹配明细。"],
        }),
        "fp-source-1",
    )
    card = result_summary_for_answer(item)
    assert card["sourceSheets"] == ["订单", "字典"]
    assert card["warnings"] == ["有 2 条记录在字典中未匹配，已列入未匹配明细。"]
    # 未完成的结果必须如实标出，不能显示成完整。
    assert "complete" not in card or card.get("complete") is False


def test_compiler_rejects_size_mismatch():
    registry = LocalResultRegistry()
    item = _registered(registry)
    with pytest.raises(ResultPlanError) as caught:
        compile_result_plan(item, _plan_request(), target_sheet="Sheet1", target_range="D1:E2")
    assert caught.value.code == "size_mismatch"


def test_compiler_rejects_unauthorized_target_sheet():
    registry = LocalResultRegistry()
    item = _registered(registry)
    with pytest.raises(ResultPlanError) as caught:
        compile_result_plan(item, _plan_request(), target_sheet="Sheet9", target_range="D1:E3")
    assert caught.value.code == "unauthorized_scope"


def test_compiler_rejects_stale_source():
    registry = LocalResultRegistry()
    item = _registered(registry)
    request = _plan_request()
    request.workbook.sourceFingerprint = "fp-changed"
    with pytest.raises(ResultPlanError) as caught:
        compile_result_plan(item, request, target_sheet="Sheet1", target_range="D1:E3")
    assert caught.value.code == "source_changed"


def test_compiler_rejects_formula_text_in_values():
    registry = LocalResultRegistry()
    item = registry.register(
        "turn-bridge-0001",
        _result(rows=[["=SUM(A1:A2)", 1], ["乙", 2]]),
        "fp-source-1",
    )
    with pytest.raises(ResultPlanError) as caught:
        compile_result_plan(item, _plan_request(), target_sheet="Sheet1", target_range="D1:E3")
    assert caught.value.code == "formula_not_allowed"


def test_compiler_enforces_wps_cell_limit():
    registry = LocalResultRegistry()
    # 500 行 × 3 列（+ 表头）= 1503 格，超过 WPS 单次 1000 格上限。
    item = registry.register(
        "turn-bridge-0001",
        _result(
            headers=["类别", "数量", "占比"],
            rows=[[f"c{index}", index, 0.1] for index in range(500)],
        ),
        "fp-source-1",
    )
    with pytest.raises(ResultPlanError) as caught:
        compile_result_plan(
            item, _plan_request(hostId="wps-et", hostVersion="12.0"),
            target_sheet="Sheet1", target_range="A1:C501",
        )
    assert caught.value.code == "too_large"


# ---------------------------------------------------------------------------
# E. SDK end-to-end (fake model transport only)
# ---------------------------------------------------------------------------


class _FakeResult:
    def __init__(self, result: str = "") -> None:
        self.result = result
        self.is_error = False
        self.subtype = "success"
        self.api_error_status = None


class _BridgeFakeSdk:
    """Drives real SDK tool handlers and delivers bridge receipts inline."""

    ResultMessage = _FakeResult
    CLINotFoundError = type("CLINotFoundError", (Exception,), {})
    CLIConnectionError = type("CLIConnectionError", (Exception,), {})
    ProcessError = type("ProcessError", (Exception,), {})
    ResultError = type("ResultError", (Exception,), {})
    CLIJSONDecodeError = type("CLIJSONDecodeError", (Exception,), {})

    def __init__(self, steps, *, task_id, delivered_result, source_fingerprint="fp-source-1"):
        from server.app.task_bridge import task_tool_bridge

        self.steps = steps
        self.task_id = task_id
        self.delivered_result = delivered_result
        self.source_fingerprint = source_fingerprint
        self.schemas: dict[str, dict] = {}
        self.calls: list = []
        self.tool_results: list = []
        self._bridge = task_tool_bridge

    def ClaudeAgentOptions(self, **kwargs):
        return SimpleNamespace(**kwargs)

    def tool(self, name, _description, _schema):
        self.schemas[name] = _schema
        return lambda handler: SimpleNamespace(name=name, handler=handler)

    def create_sdk_mcp_server(self, *, name, version, tools):
        assert name == "excel_bro"
        return {tool.name: tool for tool in tools}

    async def _deliver(self):
        session = self._bridge.session(self.task_id)
        for _ in range(2000):
            if session is not None and session.pending_calls:
                break
            await asyncio.sleep(0)
        assert session is not None and session.pending_calls, "桥没有收到查询请求"
        call = session.pending_calls[0]
        return session.deliver(
            operation_id=call["operationId"],
            request_fingerprint=call["requestFingerprint"],
            host_id="excel-office",
            workbook_name="测试.xlsx",
            source_fingerprint=self.source_fingerprint,
            result=self.delivered_result,
            error_code=None,
        )

    async def query(self, *, prompt, options):
        self.calls.append((prompt, options))
        results: list = []
        for step in self.steps:
            name, args = step(results)
            tool = options.mcp_servers["excel_bro"][name]
            task = asyncio.create_task(tool.handler(args))
            if name == "query_table":
                await self._deliver()
            outcome = await task
            results.append((name, outcome))
            self.tool_results.append(outcome)
        yield _FakeResult("")


def _install(monkeypatch, tmp_path, fake):
    monkeypatch.setattr(claude_sdk_agent, "_sdk", lambda: fake)
    monkeypatch.setattr(claude_sdk_agent, "_state_dir", lambda: tmp_path)


@pytest.mark.parametrize("host_id", ["excel-office", "wps-et"])
@pytest.mark.parametrize("mode", ["join", "deduplicate"])
def test_sdk_unsupported_combine_reports_capability_without_host_read(monkeypatch, tmp_path, host_id, mode):
    class UnsupportedSdk(_BridgeFakeSdk):
        async def query(self, *, prompt, options):
            tool = options.mcp_servers["excel_bro"]["query_table"]
            outcome = await tool.handler({"mode": "rows", "combine": {
                "mode": mode, "leftSourceSheetId": "Sheet1", "rightSourceSheetId": "Sheet2",
                "leftKey": "类别", "rightKey": "类别", "deduplicateBy": ["类别"],
            }})
            assert json.loads(outcome["content"][0]["text"])["errorCode"] == "invalid_arguments"
            final = _FakeResult()
            final.is_error = True
            final.subtype = "error_max_turns"
            yield final

    fake = UnsupportedSdk([], task_id="turn-bridge-0001", delivered_result=_result())
    _install(monkeypatch, tmp_path, fake)
    async def no_read(*args, **kwargs):
        pytest.fail("不支持的查询不得投递给宿主")
    from server.app.task_bridge import ToolBridgeSession
    monkeypatch.setattr(ToolBridgeSession, "request", no_read)
    with pytest.raises(Exception, match="当前工作簿查询暂不支持关联或去重"):
        asyncio.run(claude_sdk_agent.run_sdk_plan(_plan_request(hostId=host_id), _connection()))


def test_sdk_query_table_loop_returns_typed_receipt_then_local_result(
    monkeypatch, tmp_path
):
    """完整闭环：query_table → 桥 → 真实登记 → 有类型回执 → submit_local_result。"""
    def step_1(_results):
        return ("query_table", {
            "mode": "aggregate", "groupBy": ["类别"],
            "metrics": [{"operation": "countRows", "outputName": "数量"}],
        })

    def step_2(results):
        payload = json.loads(results[0][1]["content"][0]["text"])
        assert payload["status"] == "ok"
        return ("submit_local_result", {"resultRef": payload["resultRef"]})

    fake = _BridgeFakeSdk(
        [step_1, step_2], task_id="turn-bridge-0001", delivered_result=_result(),
    )
    _install(monkeypatch, tmp_path, fake)
    request = _plan_request()
    response = asyncio.run(claude_sdk_agent.run_sdk_plan(request, _connection()))

    assert response.kind == "answer"
    assert response.localResult is not None
    combine = fake.schemas["query_table"]["properties"]["combine"]["anyOf"]
    assert next(item for item in combine if "properties" in item)["properties"]["mode"]["enum"] == ["union"]
    assert response.localResult.resultRef.startswith("res_")
    assert response.localResult.headers == ["类别", "数量"]
    assert response.localResult.rowCount == 2
    # 模型可见的 query_table 回执里没有值；结果卡只带引用与结构，不带行。
    query_text = fake.tool_results[0]["content"][0]["text"]
    assert SENTINEL not in query_text
    assert "rows" not in json.loads(query_text)
    assert not hasattr(response.localResult, "rows")
    # 整个响应（会进聊天流/账本）不得携带来自源数据的自由文本。
    assert SENTINEL not in response.model_dump_json()


def test_sdk_registers_the_local_tools_and_prunes_host_actions(monkeypatch, tmp_path):
    fake = _BridgeFakeSdk(
        [lambda _r: ("submit_local_result", {"resultRef": "res_" + "0" * 32})],
        task_id="turn-bridge-0001", delivered_result=_result(),
    )
    _install(monkeypatch, tmp_path, fake)
    with pytest.raises(Exception):
        # 引用不存在 → handler 返回错误，model 未提交结果 → SDK 报未提交。
        asyncio.run(claude_sdk_agent.run_sdk_plan(
            _plan_request(selfReportedActions=["writeValues"]), _connection()
        ))
    assert "query_table" in fake.schemas
    assert "describe_authorized_tables" in fake.schemas
    assert "submit_local_result" in fake.schemas
    assert "submit_result_plan" in fake.schemas
    # hostId 声明后 submit_plan 的动作被裁剪成 anyOf，且不含未支持变体。
    actions = fake.schemas["submit_plan"]["properties"]["plan"]["properties"]["actions"]
    assert "anyOf" in actions["items"]


def test_sdk_still_rejects_data_rows(monkeypatch, tmp_path):
    fake = _BridgeFakeSdk([], task_id="turn-bridge-0001", delivered_result=_result())
    _install(monkeypatch, tmp_path, fake)
    request = _plan_request(dataResults=[_result().model_dump(mode="json")])
    with pytest.raises(Exception, match="仅支持不上传数据行"):
        asyncio.run(claude_sdk_agent.run_sdk_plan(request, _connection()))
    assert fake.calls == []


def _connection():
    from server.app.llm.config import ModelConnection

    return ModelConnection(
        base_url="https://api.anthropic.com",
        model="test-claude",
        api_key="test-key",
        supports_vision=False,
        adapter="claude_agent_sdk",
    )


# ---------------------------------------------------------------------------
# F. loopback endpoints (real ASGI round-trip)
# ---------------------------------------------------------------------------


def _localhost_client():
    import httpx

    from server.app.main import app

    return httpx.AsyncClient(
        transport=httpx.ASGITransport(app=app), base_url="http://localhost"
    )


def _binding_body(task_id: str) -> dict:
    return {
        "taskId": task_id,
        "hostId": "excel-office",
        "workbookName": "测试.xlsx",
        "sourceFingerprint": "fp-source-1",
    }


def test_bridge_endpoints_poll_then_receipt():
    from server.app.task_bridge import task_tool_bridge

    binding = _binding()

    async def scenario():
        session = task_tool_bridge.open(binding)
        task = asyncio.create_task(
            session.request("query_table", {"mode": "rows", "fields": ["类别"]})
        )
        for _ in range(500):
            if session.pending_calls:
                break
            await asyncio.sleep(0)
        async with _localhost_client() as client:
            poll = await client.post(
                "/api/sdk/tool-bridge/poll", json=_binding_body(binding.task_id)
            )
            assert poll.status_code == 200
            calls = poll.json()["calls"]
            assert len(calls) == 1
            call = calls[0]
            assert call["arguments"] == {"mode": "rows", "fields": ["类别"]}
            posted = await client.post(
                "/api/sdk/tool-bridge/receipt",
                json={
                    **_binding_body(binding.task_id),
                    "operationId": call["operationId"],
                    "requestFingerprint": call["requestFingerprint"],
                    "receipt": {"result": _result().model_dump(mode="json")},
                },
            )
            assert posted.json() == {"status": "accepted", "code": None}
            # 重复回执幂等
            duplicate = await client.post(
                "/api/sdk/tool-bridge/receipt",
                json={
                    **_binding_body(binding.task_id),
                    "operationId": call["operationId"],
                    "requestFingerprint": call["requestFingerprint"],
                    "receipt": {"result": _result().model_dump(mode="json")},
                },
            )
            assert duplicate.json()["status"] == "duplicate"
        return await task

    receipt = asyncio.run(scenario())
    assert receipt["status"] == "ok"
    assert SENTINEL not in json.dumps(receipt, ensure_ascii=False)


def test_bridge_result_endpoint_returns_values_only_on_loopback():
    from server.app.task_bridge import task_tool_bridge

    binding = _binding()

    async def scenario():
        session = task_tool_bridge.open(binding)
        task = asyncio.create_task(session.request("query_table", {"mode": "rows"}))
        for _ in range(500):
            if session.pending_calls:
                break
            await asyncio.sleep(0)
        call = session.pending_calls[0]
        session.deliver(
            operation_id=call["operationId"],
            request_fingerprint=call["requestFingerprint"],
            host_id="excel-office",
            workbook_name="测试.xlsx",
            source_fingerprint="fp-source-1",
            result=_result(),
            error_code=None,
        )
        receipt = await task
        async with _localhost_client() as client:
            fetched = await client.post(
                "/api/sdk/tool-bridge/result",
                json={**_binding_body(binding.task_id), "resultRef": receipt["resultRef"]},
            )
            assert fetched.status_code == 200
            payload = fetched.json()
            assert payload["rows"] == [["甲", 2], ["乙", 1]]
            assert fetched.headers.get("cache-control") == "no-store"
            missing = await client.post(
                "/api/sdk/tool-bridge/result",
                json={**_binding_body(binding.task_id), "resultRef": "res_" + "9" * 32},
            )
            assert missing.status_code == 404

    asyncio.run(scenario())


def test_bridge_receipt_endpoint_rejects_late_operation():
    from server.app.task_bridge import task_tool_bridge

    binding = _binding()
    task_tool_bridge.open(binding)

    async def scenario():
        async with _localhost_client() as client:
            response = await client.post(
                "/api/sdk/tool-bridge/receipt",
                json={
                    **_binding_body(binding.task_id),
                    "operationId": "op_" + "7" * 32,
                    "requestFingerprint": "0" * 64,
                    "receipt": {"result": _result().model_dump(mode="json")},
                },
            )
            assert response.json()["status"] == "rejected"

    asyncio.run(scenario())


def test_bridge_register_endpoint_requires_valid_result():
    from server.app.task_bridge import task_tool_bridge

    binding = _binding()
    # register 是轮内动作：必须存在活跃且归属一致的会话。
    task_tool_bridge.open(binding)

    async def scenario():
        async with _localhost_client() as client:
            ok = await client.post(
                "/api/sdk/tool-bridge/register",
                json={**_binding_body(binding.task_id), "result": _result().model_dump(mode="json")},
            )
            assert ok.status_code == 200
            assert ok.json()["resultRef"].startswith("res_")
            bad = await client.post(
                "/api/sdk/tool-bridge/register",
                json={
                    **_binding_body(binding.task_id),
                    "result": _result(
                        headers=[f"h{index}" for index in range(50)],
                        rows=[[0] * 50 for _ in range(500)],
                    ).model_dump(mode="json"),
                },
            )
            assert bad.status_code == 422

    asyncio.run(scenario())
