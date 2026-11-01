from __future__ import annotations

import asyncio
import json
from dataclasses import replace

import pytest

from server.app import claude_sdk_agent as sdk, intent, planner
from server.app.agent_runtime import sdk_agent_connection
from server.app.llm import ModelNotAvailableError, ModelNotConfiguredError
from server.app.models import IntentCheckRequest, PlanRequest
from server.app.plan_tools import compile_plan_tool, plan_compiler_definitions
from server.tests.test_claude_sdk_agent import (
    _FakeSdk, _FakeResult, _connection, _install_fake, _intent_request, _plan_request,
)


def plan_request(prompt="删除sheet1、sheet2之外的子表", **kwargs):
    request = _plan_request()
    request.prompt = prompt
    base = request.workbook.worksheets[0]
    request.workbook.worksheets = [base.model_copy(update={"name": name})
                                   for name in ("Sheet1", "Sheet2", "Sheet3")]
    return request.model_copy(update=kwargs)


@pytest.mark.parametrize("prompt", [
    "删除sheet1、sheet2之外的子表", "不要删除任何表", "清空 Sheet1!A1",
    "A1 写入 3", 'A1 写入"001"', "在C2写入=SUM(B2:B4)",
    "关联两张表，重复项怎么处理还没决定", "清空 A1 后将统计结果写到 B1",
])
@pytest.mark.parametrize("host", ["excel-office", "wps-et"])
def test_every_natural_request_reaches_sdk_unchanged(monkeypatch, tmp_path, prompt, host):
    fake = _FakeSdk(response='{"kind":"answer","message":"已收到需求"}')
    _install_fake(monkeypatch, tmp_path, fake)
    monkeypatch.setattr(planner, "selected_model_config", lambda _: _connection())
    request = plan_request(prompt, hostId=host, hostVersion="16.0" if host == "excel-office" else "12.0")
    response = asyncio.run(planner.create_plan(request))
    assert response.kind == "answer"
    assert len(fake.calls) == 1
    assert json.loads(fake.calls[0][0])["prompt"] == prompt
    assert "PRIVATE_DATA_ROW" not in fake.calls[0][0]


def test_deepseek_connection_keeps_model_key_and_vision():
    old = replace(_connection(), adapter="openai_compatible", base_url="https://api.deepseek.com/v1")
    new = sdk_agent_connection(old)
    assert new == replace(old, adapter="claude_agent_sdk", base_url="https://api.deepseek.com/anthropic")
    assert old.adapter == "openai_compatible"


def test_unmapped_connection_does_not_fall_back_to_old_agent(monkeypatch):
    old = replace(_connection(), adapter="openai_compatible", base_url="https://gateway.invalid/v1")
    monkeypatch.setattr(planner, "selected_model_config", lambda _: old)
    with pytest.raises(ModelNotAvailableError, match="已统一使用"):
        asyncio.run(planner.create_plan(plan_request("清空 A1")))


@pytest.mark.parametrize("prompt", ["清空 A1", "删除sheet1之外的子表"])
def test_no_model_never_creates_local_plan(monkeypatch, prompt):
    def missing(_):
        raise ModelNotConfiguredError("未配置模型")
    monkeypatch.setattr(planner, "selected_model_config", missing)
    with pytest.raises(ModelNotConfiguredError):
        asyncio.run(planner.create_plan(plan_request(prompt)))


def test_sdk_failure_never_restarts_on_old_agent(monkeypatch, tmp_path):
    monkeypatch.setattr(planner, "selected_model_config", lambda _: _connection())
    calls = []
    async def failed(*args, **kwargs):
        calls.append(1)
        raise RuntimeError("SDK interrupted")
    monkeypatch.setattr(sdk, "run_sdk_plan", failed)
    with pytest.raises(RuntimeError, match="SDK interrupted"):
        asyncio.run(planner.create_plan(plan_request()))
    assert calls == [1]


def test_legacy_checkpoint_is_rejected_before_sdk_call(monkeypatch):
    monkeypatch.setattr(planner, "selected_model_config", lambda _: _connection())
    # 旧检查点没有 `kind`：拒绝码必须指到"这不是本契约的现场"，而不是笼统一句
    # "不能继续"——后者会让调用方以为再点一次就好。
    with pytest.raises(ModelNotAvailableError, match="PLANNING_CONTEXT_KIND_MISMATCH"):
        asyncio.run(planner.create_plan(plan_request(), resume_checkpoint={"old": True}))


def test_intent_no_longer_rewrites_multiple_keeper_names(monkeypatch, tmp_path):
    fake = _FakeSdk(response=json.dumps({"kind": "proceed", "summary": "保留两张表", "confirmedPrompt": "保留 Sheet1 和 Sheet2，删除其他已授权表"}))
    _install_fake(monkeypatch, tmp_path, fake)
    request = _intent_request()
    request.prompt = "删除sheet1、sheet2之外的子表"
    response = asyncio.run(intent.check_intent(request, config=_connection()))
    assert response.provider == "model"
    assert json.loads(fake.calls[0][0])["prompt"] == request.prompt


def test_structured_delete_keeps_multiple_sheets_and_never_uses_global_names():
    request = plan_request()
    request.workbook.activeWorksheet = "UNAUTHORIZED"
    response = compile_plan_tool(request, "plan_delete_other_worksheets", {"keepSheets": ["sheet1", "sheet2"]})
    assert [a.sheet for a in response.plan.actions] == ["Sheet3"]
    assert all(a.type == "deleteWorksheet" for a in response.plan.actions)


@pytest.mark.parametrize("names,code", [(["Sheet1、Sheet2"], "sheet_not_authorized"), (["missing"], "sheet_not_authorized"), ([], "at least 1")])
def test_structured_delete_never_guesses(names, code):
    with pytest.raises(ValueError, match=code):
        compile_plan_tool(plan_request(), "plan_delete_other_worksheets", {"keepSheets": names})


def test_folder_ambiguous_names_and_multiple_files_fail_closed():
    request = plan_request()
    for i, sheet in enumerate(request.workbook.worksheets):
        sheet.name = f"File{i} › Sheet1"
        sheet.sourceSheet = "Sheet1"
        sheet.sourceFile = f"file{i}.xlsx"
    with pytest.raises(ValueError, match="sheet_ambiguous"):
        compile_plan_tool(request, "plan_delete_other_worksheets", {"keepSheets": ["Sheet1"]})
    with pytest.raises(ValueError, match="file_ambiguous"):
        compile_plan_tool(request, "plan_delete_other_worksheets", {"keepSheets": [request.workbook.worksheets[0].name]})


@pytest.mark.parametrize("host,version,allowed", [
    ("excel-office", "16.0", ["writeValues", "clearRange:contents"]),
    ("wps-et", "12.0", ["writeValues", "clearRange:contents"]),
])
def test_compiler_tools_respect_negotiated_host_capabilities(host, version, allowed):
    request = plan_request(hostId=host, hostVersion=version, selfReportedActions=allowed)
    names = [d["function"]["name"] for d in plan_compiler_definitions(request)]
    assert "plan_delete_other_worksheets" not in names
    assert "plan_range_edit" in names
    with pytest.raises(ValueError, match="能力范围"):
        compile_plan_tool(request, "plan_delete_other_worksheets", {"keepSheets": ["Sheet1"]})


@pytest.mark.parametrize("operation,value,key,expected", [
    ("writeValues", "001", "values", "001"), ("writeValues", False, "values", False),
    ("writeValues", None, "values", None), ("writeFormulas", "=SUM(B1:B2)", "formulas", "=SUM(B1:B2)"),
])
def test_literal_compiler_preserves_type_and_matrix(operation, value, key, expected):
    result = compile_plan_tool(plan_request(), "plan_range_edit", {"sheet": "sheet2", "range": "$C$1:C2", "operation": operation, "value": value})
    action = result.plan.actions[0]
    assert action.sheet == "Sheet2"
    assert getattr(action, key) == [[expected], [expected]]
    assert result.planOrigin == ("literal_range_edit" if operation == "writeValues" else None)


@pytest.mark.parametrize("args,code", [
    ({"sheet": "missing"}, "sheet_not_authorized"),
    ({"range": "B2:A1"}, "invalid_range"),
    ({"range": "A0"}, "invalid_range"),
    ({"range": "XFE1"}, "invalid_range"),
    ({"range": "A1:Z999999"}, "range_too_large"),
    ({"operation": "clearRange"}, "invalid_arguments"),
    ({"value": "=HYPERLINK(\"x\")"}, "formula_as_value"),
])
def test_literal_compiler_rejects_invalid_or_unauthorized_arguments(args, code):
    with pytest.raises(ValueError, match=code):
        compile_plan_tool(plan_request(), "plan_range_edit", {"sheet": "Sheet1", "range": "A1", "operation": "writeValues", "value": 1, **args})


def test_sdk_can_clarify_after_tool_error_then_continue_same_task(monkeypatch, tmp_path):
    clarification = {"id": "choose-sheets", "summary": "确认保留表", "question": "要保留哪些表？", "reason": "名称存在不同解释", "scopeLabel": "当前授权表", "options": [
        {"id": "two", "label": "保留两张", "description": "Sheet1 和 Sheet2", "resolution": "保留 Sheet1、Sheet2", "action": "resolve", "resolutionMode": "append"},
        {"id": "one", "label": "保留一张", "description": "仅 Sheet1", "resolution": "仅保留 Sheet1", "action": "resolve", "resolutionMode": "append"},
    ]}
    class SequenceSdk(_FakeSdk):
        async def query(self, *, prompt, options):
            self.calls.append((prompt, options))
            tools = options.mcp_servers["excel_bro"]
            if len(self.calls) == 1:
                receipt = await tools["plan_delete_other_worksheets"].handler({"keepSheets": ["Sheet1、Sheet2"]})
                assert receipt["is_error"] is True
                assert "sheet_not_authorized" in receipt["content"][0]["text"]
                result = await tools["submit_clarification"].handler({"clarification": clarification})
                assert not result.get("is_error")
            else:
                received = json.loads(prompt)
                assert received["taskContext"]["confirmations"][0]["value"] == "保留 Sheet1、Sheet2"
                await tools["plan_delete_other_worksheets"].handler({"keepSheets": ["Sheet1", "Sheet2"]})
            yield _FakeResult()
    fake = SequenceSdk()
    _install_fake(monkeypatch, tmp_path, fake)
    monkeypatch.setattr(planner, "selected_model_config", lambda _: _connection())
    request = plan_request(turnId="same-task")
    first = asyncio.run(planner.create_plan(request))
    assert first.kind == "clarification" and first.turnId == "same-task"
    answered = PlanRequest.model_validate({**request.model_dump(), "taskContext": {
        "schemaVersion": 1, "originalPrompt": request.prompt,
        "confirmations": [{"questionId": "choose-sheets", "label": "保留范围", "value": "保留 Sheet1、Sheet2"}],
    }})
    second = asyncio.run(planner.create_plan(answered))
    assert second.kind == "plan"
    assert [a.sheet for a in second.plan.actions] == ["Sheet3"]
    assert request.workbook == answered.workbook

def test_model_json_cannot_attest_literal_origin():
    from server.app.excel_agent import _parse_content
    result = compile_plan_tool(plan_request(), "plan_range_edit", {
        "sheet": "Sheet1", "range": "A1", "operation": "writeValues", "value": "完成",
    })
    assert result.planOrigin == "literal_range_edit"
    parsed = _parse_content(result.model_dump_json())
    assert parsed.planOrigin is None


def literal_edits():
    return [{"sheet": "Sheet1", "range": f"B{index + 2}", "value": value}
            for index, value in enumerate([111, 222, 333])]


@pytest.mark.parametrize("host,version", [("excel-office", "16.0"), ("wps-et", "12.0"), (None, None)])
def test_ordered_literal_edits_preserve_three_actions_and_trusted_origin(host, version):
    from server.app.excel_agent import _parse_content
    request = plan_request(hostId=host, hostVersion=version, selfReportedActions=["writeValues"])
    assert "plan_literal_edits" in [item["function"]["name"] for item in plan_compiler_definitions(request)]
    result = compile_plan_tool(request, "plan_literal_edits", {"edits": literal_edits()})
    assert [(a.type, a.sheet, a.range, a.values) for a in result.plan.actions] == [
        ("writeValues", "Sheet1", f"B{index + 2}", [[value]])
        for index, value in enumerate([111, 222, 333])]
    assert result.planOrigin == "literal_range_edit"
    assert [(c.range, c.expected) for c in result.plan.acceptanceCriteria if c.type == "rangeEquals"] == [
        (f"B{index + 2}", [[value]]) for index, value in enumerate([111, 222, 333])]
    assert _parse_content(result.model_dump_json()).planOrigin is None


@pytest.mark.parametrize("edit,code", [
    ({"sheet": "missing"}, "sheet_not_authorized"),
    ({"range": "XFE1"}, "invalid_range"),
    ({"value": "=SUM(A1:A2)"}, "formula_as_value"),
    ({"value": [[111]]}, "validation error"),
    ({"operation": "clearRange"}, "Extra inputs"),
])
def test_literal_edits_reject_whole_batch_on_invalid_member(edit, code):
    edits = literal_edits()
    edits[1].update(edit)
    with pytest.raises(ValueError, match=code):
        compile_plan_tool(plan_request(), "plan_literal_edits", {"edits": edits})


def test_literal_edits_enforce_total_cell_and_action_limits_and_host_capabilities(monkeypatch):
    from server.app import plan_tools
    real_limit = plan_tools.capability_int
    monkeypatch.setattr(plan_tools, "capability_int", lambda section, key:
                        2 if key == "maxLiteralActions" else real_limit(section, key))
    with pytest.raises(ValueError, match="too_many_literal_edits"):
        compile_plan_tool(plan_request(), "plan_literal_edits", {"edits": literal_edits()})
    monkeypatch.setattr(plan_tools, "capability_int", real_limit)
    with pytest.raises(ValueError, match="range_too_large"):
        compile_plan_tool(plan_request(), "plan_literal_edits", {"edits": [
            {"sheet": "Sheet1", "range": "A1:A600", "value": 1},
            {"sheet": "Sheet1", "range": "B1:B600", "value": 2}]})
    with pytest.raises(ValueError):
        compile_plan_tool(plan_request(), "plan_literal_edits", {"edits": []})
    with pytest.raises(ValueError):
        compile_plan_tool(plan_request(), "plan_literal_edits", {"edits": [{"sheet": "Sheet1", "range": "A1"}]})
    restricted = plan_request(hostId="wps-et", hostVersion="12.0", selfReportedActions=["clearRange:contents"])
    assert "plan_literal_edits" not in [item["function"]["name"] for item in plan_compiler_definitions(restricted)]
    with pytest.raises(ValueError, match="能力范围"):
        compile_plan_tool(restricted, "plan_literal_edits", {"edits": literal_edits()})


def test_sdk_submits_ordered_literals_through_production_tool_handler(monkeypatch, tmp_path):
    class LiteralSdk(_FakeSdk):
        async def query(self, *, prompt, options):
            self.calls.append((prompt, options))
            receipt = await options.mcp_servers["excel_bro"]["plan_literal_edits"].handler({"edits": literal_edits()})
            assert not receipt.get("is_error")
            yield _FakeResult()
    fake = LiteralSdk()
    _install_fake(monkeypatch, tmp_path, fake)
    monkeypatch.setattr(planner, "selected_model_config", lambda _: _connection())
    response = asyncio.run(planner.create_plan(plan_request("按顺序分别将 B2/B3/B4 写成111/222/333，不合并动作")))
    assert response.planOrigin == "literal_range_edit"
    assert [a.range for a in response.plan.actions] == ["B2", "B3", "B4"]
