from __future__ import annotations

import asyncio
import json
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

from server.app import claude_sdk_agent
from server.app.intent import check_intent
from server.app.llm.client import OpenAICompatibleClient
from server.app.llm.config import (
    CLAUDE_SDK_CATALOG_ID,
    ModelConnection,
    ModelNotAvailableError,
    model_catalog,
    selected_model_config,
)
from server.app.llm import config as llm_config
from server.app.models import IntentCheckRequest, PlanRequest
from server.app.planner import create_plan


def _connection() -> ModelConnection:
    return ModelConnection(
        base_url="https://api.anthropic.com",
        model="test-claude",
        api_key="test-key",
        supports_vision=False,
        adapter="claude_agent_sdk",
    )


class _FakeResult:
    def __init__(self, result: str = "") -> None:
        self.result = result
        self.is_error = False
        self.subtype = "success"
        self.api_error_status = None


class _FakeSdk:
    ResultMessage = _FakeResult
    CLINotFoundError = type("CLINotFoundError", (Exception,), {})
    CLIConnectionError = type("CLIConnectionError", (Exception,), {})
    ProcessError = type("ProcessError", (Exception,), {})
    ResultError = type("ResultError", (Exception,), {})
    CLIJSONDecodeError = type("CLIJSONDecodeError", (Exception,), {})

    def __init__(self, *, response: str = "", invoke: tuple[str, dict] | None = None) -> None:
        self.response = response
        self.invoke = invoke
        self.calls: list[tuple[str, SimpleNamespace]] = []
        self.tool_result: dict | None = None
        # 注册进 MCP 的**工具契约**：两条模型后端必须共用同一份，不能各写一份。
        self.schemas: dict[str, dict] = {}

    def ClaudeAgentOptions(self, **kwargs: object) -> SimpleNamespace:
        return SimpleNamespace(**kwargs)

    def tool(self, name: str, _description: str, _schema: dict):
        self.schemas[name] = _schema
        return lambda handler: SimpleNamespace(name=name, handler=handler)

    def create_sdk_mcp_server(self, *, name: str, version: str, tools: list):
        assert name == "excel_bro" and version
        return {tool.name: tool for tool in tools}

    async def query(self, *, prompt: str, options: SimpleNamespace):
        self.calls.append((prompt, options))
        if self.invoke is not None:
            name, args = self.invoke
            tool = options.mcp_servers["excel_bro"][name]
            self.tool_result = await tool.handler(args)
        yield _FakeResult(self.response)


def _intent_request() -> IntentCheckRequest:
    return IntentCheckRequest.model_validate({
        "prompt": "请按问题类型填写对应等级",
        "scope": {
            "workbookName": "测试.xlsx",
            "sourceMode": "workbook",
            "selectionMode": "auto",
            "activeWorksheet": "Sheet1",
            "totalWorksheetCount": 1,
            "worksheetNames": ["Sheet1"],
            "sheets": [{
                "name": "Sheet1", "usedRange": "A1:B3", "rowCount": 3,
                "columnCount": 2, "headers": ["类别", "等级"],
            }],
        },
        "conversation": [{"role": "user", "text": "PRIVATE_CONVERSATION_ROW"}],
        "priorResult": {
            "title": "上轮结果", "headers": ["秘密"],
            "rows": [["PRIVATE_RESULT_ROW"]], "sourceSheets": ["Sheet1"],
        },
    })


def _plan_request(*, policy: str = "none") -> PlanRequest:
    return PlanRequest.model_validate({
        "prompt": "请按类别填写等级",
        "planningReadPolicy": policy,
        "modelId": CLAUDE_SDK_CATALOG_ID,
        "workbook": {
            "name": "测试.xlsx", "capturedAt": "2026-10-01T00:00:00Z",
            "activeWorksheet": "Sheet1",
            "worksheets": [{
                "name": "Sheet1", "sourceFile": "PRIVATE_FULL_PATH",
                "usedRange": "A1:B3", "rowCount": 3, "columnCount": 2,
                "headers": ["类别", "等级"],
                "dataRows": [["PRIVATE_DATA_ROW", None], ["另一个值", None]],
            }],
        },
    })


def _install_fake(monkeypatch: pytest.MonkeyPatch, tmp_path: Path, fake: _FakeSdk) -> None:
    monkeypatch.setattr(claude_sdk_agent, "_sdk", lambda: fake)
    monkeypatch.setattr(claude_sdk_agent, "_state_dir", lambda: tmp_path)


def test_sdk_catalog_is_explicit_and_preserves_openai_default(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path,
) -> None:
    monkeypatch.setattr("server.app.llm.config._config_path", lambda: tmp_path / "config.env")
    monkeypatch.setenv("AI_BASE_URL", "https://example.invalid/v1")
    monkeypatch.setenv("AI_MODEL", "old-model")
    monkeypatch.setenv("ANTHROPIC_API_KEY", "test-key")
    monkeypatch.setenv("CLAUDE_AGENT_SDK_MODEL", "test-claude")
    monkeypatch.delenv("CLAUDE_AGENT_SDK_STATE_DIR", raising=False)
    assert CLAUDE_SDK_CATALOG_ID not in {item["id"] for item in model_catalog()["models"]}
    with pytest.raises(ModelNotAvailableError):
        selected_model_config(CLAUDE_SDK_CATALOG_ID)

    monkeypatch.setenv("CLAUDE_AGENT_SDK_STATE_DIR", ".\\sdk-state")
    assert CLAUDE_SDK_CATALOG_ID not in {item["id"] for item in model_catalog()["models"]}

    outside_repo = str(Path(__file__).resolve().anchor) + "claude-sdk-test-state"
    monkeypatch.setenv("CLAUDE_AGENT_SDK_STATE_DIR", outside_repo)
    catalog = model_catalog()
    assert catalog["defaultModelId"] == "old-model"
    assert CLAUDE_SDK_CATALOG_ID in {item["id"] for item in catalog["models"]}
    assert selected_model_config(CLAUDE_SDK_CATALOG_ID).adapter == "claude_agent_sdk"


def test_sdk_intent_sends_only_current_request_and_structure(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path,
) -> None:
    fake = _FakeSdk(response=json.dumps({
        "kind": "proceed", "summary": "填写等级",
        "confirmedPrompt": "请按问题类型填写对应等级",
    }, ensure_ascii=False))
    _install_fake(monkeypatch, tmp_path, fake)
    result = asyncio.run(check_intent(_intent_request(), config=_connection()))
    assert result.kind == "proceed"
    prompt, options = fake.calls[0]
    assert "类别" in prompt
    assert "PRIVATE_CONVERSATION_ROW" not in prompt
    assert "PRIVATE_RESULT_ROW" not in prompt
    assert options.tools == []
    assert options.setting_sources == []
    assert options.permission_mode == "dontAsk"
    assert options.cwd == str(tmp_path)
    assert options.mcp_servers == {}
    assert options.env["ANTHROPIC_BASE_URL"] == "https://api.anthropic.com"
    assert options.env["ANTHROPIC_AUTH_TOKEN"] == ""


def test_sdk_intent_rejects_images_before_query(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path,
) -> None:
    fake = _FakeSdk()
    _install_fake(monkeypatch, tmp_path, fake)
    request = _intent_request().model_copy(update={"imageCount": 1})
    with pytest.raises(ModelNotAvailableError, match="未接入图片"):
        asyncio.run(check_intent(request, config=_connection()))
    assert fake.calls == []


def test_sdk_planner_uses_validated_terminal_tool_without_exporting_rows(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path,
) -> None:
    plan = {
        "id": "test-plan", "title": "填写等级", "summary": "在 B2 写入等级",
        "actions": [{
            "type": "writeValues", "sheet": "Sheet1", "range": "B2", "values": [["甲"]],
        }],
    }
    fake = _FakeSdk(invoke=("submit_plan", {"plan": plan}))
    _install_fake(monkeypatch, tmp_path, fake)
    monkeypatch.setattr("server.app.planner.selected_model_config", lambda _id: _connection())
    result = asyncio.run(create_plan(_plan_request()))
    assert result.kind == "plan"
    assert result.plan.actions[0].type == "writeValues"
    assert fake.tool_result is not None and fake.tool_result.get("is_error") is not True
    prompt, options = fake.calls[0]
    assert "PRIVATE_DATA_ROW" not in prompt
    assert "PRIVATE_FULL_PATH" not in prompt
    assert options.tools == []
    assert options.strict_mcp_config is True
    assert "mcp__excel_bro__submit_plan" in options.allowed_tools


def test_sdk_planner_accepts_structured_answer_without_tool_call(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path,
) -> None:
    fake = _FakeSdk(response=json.dumps({"kind": "answer", "message": "这项查询没有匹配记录。"}))
    _install_fake(monkeypatch, tmp_path, fake)
    result = asyncio.run(claude_sdk_agent.run_sdk_plan(_plan_request(), _connection()))
    assert result.kind == "answer"
    assert result.message == "这项查询没有匹配记录。"


@pytest.mark.parametrize("use_tool", [True, False])
def test_sdk_planning_clarification_is_a_form_not_a_plain_answer(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, use_tool: bool,
) -> None:
    clarification = {
        "id": "missing-info", "summary": "补充统计口径", "question": "请填写两项信息",
        "reason": "当前口径不完整", "scopeLabel": "Sheet1", "options": [],
        "fields": [{"id": "period", "label": "时间范围"}, {"id": "unit", "label": "统计单位"}],
    }
    fake = _FakeSdk(invoke=("submit_clarification", {"clarification": clarification})) if use_tool else _FakeSdk(
        response=json.dumps({"kind": "clarification", "clarification": clarification}, ensure_ascii=False)
    )
    _install_fake(monkeypatch, tmp_path, fake)
    result = asyncio.run(claude_sdk_agent.run_sdk_plan(_plan_request(), _connection()))
    assert result.kind == "clarification"
    assert [field.id for field in result.clarification.fields] == ["period", "unit"]
    assert not hasattr(result, "plan")
    prompt, options = fake.calls[0]
    assert "PRIVATE_DATA_ROW" not in prompt and "PRIVATE_FULL_PATH" not in prompt
    assert "mcp__excel_bro__submit_clarification" in options.allowed_tools
    assert "不要用 submit_answer 追问" in options.system_prompt


def test_sdk_rejects_clarification_without_any_interactive_control(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path,
) -> None:
    fake = _FakeSdk(invoke=("submit_clarification", {"clarification": {
        "id": "invalid", "summary": "补充", "question": "怎么做？", "reason": "缺信息",
        "scopeLabel": "Sheet1", "options": [], "fields": [],
    }}))
    _install_fake(monkeypatch, tmp_path, fake)
    with pytest.raises(claude_sdk_agent.LLMResponseError, match="未提交有效"):
        asyncio.run(claude_sdk_agent.run_sdk_plan(_plan_request(), _connection()))
    assert fake.tool_result["is_error"] is True


def test_sdk_json_plan_still_passes_terminal_capability_check(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path,
) -> None:
    fake = _FakeSdk(response=json.dumps({"kind": "plan", "plan": {
        "id": "json-plan", "title": "写入", "summary": "写入常量",
        "actions": [{"type": "writeValues", "sheet": "Sheet1", "range": "B2", "values": [["甲"]]}],
    }}))
    _install_fake(monkeypatch, tmp_path, fake)
    checked = []

    def reject(_request, plan):
        checked.append(plan.id)
        raise ValueError("host action not allowed")

    monkeypatch.setattr("server.app.excel_agent._enforce_action_capabilities", reject)
    with pytest.raises(claude_sdk_agent.LLMResponseError, match="未提交有效"):
        asyncio.run(claude_sdk_agent.run_sdk_plan(_plan_request(), _connection()))
    assert checked == ["json-plan"]


def test_sdk_plain_prose_cannot_claim_to_be_a_preview_plan(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path,
) -> None:
    fake = _FakeSdk(response="我已经生成计划并写入了工作表。")
    _install_fake(monkeypatch, tmp_path, fake)
    with pytest.raises(claude_sdk_agent.LLMResponseError, match="未提交有效"):
        asyncio.run(claude_sdk_agent.run_sdk_plan(_plan_request(), _connection()))


def test_sdk_mapping_tool_delivers_local_rule_without_source_values(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path,
) -> None:
    fake = _FakeSdk(invoke=("map_column_values", {
        "sourceSheet": "Sheet1", "sourceRange": "A2:A3", "targetRange": "B2:B3",
        "mapping": [{"from": "甲", "to": "X"}, {"from": "乙", "to": "Y"}],
        "blankLabel": "-",
    }))
    _install_fake(monkeypatch, tmp_path, fake)
    result = asyncio.run(claude_sdk_agent.run_sdk_plan(_plan_request(), _connection()))
    assert result.kind == "local_plan_rule"
    assert result.rule.targetRange == "B2:B3"
    assert fake.tool_result.get("is_error") is not True


def test_sdk_rejects_data_bearing_plan_without_fallback(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path,
) -> None:
    fake = _FakeSdk()
    _install_fake(monkeypatch, tmp_path, fake)
    monkeypatch.setattr("server.app.planner.selected_model_config", lambda _id: _connection())
    with pytest.raises(ModelNotAvailableError, match="仅支持不上传数据行"):
        asyncio.run(create_plan(_plan_request(policy="ifNeeded")))
    assert fake.calls == []


def test_sdk_connection_cannot_be_used_as_openai_compatible() -> None:
    with pytest.raises(ModelNotAvailableError, match="尚未接入"):
        OpenAICompatibleClient(_connection(), timeout=1)


def test_sdk_reuses_saved_deepseek_connection_without_reentering_key(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path,
) -> None:
    monkeypatch.setattr(llm_config, "_config_path", lambda: tmp_path / "config.env")
    for name in (
        "AI_BASE_URL", "AI_MODEL", "OPENAI_BASE_URL", "OPENAI_MODEL",
        "ANTHROPIC_API_KEY", "CLAUDE_AGENT_SDK_MODEL",
    ):
        monkeypatch.delenv(name, raising=False)
    monkeypatch.setenv(
        "CLAUDE_AGENT_SDK_STATE_DIR",
        str(Path(__file__).resolve().anchor) + "claude-sdk-test-state",
    )
    view = llm_config.upsert_managed_connection(
        connection_id=None, label="DS", base_url="https://api.deepseek.com/v1/",
        model="deepseek-flash", api_key="private-deepseek-key",
        clear_api_key=False, supports_vision=False,
    )
    original_id = view["connections"][0]["catalogModelId"]
    derived_id = f"claude-agent-sdk:{view['connections'][0]['id']}"
    catalog = model_catalog()
    assert catalog["defaultModelId"] == original_id
    assert original_id in {item["id"] for item in catalog["models"]}
    assert derived_id in {item["id"] for item in catalog["models"]}
    assert "private-deepseek-key" not in json.dumps(catalog)

    selected = selected_model_config(derived_id)
    assert selected.adapter == "claude_agent_sdk"
    assert selected.base_url == "https://api.deepseek.com/anthropic"
    assert selected.model == "deepseek-flash"
    assert selected.api_key == "private-deepseek-key"

    fake = _FakeSdk(response=json.dumps({
        "kind": "proceed", "summary": "填写等级",
        "confirmedPrompt": "请按问题类型填写对应等级",
    }, ensure_ascii=False))
    _install_fake(monkeypatch, tmp_path, fake)
    result = asyncio.run(check_intent(_intent_request(), config=selected))
    assert result.kind == "proceed"
    _, options = fake.calls[0]
    assert options.model == "deepseek-flash"
    assert options.env["ANTHROPIC_BASE_URL"] == "https://api.deepseek.com/anthropic"
    assert options.env["ANTHROPIC_AUTH_TOKEN"] == "private-deepseek-key"
    assert options.env["ANTHROPIC_API_KEY"] == "private-deepseek-key"
    assert options.env["ANTHROPIC_DEFAULT_OPUS_MODEL"] == "deepseek-flash"
    assert options.env["ANTHROPIC_DEFAULT_SONNET_MODEL"] == "deepseek-flash"
    assert options.env["ANTHROPIC_DEFAULT_HAIKU_MODEL"] == "deepseek-flash"

    llm_config.upsert_managed_connection(
        connection_id=view["connections"][0]["id"], label="DS",
        base_url="https://api.moonshot.cn/v1", model="deepseek-flash",
        api_key=None, clear_api_key=False, supports_vision=False,
    )
    assert derived_id not in {item["id"] for item in model_catalog()["models"]}
    with pytest.raises(ModelNotAvailableError):
        selected_model_config(derived_id)


@pytest.mark.parametrize("base_url, allowed", [
    ("https://api.deepseek.com", True),
    ("https://api.deepseek.com/v1/", True),
    ("http://api.deepseek.com/v1", False),
    ("https://api.deepseek.com:444/v1", False),
    ("https://api.deepseek.com.evil.example/v1", False),
    ("https://api.deepseek.com@evil.example/v1", False),
    ("https://api.deepseek.com/v1?redirect=evil", False),
    ("https://api.deepseek.com/anthropic", False),
    ("https://api.moonshot.cn/v1", False),
])
def test_sdk_only_adapts_known_deepseek_openai_endpoint(
    base_url: str, allowed: bool,
) -> None:
    assert (llm_config._deepseek_anthropic_url(base_url) is not None) is allowed


def test_sdk_launcher_uses_saved_model_configuration_without_prompts() -> None:
    launcher = (
        Path(__file__).resolve().parents[2]
        / "scripts" / "start_claude_sdk_server.ps1"
    ).read_text(encoding="ascii")
    assert "Read-Host" not in launcher
    assert "ANTHROPIC_API_KEY" not in launcher
    assert "CLAUDE_AGENT_SDK_MODEL" not in launcher
    assert "--reload" not in launcher
    assert "--workers 1" in launcher


@pytest.mark.skipif(sys.platform != "win32", reason="Windows subprocess loop regression")
def test_sdk_rejects_selector_loop_before_starting_cli() -> None:
    fake = _FakeSdk(response="OK")
    with asyncio.Runner(loop_factory=asyncio.SelectorEventLoop) as runner:
        with pytest.raises(ModelNotAvailableError, match="Windows.*热重载") as caught:
            runner.run(claude_sdk_agent._result(fake, "fixed probe", SimpleNamespace(), 1))
    assert "start_claude_sdk_server.ps1" in str(caught.value)
    assert "AI_BASE_URL" not in str(caught.value)
    assert fake.calls == []


def test_sdk_backend_shares_the_blank_policy_contract_not_a_forked_schema(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path,
) -> None:
    """
    两种模型后端必须走到**同一套**结构化交互。

    `map_column` 的旧契约把 `blankLabel` 列为必填，模型为了填满它就只能在
    读到源值之前先问"空白怎么办" —— 那正是主流程卡住的根源。Agent SDK 后端
    是从 `excel_agent` 导入工具定义的，这里钉住它**没有另起一份**。
    """
    from server.app.excel_agent import AGENT_TOOLS_WITH_DATA

    fake = _FakeSdk(invoke=("submit_plan", {"plan": {
        "id": "test-plan", "title": "填写等级", "summary": "在 B2 写入等级",
        "actions": [{"type": "writeValues", "sheet": "Sheet1", "range": "B2", "values": [["甲"]]}],
    }}))
    _install_fake(monkeypatch, tmp_path, fake)
    monkeypatch.setattr("server.app.planner.selected_model_config", lambda _id: _connection())
    asyncio.run(create_plan(_plan_request()))

    sdk_schema = fake.schemas["map_column_values"]
    shared_schema = next(
        item["function"] for item in AGENT_TOOLS_WITH_DATA
        if item["function"]["name"] == "map_column_values"
    )["parameters"]
    # 逐字同一份，不是"看起来一样"的两份。
    assert sdk_schema == shared_schema

    properties = sdk_schema["properties"]
    assert "blankPolicy" in properties
    # 必填里没有空白字段：没读到源值之前不需要、也不该问。
    assert set(sdk_schema.get("required", [])) == {
        "sourceSheet", "sourceRange", "targetRange", "mapping"
    }
    # 旧字段还在（兼容旧规则），但描述里明确不许为了填它而先去问用户。
    assert "不要为了填它而先去问用户" in properties["blankLabel"]["description"]
    assert properties["blankPolicy"]["properties"]["kind"]["enum"] == [
        "keep", "clear", "value", "undecided"
    ]


@pytest.mark.parametrize("error_name", ["CLIConnectionError", "ProcessError"])
def test_sdk_process_failure_does_not_expose_stderr_or_claim_network_error(error_name: str) -> None:
    fake = _FakeSdk()

    async def broken_query(**_kwargs):
        raise getattr(fake, error_name)("private-key-and-prompt-stderr")
        yield  # Make this an async iterator, matching sdk.query.

    fake.query = broken_query
    with pytest.raises(ModelNotAvailableError) as caught:
        asyncio.run(claude_sdk_agent._result(fake, "fixed probe", SimpleNamespace(), 1))
    assert "Agent SDK" in str(caught.value)
    assert "private-key-and-prompt-stderr" not in str(caught.value)
    assert "AI_BASE_URL" not in str(caught.value)
