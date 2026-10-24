from __future__ import annotations

import asyncio

import json

from pathlib import Path

import httpx

import pytest

import server.app.excel_agent as excel_agent

import server.app.llm.config as llm_config

import server.app.planner as planner_module

from server.app.models import PlanRequest

from server.app.planner import create_plan

_CONFIGURED_MODEL_ID = "connection:test-primary"

_READONLY_TOOL_NAMES = excel_agent.READONLY_TOOLS

@pytest.fixture(autouse=True)
def clear_model_environment(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
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

def _request() -> PlanRequest:
    return PlanRequest.model_validate(
        {
            "prompt": "比较两个工作表",
            "workbook": {
                "name": "demo.xlsx",
                "capturedAt": "2026-07-24T00:00:00Z",
                "activeWorksheet": "7月",
                "worksheets": [
                    {
                        "name": "6月",
                        "usedRange": "A1:B3",
                        "rowCount": 3,
                        "columnCount": 2,
                        "headers": ["产品", "金额"],
                        "dataRows": [["A", 10], ["B", 20]],
                        "truncated": False,
                    },
                    {
                        "name": "7月",
                        "usedRange": "A1:B3",
                        "rowCount": 3,
                        "columnCount": 2,
                        "headers": ["产品", "金额"],
                        "dataRows": [["A", 12], ["B", 18]],
                        "truncated": False,
                    },
                ],
            },
        }
    )

def _score_request(
    prompt: str, rows: list[list[object]] | None = None
) -> PlanRequest:
    return PlanRequest.model_validate(
        {
            "prompt": prompt,
            "workbook": {
                "name": "scores.xlsx",
                "capturedAt": "2026-07-24T00:00:00Z",
                "activeWorksheet": "Sheet1",
                "worksheets": [
                    {
                        "name": "Sheet1",
                        "usedRange": "Sheet1!A1:B4",
                        "rowCount": 4,
                        "columnCount": 2,
                        "headers": ["人员", "得分"],
                        "dataRows": rows
                        if rows is not None
                        else [["嘟嘟嘟", 33], ["阿里", 44], ["企鹅", -3]],
                        "truncated": False,
                    }
                ],
            },
        }
    )

def _find_fields_agent_handler(
    call_log: list[dict[str, object]],
):
    """两轮：先 find_fields，再 submit_plan（写公式）。"""

    async def handler(request: httpx.Request) -> httpx.Response:
        body = json.loads(request.content)
        call_log.append(body)
        # 每次 create_plan 都会重新从第 1 轮开始，故按“本次请求内的顺序”判断。
        tool_messages = [
            message
            for message in body["messages"]
            if isinstance(message, dict) and message.get("role") == "tool"
        ]
        if not tool_messages:
            return httpx.Response(
                200,
                json={
                    "choices": [
                        {
                            "message": {
                                "content": None,
                                "tool_calls": [
                                    {
                                        "id": "field-call",
                                        "type": "function",
                                        "function": {
                                            "name": "find_fields",
                                            "arguments": json.dumps(
                                                {"query": "得分"},
                                                ensure_ascii=False,
                                            ),
                                        },
                                    }
                                ],
                            }
                        }
                    ]
                },
            )
        return httpx.Response(
            200,
            json={
                "choices": [
                    {
                        "message": {
                            "content": None,
                            "tool_calls": [
                                {
                                    "id": "submit-call",
                                    "type": "function",
                                    "function": {
                                        "name": "submit_plan",
                                        "arguments": json.dumps(
                                            {
                                                "plan": {
                                                    "id": "formula-plan",
                                                    "title": "写入得分合计公式",
                                                    "summary": "在 A5 写入得分合计公式",
                                                    "actions": [
                                                        {
                                                            "type": "writeFormulas",
                                                            "sheet": "Sheet1",
                                                            "range": "A5",
                                                            "formulas": [
                                                                ["=SUM(B2:B4)"]
                                                            ],
                                                        }
                                                    ],
                                                }
                                            },
                                            ensure_ascii=False,
                                        ),
                                    },
                                }
                            ],
                        }
                    }
                ]
            },
        )

    return handler

def _tool_names(body: dict[str, object]) -> set[str]:
    tools = body.get("tools") or []
    names: set[str] = set()
    for tool in tools:
        if isinstance(tool, dict):
            function = tool.get("function") or {}
            name = function.get("name")
            if isinstance(name, str):
                names.add(name)
    return names

def _answer_handler(call_log: list[dict[str, object]]):
    """单轮：模型直接 submit_answer（数据已在手时的典型路径）。"""

    async def handler(request: httpx.Request) -> httpx.Response:
        body = json.loads(request.content)
        call_log.append(body)
        return httpx.Response(
            200,
            json={
                "choices": [
                    {
                        "message": {
                            "content": None,
                            "tool_calls": [
                                {
                                    "id": "answer-call",
                                    "type": "function",
                                    "function": {
                                        "name": "submit_answer",
                                        "arguments": json.dumps(
                                            {"message": "得分最高的是阿里，44 分。"},
                                            ensure_ascii=False,
                                        ),
                                    },
                                }
                            ],
                        }
                    }
                ]
            },
        )

    return handler

def _data_result_request() -> PlanRequest:
    return PlanRequest.model_validate(
        {
            "prompt": "谁的得分最高",
            "workbook": {
                "name": "scores.xlsx",
                "capturedAt": "2026-07-24T00:00:00Z",
                "activeWorksheet": "Sheet1",
                "worksheets": [
                    {
                        "name": "Sheet1",
                        "usedRange": "Sheet1!A1:B4",
                        "rowCount": 4,
                        "columnCount": 2,
                        "headers": ["人员", "得分"],
                        "dataRows": [["嘟嘟嘟", 33], ["阿里", 44]],
                        "truncated": False,
                    }
                ],
            },
            "dataResults": [
                {
                    "requestId": "q1",
                    "tool": "query_table",
                    "title": "得分排序",
                    "headers": ["人员", "得分"],
                    "rows": [["阿里", 44], ["嘟嘟嘟", 33]],
                    "sourceSheets": ["Sheet1"],
                    "scannedRows": 3,
                    "complete": True,
                    "calculation": "按得分降序排列。",
                }
            ],
        }
    )

def _install_mock_model(
    monkeypatch: pytest.MonkeyPatch, handler
) -> None:
    monkeypatch.setenv("AI_BASE_URL", "http://127.0.0.1:11434/v1")
    monkeypatch.setenv("AI_MODEL", "tool-model")
    transport = httpx.MockTransport(handler)
    original_client = httpx.AsyncClient
    monkeypatch.setattr(
        httpx,
        "AsyncClient",
        lambda **kwargs: original_client(transport=transport, **kwargs),
    )

@pytest.mark.parametrize(
    "prompt",
    [
        "比较两个工作表",
        "企鹅是多少分",
        "谁的得分最高",
        "计算一下平均金额",
        "按人员去重",
        "把企鹅的得分改为-3",
        "针对得分列生成求和公式，放到 A5",
        "这份表是谁做的",
    ],
)
def test_semantic_requests_raise_without_a_model(prompt: str) -> None:
    """语义请求在无模型时必须抛稳定错误，不得静默降级为基础模式。

    旧行为是返回 kind="answer" + provider="local" + 「基础模式不支持此操作」，
    也就是把一次模型失败伪装成一次正常回答。现在必须是可识别的异常，
    由上层映射成 409 MODEL_NOT_CONFIGURED。
    """

    request = _score_request(prompt)
    request.modelId = None

    with pytest.raises(llm_config.ModelNotConfiguredError) as failure:
        asyncio.run(create_plan(request))

    message = str(failure.value)
    # 文案给出可执行的下一步：去「添加模型」填服务地址 / 模型 ID / API Key。
    assert "添加模型" in message
    assert "服务地址" in message
    assert "API Key" in message

def test_legacy_local_model_id_is_rejected_like_no_model() -> None:
    """旧客户端缓存的 "local" 等同「没有选择模型」，不能回落到规则模式。"""

    request = _score_request("比较两个工作表")
    request.modelId = "local"

    with pytest.raises(llm_config.ModelNotConfiguredError):
        asyncio.run(create_plan(request))

def test_legacy_local_selection_is_rejected_when_models_are_configured(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """已配置真实连接时仍携带旧 "local"，必须显式报「模型不可用」。

    旧行为是返回 provider="local" 的离线兜底计划，等于把「用户没有可用模型」
    伪装成「已经理解并规划好了」。现在只允许失败，并由前端把用户带回模型选择。
    注意这里刻意不装 mock transport：如果请求真的走到了模型调用，
    抛出的会是 LLMConnectError，测试就会以另一种方式失败。
    """

    monkeypatch.setenv("AI_BASE_URL", "http://127.0.0.1:11434/v1")
    monkeypatch.setenv("AI_MODEL", "configured-model")
    request = _request()
    request.modelId = "local"

    with pytest.raises(llm_config.ModelNotAvailableError) as failure:
        asyncio.run(create_plan(request))

    assert "基础模式" in str(failure.value)

def test_deleted_connection_id_reports_model_not_available() -> None:
    """指定了已删除的连接时给出 MODEL_NOT_AVAILABLE，不偷偷改用别的供应商。

    与「一条连接都没有」必须区分开：那种情况是 MODEL_NOT_CONFIGURED、
    界面提示去「添加模型」；而这里是"你选的那个已失效"，
    界面该提示重新选择，而不是让用户以为从没配置过。
    所以先落一条还在的连接，再请求一个不存在的 id。
    """

    llm_config.upsert_managed_connection(
        connection_id=None,
        label="仍在的连接",
        base_url="https://api.example.test/v1",
        model="still-there",
        api_key="secret",
        clear_api_key=False,
        supports_vision=False,
    )

    request = _score_request("比较两个工作表")
    request.modelId = "connection:model-deleted-earlier"

    with pytest.raises(llm_config.ModelNotAvailableError) as failure:
        asyncio.run(create_plan(request))

    assert "不存在" in str(failure.value) or "已删除" in str(failure.value)

def test_no_connections_at_all_reports_model_not_configured() -> None:
    """一条连接都没有时是 MODEL_NOT_CONFIGURED（去添加模型），不是 NOT_AVAILABLE。"""

    request = _score_request("比较两个工作表")
    request.modelId = None

    with pytest.raises(llm_config.ModelNotConfiguredError):
        asyncio.run(create_plan(request))

def test_request_rejects_model_outside_allowlist(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("AI_BASE_URL", "http://127.0.0.1:11434/v1")
    monkeypatch.setenv("AI_MODEL", "allowed-model")
    request = _request()
    request.modelId = "untrusted-model"

    with pytest.raises(ValueError, match="允许列表"):
        asyncio.run(create_plan(request))

def test_agent_tool_schemas_do_not_contain_local_references() -> None:
    from server.app.excel_agent import AGENT_TOOLS

    serialized = json.dumps(AGENT_TOOLS)
    assert '"$defs"' not in serialized
    assert '"$ref"' not in serialized

def test_submit_plan_schema_is_compacted_but_preserves_action_fields() -> None:
    from server.app.excel_agent import AGENT_TOOLS

    plan_tool = next(
        tool
        for tool in AGENT_TOOLS
        if tool["function"]["name"] == "submit_plan"
    )
    plan_schema = plan_tool["function"]["parameters"]["properties"]["plan"]
    plan_properties = plan_schema["properties"]

    # 业务字段 title 必须保留（曾因元数据剥离误删导致必填校验失败）。
    assert "title" in plan_properties
    # acceptanceCriteria 由验证器自动推断，不应再向模型公布这段最大子 schema。
    assert "acceptanceCriteria" not in plan_properties
    # 动作字段级结构必须完整保留，模型才知道每种动作的字段。
    assert "items" in plan_properties["actions"]
    # Pydantic 自动生成的 title 元数据应已剥离（顶层不再有裸 title 注解）。
    assert "title" not in plan_schema

    serialized = json.dumps(plan_tool, ensure_ascii=False)

    # 真约束必须保留：剥的是注解，不是模型需要遵守的边界。
    assert "maxLength" in serialized
    assert "minLength" in serialized
    # default 是服务端缺省行为，不是模型要满足的条件，应已剥离。
    assert '"default"' not in serialized

    # 精简后 token 体积应显著低于未精简的完整内联 schema。
    #
    # 这道闸是真闸，别为了让它变绿而抬上限。目前能剥的注解（title /
    # description / default）已经剥完，可压缩的余量基本见底 —— 下次新增动作
    # 把它顶破时，诚实的做法是抬上限并在这里写清「已无注解可剥」，而不是去削
    # maxLength / enum 这类模型真正需要看见的约束。
    assert len(serialized) < 26000

def test_plan_without_acceptance_criteria_still_validates() -> None:
    from server.app.models import AnalysisPlan

    plan = AnalysisPlan.model_validate(
        {
            "id": "t1",
            "title": "写入测试",
            "summary": "在 A1 写入一个值",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": "Sheet1",
                    "range": "A1",
                    "values": [[1]],
                }
            ],
        }
    )
    # 模型省略 acceptanceCriteria 时，验证器仍能自动推断出确定性验收条件。
    assert plan.acceptanceCriteria
