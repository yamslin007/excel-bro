from __future__ import annotations

import asyncio
import json
from pathlib import Path

import httpx
import pytest

import server.app.llm.config as llm_config
from server.app.models import PlanRequest
from server.app.planner import create_plan


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


def test_direct_cell_write_works_even_when_workbook_has_no_data() -> None:
    request = PlanRequest.model_validate(
        {
            "prompt": "在A5单元格填入数字4",
            "workbook": {
                "name": "empty.xlsx",
                "capturedAt": "2026-07-24T00:00:00Z",
                "activeWorksheet": "Sheet1",
                "worksheets": [
                    {
                        "name": "Sheet1",
                        "usedRange": None,
                        "rowCount": 0,
                        "columnCount": 0,
                        "headers": [],
                        "dataRows": [],
                        "truncated": False,
                    }
                ],
            },
        }
    )

    response = asyncio.run(create_plan(request))

    assert response.kind == "plan"
    assert response.plan.actions[0].type == "writeValues"
    assert response.plan.actions[0].sheet == "Sheet1"
    assert response.plan.actions[0].range == "A5"
    assert response.plan.actions[0].values == [[4]]


def test_direct_formula_and_clear_commands_create_focused_plans() -> None:
    formula = asyncio.run(create_plan(_score_request("在C2单元格写入公式=SUM(B2:B4)")))
    clearing = asyncio.run(create_plan(_score_request("清空 Sheet1!B2:B4")))

    assert formula.kind == "plan"
    assert formula.plan.actions[0].type == "writeFormulas"
    assert formula.plan.actions[0].formulas == [["=sum(b2:b4)"]]
    assert clearing.kind == "plan"
    assert clearing.plan.actions[0].type == "clearRange"
    assert clearing.plan.actions[0].range == "B2:B4"


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
def test_semantic_requests_decline_without_a_model(prompt: str) -> None:
    response = asyncio.run(create_plan(_score_request(prompt)))

    assert response.kind == "answer"
    assert response.provider == "local"
    assert "基础模式不支持此操作" in response.message
    assert "/help" in response.message
    assert "配置 AI 模型" in response.message


def test_keyless_local_gateway_can_be_used(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("AI_BASE_URL", "http://127.0.0.1:11434/v1/")
    monkeypatch.setenv("AI_MODEL", "local-model")
    captured_headers: dict[str, str] = {}

    async def handler(request: httpx.Request) -> httpx.Response:
        captured_headers.update(dict(request.headers))
        return httpx.Response(
            200,
            json={
                "choices": [
                    {
                        "message": {
                            "content": json.dumps(
                                {"kind": "answer", "message": "本地模型回答"}
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

    response = asyncio.run(create_plan(_request()))

    assert response.kind == "answer"
    assert response.provider == "model"
    assert "authorization" not in captured_headers


def test_request_can_select_an_allowlisted_model(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("AI_BASE_URL", "http://127.0.0.1:11434/v1")
    monkeypatch.setenv("AI_MODEL", "default-model")
    monkeypatch.setenv("AI_MODELS", "fast-model,reasoning-model")
    selected_models: list[str] = []

    async def handler(request: httpx.Request) -> httpx.Response:
        body = json.loads(request.content)
        selected_models.append(body["model"])
        return httpx.Response(
            200,
            json={
                "choices": [
                    {
                        "message": {
                            "content": json.dumps(
                                {"kind": "answer", "message": "已切换模型"}
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
    request = _request()
    request.modelId = "reasoning-model"

    response = asyncio.run(create_plan(request))

    assert response.provider == "model"
    assert selected_models == ["reasoning-model"]


def test_kimi_k3_receives_images_as_multimodal_content(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("AI_BASE_URL", "https://api.moonshot.cn/v1")
    monkeypatch.setenv("AI_MODEL", "kimi-k3")
    monkeypatch.setenv("AI_VISION_MODELS", "kimi-k3")
    captured_messages: list[dict[str, object]] = []

    async def handler(request: httpx.Request) -> httpx.Response:
        body = json.loads(request.content)
        captured_messages.extend(body["messages"])
        return httpx.Response(
            200,
            json={
                "choices": [
                    {
                        "message": {
                            "content": json.dumps(
                                {"kind": "answer", "message": "已看懂截图"}
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
    request = PlanRequest.model_validate(
        {
            **_request().model_dump(),
            "images": [
                {
                    "name": "错误截图.png",
                    "mediaType": "image/png",
                    "data": "aGVsbG8=",
                }
            ],
        }
    )

    response = asyncio.run(create_plan(request))

    assert response.kind == "answer"
    user_content = captured_messages[1]["content"]
    assert isinstance(user_content, list)
    assert user_content[0]["type"] == "text"
    assert user_content[1] == {
        "type": "image_url",
        "image_url": {"url": "data:image/png;base64,aGVsbG8="},
    }


def test_images_require_a_vision_model(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("AI_BASE_URL", "https://api.example.test/v1")
    monkeypatch.setenv("AI_MODEL", "text-only-model")
    request = PlanRequest.model_validate(
        {
            **_request().model_dump(),
            "images": [
                {
                    "name": "截图.png",
                    "mediaType": "image/png",
                    "data": "aGVsbG8=",
                }
            ],
        }
    )

    with pytest.raises(ValueError, match="AI_VISION_MODELS"):
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

    # 精简后 token 体积应显著低于未精简的完整内联 schema。
    serialized = json.dumps(plan_tool, ensure_ascii=False)
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


def test_request_can_force_local_fallback(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("AI_BASE_URL", "http://127.0.0.1:11434/v1")
    monkeypatch.setenv("AI_MODEL", "configured-model")
    request = _request()
    request.modelId = "local"

    response = asyncio.run(create_plan(request))

    assert response.provider == "local"


def test_model_agent_inspects_fields_before_submitting_formula_plan(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("AI_BASE_URL", "http://127.0.0.1:11434/v1")
    monkeypatch.setenv("AI_MODEL", "tool-model")
    requests: list[dict[str, object]] = []

    async def handler(request: httpx.Request) -> httpx.Response:
        body = json.loads(request.content)
        requests.append(body)
        if len(requests) == 1:
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

    transport = httpx.MockTransport(handler)
    original_client = httpx.AsyncClient
    monkeypatch.setattr(
        httpx,
        "AsyncClient",
        lambda **kwargs: original_client(transport=transport, **kwargs),
    )

    request = _score_request("请在A5里面写一个计算总和的公式，针对得分的")
    response = asyncio.run(create_plan(request))

    assert response.kind == "plan"
    assert response.provider == "model"
    assert response.plan.actions[0].type == "writeFormulas"
    assert len(requests) == 2
    second_messages = requests[1]["messages"]
    assert isinstance(second_messages, list)
    tool_message = next(
        message for message in second_messages if message["role"] == "tool"
    )
    tool_result = json.loads(tool_message["content"])
    assert tool_result["matches"][0]["dataRange"] == "B2:B4"


def test_openai_environment_names_are_supported(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("OPENAI_BASE_URL", "http://127.0.0.1:1234/v1")
    monkeypatch.setenv("OPENAI_MODEL", "compatible-model")
    monkeypatch.setenv("OPENAI_API_KEY", "test-key")
    captured_headers: dict[str, str] = {}

    async def handler(request: httpx.Request) -> httpx.Response:
        captured_headers.update(dict(request.headers))
        return httpx.Response(
            200,
            json={
                "choices": [
                    {
                        "message": {
                            "content": json.dumps(
                                {"kind": "answer", "message": "兼容变量可用"}
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

    response = asyncio.run(create_plan(_request()))

    assert response.provider == "model"
    assert captured_headers["authorization"] == "Bearer test-key"


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


def test_tool_cache_reuses_readonly_results_across_retries(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("AI_BASE_URL", "http://127.0.0.1:11434/v1")
    monkeypatch.setenv("AI_MODEL", "tool-model")

    call_log: list[dict[str, object]] = []
    transport = httpx.MockTransport(_find_fields_agent_handler(call_log))
    original_client = httpx.AsyncClient
    monkeypatch.setattr(
        httpx,
        "AsyncClient",
        lambda **kwargs: original_client(transport=transport, **kwargs),
    )

    import server.app.excel_agent as excel_agent

    find_calls = {"count": 0}
    real_find_fields = excel_agent._find_fields

    def spy_find_fields(request, query):
        find_calls["count"] += 1
        return real_find_fields(request, query)

    monkeypatch.setattr(excel_agent, "_find_fields", spy_find_fields)

    request = _score_request("请在A5里面写一个计算总和的公式，针对得分的")
    tool_cache: dict[str, object] = {}

    first = asyncio.run(create_plan(request, tool_cache=tool_cache))
    second = asyncio.run(create_plan(request, tool_cache=tool_cache))

    assert first.kind == "plan"
    assert second.kind == "plan"
    # find_fields 只真正执行一次；第二轮 create_plan 命中缓存。
    assert find_calls["count"] == 1
    assert any("find_fields" in json.dumps(body) for body in call_log)


def test_tool_cache_isolates_distinct_arguments(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import server.app.excel_agent as excel_agent

    request = _score_request("测试")
    cache: dict[str, object] = {}
    key_a = excel_agent._tool_cache_key("find_fields", {"query": "得分"})
    key_b = excel_agent._tool_cache_key("find_fields", {"query": "人员"})

    assert key_a != key_b


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


def test_agent_drops_readonly_tools_when_data_results_present(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    call_log: list[dict[str, object]] = []
    _install_mock_model(monkeypatch, _answer_handler(call_log))

    response = asyncio.run(create_plan(_data_result_request()))

    assert response.kind == "answer"
    assert call_log, "模型应至少被调用一次"
    names = _tool_names(call_log[0])
    assert "get_workbook_context" not in names
    assert "find_fields" not in names
    assert "read_range" not in names
    assert {"submit_answer", "submit_plan"} <= names


def test_agent_keeps_readonly_tools_without_data_results(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    call_log: list[dict[str, object]] = []
    _install_mock_model(monkeypatch, _answer_handler(call_log))

    response = asyncio.run(create_plan(_score_request("谁的得分最高")))

    assert response.kind == "answer"
    assert call_log
    names = _tool_names(call_log[0])
    assert {
        "get_workbook_context",
        "find_fields",
        "read_range",
        "submit_answer",
        "submit_plan",
    } <= names

