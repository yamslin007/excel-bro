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


# Historical loop regression only. No production route calls this helper.
async def create_plan(request, **kwargs):
    return await excel_agent.run_excel_agent(request,
        connection=llm_config.selected_model_config(request.modelId), timeout=180, **kwargs)



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
    # 断言的是"只读工具已撤下"，不是"终局工具集恰好这两个"——
    # 后者会在每次新增终局工具时假报回归。
    assert _READONLY_TOOL_NAMES.isdisjoint(_tool_names(requests[1]))
    assert {"select_plan_actions", "submit_answer"} <= _tool_names(requests[1])
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
    assert {"submit_answer", "select_plan_actions"} <= names


def test_confirmed_split_plan_submits_once_without_reading_snapshot_rows(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    call_log: list[dict[str, object]] = []

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
                                    "id": "split-submit",
                                    "type": "function",
                                    "function": {
                                        "name": "submit_plan",
                                        "arguments": json.dumps(
                                            {
                                                "plan": {
                                                    "id": "split-by-movie",
                                                    "title": "按影片拆分并按影院汇总",
                                                    "summary": "按影片建表，表内按影院统计记录数和占比",
                                                    "actions": [
                                                        {
                                                            "type": "splitGroupAggregate",
                                                            "sheet": "sheet-dc1",
                                                            "splitBy": "影片名称",
                                                            "groupBy": [
                                                                "影院名称",
                                                                "影院编码",
                                                            ],
                                                            "metrics": [
                                                                {
                                                                    "operation": "countRows",
                                                                    "outputName": "计数项:影厅",
                                                                    "ratioOutputName": "占比",
                                                                }
                                                            ],
                                                            "existingSheetPolicy": "rename",
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

    _install_mock_model(monkeypatch, handler)
    request = PlanRequest.model_validate(
        {
            "prompt": (
                "按影片名称建表，每张表内按影院名称和影院编码汇总；"
                "计数项:影厅为该影片在该影院的记录数，占比为该记录数"
                "除以该影院全部影片的记录总数。"
            ),
            "planningReadPolicy": "none",
            "workbook": {
                "name": "影院放映成绩实时报表.xlsx",
                "capturedAt": "2026-09-01T00:00:00Z",
                "activeWorksheet": "sheet-dc1",
                "worksheets": [
                    {
                        "name": "sheet-dc1",
                        "usedRange": "sheet-dc1!A1:W811",
                        "rowCount": 811,
                        "columnCount": 23,
                        "headers": [
                            "影院名称",
                            "影院编码",
                            "影片名称",
                            "影厅",
                        ],
                        "dataRows": [],
                        "truncated": True,
                    }
                ],
            },
        }
    )

    response = asyncio.run(create_plan(request))

    assert response.kind == "plan"
    assert response.plan.actions[0].type == "splitGroupAggregate"
    assert response.plan.actions[0].splitBy == "影片名称"
    assert response.plan.actions[0].groupBy == ["影院名称", "影院编码"]
    assert response.plan.actions[0].metrics[0].ratioOutputName == "占比"
    assert len(call_log) == 1
    names = _tool_names(call_log[0])
    assert _READONLY_TOOL_NAMES.isdisjoint(names)
    assert {"select_plan_actions", "submit_answer"} <= names
    user_message = call_log[0]["messages"][1]
    assert isinstance(user_message, dict)
    assert "不得调用工作簿只读工具" in str(user_message["content"])


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
        "select_plan_actions",
    } <= names
