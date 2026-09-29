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


# 测试里反复使用的「用户选了一个模型 id」标记。
# 注意：它不是在测试环境下真实存在的连接，只是用来表达"请求里带了这个 id"。
_CONFIGURED_MODEL_ID = "connection:test-primary"

# 只读工具名取自生产定义，新增终局工具时不必同步改测试。
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


def test_direct_cell_write_works_even_when_workbook_has_no_data() -> None:
    """空工作簿上的字面地址写值：不依赖模型，也不需要任何模型配置。

    这条路径由 `_local_address_plan` 在模型解析之前收口，所以请求里
    不带 modelId 也必须成功——"没有配置模型"不该让已固化的确定性能力失效。
    """

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
    """字面公式与清空区域同样走确定性快路径，不需要模型。"""

    formula = asyncio.run(create_plan(_score_request("在C2单元格写入公式=SUM(B2:B4)")))
    clearing = asyncio.run(create_plan(_score_request("清空 Sheet1!B2:B4")))

    assert formula.kind == "plan"
    assert formula.plan.actions[0].type == "writeFormulas"
    assert formula.plan.actions[0].formulas == [["=sum(b2:b4)"]]
    assert clearing.kind == "plan"
    assert clearing.plan.actions[0].type == "clearRange"
    assert clearing.plan.actions[0].range == "B2:B4"


def test_bulk_sheet_deletion_creates_only_authorized_delete_actions() -> None:
    request = PlanRequest.model_validate(
        {
            "prompt": "清除sheet-dc1外的所有子表",
            "workbook": {
                "name": "movies.xlsx",
                "capturedAt": "2026-09-01T00:00:00Z",
                "activeWorksheet": "sheet-dc1",
                "worksheets": [
                    {
                        "name": name,
                        "usedRange": "A1:A1",
                        "rowCount": 1,
                        "columnCount": 1,
                        "headers": ["字段"],
                        "dataRows": [],
                        "truncated": False,
                    }
                    for name in ("sheet-dc1", "恐怖游轮", "痴迷")
                ],
            },
        }
    )

    response = asyncio.run(create_plan(request))

    assert response.kind == "plan"
    assert response.provider == "local"
    assert [action.type for action in response.plan.actions] == [
        "deleteWorksheet",
        "deleteWorksheet",
    ]
    assert [action.sheet for action in response.plan.actions] == ["恐怖游轮", "痴迷"]
    assert [criterion.type for criterion in response.plan.acceptanceCriteria] == [
        "worksheetMissing",
        "worksheetMissing",
    ]


def test_bulk_sheet_deletion_blocks_when_keeper_is_not_authorized() -> None:
    request = PlanRequest.model_validate(
        {
            "prompt": "删除除 sheet-dc1 外的所有工作表",
            "workbook": {
                "name": "movies.xlsx",
                "capturedAt": "2026-09-01T00:00:00Z",
                "activeWorksheet": "恐怖游轮",
                "worksheets": [
                    {
                        "name": "恐怖游轮",
                        "usedRange": "A1:A1",
                        "rowCount": 1,
                        "columnCount": 1,
                        "headers": ["字段"],
                        "dataRows": [],
                        "truncated": False,
                    }
                ],
            },
        }
    )

    response = asyncio.run(create_plan(request))

    assert response.kind == "answer"
    assert "不在当前已授权范围" in response.message
    assert "选择全部涉及的工作表" in response.message


def test_bulk_sheet_deletion_resolves_folder_display_names_by_source_sheet() -> None:
    request = PlanRequest.model_validate(
        {
            "prompt": "清除sheet-dc1外的所有子表",
            "workbook": {
                "name": "文件夹：tst",
                "capturedAt": "2026-09-01T00:00:00Z",
                "activeWorksheet": "movies.xlsx › sheet-dc1",
                "worksheets": [
                    {
                        "name": f"movies.xlsx › {name}",
                        "sourceFileId": "movies-file",
                        "sourceSheetId": f"sheet-{index}",
                        "sourceFile": "movies.xlsx",
                        "sourceSheet": name,
                        "usedRange": "A1:A1",
                        "rowCount": 1,
                        "columnCount": 1,
                        "headers": ["字段"],
                        "dataRows": [],
                        "truncated": False,
                    }
                    for index, name in enumerate(
                        ("sheet-dc1", "恐怖游轮", "痴迷")
                    )
                ],
            },
        }
    )

    response = asyncio.run(create_plan(request))

    assert response.kind == "plan"
    assert response.provider == "local"
    assert [action.sheet for action in response.plan.actions] == [
        "movies.xlsx › 恐怖游轮",
        "movies.xlsx › 痴迷",
    ]
    assert "保留「sheet-dc1」" in response.plan.summary


def test_bulk_sheet_deletion_rejects_mixed_folder_files() -> None:
    request = PlanRequest.model_validate(
        {
            "prompt": "删除除 sheet-dc1 外的所有工作表",
            "workbook": {
                "name": "文件夹：tst",
                "capturedAt": "2026-09-01T00:00:00Z",
                "activeWorksheet": "a.xlsx › sheet-dc1",
                "worksheets": [
                    {
                        "name": "a.xlsx › sheet-dc1",
                        "sourceFileId": "file-a",
                        "sourceSheetId": "sheet-a",
                        "sourceFile": "a.xlsx",
                        "sourceSheet": "sheet-dc1",
                        "usedRange": "A1:A1",
                        "rowCount": 1,
                        "columnCount": 1,
                        "headers": ["字段"],
                        "dataRows": [],
                        "truncated": False,
                    },
                    {
                        "name": "b.xlsx › 结果",
                        "sourceFileId": "file-b",
                        "sourceSheetId": "sheet-b",
                        "sourceFile": "b.xlsx",
                        "sourceSheet": "结果",
                        "usedRange": "A1:A1",
                        "rowCount": 1,
                        "columnCount": 1,
                        "headers": ["字段"],
                        "dataRows": [],
                        "truncated": False,
                    },
                ],
            },
        }
    )

    response = asyncio.run(create_plan(request))

    assert response.kind == "answer"
    assert "一次只能作用于一个文件" in response.message


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


def test_deterministic_direct_commands_still_work_without_a_model() -> None:
    """纯本地确定性指令（清空指定区域）不依赖模型，必须继续可用。

    这是「保留确定性能力」的红线：移除的是基础模式这个产品入口，
    不是这些已经固化下来的本地工具。因此这里不带任何模型也必须成功。
    """

    request = _score_request("清空 Sheet1!B2:B4")
    request.modelId = None

    response = asyncio.run(create_plan(request))

    assert response.kind == "plan"
    assert response.plan.actions[0].type == "clearRange"


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


def test_deterministic_address_command_does_not_consult_the_model(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """已经配置模型时，字面地址命令仍走本地确定性路径，不消耗模型调用。

    这里把 `selected_model_config` 换成一个「一旦被调用就断言失败」的桩：
    清空指令必须在模型解析之前就返回计划，否则说明确定性快路径被架空了。
    """

    def explode(model_id: str | None = None):  # pragma: no cover - 不应被调用
        raise AssertionError(
            f"确定性地址命令不应解析模型，却收到了 modelId={model_id!r}"
        )

    monkeypatch.setattr(planner_module, "selected_model_config", explode)

    request = _score_request("清空 Sheet1!B2:B4")
    request.modelId = _CONFIGURED_MODEL_ID

    response = asyncio.run(create_plan(request))

    assert response.kind == "plan"
    assert response.provider == "local"
    assert response.plan.actions[0].type == "clearRange"
    assert response.plan.actions[0].range == "B2:B4"


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
    assert {"submit_plan", "submit_answer"} <= _tool_names(requests[1])
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
    assert {"submit_plan", "submit_answer"} <= names
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
        "submit_plan",
    } <= names

