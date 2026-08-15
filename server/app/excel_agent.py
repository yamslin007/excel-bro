from __future__ import annotations

import json
import re
from collections.abc import Awaitable, Callable
from typing import Any

from pydantic import ValidationError

from .capabilities import capability_int
from .llm import ModelConnection, OpenAICompatibleClient
from .models import (
    AnalysisPlan,
    AnswerResponse,
    AssistantResponse,
    PlanRequest,
    PlanResponse,
    ResultContext,
)

MAX_AGENT_TURNS = capability_int("agent", "maxTurns")
MAX_READ_ROWS = capability_int("agent", "maxReadRows")
MAX_READ_COLUMNS = capability_int("agent", "maxReadColumns")
MAX_OUTPUT_TOKENS = capability_int("agent", "maxOutputTokens")

AGENT_SYSTEM_PROMPT = """
你是 Excel Bro 的规划 Agent。你的工作是理解目标、按需查看工作簿，再提交一个
可预览、可验证的 Excel 操作计划；你不能直接声称已经修改工作簿。

工作原则：
- 先理解用户想得到的结果，再决定是否需要调用工作簿工具。
- 不根据样本猜测字段地址；涉及字段、数据范围、公式引用时，先调用 find_fields
  或 read_range。
- 尊重“写公式”和“写计算结果”的区别。用户要求公式时使用 writeFormulas，
  不能用服务端算出的静态值代替。
- 根据用户表达组合通用动作，不依赖固定句式。
- 删除、清空、覆盖、移动数据必须在 warnings 中写清影响。
- 写类动作（removeDuplicates / clearRange / deleteRange / copyRange）支持用
  filters 做“按值选行”：filters 的 field 是列的表头名（不是列号），operator
  通常为 equals，value 是用户点名的取值。filters 只对命中行生效，未命中行
  （含其重复）原样保留、顺序不变。
- 当用户说“把 XX 的数据去重”“删除 XX 的行”“清空 XX 的行”且 XX 是数据中某列的
  取值（而非表头名或工作表名）时，属于按值选行，应填 filters，不要反问范围，
  也不要对整表去重/删除/清空。先 read_range 看数据判断 XX 是“值”还是“字段/范围”。
- 去重带 filters 时，只对命中 filters 的行之间按 columns 判重，每个命中值保留
  第一次出现的行；未命中行全部保留。
- 工作簿快照可能截断；不能把截断样本说成完整数据。
- 当用户要求按某字段拆成多个工作表，并在每张结果表中分组计数、求和或计算
  组内占比时，优先使用 splitGroupAggregate。该动作会在 Excel 端读取源工作表
  的完整使用区域，不受快照 200 行限制；不要根据截断样本生成大量 writeTable。
- splitGroupAggregate 的 splitBy 是输出工作表的拆分字段，groupBy 是每张结果表
  内的分组字段。countRows 表示记录行数；为指标填写 ratioOutputName 时，占比的
  分母是同一 groupBy 分组在全部非空 splitBy 值中的该指标总量。
- 源工作表可能在实际表头前有标题行；splitGroupAggregate 会自动识别同时包含
  所需字段的表头行。existingSheetPolicy 默认使用 rename，避免覆盖已有工作表。
- 你只规划白名单 Excel 动作，不执行 VBA、宏、脚本或外部程序。
- 若请求包含“本地数据工具结果”，该结果由 Excel 端对用户授权范围进行确定性
  计算。优先直接使用其 headers、rows、calculation 和完整性标记，不要再次调用
  read_range 重复读取，也不要自行改变计算口径。

完成分析后优先调用 submit_plan 或 submit_answer。若模型不支持工具调用，也可以
直接返回符合 AssistantResponse JSON Schema 的 JSON。
""".strip()


def _column_number(name: str) -> int:
    result = 0
    for character in name.upper():
        result = result * 26 + ord(character) - 64
    return result


def _column_name(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _used_origin(address: str | None) -> tuple[int, int]:
    if not address:
        return 1, 1
    cell = address.rsplit("!", 1)[-1].split(":", 1)[0]
    match = re.fullmatch(r"\$?([A-Za-z]+)\$?(\d+)", cell)
    if not match:
        return 1, 1
    return _column_number(match.group(1)), int(match.group(2))


def _sheet(request: PlanRequest, name: str):
    normalized = name.strip().casefold()
    for sheet in request.workbook.worksheets:
        if sheet.name.casefold() == normalized:
            return sheet
        if sheet.sourceSheet and sheet.sourceSheet.casefold() == normalized:
            return sheet
    raise ValueError(f"未找到工作表「{name}」")


def _workbook_context(request: PlanRequest) -> dict[str, Any]:
    return {
        "name": request.workbook.name,
        "activeWorksheet": request.workbook.activeWorksheet,
        "selectedRange": request.workbook.selectedRange,
        "worksheets": [
            {
                "name": sheet.name,
                "sourceFile": sheet.sourceFile,
                "sourceSheet": sheet.sourceSheet,
                "usedRange": sheet.usedRange,
                "rowCount": sheet.rowCount,
                "columnCount": sheet.columnCount,
                "headers": sheet.headers,
                "truncated": sheet.truncated,
            }
            for sheet in request.workbook.worksheets
        ],
    }


def _read_range(request: PlanRequest, sheet_name: str, address: str) -> dict[str, Any]:
    sheet = _sheet(request, sheet_name)
    match = re.fullmatch(
        r"\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?",
        address.strip(),
    )
    if not match:
        raise ValueError("range 必须是 A1 或 A1:B10 格式")
    min_column = _column_number(match.group(1))
    min_row = int(match.group(2))
    max_column = _column_number(match.group(3) or match.group(1))
    max_row = int(match.group(4) or match.group(2))
    if max_column < min_column or max_row < min_row:
        raise ValueError("range 的结束位置不能早于开始位置")
    if (
        max_row - min_row + 1 > MAX_READ_ROWS
        or max_column - min_column + 1 > MAX_READ_COLUMNS
    ):
        raise ValueError(
            f"单次最多读取 {MAX_READ_ROWS} 行、{MAX_READ_COLUMNS} 列"
        )

    origin_column, origin_row = _used_origin(sheet.usedRange)
    values = [sheet.headers, *sheet.dataRows]
    result: list[list[object]] = []
    for row in range(min_row, max_row + 1):
        output_row: list[object] = []
        for column in range(min_column, max_column + 1):
            row_index = row - origin_row
            column_index = column - origin_column
            value = None
            if (
                0 <= row_index < len(values)
                and 0 <= column_index < len(values[row_index])
            ):
                value = values[row_index][column_index]
            output_row.append(value)
        result.append(output_row)
    return {
        "sheet": sheet.name,
        "range": address.upper().replace("$", ""),
        "values": result,
        "snapshotTruncated": sheet.truncated,
    }


def _find_fields(request: PlanRequest, query: str) -> dict[str, Any]:
    normalized = re.sub(r"[\W_]+", "", query).casefold()
    matches: list[dict[str, Any]] = []
    for sheet in request.workbook.worksheets:
        origin_column, origin_row = _used_origin(sheet.usedRange)
        for index, header in enumerate(sheet.headers):
            header_text = re.sub(r"[\W_]+", "", str(header)).casefold()
            if not header_text or (
                normalized not in header_text and header_text not in normalized
            ):
                continue
            column = _column_name(origin_column + index)
            populated_rows = [
                row_index + origin_row + 1
                for row_index, row in enumerate(sheet.dataRows)
                if index < len(row) and row[index] not in (None, "")
            ]
            data_start = min(populated_rows) if populated_rows else origin_row + 1
            data_end = max(populated_rows) if populated_rows else data_start
            full_data_end = max(origin_row + 1, origin_row + sheet.rowCount - 1)
            matches.append(
                {
                    "sheet": sheet.name,
                    "header": header,
                    "headerCell": f"{column}{origin_row}",
                    "dataRange": f"{column}{origin_row + 1}:{column}{full_data_end}",
                    "observedNonEmptyRange": (
                        f"{column}{data_start}:{column}{data_end}"
                    ),
                    "nonEmptyCountInSnapshot": len(populated_rows),
                    "snapshotTruncated": sheet.truncated,
                }
            )
    return {"query": query, "matches": matches}


def _inline_json_schema(schema: dict[str, Any]) -> dict[str, Any]:
    """Expand Pydantic $defs references for model APIs with stricter validators."""
    definitions = schema.get("$defs", {})

    def expand(node: Any) -> Any:
        if isinstance(node, list):
            return [expand(item) for item in node]
        if not isinstance(node, dict):
            return node

        reference = node.get("$ref")
        if isinstance(reference, str) and reference.startswith("#/$defs/"):
            definition_name = reference.removeprefix("#/$defs/")
            target = definitions.get(definition_name)
            if target is None:
                raise ValueError(f"JSON Schema 引用了未知定义：{reference}")
            resolved = expand(target)
            siblings = {
                key: expand(value)
                for key, value in node.items()
                if key != "$ref"
            }
            return {**resolved, **siblings}

        return {
            key: expand(value)
            for key, value in node.items()
            if key != "$defs"
        }

    return expand(schema)


def _strip_metadata(node: Any) -> Any:
    """Remove Pydantic-generated title/description annotations that bloat schema.

    模型需要的是字段结构和取值约束，不是自动生成的字段名标题。剥离后可显著
    降低每轮 Agent 请求携带的 schema token，不改变任何字段或校验规则。

    注意：只能剥离作为 JSON Schema 元数据的 title/description，不能误删
    properties 内部同名的业务字段（例如 AnalysisPlan 自身就有 title 字段）。
    因此进入 properties / $defs 等“键即字段名”的容器时不再当作元数据处理。
    """
    if isinstance(node, list):
        return [_strip_metadata(item) for item in node]
    if not isinstance(node, dict):
        return node
    result: dict[str, Any] = {}
    for key, value in node.items():
        if key in ("title", "description") and not isinstance(value, (dict, list)):
            # 标量的 title/description 是 schema 注解，安全剥离。
            continue
        if key in ("properties", "$defs", "definitions") and isinstance(value, dict):
            # 这些容器的键是字段名，值继续递归但键本身保留。
            result[key] = {
                field: _strip_metadata(field_schema)
                for field, field_schema in value.items()
            }
        else:
            result[key] = _strip_metadata(value)
    return result


def _plan_tool_schema() -> dict[str, Any]:
    """Compact submit_plan schema.

    在完整内联 schema 基础上做两处安全裁剪：
    1. 移除 acceptanceCriteria 字段——它由 AnalysisPlan 的
       add_deterministic_acceptance_criteria 验证器自动推断；模型即便仍主动
       提交，Pydantic 的 default_factory 也照常接受，因此不必在工具里向模型
       公布这段最大的子 schema。
    2. 剥离 Pydantic 自动生成的 title 元数据。
    动作(actions)的字段级结构完整保留，模型仍能看到全部动作类型及其字段。
    """
    schema = _inline_json_schema(AnalysisPlan.model_json_schema())
    properties = schema.get("properties")
    if isinstance(properties, dict):
        properties.pop("acceptanceCriteria", None)
    required = schema.get("required")
    if isinstance(required, list) and "acceptanceCriteria" in required:
        required.remove("acceptanceCriteria")
    return _strip_metadata(schema)


def _readonly_tools() -> list[dict[str, Any]]:
    return [
        {
            "type": "function",
            "function": {
                "name": "get_workbook_context",
                "description": "查看工作簿、工作表、表头、活动表和当前选区。",
                "parameters": {"type": "object", "properties": {}},
            },
        },
        {
            "type": "function",
            "function": {
                "name": "find_fields",
                "description": "按业务字段名称查找表头单元格和实际非空数据范围。",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "query": {
                            "type": "string",
                            "description": "用户提到的实际字段名称。",
                        }
                    },
                    "required": ["query"],
                    "additionalProperties": False,
                },
            },
        },
        {
            "type": "function",
            "function": {
                "name": "read_range",
                "description": "从已捕获的工作簿快照读取一个明确区域。",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "sheet": {"type": "string"},
                        "range": {"type": "string"},
                    },
                    "required": ["sheet", "range"],
                    "additionalProperties": False,
                },
            },
        },
    ]


def _terminal_tools() -> list[dict[str, Any]]:
    plan_schema = _plan_tool_schema()
    result_context_schema = _strip_metadata(
        _inline_json_schema(ResultContext.model_json_schema())
    )
    return [
        {
            "type": "function",
            "function": {
                "name": "submit_plan",
                "description": "提交需要用户预览并确认的 Excel 操作计划。",
                "parameters": {
                    "type": "object",
                    "properties": {"plan": plan_schema},
                    "required": ["plan"],
                    "additionalProperties": False,
                },
            },
        },
        {
            "type": "function",
            "function": {
                "name": "submit_answer",
                "description": "提交无需修改 Excel 的直接回答。",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "message": {"type": "string"},
                        "resultContext": {
                            "anyOf": [result_context_schema, {"type": "null"}]
                        },
                    },
                    "required": ["message"],
                    "additionalProperties": False,
                },
            },
        },
    ]


# 完整工具集：需要模型自行查数据时使用。
AGENT_TOOLS = [*_readonly_tools(), *_terminal_tools()]
# 数据已由前端确定性算好并随 dataResults 带入时，去掉三个只读工具，
# 让模型直接用现成结果提交回答或计划。减少每轮携带的 token，也避免模型
# 多绕几轮重复读取，从而缩短真实耗时、降低撞上超时中断的概率。
AGENT_TOOLS_WITH_DATA = _terminal_tools()


def _parse_content(content: str) -> AssistantResponse:
    text = content.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    start = text.find("{")
    end = text.rfind("}")
    if start < 0 or end <= start:
        raise ValueError("模型没有提交计划或回答")
    payload = json.loads(text[start : end + 1])
    if payload.get("kind") == "answer":
        return AnswerResponse.model_validate({**payload, "provider": "model"})
    if payload.get("kind") == "plan":
        return PlanResponse.model_validate({**payload, "provider": "model"})
    raise ValueError("模型返回了未知响应类型")


def _execute_tool(
    request: PlanRequest, name: str, arguments: dict[str, Any]
) -> AssistantResponse | dict[str, Any]:
    if name == "get_workbook_context":
        return _workbook_context(request)
    if name == "find_fields":
        return _find_fields(request, str(arguments.get("query", "")))
    if name == "read_range":
        return _read_range(
            request,
            str(arguments.get("sheet", "")),
            str(arguments.get("range", "")),
        )
    if name == "submit_plan":
        return PlanResponse(
            provider="model",
            plan=AnalysisPlan.model_validate(arguments.get("plan")),
        )
    if name == "submit_answer":
        return AnswerResponse(
            provider="model",
            message=str(arguments.get("message", "")),
            resultContext=arguments.get("resultContext"),
        )
    raise ValueError(f"未知工具：{name}")


READONLY_TOOLS = frozenset(
    {"get_workbook_context", "find_fields", "read_range"}
)


def _tool_cache_key(name: str, arguments: dict[str, Any]) -> str:
    return f"{name}:{json.dumps(arguments, ensure_ascii=False, sort_keys=True)}"


def _tool_step(name: str, arguments: dict[str, Any]) -> str | None:
    """把只读工具调用翻译成用户可读的步骤文案。"""
    if name == "get_workbook_context":
        return "正在查看工作簿结构"
    if name == "find_fields":
        query = str(arguments.get("query", "")).strip()
        return f"正在查找字段「{query}」" if query else "正在查找字段"
    if name == "read_range":
        sheet = str(arguments.get("sheet", "")).strip()
        cell_range = str(arguments.get("range", "")).strip()
        if sheet and cell_range:
            return f"正在读取 {sheet}!{cell_range}"
        return "正在读取数据范围"
    return None


async def run_excel_agent(
    request: PlanRequest,
    *,
    connection: ModelConnection,
    timeout: float,
    tool_cache: dict[str, Any] | None = None,
    on_event: Callable[[dict[str, Any]], Awaitable[None]] | None = None,
) -> AssistantResponse:
    cache = tool_cache

    async def emit(
        title: str, detail: str | None = None, completed: str | None = None
    ) -> None:
        if on_event is None:
            return
        event: dict[str, Any] = {"phase": "planning", "title": title}
        if detail is not None:
            event["detail"] = detail
        if completed is not None:
            event["completedStep"] = completed
        await on_event(event)
    user_text = (
        f"用户需求：\n{request.prompt}\n\n"
        f"初始工作簿上下文：\n"
        f"{json.dumps(_workbook_context(request), ensure_ascii=False)}\n\n"
        f"上一轮结构化结果：\n"
        f"{request.lastResult.model_dump_json() if request.lastResult else 'null'}\n\n"
        f"本地数据工具结果：\n"
        f"{json.dumps([result.model_dump() for result in request.dataResults], ensure_ascii=False)}\n\n"
        "请按需调用工具，不要猜测数据地址。"
    )
    user_content: str | list[dict[str, Any]] = user_text
    if request.images:
        user_content = [{"type": "text", "text": user_text}]
        user_content.extend(
            {
                "type": "image_url",
                "image_url": {
                    "url": f"data:{image.mediaType};base64,{image.data}"
                },
            }
            for image in request.images
        )

    messages: list[dict[str, Any]] = [
        {"role": "system", "content": AGENT_SYSTEM_PROMPT},
        {
            "role": "user",
            "content": user_content,
        },
    ]

    await emit("正在理解需求")
    last_step: str | None = None
    # 前端已带回确定性查询结果时，去掉只读工具，直接进入提交。
    tools = AGENT_TOOLS_WITH_DATA if request.dataResults else AGENT_TOOLS
    async with OpenAICompatibleClient(connection, timeout=timeout) as client:
        for turn_index in range(MAX_AGENT_TURNS):
            await emit(
                "正在规划操作"
                if turn_index == 0
                else f"正在规划操作（第 {turn_index + 1} 步）",
                completed=last_step,
            )
            payload = await client.chat_completions(
                messages=messages,
                tools=tools,
                tool_choice="auto",
                max_tokens=MAX_OUTPUT_TOKENS,
            )
            message = payload["choices"][0]["message"]
            tool_calls = message.get("tool_calls") or []
            if not tool_calls:
                content = message.get("content") or ""
                try:
                    return _parse_content(content)
                except (ValueError, ValidationError, json.JSONDecodeError) as error:
                    messages.append({"role": "assistant", "content": content})
                    messages.append(
                        {
                            "role": "user",
                            "content": (
                                f"你的提交没有通过结构校验：{error}。"
                                "请检查工具返回内容，并使用 submit_plan 或 "
                                "submit_answer 重新提交。"
                            ),
                        }
                    )
                    continue

            assistant_message = dict(message)
            assistant_message.setdefault("role", "assistant")
            messages.append(assistant_message)
            for tool_call in tool_calls:
                function = tool_call.get("function") or {}
                name = str(function.get("name", ""))
                try:
                    arguments = json.loads(function.get("arguments") or "{}")
                    step_text = _tool_step(name, arguments)
                    if step_text is not None:
                        await emit(step_text, completed=last_step)
                        last_step = step_text
                    cache_key: str | None = None
                    if cache is not None and name in READONLY_TOOLS:
                        cache_key = _tool_cache_key(name, arguments)
                    if cache_key is not None and cache_key in cache:
                        tool_content = cache[cache_key]
                    else:
                        result = _execute_tool(request, name, arguments)
                        if isinstance(result, (PlanResponse, AnswerResponse)):
                            return result
                        tool_content = json.dumps(result, ensure_ascii=False)
                        if cache_key is not None:
                            cache[cache_key] = tool_content
                except (ValueError, ValidationError, json.JSONDecodeError) as error:
                    tool_content = json.dumps(
                        {"error": str(error)}, ensure_ascii=False
                    )
                messages.append(
                    {
                        "role": "tool",
                        "tool_call_id": tool_call.get("id", ""),
                        "name": name,
                        "content": tool_content,
                    }
                )
    raise ValueError("Agent 在限定轮次内没有提交计划或回答")
