from __future__ import annotations

import json
import re
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


def _tools() -> list[dict[str, Any]]:
    plan_schema = _inline_json_schema(AnalysisPlan.model_json_schema())
    result_context_schema = _inline_json_schema(ResultContext.model_json_schema())
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


AGENT_TOOLS = _tools()


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


async def run_excel_agent(
    request: PlanRequest,
    *,
    connection: ModelConnection,
    timeout: float,
) -> AssistantResponse:
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

    async with OpenAICompatibleClient(connection, timeout=timeout) as client:
        for _ in range(MAX_AGENT_TURNS):
            payload = await client.chat_completions(
                messages=messages,
                tools=AGENT_TOOLS,
                tool_choice="auto",
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
                    result = _execute_tool(request, name, arguments)
                    if isinstance(result, (PlanResponse, AnswerResponse)):
                        return result
                    tool_content = json.dumps(result, ensure_ascii=False)
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
