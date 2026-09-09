from __future__ import annotations

import copy
import json
import logging
import re
import time
from collections.abc import Awaitable, Callable
from typing import Any

from pydantic import ValidationError

from .agent_loop_protection import (
    LoopConfig,
    LoopState,
    detect_action_conflict,
    detect_path_loop,
    get_warning_prompt,
    safe_degradation,
    should_stop,
    update_loop_state,
)
from .capabilities import capability_float, capability_int
from .llm import ModelConnection, OpenAICompatibleClient
from .models import (
    AnalysisPlan,
    AnswerResponse,
    AssistantResponse,
    PlanRequest,
    PlanResponse,
    ResultContext,
)

MAX_READ_ROWS = capability_int("agent", "maxReadRows")
MAX_READ_COLUMNS = capability_int("agent", "maxReadColumns")
MAX_OUTPUT_TOKENS = capability_int("agent", "maxOutputTokens")
REPAIR_MAX_OUTPUT_TOKENS = capability_int("agent", "repairMaxOutputTokens")

logger = logging.getLogger(__name__)


def _default_loop_config() -> LoopConfig:
    """从共享能力配置构造生产环境 Agent 循环限额。"""
    return LoopConfig(
        max_turns=capability_int("agent", "maxTurns"),
        max_total_tokens=capability_int("agent", "maxTotalTokens"),
        max_wall_time_seconds=capability_float(
            "taskRecovery", "planningHardWallTimeSeconds"
        ),
        max_consecutive_errors=capability_int(
            "agent", "maxConsecutiveErrors"
        ),
        max_cache_hits_per_query=capability_int(
            "agent", "maxCacheHitsPerQuery"
        ),
        max_tool_path_loop_length=capability_int(
            "agent", "maxToolPathLoopLength"
        ),
        warning_turn=capability_int("agent", "warningTurn"),
    )

AGENT_SYSTEM_PROMPT = """
你是 Excel Bro 的规划 Agent。你的工作是理解目标、按需查看工作簿，再提交一个
可预览、可验证的 Excel 操作计划；你不能直接声称已经修改工作簿。

工作原则：
- 先理解用户想得到的结果，再决定是否需要调用工作簿工具。
- 初始用户消息中的工作簿上下文已经包含已授权工作表的准确表头。所需字段与表头
  完全匹配，且动作按表头寻址（如 computeValues、splitGroupAggregate）时，直接
  提交计划，不要再调用 get_workbook_context、find_fields 或 read_range。只有字段
  不在表头中、名称存在歧义，或确实需要定位单元格地址/判断取值时才使用只读工具。
- 不根据样本猜测字段地址；涉及单元格范围、公式引用，且上下文没有足够地址信息时，
  先调用 find_fields 或 read_range。
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
- splitBy 字段会由执行器自动写入每张结果表，绝不能为了让该字段显示在结果中
  又把它放进 groupBy。groupBy 只填写除 splitBy 之外、用于确定表内每一行的字段。
- 输出工作表名称来自哪个字段，该字段就是 splitBy；每张结果表内每行由哪些字段
  确定，这些字段就是 groupBy。splitGroupAggregate 在执行时读取完整使用区域，
  不要调用 read_range 抽样计数，也不要先计算各组结果。
- 源工作表可能在实际表头前有标题行；splitGroupAggregate 会自动识别同时包含
  所需字段的表头行。existingSheetPolicy 默认使用 rename，避免覆盖已有工作表。
- 有默认值的列表字段没有内容时应省略或提交空列表，不要提交 null。
- 你只规划白名单 Excel 动作，不执行 VBA、宏、脚本或外部程序。

写入方式的优先级（从高到低，只有前一档表达不了才降级）：
1. writeFormulas —— 公式引用单元格，源数据一变 Excel 自己重算。表达得了就用它。
2. computeValues / splitGroupAggregate —— 执行时求值的动作：动作体里**只带查询
   配置，不带结果**，Excel 端在执行那一刻读取源工作表的实时数据算出来。需要
   筛选、分组、聚合、排序、取前 N 这类查询时用 computeValues。
3. writeValues / writeTable —— 只在写入的内容确实是**固定值**时使用：用户点名的
   常量、表头文字、说明文字。

不要把算出来的结果塞进 writeValues.values。那是一张一次性快照：源数据改一格
结果就过期了，而计划本身无法重算，用户只能重新提问。同样的需求用
computeValues 表达，改完数据直接重跑就是新结果。

- 若请求包含“本地数据工具结果”，该结果由 Excel 端对用户授权范围进行确定性
  计算，可以直接引用它的 headers、rows、calculation 和完整性标记来**回答问题**
  （submit_answer），不必再调 read_range 重复读取，也不要自行改变计算口径。
  但要**写回工作表**时，不要把这些 rows 当静态值塞进 writeValues —— 改用
  computeValues 复述同一套查询条件，让 Excel 在执行时重算。

完成分析后优先调用 submit_plan 或 submit_answer。若模型不支持工具调用，也可以
直接返回符合 AssistantResponse JSON Schema 的 JSON。
""".strip()


class AgentBudgetPause(RuntimeError):
    """A complete planning boundary was saved and now awaits user budget approval."""

    def __init__(self, checkpoint: dict[str, Any], approved_seconds: float) -> None:
        super().__init__("本轮处理时间已用完，任务现场已保存。")
        self.checkpoint = checkpoint
        self.approved_seconds = approved_seconds


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
    """Remove Pydantic-generated title/description/default annotations.

    模型需要的是字段结构和取值约束，不是自动生成的字段名标题。剥离后可显著
    降低每轮 Agent 请求携带的 schema token，不改变任何字段或校验规则。

    default 一并剥离：它描述的是**服务端在字段缺失时的行为**，不是模型必须
    满足的条件。剥离后模型能生成的合法计划集合一个都没少 —— maxLength /
    minLength / enum / anyOf 这些真约束全部保留。这条界线要守住：削注解可以，
    削约束不行，后者会让模型看不见真实边界从而反复提交非法计划。

    注意：只能剥离作为 JSON Schema 元数据的标量注解，不能误删 properties
    内部同名的业务字段（例如 AnalysisPlan 自身就有 title 字段）。因此进入
    properties / $defs 等“键即字段名”的容器时不再当作元数据处理。
    """
    if isinstance(node, list):
        return [_strip_metadata(item) for item in node]
    if not isinstance(node, dict):
        return node
    result: dict[str, Any] = {}
    for key, value in node.items():
        if key in ("title", "description", "default") and not isinstance(
            value, (dict, list)
        ):
            # 标量的 title/description/default 是 schema 注解，安全剥离。
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
                "description": (
                    "提交 Excel 操作计划。写入优先级：writeFormulas > "
                    "computeValues（执行时按实时数据求值）> writeValues（仅固定值）。"
                    "splitGroupAggregate 的 splitBy 会自动成为结果列，groupBy 不得"
                    "重复包含 splitBy。无内容的默认列表字段请省略或传空列表，不要"
                    "传 null。"
                ),
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
PLAN_ONLY_TOOLS = [
    tool
    for tool in AGENT_TOOLS_WITH_DATA
    if tool.get("function", {}).get("name") == "submit_plan"
]


def _normalize_submitted_plan(value: Any) -> Any:
    """修正模型计划中不改变语义、且可确定消除的协议冗余。

    splitGroupAggregate 的 splitBy 本身就是输出列。模型偶尔会为了“保留该字段”
    又把它放进 groupBy；这不会增加任何信息，却会让合法计划在模型层来回重试。
    这里只移除与 splitBy 完全同名的 groupBy 项，其他无效配置仍交给 Pydantic
    拒绝，避免本地层猜测业务语义。
    """
    if not isinstance(value, dict):
        return value

    normalized = copy.deepcopy(value)
    # 模型常把“可省略且默认由服务端创建的集合”输出为 null。对所有
    # AnalysisPlan 顶层 default_factory 字段统一按“省略”处理，不绑定具体字段名，
    # 也不改变任何业务语义；必填字段和没有默认工厂的字段仍照常严格校验。
    for field_name, field_info in AnalysisPlan.model_fields.items():
        if (
            normalized.get(field_name) is None
            and field_info.default_factory is not None
        ):
            normalized.pop(field_name, None)

    actions = normalized.get("actions")
    if not isinstance(actions, list):
        return normalized

    removed_fields: list[str] = []
    for action in actions:
        if not isinstance(action, dict) or action.get("type") != "splitGroupAggregate":
            continue
        split_by = action.get("splitBy")
        group_by = action.get("groupBy")
        if not isinstance(split_by, str) or not isinstance(group_by, list):
            continue
        normalized_split = split_by.strip()
        filtered_group_by = [
            field
            for field in group_by
            if not (
                isinstance(field, str)
                and field.strip() == normalized_split
            )
        ]
        if len(filtered_group_by) == len(group_by):
            continue
        action["groupBy"] = filtered_group_by
        if normalized_split and normalized_split not in removed_fields:
            removed_fields.append(normalized_split)

    if not removed_fields:
        return normalized

    warnings = normalized.get("warnings")
    if warnings is None:
        warnings = []
        normalized["warnings"] = warnings
    if isinstance(warnings, list) and len(warnings) < 20:
        fields = "、".join(f"「{field}」" for field in removed_fields)
        warning = (
            f"已从 groupBy 移除与 splitBy 重复的字段 {fields}；"
            "拆分字段会自动保留在结果列中。"
        )
        if warning not in warnings:
            warnings.append(warning)
    return normalized


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
        return PlanResponse.model_validate(
            {
                **payload,
                "provider": "model",
                "plan": _normalize_submitted_plan(payload.get("plan")),
            }
        )
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
            plan=AnalysisPlan.model_validate(
                _normalize_submitted_plan(arguments.get("plan"))
            ),
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


def _safe_progress_text(value: object, limit: int = 1000) -> str:
    """把可公开的确定性摘要限制在进度协议允许的长度内。"""
    text = re.sub(r"\s+", " ", str(value)).strip()
    if len(text) <= limit:
        return text
    return text[: max(0, limit - 1)] + "…"


def _progress_fact(
    label: str, value: object, *, style: str = "text"
) -> dict[str, str] | None:
    safe_label = _safe_progress_text(label, 40)
    safe_value = _safe_progress_text(value, 500)
    if not safe_label or not safe_value:
        return None
    return {
        "label": safe_label,
        "value": safe_value,
        "style": "code" if style == "code" else "text",
    }


def _tool_argument_facts(
    name: str, arguments: dict[str, Any]
) -> list[dict[str, str]]:
    facts: list[dict[str, str] | None] = []
    if name == "find_fields":
        facts.append(_progress_fact("查询", arguments.get("query", "")))
    elif name == "read_range":
        facts.extend(
            [
                _progress_fact("工作表", arguments.get("sheet", "")),
                _progress_fact("范围", arguments.get("range", ""), style="code"),
            ]
        )
    elif name == "submit_plan":
        plan = arguments.get("plan")
        if isinstance(plan, dict):
            actions = plan.get("actions")
            action_items = actions if isinstance(actions, list) else []
            facts.append(_progress_fact("动作数", len(action_items)))
    return [fact for fact in facts if fact is not None]


def _readonly_tool_facts(
    name: str,
    arguments: dict[str, Any],
    result: object,
    *,
    cached: bool,
) -> list[dict[str, str]]:
    payload = result if isinstance(result, dict) else {}
    facts: list[dict[str, str] | None] = [
        _progress_fact("读取方式", "复用本轮结果" if cached else "实时读取")
    ]
    if name == "get_workbook_context":
        worksheets = payload.get("worksheets")
        worksheet_items = worksheets if isinstance(worksheets, list) else []
        field_count = sum(
            len(item.get("headers", []))
            for item in worksheet_items
            if isinstance(item, dict) and isinstance(item.get("headers"), list)
        )
        facts.extend(
            [
                _progress_fact("工作表数", len(worksheet_items)),
                _progress_fact("字段数", field_count),
            ]
        )
    elif name == "find_fields":
        matches = payload.get("matches")
        match_items = matches if isinstance(matches, list) else []
        facts.extend(
            [
                _progress_fact("查询", arguments.get("query", "")),
                _progress_fact("匹配数", len(match_items)),
            ]
        )
    elif name == "read_range":
        values = payload.get("values")
        rows = values if isinstance(values, list) else []
        column_count = max(
            (len(row) for row in rows if isinstance(row, list)), default=0
        )
        facts.extend(
            [
                _progress_fact(
                    "工作表", payload.get("sheet", arguments.get("sheet", ""))
                ),
                _progress_fact(
                    "范围",
                    payload.get("range", arguments.get("range", "")),
                    style="code",
                ),
                _progress_fact("结果", f"{len(rows)} 行 × {column_count} 列"),
                _progress_fact(
                    "快照",
                    "已截断"
                    if payload.get("snapshotTruncated") is True
                    else "完整返回",
                ),
            ]
        )
    return [fact for fact in facts if fact is not None]


def _plan_progress_facts(plan: AnalysisPlan) -> list[dict[str, str]]:
    action_types = [action.type for action in plan.actions]
    facts: list[dict[str, str] | None] = [
        _progress_fact("动作数", len(action_types)),
        _progress_fact("动作", "、".join(action_types), style="code"),
    ]
    return [fact for fact in facts if fact is not None]


def _readonly_tool_completion(
    name: str,
    arguments: dict[str, Any],
    result: object,
    *,
    cached: bool,
) -> tuple[str, str, dict[str, Any], dict[str, str] | None]:
    """只从工具元数据生成用户可见摘要，绝不回传 read_range.values。"""
    suffix = "（复用本轮已读取结果）" if cached else ""
    payload = result if isinstance(result, dict) else {}
    target: dict[str, str] | None = None
    if name == "get_workbook_context":
        worksheets = payload.get("worksheets")
        worksheet_items = worksheets if isinstance(worksheets, list) else []
        field_count = sum(
            len(item.get("headers", []))
            for item in worksheet_items
            if isinstance(item, dict) and isinstance(item.get("headers"), list)
        )
        truncated_count = sum(
            1
            for item in worksheet_items
            if isinstance(item, dict) and item.get("truncated") is True
        )
        summary = (
            f"已读取 {len(worksheet_items)} 张已授权工作表的结构，共 {field_count} 个字段"
        )
        if truncated_count:
            summary += f"；{truncated_count} 张表的数据快照已截断"
        summary += suffix
        return (
            "已查看工作簿结构",
            summary,
            {
                "kind": "evidence",
                "summary": summary,
                "basis": "工作簿结构工具返回的工作表数、字段数与快照状态",
            },
            None,
        )
    if name == "find_fields":
        matches = payload.get("matches")
        match_items = matches if isinstance(matches, list) else []
        query = _safe_progress_text(arguments.get("query", ""), 120)
        locations = [
            f"{item.get('sheet')}!{item.get('headerCell')}"
            for item in match_items[:5]
            if isinstance(item, dict) and item.get("sheet") and item.get("headerCell")
        ]
        summary = f"字段「{query}」找到 {len(match_items)} 处匹配{suffix}"
        basis = "、".join(locations) if locations else "没有发现匹配位置"
        return (
            "已完成字段定位",
            summary,
            {"kind": "evidence", "summary": summary, "basis": basis},
            None,
        )
    if name == "read_range":
        values = payload.get("values")
        rows = values if isinstance(values, list) else []
        row_count = len(rows)
        column_count = max(
            (len(row) for row in rows if isinstance(row, list)), default=0
        )
        sheet = _safe_progress_text(payload.get("sheet", arguments.get("sheet", "")), 500)
        cell_range = _safe_progress_text(
            payload.get("range", arguments.get("range", "")), 80
        )
        truncated = payload.get("snapshotTruncated") is True
        summary = f"已读取 {sheet}!{cell_range} 的 {row_count} 行 × {column_count} 列"
        if truncated:
            summary += "；结果来自截断快照"
        summary += suffix
        target = {"sheet": sheet, "range": cell_range}
        return (
            "已读取数据范围",
            summary,
            {
                "kind": "evidence",
                "summary": summary,
                "basis": "只展示范围和尺寸，不展示单元格原始内容",
            },
            target,
        )
    summary = f"受控工具 {name} 已完成{suffix}"
    return (
        "工具调用已完成",
        summary,
        {"kind": "evidence", "summary": summary},
        None,
    )


def _append_warning_to_last_user_message(
    messages: list[dict[str, Any]], warning: str
) -> None:
    """把收尾提示追加到最后一条用户消息，兼容文本和多模态内容。"""
    for message in reversed(messages):
        if message.get("role") != "user":
            continue
        content = message.get("content")
        if isinstance(content, str):
            message["content"] = content + warning
            return
        if isinstance(content, list):
            for item in reversed(content):
                if isinstance(item, dict) and item.get("type") == "text":
                    item["text"] = str(item.get("text", "")) + warning
                    return
            content.append({"type": "text", "text": warning})
        return


def _inject_tool_warning(tool_content: str, warning: str) -> str:
    """向 JSON 工具结果附加警告，并保留同一结果中已有的警告。"""
    try:
        result = json.loads(tool_content)
    except (TypeError, json.JSONDecodeError):
        result = {"result": tool_content}
    if not isinstance(result, dict):
        result = {"result": result}

    existing_warning = result.get("_warning")
    if isinstance(existing_warning, str) and existing_warning:
        if warning not in existing_warning:
            result["_warning"] = f"{existing_warning}\n{warning}"
    else:
        result["_warning"] = warning
    return json.dumps(result, ensure_ascii=False)


def _contains_snapshot_truncation(value: Any) -> bool:
    """递归识别工具结果中的快照截断标记。"""
    if isinstance(value, dict):
        if value.get("snapshotTruncated") is True or value.get("truncated") is True:
            return True
        return any(_contains_snapshot_truncation(item) for item in value.values())
    if isinstance(value, list):
        return any(_contains_snapshot_truncation(item) for item in value)
    return False


async def run_excel_agent(
    request: PlanRequest,
    *,
    connection: ModelConnection,
    timeout: float,
    tool_cache: dict[str, Any] | None = None,
    on_event: Callable[[dict[str, Any]], Awaitable[None]] | None = None,
    on_checkpoint: Callable[[dict[str, Any]], Awaitable[None]] | None = None,
    resume_checkpoint: dict[str, Any] | None = None,
    approved_wall_time_seconds: float | None = None,
    config: LoopConfig | None = None,
) -> AssistantResponse:
    """运行受预算和循环检测保护的 Excel 规划 Agent。"""
    cache = tool_cache
    loop_config = config or _default_loop_config()
    if resume_checkpoint is not None:
        if (
            resume_checkpoint.get("schemaVersion") not in {1, 2}
            or resume_checkpoint.get("kind") != "planning"
        ):
            raise ValueError("规划检查点版本或类型不受支持")
        raw_loop_state = resume_checkpoint.get("loopState")
        if not isinstance(raw_loop_state, dict):
            raise ValueError("规划检查点缺少循环状态")
        state = LoopState.from_checkpoint(raw_loop_state)
    else:
        state = LoopState()
    elapsed_before_run = state.elapsed_time
    start_time = time.monotonic()

    async def emit(
        step_id: str,
        title: str,
        *,
        status: str = "progress",
        summary: str | None = None,
        detail: str | None = None,
        trace: dict[str, Any] | None = None,
        target: dict[str, str] | None = None,
        tool_name: str | None = None,
        facts: list[dict[str, str]] | None = None,
    ) -> None:
        if on_event is None:
            return
        event: dict[str, Any] = {
            "phase": "planning",
            "stepId": step_id,
            "status": status,
            "title": title,
        }
        if summary is not None:
            event["summary"] = _safe_progress_text(summary)
        if detail is not None:
            event["detail"] = _safe_progress_text(detail, 2000)
        if trace is not None:
            event["trace"] = trace
        if target is not None:
            event["target"] = target
        if tool_name is not None:
            event["toolName"] = _safe_progress_text(tool_name, 100)
        if facts:
            event["facts"] = facts[:12]
        await on_event(event)
    user_text = (
        f"用户需求：\n{request.prompt}\n\n"
        f"初始工作簿上下文：\n"
        f"{json.dumps(_workbook_context(request), ensure_ascii=False)}\n\n"
        f"上一轮结构化结果：\n"
        f"{request.lastResult.model_dump_json() if request.lastResult else 'null'}\n\n"
        f"本地数据工具结果：\n"
        f"{json.dumps([result.model_dump() for result in request.dataResults], ensure_ascii=False)}\n\n"
        + (
            "用户已经从结构化确认选项中确定了最终口径；现有准确表头足以生成计划。"
            "不得调用工作簿只读工具，直接提交计划。"
            if request.planningReadPolicy == "none"
            else "请按需调用工具，不要猜测数据地址。"
        )
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

    if resume_checkpoint is not None:
        raw_messages = resume_checkpoint.get("messages")
        if not isinstance(raw_messages, list) or not all(
            isinstance(item, dict) for item in raw_messages
        ):
            raise ValueError("规划检查点缺少完整消息现场")
        messages = [dict(item) for item in raw_messages]
        if cache is not None and isinstance(resume_checkpoint.get("toolCache"), dict):
            cache.clear()
            cache.update(dict(resume_checkpoint["toolCache"]))
    else:
        messages = [
            {"role": "system", "content": AGENT_SYSTEM_PROMPT},
            {
                "role": "user",
                "content": user_content,
            },
        ]

    await emit("planning.understand", "正在理解需求", status="started")
    # 前端已带回确定性查询结果，或用户已经从结构化选项确认了最终口径时，
    # 去掉只读工具，直接进入提交。运行时聚合动作会在执行阶段读取完整数据。
    initial_tools = (
        AGENT_TOOLS_WITH_DATA
        if request.dataResults or request.planningReadPolicy == "none"
        else AGENT_TOOLS
    )
    tool_mode = (
        str(resume_checkpoint.get("toolMode"))
        if resume_checkpoint is not None
        else ("terminal" if initial_tools is AGENT_TOOLS_WITH_DATA else "full")
    )
    tools = {
        "full": AGENT_TOOLS,
        "terminal": AGENT_TOOLS_WITH_DATA,
        "plan_only": PLAN_ONLY_TOOLS,
    }.get(tool_mode)
    if tools is None:
        raise ValueError("规划检查点包含未知工具集合")
    repairing_submission = bool(
        resume_checkpoint.get("repairingSubmission")
        if resume_checkpoint is not None
        else False
    )
    next_turn_index = max(
        int(resume_checkpoint.get("nextTurnIndex", 0) or 0)
        if resume_checkpoint is not None
        else 0,
        0,
    )
    pending_assistant_message = (
        dict(resume_checkpoint["pendingAssistantMessage"])
        if resume_checkpoint is not None
        and isinstance(resume_checkpoint.get("pendingAssistantMessage"), dict)
        else None
    )
    next_tool_index = max(
        int(resume_checkpoint.get("nextToolIndex", 0) or 0)
        if resume_checkpoint is not None
        else 0,
        0,
    )
    restored_tool_batch_state = (
        dict(resume_checkpoint["toolBatchState"])
        if resume_checkpoint is not None
        and isinstance(resume_checkpoint.get("toolBatchState"), dict)
        else {}
    )

    def update_elapsed() -> None:
        state.elapsed_time = elapsed_before_run + (time.monotonic() - start_time)

    def current_tool_mode() -> str:
        if tools is PLAN_ONLY_TOOLS:
            return "plan_only"
        if tools is AGENT_TOOLS_WITH_DATA:
            return "terminal"
        return "full"

    async def save_boundary(
        stage: str,
        *,
        upcoming_turn: int,
        pending_message: dict[str, Any] | None = None,
        upcoming_tool: int = 0,
        tool_batch_state: dict[str, bool] | None = None,
    ) -> dict[str, Any]:
        update_elapsed()
        checkpoint = {
            "schemaVersion": 2,
            "kind": "planning",
            "stage": stage,
            "messages": messages,
            "toolCache": dict(cache or {}),
            "loopState": state.to_checkpoint(),
            "toolMode": current_tool_mode(),
            "repairingSubmission": repairing_submission,
            "nextTurnIndex": upcoming_turn,
            "pendingAssistantMessage": pending_message,
            "nextToolIndex": upcoming_tool,
            "toolBatchState": dict(tool_batch_state or {}),
        }
        if on_checkpoint is not None:
            await on_checkpoint(checkpoint)
        return checkpoint

    async def pause_if_needed(
        stage: str,
        *,
        upcoming_turn: int,
        pending_message: dict[str, Any] | None = None,
        upcoming_tool: int = 0,
        tool_batch_state: dict[str, bool] | None = None,
    ) -> None:
        # Every resumable boundary is committed, not only the boundary that
        # happens to cross the soft time budget.  A process can exit before the
        # reminder is due, and recovery still needs the latest complete model
        # response or tool batch instead of falling back to a fresh run.
        checkpoint = await save_boundary(
            stage,
            upcoming_turn=upcoming_turn,
            pending_message=pending_message,
            upcoming_tool=upcoming_tool,
            tool_batch_state=tool_batch_state,
        )
        if approved_wall_time_seconds is None:
            return
        update_elapsed()
        if state.elapsed_time < approved_wall_time_seconds:
            return
        await emit(
            "planning.budget_pause",
            "已安全保存当前进度",
            status="warning",
            summary="本轮处理时间已用完；继续前不会启动新的模型或工具调用。",
        )
        raise AgentBudgetPause(checkpoint, approved_wall_time_seconds)

    async with OpenAICompatibleClient(connection, timeout=timeout) as client:
        for turn_index in range(next_turn_index, loop_config.max_turns):
            update_elapsed()
            state.turn_index = turn_index
            stop_flag, stop_reason = should_stop(state, loop_config)
            if stop_flag:
                assert stop_reason is not None
                logger.warning(
                    "Agent loop stopped: reason=%s, turns=%s, tokens=%s, "
                    "errors=%s, loop=%s",
                    stop_reason,
                    state.turn_index,
                    state.total_tokens,
                    state.consecutive_errors,
                    state.detected_path_loop,
                )
                await emit(
                    "planning.guard",
                    "规划已安全停止",
                    status="warning",
                    summary="达到规划循环保护条件，已停止继续尝试。",
                    trace={
                        "kind": "correction",
                        "summary": "达到规划循环保护条件，已停止继续尝试。",
                        "basis": _safe_progress_text(stop_reason, 1000),
                    },
                )
                return safe_degradation(state, stop_reason)

            if turn_index == loop_config.warning_turn:
                warning_prompt = get_warning_prompt(state)
                if warning_prompt:
                    _append_warning_to_last_user_message(
                        messages, warning_prompt
                    )

            if turn_index == 0 and resume_checkpoint is None:
                authorized_count = len(request.workbook.worksheets)
                understanding_summary = (
                    f"已整理用户目标，并将规划范围限制在 {authorized_count} 张已授权工作表内。"
                )
                await emit(
                    "planning.understand",
                    "已理解需求与授权范围",
                    status="completed",
                    summary=understanding_summary,
                    trace={
                        "kind": "understanding",
                        "summary": understanding_summary,
                        "basis": "用户当前请求与已授权工作表元数据",
                    },
                )
            model_step_id = f"planning.model.{turn_index + 1}"
            resuming_tool_batch = False
            if pending_assistant_message is not None:
                message = pending_assistant_message
                pending_assistant_message = None
                resuming_tool_batch = next_tool_index > 0
                await emit(
                    model_step_id,
                    "正在继续已保存的工具步骤",
                    status="completed",
                    summary="已复用暂停前完整返回的模型响应。",
                )
            else:
                await pause_if_needed(
                    "before_model", upcoming_turn=turn_index
                )
                await emit(
                    model_step_id,
                    "正在生成操作方案"
                    if turn_index == 0
                    else f"正在修正操作方案（第 {turn_index + 1} 轮）",
                    status="started",
                )
                payload, usage = await client.chat_completions(
                    messages=messages,
                    tools=tools,
                    tool_choice="auto",
                    max_tokens=(
                        REPAIR_MAX_OUTPUT_TOKENS
                        if repairing_submission
                        else MAX_OUTPUT_TOKENS
                    ),
                )
                state.total_tokens += usage.get("total_tokens", 0)
                message = payload["choices"][0]["message"]
            assistant_message = dict(message)
            assistant_message.setdefault("role", "assistant")
            await pause_if_needed(
                "model_response_saved",
                upcoming_turn=turn_index,
                pending_message=assistant_message,
                upcoming_tool=next_tool_index if resuming_tool_batch else 0,
                tool_batch_state=(
                    restored_tool_batch_state if resuming_tool_batch else None
                ),
            )
            tool_calls = assistant_message.get("tool_calls") or []
            if not tool_calls:
                content = assistant_message.get("content") or ""
                try:
                    parsed = _parse_content(content)
                    if isinstance(parsed, PlanResponse):
                        action_types = "、".join(
                            action.type for action in parsed.plan.actions
                        )
                        await emit(
                            model_step_id,
                            "操作计划已生成",
                            status="completed",
                            summary=parsed.plan.summary,
                            trace={
                                "kind": "decision",
                                "summary": parsed.plan.summary,
                                "basis": f"共 {len(parsed.plan.actions)} 步：{action_types}",
                            },
                        )
                    else:
                        await emit(
                            model_step_id,
                            "回答已生成",
                            status="completed",
                            summary="已根据受控上下文生成回答。",
                        )
                    return parsed
                except (ValueError, ValidationError, json.JSONDecodeError) as error:
                    repairing_submission = True
                    state.last_error_message = str(error)[:200]
                    update_loop_state(
                        state,
                        None,
                        None,
                        error_occurred=True,
                    )
                    # 无效自由文本可能接近输出上限；错误原因已单独反馈，历史中只
                    # 保留最小占位，避免修复轮再次携带整段无效输出。
                    messages.append(
                        {
                            "role": "assistant",
                            "content": "未提交有效的结构化计划或回答。",
                        }
                    )
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
                    correction_summary = "模型提交内容未通过结构校验，正在要求重新提交。"
                    await emit(
                        model_step_id,
                        "提交格式未通过，正在修正",
                        status="warning",
                        summary=correction_summary,
                        trace={
                            "kind": "correction",
                            "summary": correction_summary,
                            "basis": _safe_progress_text(error, 1000),
                        },
                    )
                    continue

            if not resuming_tool_batch and (
                not messages or messages[-1] != assistant_message
            ):
                messages.append(assistant_message)
            requested_tool_names = [
                str((tool_call.get("function") or {}).get("name", "")).strip()
                for tool_call in tool_calls
            ]
            requested_tool_names = [
                name for name in requested_tool_names if name
            ]
            await emit(
                model_step_id,
                "已确定下一步受控操作",
                status="completed",
                summary=(
                    f"模型请求执行 {len(tool_calls)} 个白名单工具调用"
                    + (
                        f"：{'、'.join(requested_tool_names)}。"
                        if requested_tool_names
                        else "。"
                    )
                ),
                facts=[
                    fact
                    for fact in [
                        _progress_fact(
                            "工具",
                            "、".join(requested_tool_names),
                            style="code",
                        )
                    ]
                    if fact is not None
                ],
            )
            successful_readonly_batch = bool(
                restored_tool_batch_state.get("successfulReadonly")
                if resuming_tool_batch
                else False
            )
            readonly_batch_had_error = bool(
                restored_tool_batch_state.get("readonlyHadError")
                if resuming_tool_batch
                else False
            )
            tool_start_index = next_tool_index if resuming_tool_batch else 0
            if tool_start_index > len(tool_calls):
                raise ValueError("规划检查点的工具序号超出完整模型响应范围")
            for tool_index, tool_call in enumerate(tool_calls[tool_start_index:], start=tool_start_index):
                function = tool_call.get("function") or {}
                name = str(function.get("name", ""))
                arguments: dict[str, Any] = {}
                tool_step_id = (
                    f"planning.tool.{turn_index + 1}.{tool_index + 1}"
                )
                try:
                    parsed_arguments = json.loads(
                        function.get("arguments") or "{}"
                    )
                    if not isinstance(parsed_arguments, dict):
                        raise ValueError("工具参数必须是 JSON 对象")
                    arguments = parsed_arguments
                    step_text = _tool_step(name, arguments)
                    if step_text is None:
                        if name == "submit_plan":
                            step_text = "正在校验操作计划"
                        elif name == "submit_answer":
                            step_text = "正在校验回答"
                        else:
                            step_text = "正在调用受控工具"
                    await emit(
                        tool_step_id,
                        step_text,
                        status="started",
                        tool_name=name or None,
                        facts=_tool_argument_facts(name, arguments),
                    )
                    cache_key: str | None = None
                    if cache is not None and name in READONLY_TOOLS:
                        cache_key = _tool_cache_key(name, arguments)
                    if cache_key is not None and cache_key in cache:
                        cached_value = cache[cache_key]
                        if isinstance(cached_value, str):
                            tool_content = cached_value
                            try:
                                truncation_value = json.loads(cached_value)
                            except json.JSONDecodeError:
                                truncation_value = cached_value
                        else:
                            truncation_value = cached_value
                            tool_content = json.dumps(
                                cached_value, ensure_ascii=False
                            )
                        summary_value = truncation_value
                        used_cache = True
                        if _contains_snapshot_truncation(truncation_value):
                            state.encountered_truncation = True
                        conflict_warning = detect_action_conflict(
                            state,
                            cache_key,
                            loop_config.max_cache_hits_per_query,
                        )
                        if conflict_warning:
                            tool_content = _inject_tool_warning(
                                tool_content, conflict_warning
                            )
                    else:
                        result = _execute_tool(request, name, arguments)
                        if isinstance(result, (PlanResponse, AnswerResponse)):
                            if isinstance(result, PlanResponse):
                                action_types = "、".join(
                                    action.type for action in result.plan.actions
                                )
                                await emit(
                                    tool_step_id,
                                    "操作计划已通过结构校验",
                                    status="completed",
                                    summary=result.plan.summary,
                                    tool_name=name,
                                    facts=_plan_progress_facts(result.plan),
                                    trace={
                                        "kind": "decision",
                                        "summary": result.plan.summary,
                                        "basis": (
                                            f"共 {len(result.plan.actions)} 步：{action_types}"
                                        ),
                                    },
                                )
                            else:
                                await emit(
                                    tool_step_id,
                                    "回答已通过结构校验",
                                    status="completed",
                                    summary="已根据受控上下文生成回答。",
                                    tool_name=name,
                                    facts=[
                                        fact
                                        for fact in [
                                            _progress_fact("结果", "回答协议校验通过")
                                        ]
                                        if fact is not None
                                    ],
                                )
                            return result
                        summary_value = result
                        used_cache = False
                        if _contains_snapshot_truncation(result):
                            state.encountered_truncation = True
                        tool_content = json.dumps(result, ensure_ascii=False)
                        if cache_key is not None:
                            cache[cache_key] = tool_content

                    update_loop_state(
                        state,
                        name,
                        cache_key,
                        error_occurred=False,
                    )
                    if name in READONLY_TOOLS:
                        successful_readonly_batch = True
                    path_warning = detect_path_loop(
                        state, loop_config.max_tool_path_loop_length
                    )
                    if path_warning:
                        tool_content = _inject_tool_warning(
                            tool_content, path_warning
                        )
                        state.last_warning = path_warning
                    if name in READONLY_TOOLS:
                        (
                            completed_title,
                            completed_summary,
                            completed_trace,
                            completed_target,
                        ) = _readonly_tool_completion(
                            name,
                            arguments,
                            summary_value,
                            cached=used_cache,
                        )
                        await emit(
                            tool_step_id,
                            completed_title,
                            status="completed",
                            summary=completed_summary,
                            trace=completed_trace,
                            target=completed_target,
                            tool_name=name,
                            facts=_readonly_tool_facts(
                                name,
                                arguments,
                                summary_value,
                                cached=used_cache,
                            ),
                        )
                except (ValueError, ValidationError, json.JSONDecodeError) as error:
                    if name in READONLY_TOOLS:
                        readonly_batch_had_error = True
                    if name == "submit_plan":
                        # 失败的工具参数可能接近输出上限。保留完整原文会让下一轮
                        # 再次携带数千 token；工具错误已经包含校验原因，因此把历史
                        # 调用压缩为合法的最小占位，并只暴露 submit_plan 供定向修复。
                        function["arguments"] = json.dumps(
                            {"plan": None}, ensure_ascii=False
                        )
                        tools = PLAN_ONLY_TOOLS
                        repairing_submission = True
                        failed_plan = arguments.get("plan")
                        state.last_failed_plan = (
                            dict(failed_plan)
                            if isinstance(failed_plan, dict)
                            else None
                        )
                    state.last_error_message = str(error)[:200]
                    update_loop_state(
                        state,
                        name,
                        None,
                        error_occurred=True,
                    )
                    tool_content = json.dumps(
                        {"error": str(error)}, ensure_ascii=False
                    )
                    correction_summary = (
                        "操作计划未通过结构校验，正在修正。"
                        if name == "submit_plan"
                        else f"受控工具 {name or '调用'} 未成功，正在调整下一步。"
                    )
                    await emit(
                        tool_step_id,
                        "检查未通过，正在修正",
                        status="warning",
                        summary=correction_summary,
                        tool_name=name or None,
                        facts=[
                            fact
                            for fact in [
                                _progress_fact("结果", "未通过结构校验")
                            ]
                            if fact is not None
                        ],
                        trace={
                            "kind": "correction",
                            "summary": correction_summary,
                            "basis": _safe_progress_text(error, 1000),
                        },
                    )
                messages.append(
                    {
                        "role": "tool",
                        "tool_call_id": tool_call.get("id", ""),
                        "name": name,
                        "content": tool_content,
                    }
                )
                await pause_if_needed(
                    "tool_completed",
                    upcoming_turn=turn_index,
                    pending_message=assistant_message,
                    upcoming_tool=tool_index + 1,
                    tool_batch_state={
                        "successfulReadonly": successful_readonly_batch,
                        "readonlyHadError": readonly_batch_had_error,
                    },
                )
            if (
                successful_readonly_batch
                and not readonly_batch_had_error
                and tools is AGENT_TOOLS
            ):
                tools = AGENT_TOOLS_WITH_DATA
                _append_warning_to_last_user_message(
                    messages,
                    "工作簿只读探索已完成。下一轮必须基于现有工作簿上下文和"
                    "工具结果直接调用 submit_plan 或 submit_answer；不要继续读取数据。",
                )
            await pause_if_needed(
                "tool_batch_completed", upcoming_turn=turn_index + 1
            )
            next_tool_index = 0
            restored_tool_batch_state = {}
    state.turn_index = loop_config.max_turns
    update_elapsed()
    logger.warning(
        "Agent loop exhausted: turns=%s, tokens=%s, errors=%s, loop=%s",
        state.turn_index,
        state.total_tokens,
        state.consecutive_errors,
        state.detected_path_loop,
    )
    await emit(
        "planning.guard",
        "规划已安全停止",
        status="warning",
        summary="已达到最大规划轮次，停止继续尝试。",
        trace={
            "kind": "correction",
            "summary": "已达到最大规划轮次，停止继续尝试。",
        },
    )
    return safe_degradation(state, "max_turns")
