"""Shared agent tool catalog: one source for terminal / non-terminal dispatch,
host capability pruning and typed outbound projections.

Both model backends must describe the same local table tools, validate them the
same way and project the same receipts:

* the normal OpenAI-compatible agent loop (``excel_agent``), and
* the Claude Agent SDK adapter (``claude_sdk_agent``).

This module owns the *local query* tool surface and its **model-safe**
projections so that neither adapter forks business rules. It deliberately
imports only ``models`` / ``capabilities`` at module import time; the one place
that needs the agent loop (host action pruning) imports it lazily to avoid an
import cycle.

Boundaries kept here:

* The model never receives data rows. ``query_table`` returns an opaque
  ``resultRef`` plus a fixed status and the *allowed structure*; the value
  matrix stays in the trusted local registry and is shown to the user by the
  local result card.
* Projections are explicit whitelists. A field being named ``headers`` does not
  by itself authorize it, so every emitted key is listed here on purpose.
"""

from __future__ import annotations

import copy
from typing import Any, Callable

from .capabilities import capability_int
from .models import PlanRequest, QueryTableArguments


def _deepcopy(node: Any) -> Any:
    return copy.deepcopy(node)

# --------------------------------------------------------------------------
# Tool names
# --------------------------------------------------------------------------

DESCRIBE_TABLES_TOOL = "describe_authorized_tables"
QUERY_TABLE_TOOL = "query_table"
PREPARE_PYTHON_DATA_TOOL = "prepare_python_data"
SUBMIT_LOCAL_RESULT_TOOL = "submit_local_result"
SUBMIT_RESULT_PLAN_TOOL = "submit_result_plan"

# Non-terminal tools: they return a typed receipt and the model keeps working
# inside the same ``query()`` / agent loop.
LOCAL_QUERY_TOOL_NAMES = frozenset({DESCRIBE_TABLES_TOOL, QUERY_TABLE_TOOL})

# 隔离 Python 数据登记工具：非终结，走桥让宿主读授权表并登记值矩阵。
# 由 python_execution_enabled 总开关决定是否下发（见 claude_sdk_agent 注册处）。
PREPARE_PYTHON_DATA_TOOL_NAMES = frozenset({PREPARE_PYTHON_DATA_TOOL})

# Terminal tools specific to the local result / result-plan path. They join the
# existing terminal set (``submit_plan`` / ``submit_answer`` /
# ``submit_clarification`` / ``map_column_values``) owned by ``excel_agent``.
LOCAL_TERMINAL_TOOL_NAMES = frozenset(
    {SUBMIT_LOCAL_RESULT_TOOL, SUBMIT_RESULT_PLAN_TOOL}
)

# Fixed receipt / error codes. The model only ever sees one of these strings,
# never free text coming from the workbook or from an exception.
QUERY_STATUS_OK = "ok"
QUERY_STATUS_ERROR = "error"

QUERY_ERROR_CODES = frozenset(
    {
        "unauthorized_scope",
        "invalid_arguments",
        "source_changed",
        "cancelled",
        "timeout",
        "host_unavailable",
        "result_too_large",
        "budget_exhausted",
        "busy",
        "internal",
    }
)


# --------------------------------------------------------------------------
# Schema helpers (kept local so this module stays import-light)
# --------------------------------------------------------------------------


def _strip_metadata(node: Any) -> Any:
    """Drop pydantic's auto-generated ``title`` keys; keep everything else."""
    if isinstance(node, dict):
        return {
            key: _strip_metadata(value)
            for key, value in node.items()
            if key != "title"
        }
    if isinstance(node, list):
        return [_strip_metadata(item) for item in node]
    return node


def _inline_json_schema(schema: dict[str, Any]) -> dict[str, Any]:
    """Inline ``$defs`` references so a tool schema is self-contained."""
    defs = schema.get("$defs", {})
    seen: set[str] = set()

    def resolve(node: Any) -> Any:
        if isinstance(node, dict):
            ref = node.get("$ref")
            if isinstance(ref, str) and ref.startswith("#/$defs/"):
                name = ref.rsplit("/", 1)[-1]
                if name not in defs or name in seen:
                    return {}
                seen.add(name)
                resolved = resolve(defs[name])
                seen.discard(name)
                return resolved
            return {key: resolve(value) for key, value in node.items() if key != "$defs"}
        if isinstance(node, list):
            return [resolve(item) for item in node]
        return node

    return resolve(schema)


def _query_table_parameters() -> dict[str, Any]:
    """``query_table`` parameter schema, generated from the shared model.

    Reuse shared arguments, then restrict combine to the union mode implemented
    by both live-workbook executors. Folder join/deduplication are separate paths.
    """
    schema = _inline_json_schema(QueryTableArguments.model_json_schema())
    for variant in schema["properties"]["combine"]["anyOf"]:
        if "properties" in variant:
            variant["properties"]["mode"]["enum"] = ["union"]
    return _strip_metadata(schema)


# --------------------------------------------------------------------------
# Tool definitions
# --------------------------------------------------------------------------


def local_query_tool_definitions() -> list[dict[str, Any]]:
    """Non-terminal local table tools available inside one SDK turn."""
    return [
        {
            "type": "function",
            "function": {
                "name": DESCRIBE_TABLES_TOOL,
                "description": (
                    "查看本任务已授权工作表的结构（表名、使用范围、行列数、表头）。"
                    "只返回授权范围内的结构信息，不含任何数据行。"
                ),
                "parameters": {"type": "object", "properties": {}},
            },
        },
        {
            "type": "function",
            "function": {
                "name": QUERY_TABLE_TOOL,
                "description": (
                    "在本任务已授权的数据上做确定性查询：筛选、分组汇总、纵向追加合并"
                    "（union，保留重复记录）、画像。当前工作簿不支持关联（join）或去重，"
                    "遇到此需求应明确说明限制，不得擅自改成追加合并。"
                    "计算完全在本地完成，回执只给出固定状态、"
                    "结果引用和结果结构，**不含任何数值**；真实结果由本地结果卡展示。"
                    "需要多步查询时可在同一任务内多次调用，也可以根据回执修正参数后"
                    "再次调用。"
                ),
                "parameters": _query_table_parameters(),
            },
        },
    ]


def prepare_python_data_tool_definition() -> dict[str, Any]:
    """``prepare_python_data`` 工具 schema：把授权表登记为隔离 Python 的数据集。

    模型只拿到 ``dataset://`` 句柄字符串；值矩阵与 fingerprint 都由可信侧管理，
    模型自报句柄不能当授权凭据。本工具受 ``python_execution_enabled`` 总开关约束，
    关闭时不随工具目录下发。
    """
    return {
        "type": "function",
        "function": {
            "name": PREPARE_PYTHON_DATA_TOOL,
            "description": (
                "把本任务已授权工作表的指定区域登记为隔离 Python 分析可用的数据集。"
                "回执只给出不透明数据集句柄（dataset://…），**不含任何数值**；"
                "真实值由本地可信侧管理，模型自报句柄不能当授权凭据。"
                "登记后即可在 run_python_analysis 的 datasetHandles 里引用这些句柄。"
                "每个 source 的 sheet 必须是已授权工作表名，range 是 A1 区域（如 A1:D20）。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "sources": {
                        "type": "array",
                        "minItems": 1,
                        "maxItems": 8,
                        "items": {
                            "type": "object",
                            "properties": {
                                "sheet": {"type": "string", "minLength": 1, "maxLength": 31},
                                "range": {"type": "string", "minLength": 1, "maxLength": 40},
                            },
                            "required": ["sheet", "range"],
                            "additionalProperties": False,
                        },
                    },
                },
                "required": ["sources"],
                "additionalProperties": False,
            },
        },
    }


def local_terminal_tool_definitions() -> list[dict[str, Any]]:
    """Terminal tools that turn a registered local result into user-visible output."""
    return [
        {
            "type": "function",
            "function": {
                "name": SUBMIT_LOCAL_RESULT_TOOL,
                "description": (
                    "把一个**已登记**的本地查询结果作为**只读结果卡**展示给用户。"
                    "只传结果引用，不要、也无法重新生成结果矩阵。"
                    "调用前必须判断本轮用户要的是不是写入：若用户要的是"
                    "「把结果写进某张表 / 生成写入预览」，那只是中间成果，"
                    "不要把 planningGoal 填成 write_preview 后再用本工具收尾 —— "
                    "应改用 submit_result_plan 提交结果引用与目标区域，"
                    "否则本轮规划目标不算完成。"
                ),
                "parameters": {
                    "type": "object",
                    "properties": {
                        "resultRef": {"type": "string", "minLength": 1, "maxLength": 200},
                        "message": {"type": "string", "maxLength": 1000},
                        "planningGoal": {
                            "type": "string",
                            "enum": ["read_only", "write_preview"],
                            "description": (
                                "你从用户需求里读出的**本轮规划目标**："
                                "read_only = 用户只要看到分析结果；"
                                "write_preview = 用户要求把结果生成写入预览。"
                                "声明 write_preview 时本工具会被拒绝，"
                                "请改用 submit_result_plan。"
                            ),
                        },
                    },
                    "required": ["resultRef", "planningGoal"],
                    "additionalProperties": False,
                },
            },
        },
        {
            "type": "function",
            "function": {
                "name": SUBMIT_RESULT_PLAN_TOOL,
                "description": (
                    "提出把一个**已登记**的本地结果写入一个已明确的目标区域。"
                    "本地层会校验结果类型、尺寸、目标范围、重叠与宿主限额，"
                    "只编译成已有白名单动作；提交成功不代表写入完成。"
                    "写入任务由宿主复核授权、来源与恢复准备后执行一次，再回读供审阅；确认变更不重复执行。"
                ),
                "parameters": {
                    "type": "object",
                    "properties": {
                        "resultRef": {"type": "string", "minLength": 1, "maxLength": 200},
                        "target": {
                            "type": "object",
                            "properties": {
                                "sheet": {"type": "string", "minLength": 1, "maxLength": 100},
                                "range": {"type": "string", "minLength": 1, "maxLength": 40},
                            },
                            "required": ["sheet", "range"],
                            "additionalProperties": False,
                        },
                        "title": {"type": "string", "maxLength": 200},
                        "includeHeader": {
                            "type": "boolean",
                            "description": (
                                "是否把结果表头写入目标区域的第 1 行。默认 true。"
                                "目标区域行列数必须与要写入的矩阵完全一致。"
                            ),
                        },
                    },
                    "required": ["resultRef", "target"],
                    "additionalProperties": False,
                },
            },
        },
    ]


def host_pruned_terminal_definitions(request: PlanRequest) -> list[dict[str, Any]]:
    """Terminal definitions with ``submit_plan`` pruned to the host bound.

    When the caller declares a ``hostId``, the ``submit_plan`` action variants
    are limited to that host's allowed set so unsupported variants (e.g.
    ``clearRange.applyTo``) never appear as candidates. When no ``hostId`` is
    declared the definition is returned **unchanged**, preserving the existing
    byte-for-byte shared contract and the post-submit capability check.

    **收窄结果为空集时不得产出 `{"anyOf": []}`。** 空 `anyOf` 是**非法 JSON
    Schema**（多数模型网关会直接拒绝整轮请求：``[] has less than 1 item``），
    后果不是"这个工具不可用"，而是**整次规划请求 400** —— 客户端漏报
    ``selfReportedActions``、``hostVersion`` 不匹配、宿主开关关闭都会走到
    这里，若产空 `anyOf` 就表现为"请求莫名失败"，掩盖真正的协商失败原因。

    空集时省略 submit_plan，保留回答/澄清等工具。既不发非法 schema，
    也不向模型宣称未协商出的动作可用。提交校验与执行预检仍独立存在。
    """
    from .excel_agent import (
        AGENT_TOOLS_WITH_DATA,
        _action_schemas,
        _allowed_action_types,
    )

    if not request.hostId:
        return [tool for tool in AGENT_TOOLS_WITH_DATA]

    variants = _action_schemas()
    selected = sorted(set(variants) & _allowed_action_types(request))
    candidates = [variants[name] for name in selected]
    result: list[dict[str, Any]] = []
    for tool in AGENT_TOOLS_WITH_DATA:
        if tool["function"]["name"] != "submit_plan":
            result.append(tool)
            continue
        if not candidates:
            continue
        projected = _deepcopy(tool)
        items = {"anyOf": candidates}
        projected["function"]["parameters"]["properties"]["plan"]["properties"]["actions"]["items"] = items
        result.append(projected)
    return result


# --------------------------------------------------------------------------
# Authorized structure projection
# --------------------------------------------------------------------------


def authorized_structure_payload(request: PlanRequest) -> list[dict[str, Any]]:
    """Whitelist rebuild of the authorized table structure.

    Never returns workbook / snapshot objects: only names, used range, sizes and
    headers, all of which the prompt is already allowed to carry.
    """
    tables: list[dict[str, Any]] = []
    for sheet in request.workbook.worksheets:
        tables.append(
            {
                "name": sheet.name,
                "sourceSheet": sheet.sourceSheet,
                "usedRange": sheet.usedRange,
                "rowCount": sheet.rowCount,
                "columnCount": sheet.columnCount,
                "headers": [
                    header if isinstance(header, str) else str(header)
                    for header in sheet.headers
                ],
            }
        )
    return tables


def _bounded_str(value: Any, limit: int) -> str | None:
    if not isinstance(value, str):
        return None
    text = value.strip()
    if not text or len(text) > limit:
        return None
    return text


# --------------------------------------------------------------------------
# Outbound projections (the model-facing gate)
# --------------------------------------------------------------------------


def project_describe_tables(result: dict[str, Any]) -> dict[str, Any]:
    """Projection for ``describe_authorized_tables``.

    The payload is already a whitelist rebuild. This re-validates the shape so
    a malformed result fails closed instead of leaking whatever it happens to
    contain.
    """
    if not isinstance(result, dict):
        raise ValueError("表结构结果必须是对象")
    tables = result.get("tables")
    if not isinstance(tables, list):
        raise ValueError("表结构结果缺少 tables")
    projected: list[dict[str, Any]] = []
    for table in tables:
        if not isinstance(table, dict):
            raise ValueError("表结构项必须是对象")
        name = _bounded_str(table.get("name"), 255)
        if name is None:
            raise ValueError("表结构项缺少有效名称")
        headers = table.get("headers")
        if headers is not None and not isinstance(headers, list):
            raise ValueError("表头必须是数组")
        safe_headers = []
        for header in headers or []:
            text = header if isinstance(header, str) else str(header)
            safe_headers.append(text[:200])
        projected.append(
            {
                "name": name,
                "sourceSheet": _bounded_str(table.get("sourceSheet"), 100),
                "usedRange": _bounded_str(table.get("usedRange"), 40),
                "rowCount": int(table.get("rowCount") or 0),
                "columnCount": int(table.get("columnCount") or 0),
                "headers": safe_headers,
            }
        )
    return {"tool": DESCRIBE_TABLES_TOOL, "status": QUERY_STATUS_OK, "tables": projected}


def project_query_receipt(result: dict[str, Any]) -> dict[str, Any]:
    """Projection for ``query_table``.

    Emitted keys are an explicit whitelist: fixed status, opaque result ref,
    allowed structure and a fixed error code. Values, ``title``, ``calculation``,
    ``warnings``, ``scannedRows`` and raw exception text are **not** emitted,
    even though the trusted local receipt carries them.
    """
    if not isinstance(result, dict):
        raise ValueError("查询回执必须是对象")
    status = result.get("status")
    if status not in (QUERY_STATUS_OK, QUERY_STATUS_ERROR):
        raise ValueError("查询回执状态无效")

    projected: dict[str, Any] = {"tool": QUERY_TABLE_TOOL, "status": status}

    if status == QUERY_STATUS_ERROR:
        code = result.get("errorCode")
        if code not in QUERY_ERROR_CODES:
            code = "internal"
        projected["errorCode"] = code
        return projected

    ref = _bounded_str(result.get("resultRef"), 200)
    if ref is None:
        raise ValueError("成功回执缺少结果引用")
    projected["resultRef"] = ref

    structure = result.get("structure")
    if structure is None:
        structure = {}
    if not isinstance(structure, dict):
        raise ValueError("结果结构必须是对象")
    headers = structure.get("headers")
    if headers is not None and not isinstance(headers, list):
        raise ValueError("结果表头必须是数组")
    safe_headers = []
    for header in headers or []:
        text = header if isinstance(header, str) else str(header)
        safe_headers.append(text[:200])
    projected["structure"] = {
        "headers": safe_headers,
        "rowCount": int(structure.get("rowCount") or 0),
        "columnCount": int(structure.get("columnCount") or 0),
        "complete": bool(structure.get("complete")),
        "truncated": bool(structure.get("truncated")),
    }
    return projected


def project_prepare_python_data(result: dict[str, Any]) -> dict[str, Any]:
    """Projection for ``prepare_python_data``.

    只出站不透明 ``dataset://`` 句柄列表与固定状态；绝不透出值矩阵或 fingerprint。
    """
    if not isinstance(result, dict):
        raise ValueError("数据登记回执必须是对象")
    status = result.get("status")
    if status not in (QUERY_STATUS_OK, QUERY_STATUS_ERROR):
        raise ValueError("数据登记回执状态无效")

    projected: dict[str, Any] = {"tool": PREPARE_PYTHON_DATA_TOOL, "status": status}
    if status == QUERY_STATUS_ERROR:
        code = result.get("errorCode")
        if code not in QUERY_ERROR_CODES:
            code = "internal"
        projected["errorCode"] = code
        return projected

    handles = result.get("datasetHandles")
    if not isinstance(handles, list) or not handles:
        raise ValueError("成功回执缺少数据集句柄")
    safe_handles = []
    for handle in handles:
        text = _bounded_str(handle, 256)
        if text is None:
            raise ValueError("数据集句柄非法")
        safe_handles.append(text)
    projected["datasetHandles"] = safe_handles
    return projected


def project_terminal_result(result: dict[str, Any]) -> dict[str, Any]:
    """Projection for the local result terminal tools.

    These normally return an ``AssistantResponse`` and never reach the
    non-terminal projection gate; this exists only so the registry is complete
    and a malformed dict fails closed.
    """
    raise ValueError("终局工具结果不出站")


LOCAL_NON_TERMINAL_PROJECTIONS: dict[str, Callable[[dict[str, Any]], dict[str, Any]]] = {
    DESCRIBE_TABLES_TOOL: project_describe_tables,
    QUERY_TABLE_TOOL: project_query_receipt,
    PREPARE_PYTHON_DATA_TOOL: project_prepare_python_data,
}


def max_bridge_tool_calls() -> int:
    """Per-task cap on bridge tool calls, from shared capabilities."""
    return capability_int("agentToolBridge", "maxToolCallsPerTask")
