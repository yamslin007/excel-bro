"""Structured local plan compilers. No prompt parsing or workbook mutation."""
from __future__ import annotations

import uuid
from typing import Any, Literal

from openpyxl.utils.cell import range_boundaries
from pydantic import BaseModel, ConfigDict, Field

from .capabilities import capability_int
from .models import AnalysisPlan, PlanRequest, PlanResponse, _range_dimensions


class DeleteOtherSheets(BaseModel):
    model_config = ConfigDict(extra="forbid")
    keepSheets: list[str] = Field(min_length=1, max_length=256)


class RangeEdit(BaseModel):
    model_config = ConfigDict(extra="forbid")
    sheet: str = Field(min_length=1, max_length=100)
    range: str = Field(min_length=1, max_length=100)
    operation: Literal["writeValues", "writeFormulas", "clearRange"]
    # A single explicit literal repeated over the range; calculations belong to
    # query/Python tools and result references, never to guessed model values.
    value: str | int | float | bool | None = None


PLAN_COMPILERS = {
    "plan_delete_other_worksheets": (
        DeleteOtherSheets,
        "按明确的 keepSheets 列表保留工作表，仅为同一文件当前已授权的其余表生成删除计划。"
        "不能解析自然语言，不能扩大授权，不执行删除；名称不明确时返回错误供澄清。",
    ),
    "plan_range_edit": (
        RangeEdit,
        "按明确表名、A1 区域和字面常量/公式生成写值、写公式或清空内容计划。"
        "不读取数据、不计算结果、不执行；统计结果应使用 submit_result_plan。",
    ),
}


def plan_compiler_definitions(request: PlanRequest) -> list[dict[str, Any]]:
    from .excel_agent import _allowed_action_types
    allowed = _allowed_action_types(request) if request.hostId else None
    definitions = []
    for name, (model, description) in PLAN_COMPILERS.items():
        if allowed is not None:
            if name == "plan_delete_other_worksheets" and "deleteWorksheet" not in allowed:
                continue
            if name == "plan_range_edit" and not allowed.intersection(
                {"writeValues", "writeFormulas", "clearRange:contents"}
            ):
                continue
        schema = model.model_json_schema()
        if name == "plan_range_edit" and allowed is not None:
            schema["properties"]["operation"]["enum"] = [
                op for op in schema["properties"]["operation"]["enum"]
                if ("clearRange:contents" if op == "clearRange" else op) in allowed
            ]
        definitions.append({"type": "function", "function": {
            "name": name, "description": description, "parameters": schema,
        }})
    return definitions


def _sheet(request: PlanRequest, name: str):
    key = name.strip().casefold()
    exact = [s for s in request.workbook.worksheets if s.name.casefold() == key]
    candidates = exact or [s for s in request.workbook.worksheets
                           if s.sourceSheet and s.sourceSheet.casefold() == key]
    if not candidates:
        raise ValueError("sheet_not_authorized：名称不在当前授权清单；请核对名称或向用户请求调整范围。")
    if len(candidates) != 1:
        raise ValueError("sheet_ambiguous：多个已授权文件中存在同名表，请澄清具体文件并使用完整显示名。")
    return candidates[0]


def compile_plan_tool(request: PlanRequest, name: str, arguments: dict[str, Any]) -> PlanResponse:
    from .excel_agent import _enforce_action_capabilities
    if name == "plan_delete_other_worksheets":
        args = DeleteOtherSheets.model_validate(arguments)
        keep = {_sheet(request, name).name for name in args.keepSheets}
        # Folder display names may refer to different workbooks. Never infer a
        # target file from the first matching sheet or delete across files.
        files = {s.sourceFileId or s.sourceFile for s in request.workbook.worksheets}
        if len(files) != 1:
            raise ValueError("file_ambiguous：本工具一次只处理一个文件，请澄清目标文件并调整授权范围。")
        actions = [{"type": "deleteWorksheet", "sheet": s.name}
                   for s in request.workbook.worksheets if s.name not in keep]
        if not actions:
            raise ValueError("no_changes：当前授权范围内没有其他工作表可删除，无需执行。")
        title = "保留指定工作表并删除已授权的其他表"
    elif name == "plan_range_edit":
        args = RangeEdit.model_validate(arguments)
        sheet = _sheet(request, args.sheet)
        address = args.range.upper().replace("$", "")
        shape = _range_dimensions(address)
        if shape is None or min(shape) <= 0:
            raise ValueError("invalid_range：请提供单个有效 A1 矩形区域。")
        left, top, right, bottom = range_boundaries(address)
        if not (1 <= left <= right <= 16384 and 1 <= top <= bottom <= 1048576):
            raise ValueError("invalid_range：区域超出工作表行列边界。")
        if shape[0] * shape[1] > capability_int("agentPlanTools", "maxLiteralCells"):
            raise ValueError("range_too_large：区域超过本地字面写入工具限额。")
        action: dict[str, Any] = {"type": args.operation, "sheet": sheet.name, "range": address}
        if args.operation == "clearRange":
            if "value" in args.model_fields_set:
                raise ValueError("invalid_arguments：清空操作不能同时写入值。")
            action["applyTo"] = "contents"
        else:
            if "value" not in args.model_fields_set:
                raise ValueError("missing_value：写入值必须明确指定；需要用户决定时请澄清。")
            if args.operation == "writeFormulas" and (
                not isinstance(args.value, str) or not args.value.startswith("=")
            ):
                raise ValueError("invalid_formula：公式必须为以等号开头的字符串。")
            if args.operation == "writeValues" and isinstance(args.value, str) and args.value.lstrip().startswith(("=", "+", "-", "@")):
                raise ValueError("formula_as_value：疑似公式不能通过常量写入工具执行。")
            key = "values" if args.operation == "writeValues" else "formulas"
            action[key] = [[args.value] * shape[1] for _ in range(shape[0])]
        actions = [action]
        title = "按指定区域生成编辑计划"
    else:
        raise ValueError("unknown_plan_compiler")
    plan = AnalysisPlan.model_validate({
        "id": f"sdk-local-{uuid.uuid4().hex[:12]}", "title": title,
        "summary": title, "actions": actions,
    })
    _enforce_action_capabilities(request, plan)
    return PlanResponse(provider="model", plan=plan)
