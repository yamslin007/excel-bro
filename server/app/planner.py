from __future__ import annotations

import re
import uuid
from collections.abc import Awaitable, Callable
from typing import Any

from pydantic import ValidationError

from .capabilities import model_timeout_seconds
from .excel_agent import run_excel_agent
from .llm import selected_model_config
from .models import (
    AnalysisPlan,
    AnswerResponse,
    AssistantResponse,
    PlanRequest,
    PlanResponse,
    ResultContext,
)


COMMAND_VARIANT_TRANSLATION = str.maketrans(
    {
        "計": "计",
        "輸": "输",
        "將": "将",
        "結": "结",
        "應": "应",
        "當": "当",
        "選": "选",
        "擇": "择",
        "寫": "写",
        "錄": "录",
        "儲": "储",
        "彙": "汇",
        "匯": "汇",
        "總": "总",
        "統": "统",
        "對": "对",
        "較": "较",
        "異": "异",
        "數": "数",
        "據": "据",
        "間": "间",
        "頁": "页",
        "裡": "里",
        "個": "个",
        "兩": "两",
        "項": "项",
        "庫": "库",
        "這": "这",
        "剛": "刚",
        "該": "该",
        "顯": "显",
        "範": "范",
        "圍": "围",
        "與": "与",
        "從": "从",
        "後": "后",
        "還": "还",
        "為": "为",
        "開": "开",
        "關": "关",
        "現": "现",
        "檔": "档",
        "單": "单",
        "刪": "删",
        "除": "除",
        "貼": "贴",
        "值": "值",
    }
)


def _canonical_text(value: object) -> str:
    return str(value).translate(COMMAND_VARIANT_TRANSLATION).casefold()


def _strip_edit_value(value: str) -> str:
    return value.strip().strip("\"'“”‘’` ")


def _parse_edit_value(value: str) -> object:
    cleaned = _strip_edit_value(value)
    numeric = _numeric_value(cleaned)
    if numeric is not None:
        return int(numeric) if numeric.is_integer() else numeric
    lowered = cleaned.casefold()
    if lowered in {"true", "是"}:
        return True
    if lowered in {"false", "否"}:
        return False
    if lowered in {"空白", "空", "null", "none"}:
        return None
    return cleaned


def _resolve_target_sheet(request: PlanRequest, requested: str | None) -> str | None:
    sheets = request.workbook.worksheets
    if requested:
        normalized = _canonical_text(requested).strip("'\" ")
        exact = next(
            (sheet.name for sheet in sheets if _canonical_text(sheet.name) == normalized),
            None,
        )
        if exact:
            return exact
        source_match = next(
            (
                sheet.name
                for sheet in sheets
                if sheet.sourceSheet
                and _canonical_text(sheet.sourceSheet) == normalized
            ),
            None,
        )
        return source_match
    if any(sheet.name == request.workbook.activeWorksheet for sheet in sheets):
        return request.workbook.activeWorksheet
    return sheets[0].name if sheets else None


def _local_address_plan(request: PlanRequest) -> AssistantResponse | None:
    text = _canonical_text(request.prompt).strip()
    clear_match = re.search(
        r"(?:清空|清除)\s*(?:(?P<sheet>[^!，,\s]+)!)?"
        r"(?P<range>\$?[a-z]{1,3}\$?\d+(?::\$?[a-z]{1,3}\$?\d+)?)"
        r"(?:\s*(?:单元格|区域|范围))?",
        text,
    )
    if clear_match:
        sheet_name = _resolve_target_sheet(request, clear_match.group("sheet"))
        if not sheet_name:
            return AnswerResponse(provider="local", message="没有找到要清空的工作表。")
        address = clear_match.group("range").upper().replace("$", "")
        plan = AnalysisPlan.model_validate(
            {
                "id": f"local-clear-{uuid.uuid4().hex[:10]}",
                "title": f"清空 {sheet_name}!{address}",
                "summary": f"清空「{sheet_name}」{address} 中的内容。",
                "warnings": [f"将清除「{sheet_name}」{address} 的现有内容。"],
                "actions": [
                    {
                        "type": "clearRange",
                        "sheet": sheet_name,
                        "range": address,
                        "applyTo": "contents",
                    }
                ],
            }
        )
        return PlanResponse(plan=plan, provider="local")

    match = re.search(
        r"(?:把|将|在|向)?\s*(?:(?P<sheet>[^!，,\s]+)!)?"
        r"(?P<range>\$?[a-z]{1,3}\$?\d+(?::\$?[a-z]{1,3}\$?\d+)?)"
        r"(?:\s*(?:单元格|区域|范围))?\s*"
        r"(?:填入|写入|输入|设置为|设为|填上|写上)\s*"
        r"(?:(?:数字|数值|内容|公式)\s*)?(?P<value>.+?)\s*[。！!？?]?$",
        text,
    )
    if not match:
        return None
    sheet_name = _resolve_target_sheet(request, match.group("sheet"))
    if not sheet_name:
        return AnswerResponse(provider="local", message="没有找到要写入的工作表。")
    address = match.group("range").upper().replace("$", "")
    raw_value = _strip_edit_value(match.group("value"))
    if not raw_value:
        return AnswerResponse(provider="local", message="请告诉我要写入的具体内容。")
    range_match = re.fullmatch(
        r"([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?", address
    )
    row_count = 1
    column_count = 1
    if range_match and range_match.group(3):
        row_count = int(range_match.group(4)) - int(range_match.group(2)) + 1
        column_count = (
            sum(
                (ord(character) - 64) * (26 ** index)
                for index, character in enumerate(
                    reversed(range_match.group(3))
                )
            )
            - sum(
                (ord(character) - 64) * (26 ** index)
                for index, character in enumerate(
                    reversed(range_match.group(1))
                )
            )
            + 1
        )
    is_formula = raw_value.startswith("=")
    matrix = [
        [raw_value if is_formula else _parse_edit_value(raw_value)] * column_count
        for _ in range(row_count)
    ]
    action = (
        {
            "type": "writeFormulas",
            "sheet": sheet_name,
            "range": address,
            "formulas": matrix,
        }
        if is_formula
        else {
            "type": "writeValues",
            "sheet": sheet_name,
            "range": address,
            "values": matrix,
        }
    )
    plan = AnalysisPlan.model_validate(
        {
            "id": f"local-write-{uuid.uuid4().hex[:10]}",
            "title": f"写入 {sheet_name}!{address}",
            "summary": f"准备向「{sheet_name}」{address} 写入指定内容。",
            "warnings": ["如果目标单元格已有内容，执行后会被覆盖。"],
            "actions": [action],
        }
    )
    return PlanResponse(plan=plan, provider="local")


def _numeric_value(value: object) -> float | None:
    if isinstance(value, bool) or value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        candidate = value.strip().replace(",", "")
        if re.fullmatch(r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)", candidate):
            return float(candidate)
    return None


def _local_analysis(request: PlanRequest) -> AssistantResponse:
    address_plan = _local_address_plan(request)
    if address_plan is not None:
        return address_plan
    return AnswerResponse(
        provider="local",
        message=(
            "基础模式不支持此操作。\n\n"
            "💡 输入 /help 查看支持的命令\n"
            "或配置 AI 模型解锁完整能力。"
        ),
    )


def _local_data_tool_response(request: PlanRequest) -> AnswerResponse | None:
    if not request.dataResults:
        return None
    result = request.dataResults[-1]
    preview_lines = [
        " | ".join(str(value) for value in row)
        for row in result.rows[:10]
    ]
    remainder = (
        f"\n另有 {len(result.rows) - 10} 行未在消息中展开。"
        if len(result.rows) > 10
        else ""
    )
    warnings = (
        f"\n注意：{'；'.join(result.warnings)}"
        if result.warnings
        else ""
    )
    return AnswerResponse(
        provider="local",
        message=(
            f"{result.title}。{result.calculation}\n"
            f"已扫描 {result.scannedRows} 行，结果 {len(result.rows)} 行。\n"
            f"{chr(10).join(preview_lines)}{remainder}{warnings}"
        ),
        resultContext=ResultContext(
            title=result.title,
            headers=result.headers,
            rows=result.rows,
            sourceSheets=result.sourceSheets,
            warnings=result.warnings,
        ),
    )


async def create_plan(
    request: PlanRequest,
    *,
    tool_cache: dict[str, Any] | None = None,
    on_event: Callable[[dict[str, Any]], Awaitable[None]] | None = None,
) -> AssistantResponse:
    config = selected_model_config(request.modelId)
    if config is None:
        if request.images:
            raise ValueError("基础模式不支持图片，请选择支持视觉输入的模型")
        tool_response = _local_data_tool_response(request)
        if tool_response is not None:
            return tool_response
        return _local_analysis(request)
    if request.images and not config.supports_vision:
        raise ValueError(
            f"模型「{config.model}」未配置图片能力；请在 AI_VISION_MODELS 中声明"
        )
    timeout = model_timeout_seconds()
    try:
        return await run_excel_agent(
            request,
            connection=config,
            timeout=timeout,
            tool_cache=tool_cache,
            on_event=on_event,
        )
    except (KeyError, IndexError, TypeError, ValidationError) as error:
        raise ValueError(f"模型计划无法通过安全校验：{error}") from error


