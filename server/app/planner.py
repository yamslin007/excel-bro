from __future__ import annotations

import re
import uuid
from collections.abc import Awaitable, Callable
from typing import Any

from pydantic import ValidationError

from .capabilities import model_timeout_seconds
from .excel_agent import run_excel_agent
from .llm import ModelNotConfiguredError, selected_model_config
from .models import (
    AnalysisPlan,
    AnswerResponse,
    AssistantResponse,
    PlanRequest,
    PlanResponse,
    ResultContext,
)
from .structural_intent import parse_bulk_worksheet_deletion


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
    """纯语法地址命令：清空区域 / 地址写值 / 地址写公式。

    这层只做字面解析，不做任何语义理解，因此不依赖外部模型，也不构成
    “基础模式”这个产品入口的回归——入口已从模型列表和首次引导中移除，
    模型选择里不会再出现它。

    匹配不上时返回 ``None``，把请求交回模型路径；匹配上时产出的仍是
    白名单 ``AnalysisPlan``，照旧先预览、再由用户确认后才执行。
    """

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


def _local_bulk_worksheet_deletion_plan(
    request: PlanRequest,
) -> AssistantResponse | None:
    deletion = parse_bulk_worksheet_deletion(request.prompt)
    if deletion is None:
        return None

    keeper_key = _canonical_text(deletion.keeper_name).strip()
    exact_candidates = [
        sheet
        for sheet in request.workbook.worksheets
        if _canonical_text(sheet.name).strip() == keeper_key
    ]
    source_candidates = [
        sheet
        for sheet in request.workbook.worksheets
        if sheet.sourceSheet
        and _canonical_text(sheet.sourceSheet).strip() == keeper_key
    ]
    keeper_candidates = exact_candidates or source_candidates
    if not keeper_candidates:
        return AnswerResponse(
            provider="local",
            message=(
                f"要保留的工作表「{deletion.keeper_name}」不在当前已授权范围内。"
                "请先在“数据范围”中选择全部涉及的工作表，再重新提交。"
            ),
        )

    if len(keeper_candidates) > 1:
        candidate_names = "、".join(
            f"「{sheet.name}」" for sheet in keeper_candidates[:5]
        )
        return AnswerResponse(
            provider="local",
            message=(
                f"已选择的多个文件中都存在工作表「{deletion.keeper_name}」："
                f"{candidate_names}。请只选择一个文件的工作表，或使用“文件 › 工作表”"
                "的完整名称指定要保留的工作表。"
            ),
        )

    keeper = keeper_candidates[0]
    keeper_name = keeper.name
    keeper_label = keeper.sourceSheet or keeper.name
    keeper_source_key = keeper.sourceFileId or keeper.sourceFile
    if keeper_source_key is not None:
        foreign_sheets = [
            sheet
            for sheet in request.workbook.worksheets
            if (sheet.sourceFileId or sheet.sourceFile) != keeper_source_key
        ]
        if foreign_sheets:
            return AnswerResponse(
                provider="local",
                message=(
                    "批量保留/删除工作表一次只能作用于一个文件。"
                    "请在“数据范围”中只选择目标文件的工作表后重新提交。"
                ),
            )

    targets = [
        sheet
        for sheet in request.workbook.worksheets
        if sheet.name.casefold() != keeper_name.casefold()
    ]
    if not targets:
        return AnswerResponse(
            provider="local",
            message=f"当前已授权范围中没有「{keeper_label}」之外的工作表可删除。",
        )
    if len(targets) > 50:
        return AnswerResponse(
            provider="local",
            message=(
                f"当前有 {len(targets)} 张工作表需要删除，超过单次计划最多 50 个动作的限制。"
                "请分批选择并删除。"
            ),
        )

    target_preview = "、".join(
        f"「{sheet.sourceSheet or sheet.name}」" for sheet in targets[:8]
    )
    if len(targets) > 8:
        target_preview += f"等 {len(targets)} 张工作表"
    plan = AnalysisPlan.model_validate(
        {
            "id": f"local-delete-sheets-{uuid.uuid4().hex[:10]}",
            "title": f"保留 {keeper_label} 并删除其他工作表",
            "summary": (
                f"保留「{keeper_label}」，删除当前已授权的其余 {len(targets)} 张工作表。"
            ),
            "warnings": [f"将永久删除：{target_preview}。"],
            "actions": [
                {"type": "deleteWorksheet", "sheet": sheet.name}
                for sheet in targets
            ],
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
    on_checkpoint: Callable[[dict[str, Any]], Awaitable[None]] | None = None,
    resume_checkpoint: dict[str, Any] | None = None,
    approved_wall_time_seconds: float | None = None,
) -> AssistantResponse:
    structural_plan = _local_bulk_worksheet_deletion_plan(request)
    if structural_plan is not None:
        return structural_plan
    # 纯语法地址命令（清空区域 / 地址写值 / 地址写公式）是已固化下来的确定性
    # 能力：只做字面解析，不需要模型理解语义，也不该因为“没有选模型”或
    # “选了模型”而变成两条不同行为。它必须排在模型解析之前。
    address_plan = _local_address_plan(request)
    if address_plan is not None:
        return address_plan
    # 其余请求都需要外部模型理解需求。没有可用模型时必须显式失败：
    # 不静默进入规则模式，也不擅自改用其他供应商。
    try:
        config = selected_model_config(request.modelId)
    except ModelNotConfiguredError as error:
        if request.images:
            raise ModelNotConfiguredError(
                "尚未配置模型。图片输入需要支持视觉的外部模型，"
                "请先点击「添加模型」完成配置。"
            ) from error
        tool_response = _local_data_tool_response(request)
        if tool_response is not None:
            # 本条路径的全部计算都已在本地完成，只是把结果读出来，
            # 不涉及语义理解，也不代表用户“选好了模型”。
            return tool_response
        raise ModelNotConfiguredError(
            "尚未配置模型，无法理解这条需求。请点击「添加模型」，"
            "填写服务地址、模型 ID 和 API Key 后重试。"
        ) from error
    if request.images and not config.supports_vision:
        raise ValueError(
            f"模型「{config.model}」未配置图片能力；请在「模型设置」里为该连接"
            "勾选「支持图片输入」（用 .env 配置时改为把它加进 AI_VISION_MODELS）"
        )
    if config.adapter == "claude_agent_sdk":
        if resume_checkpoint is not None:
            from .llm import ModelNotAvailableError

            raise ModelNotAvailableError(
                "Claude Agent SDK 不能恢复另一规划后端的检查点，请重新发起预览。"
            )
        from .claude_sdk_agent import run_sdk_plan

        return await run_sdk_plan(request, config, on_event=on_event)
    timeout = model_timeout_seconds()
    try:
        return await run_excel_agent(
            request,
            connection=config,
            timeout=timeout,
            tool_cache=tool_cache,
            on_event=on_event,
            on_checkpoint=on_checkpoint,
            resume_checkpoint=resume_checkpoint,
            approved_wall_time_seconds=approved_wall_time_seconds,
        )
    except (KeyError, IndexError, TypeError, ValidationError) as error:
        raise ValueError(f"模型计划无法通过安全校验：{error}") from error


