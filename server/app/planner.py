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

GENERATED_SHEET_NAMES = {"AI分析计划", "AI分析结果"}


def _has_data(sheet: object) -> bool:
    data_rows = getattr(sheet, "dataRows", [])
    return any(any(value not in (None, "") for value in row) for row in data_rows)


def _column_name(headers: list[object], index: int) -> str:
    if index < len(headers) and headers[index] not in (None, ""):
        return str(headers[index])
    return f"第 {index + 1} 列"


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


def _normalize_lookup_text(value: object) -> str:
    return re.sub(r"[\W_]+", "", _canonical_text(value))


def _header_matches_prompt(header: object, prompt: str) -> bool:
    name = _normalize_lookup_text(header)
    return bool(name and name in prompt)


def _format_cell_value(value: object) -> str:
    if value is None or value == "":
        return "空白"
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.4f}".rstrip("0").rstrip(".")
    return str(value)


def _column_letter(index: int) -> str:
    result = ""
    current = index + 1
    while current:
        current, remainder = divmod(current - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _used_range_origin(address: str | None) -> tuple[int, int]:
    if not address:
        return 0, 1
    cell_address = address.rsplit("!", 1)[-1]
    match = re.match(r"\$?([A-Z]+)\$?(\d+)", cell_address.upper())
    if not match:
        return 0, 1
    column = 0
    for character in match.group(1):
        column = column * 26 + ord(character) - 64
    return column - 1, int(match.group(2))


def _local_location(
    request: PlanRequest, source_sheets: list[object]
) -> AnswerResponse | None:
    prompt = _normalize_lookup_text(request.prompt)
    location_markers = (
        "在哪里",
        "在哪儿",
        "在哪",
        "什么位置",
        "哪个位置",
        "哪一行",
        "第几行",
        "哪个单元格",
        "什么单元格",
    )
    if not any(marker in prompt for marker in location_markers):
        return None

    matches: list[tuple[int, str, str, str, int]] = []
    for sheet in source_sheets:
        start_column, start_row = _used_range_origin(sheet.usedRange)
        rows = [sheet.headers, *sheet.dataRows]
        for row_offset, row in enumerate(rows):
            for column_offset, value in enumerate(row):
                value_text = _normalize_lookup_text(value)
                if (
                    value in (None, "")
                    or len(value_text) < 2
                    or value_text not in prompt
                ):
                    continue
                row_number = start_row + row_offset
                address = f"{_column_letter(start_column + column_offset)}{row_number}"
                score = len(value_text)
                if sheet.name == request.workbook.activeWorksheet:
                    score += 1
                matches.append(
                    (score, str(value), sheet.name, address, row_number)
                )

    if not matches:
        return None

    best_score = max(match[0] for match in matches)
    best_matches = [match for match in matches if match[0] == best_score]
    lines = [
        f"{value}在「{sheet_name}」的 {address}（第 {row_number} 行）。"
        for _, value, sheet_name, address, row_number in best_matches
    ]
    return AnswerResponse(provider="local", message="\n".join(lines))


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


def _cell_in_range(cell: str, address: str) -> bool:
    cell_match = re.fullmatch(r"([A-Z]+)(\d+)", cell)
    range_match = re.fullmatch(
        r"([A-Z]+)(\d+):([A-Z]+)(\d+)", address
    )
    if not cell_match or not range_match:
        return False
    column = _column_number(cell_match.group(1))
    row = int(cell_match.group(2))
    return (
        _column_number(range_match.group(1))
        <= column
        <= _column_number(range_match.group(3))
        and int(range_match.group(2)) <= row <= int(range_match.group(4))
    )


def _local_formula_plan(
    request: PlanRequest, source_sheets: list[object]
) -> AssistantResponse | None:
    canonical = _canonical_text(request.prompt)
    normalized = _normalize_lookup_text(canonical)
    if "公式" not in normalized:
        return None

    operation = _match_aggregate_operation(normalized)
    if operation is None:
        return None
    _op_id, op_label, op_formula = operation

    destination_match = re.search(
        r"(?:(?P<sheet>[^!，,\s]+)!)?"
        r"(?P<cell>\$?[A-Za-z]{1,3}\$?\d+)",
        canonical,
    )
    if not destination_match:
        return AnswerResponse(
            provider="local",
            message="你要求写入公式，但没有提供目标单元格。请指定例如 A5。",
        )
    destination_sheet = _resolve_target_sheet(
        request, destination_match.group("sheet")
    )
    if not destination_sheet:
        return AnswerResponse(provider="local", message="没有找到公式目标工作表。")
    destination_cell = (
        destination_match.group("cell").upper().replace("$", "")
    )

    field_matches: list[tuple[object, int, str]] = []
    for sheet in source_sheets:
        for column_index, header in enumerate(sheet.headers):
            if header in (None, "") or not _header_matches_prompt(header, normalized):
                continue
            field_matches.append((sheet, column_index, str(header)))
    if len(field_matches) != 1:
        if not field_matches:
            return AnswerResponse(
                provider="local",
                message="没有找到公式所指的字段，请明确写出工作表和字段名称。",
            )
        locations = "、".join(
            f"{sheet.name} 的「{header}」"
            for sheet, _, header in field_matches[:8]
        )
        return AnswerResponse(
            provider="local",
            message=f"找到多个可能字段：{locations}。请指定公式要引用哪一个。",
        )

    source_sheet, column_index, header = field_matches[0]
    start_column, start_row = _used_range_origin(source_sheet.usedRange)
    source_column = _column_letter(start_column + column_index)
    data_start_row = start_row + 1
    data_end_row = max(data_start_row, start_row + source_sheet.rowCount - 1)
    source_range = f"{source_column}{data_start_row}:{source_column}{data_end_row}"
    qualified_source = (
        source_range
        if source_sheet.name == destination_sheet
        else f"'{source_sheet.name.replace(chr(39), chr(39) * 2)}'!{source_range}"
    )
    formula = f"={op_formula}({qualified_source})"
    if (
        source_sheet.name == destination_sheet
        and _cell_in_range(destination_cell, source_range)
    ):
        return AnswerResponse(
            provider="local",
            message=(
                f"目标 {destination_cell} 位于公式引用范围 {source_range} 内，"
                "会产生循环引用。请换一个目标单元格。"
            ),
        )

    plan = AnalysisPlan.model_validate(
        {
            "id": f"local-formula-{uuid.uuid4().hex[:10]}",
            "title": f"写入{op_label}公式",
            "summary": (
                f"在「{destination_sheet}」{destination_cell} 写入针对"
                f"「{source_sheet.name}」字段「{header}」的公式 {formula}。"
            ),
            "assumptions": ["首行是字段名称，公式覆盖该字段的实际数据行。"],
            "warnings": (
                ["源工作表快照被截断；公式仍按工作表记录的完整已用行数引用。"]
                if source_sheet.truncated
                else []
            ),
            "actions": [
                {
                    "type": "writeFormulas",
                    "sheet": destination_sheet,
                    "range": destination_cell,
                    "formulas": [[formula]],
                },
                {"type": "activateWorksheet", "sheet": destination_sheet},
            ],
        }
    )
    return PlanResponse(plan=plan, provider="local")


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


def _local_edit_plan(
    request: PlanRequest, source_sheets: list[object]
) -> AssistantResponse | None:
    match = re.search(
        r"(?:把|将)\s*(?:这个|该)?\s*(.+?)\s*"
        r"(?:修改为|改为|改成|替换为|替换成)\s*(.+?)\s*[。！!？?]?$",
        request.prompt.strip(),
    )
    if not match:
        return None

    old_value = _strip_edit_value(match.group(1))
    new_value = _parse_edit_value(match.group(2))
    if not old_value:
        return None

    matches: list[tuple[str, str, object]] = []
    for sheet in source_sheets:
        start_column, start_row = _used_range_origin(sheet.usedRange)
        rows = [sheet.headers, *sheet.dataRows]
        for row_offset, row in enumerate(rows):
            for column_offset, value in enumerate(row):
                if str(value).strip() != old_value:
                    continue
                address = (
                    f"{_column_letter(start_column + column_offset)}"
                    f"{start_row + row_offset}"
                )
                matches.append((sheet.name, address, value))

    if not matches:
        selector = _normalize_lookup_text(old_value)
        for sheet in source_sheets:
            start_column, start_row = _used_range_origin(sheet.usedRange)
            for column_index, header in enumerate(sheet.headers):
                header_text = _normalize_lookup_text(header)
                if not header_text or header_text not in selector:
                    continue
                for row_offset, row in enumerate(sheet.dataRows, start=1):
                    if column_index >= len(row):
                        continue
                    for label_index, label in enumerate(row):
                        label_text = _normalize_lookup_text(label)
                        if (
                            label_index == column_index
                            or len(label_text) < 2
                            or label_text not in selector
                        ):
                            continue
                        address = (
                            f"{_column_letter(start_column + column_index)}"
                            f"{start_row + row_offset}"
                        )
                        matches.append(
                            (sheet.name, address, row[column_index])
                        )

    if not matches:
        return AnswerResponse(
            provider="local",
            message=f"在已选工作表中没有找到值“{old_value}”，因此没有修改。",
        )
    if len(matches) > 1:
        locations = "、".join(
            f"{sheet_name}!{address}" for sheet_name, address, _ in matches
        )
        return AnswerResponse(
            provider="local",
            message=(
                f"找到了 {len(matches)} 个值“{old_value}”：{locations}。"
                "请指定要修改的工作表或单元格，避免误改。"
            ),
        )

    sheet_name, address, original = matches[0]
    plan = AnalysisPlan.model_validate(
        {
            "id": f"local-edit-{uuid.uuid4().hex[:10]}",
            "title": f"修改 {sheet_name}!{address}",
            "summary": (
                f"准备把「{sheet_name}」{address} 中的"
                f"“{_format_cell_value(original)}”改为"
                f"“{_format_cell_value(new_value)}”。"
            ),
            "assumptions": [],
            "warnings": ["此操作会修改现有单元格；请确认新值无误。"],
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": sheet_name,
                    "range": address,
                    "values": [[new_value]],
                }
            ],
        }
    )
    return PlanResponse(plan=plan, provider="local")


# 聚合操作单一来源：id（内部标识）、formula（Excel 函数名）、
# markers（匹配用户输入的关键词，取并集避免各命令漂移）、label（展示中文名）。
# 新增操作只改这里，formula / aggregate / 分析请求判定三处共享。
_AGGREGATE_OPERATIONS: tuple[tuple[str, str, tuple[str, ...], str], ...] = (
    ("sum", "SUM", ("合计", "总计", "总和", "求和", "一共"), "合计"),
    ("average", "AVERAGE", ("平均", "均值"), "平均值"),
    ("max", "MAX", ("最大", "最高"), "最大值"),
    ("min", "MIN", ("最小", "最低", "最少"), "最小值"),
    ("count", "COUNT", ("计数", "数量", "个数"), "计数"),
)


def _match_aggregate_operation(
    normalized: str, supported: frozenset[str] | None = None
) -> tuple[str, str, str] | None:
    """匹配归一化文本中的聚合操作，返回 (id, label, formula) 或 None。

    supported 限定可选操作（例如 aggregate 只支持 sum/average/max/min）。
    """
    for op_id, formula, markers, label in _AGGREGATE_OPERATIONS:
        if supported is not None and op_id not in supported:
            continue
        if any(marker in normalized for marker in markers):
            return op_id, label, formula
    return None


_REMOVE_DUPLICATES_MARKERS = ("去重", "删除重复", "去除重复", "删除重复行", "去重复")


def _local_remove_duplicates_plan(
    request: PlanRequest, source_sheets: list[object]
) -> AssistantResponse | None:
    normalized = _normalize_lookup_text(request.prompt)
    if not any(marker in normalized for marker in _REMOVE_DUPLICATES_MARKERS):
        return None

    target_sheet = _resolve_target_sheet(request, None)
    sheet = next(
        (item for item in source_sheets if item.name == target_sheet),
        None,
    )
    if sheet is None:
        if len(source_sheets) == 1:
            sheet = source_sheets[0]
        else:
            return AnswerResponse(
                provider="local",
                message="当前有多个工作表，请指定要在哪张表上去重。",
            )

    if sheet.truncated:
        return AnswerResponse(
            provider="local",
            message=(
                f"「{sheet.name}」的数据快照被截断，无法在本地精确去重。"
                "请缩小数据范围或配置模型后重试。"
            ),
        )

    matched_columns = [
        column_index
        for column_index, header in enumerate(sheet.headers)
        if header not in (None, "") and _header_matches_prompt(header, normalized)
    ]

    column_count = max(
        len(sheet.headers),
        max((len(row) for row in sheet.dataRows), default=0),
    )
    if column_count == 0:
        return AnswerResponse(
            provider="local", message=f"「{sheet.name}」没有可去重的数据。"
        )

    filters: list[dict[str, object]] = []
    filter_column: int | None = None
    filter_value_text: str | None = None
    if not matched_columns:
        # 未命中表头时，尝试把提示里出现的值当作过滤条件（如「阿里去重」→ 人员=阿里）
        for column_index in range(column_count):
            header = (
                sheet.headers[column_index]
                if column_index < len(sheet.headers)
                else None
            )
            if header in (None, ""):
                continue
            for row in sheet.dataRows:
                value = row[column_index] if column_index < len(row) else None
                if value in (None, ""):
                    continue
                value_text = _normalize_lookup_text(value)
                if len(value_text) >= 2 and value_text in normalized:
                    filter_column = column_index
                    filter_value_text = value_text
                    filters = [
                        {
                            "field": str(header or ""),
                            "operator": "equals",
                            "value": value,
                        }
                    ]
                    break
            if filter_column is not None:
                break

    if matched_columns:
        dedupe_columns = matched_columns
        by_text = "、".join(str(sheet.headers[i]) for i in matched_columns)
    elif filter_column is not None:
        dedupe_columns = [filter_column]
        by_text = f"「{filters[0]['field']}」等于「{filters[0]['value']}」的行"
    else:
        dedupe_columns = list(range(column_count))
        by_text = "整行"

    def pad_row(row: list[object]) -> list[object]:
        return [row[i] if i < len(row) else None for i in range(column_count)]

    def matches_filters(row: list[object]) -> bool:
        if filter_column is None:
            return True
        value = row[filter_column] if filter_column < len(row) else None
        if value in (None, ""):
            return False
        return _normalize_lookup_text(value) == filter_value_text

    seen: set[tuple[object, ...]] = set()
    unique_rows: list[list[object]] = []
    for row in sheet.dataRows:
        if not matches_filters(row):
            # 未命中过滤条件的行原样保留（含其重复）
            unique_rows.append(row)
            continue
        key = tuple(row[i] if i < len(row) else None for i in dedupe_columns)
        if key in seen:
            continue
        seen.add(key)
        unique_rows.append(row)

    removed_count = len(sheet.dataRows) - len(unique_rows)
    if removed_count == 0:
        return AnswerResponse(
            provider="local",
            message=f"「{sheet.name}」的当前数据里没有重复行，无需去重。",
        )

    start_column, start_row = _used_range_origin(sheet.usedRange)
    start_letter = _column_letter(start_column)
    end_letter = _column_letter(start_column + column_count - 1)
    range_address = (
        f"{start_letter}{start_row}:{end_letter}{start_row + len(sheet.dataRows)}"
    )
    result_range = (
        f"{start_letter}{start_row}:{end_letter}{start_row + len(unique_rows)}"
    )
    result_rows = [pad_row(sheet.headers), *[pad_row(row) for row in unique_rows]]

    plan = AnalysisPlan.model_validate(
        {
            "id": f"local-dedupe-{uuid.uuid4().hex[:10]}",
            "title": f"去重「{sheet.name}」",
            "summary": (
                f"准备对「{sheet.name}」按 {by_text} 去重，"
                f"预计删除 {removed_count} 行重复数据。"
            ),
            "assumptions": (
                ["首行是字段名称，重复数据删除整行。"]
                if not filters
                else [
                    "首行是字段名称，重复数据删除整行。",
                    "仅对命中过滤条件的行去重，未命中行原样保留。",
                ]
            ),
            "warnings": ["去重会删除整行数据；请确认被删除的是多余重复行。"],
            "actions": [
                {
                    "type": "removeDuplicates",
                    "sheet": sheet.name,
                    "range": range_address,
                    "columns": dedupe_columns,
                    "hasHeaders": True,
                    **({"filters": filters} if filters else {}),
                }
            ],
            "acceptanceCriteria": [
                {
                    "type": "rangeEquals",
                    "sheet": sheet.name,
                    "range": result_range,
                    "expected": result_rows,
                }
            ],
        }
    )
    return PlanResponse(plan=plan, provider="local")


def _local_lookup(
    request: PlanRequest, source_sheets: list[object]
) -> AnswerResponse | None:
    prompt = _normalize_lookup_text(request.prompt)
    question_markers = ("多少", "几", "是什么", "是多少", "查询", "查找", "告诉我")
    if not any(marker in prompt for marker in question_markers):
        return None

    matches: list[tuple[int, str, str, object]] = []
    for sheet in source_sheets:
        matching_headers = [
            (index, header)
            for index, header in enumerate(sheet.headers)
            if header not in (None, "") and _header_matches_prompt(header, prompt)
        ]
        for row in sheet.dataRows:
            for label_index, label in enumerate(row):
                label_text = _normalize_lookup_text(label)
                if (
                    not label_text
                    or len(label_text) < 2
                    or label_text not in prompt
                ):
                    continue
                for column_index, header in matching_headers:
                    if column_index == label_index or column_index >= len(row):
                        continue
                    score = len(label_text) + len(_normalize_lookup_text(header))
                    if sheet.name == request.workbook.activeWorksheet:
                        score += 1
                    matches.append(
                        (score, sheet.name, f"{label}的{header}", row[column_index])
                    )

    if not matches:
        return None

    best_score = max(item[0] for item in matches)
    best_matches = [item for item in matches if item[0] == best_score]
    if len(best_matches) == 1:
        _, sheet_name, label, value = best_matches[0]
        return AnswerResponse(
            provider="local",
            message=f"{label}是 {_format_cell_value(value)}。",
            resultContext=ResultContext(
                title="查询结果",
                headers=["项目", "值"],
                rows=[[label, value]],
                primaryValueColumn=1,
                sourceSheets=[sheet_name],
            ),
        )

    lines = [
        f"「{sheet_name}」中，{label}是 {_format_cell_value(value)}。"
        for _, sheet_name, label, value in best_matches
    ]
    return AnswerResponse(
        provider="local",
        message="\n".join(lines),
        resultContext=ResultContext(
            title="查询结果",
            headers=["项目", "值"],
            rows=[
                [f"{sheet_name} › {label}", value]
                for _, sheet_name, label, value in best_matches
            ],
            primaryValueColumn=1,
            sourceSheets=list(
                dict.fromkeys(sheet_name for _, sheet_name, _, _ in best_matches)
            ),
        ),
    )


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


WRITE_OUTPUT_MARKERS = (
    "写入",
    "写到",
    "写在",
    "写进",
    "填入",
    "填到",
    "填在",
    "放到",
    "放在",
    "输出到",
    "输出至",
    "输出在",
    "保存到",
    "保存在",
    "存到",
    "记录到",
    "记录在",
    "插入到",
    "添加到",
    "生成结果表",
)

RESULT_REFERENCE_MARKERS = (
    "这个结果",
    "这些结果",
    "这个值",
    "这些值",
    "这个平均值",
    "这个合计",
    "刚才",
    "上一个结果",
    "上一轮结果",
    "前面的结果",
    "上述结果",
    "那个结果",
)


def _wants_written_output(prompt: str) -> bool:
    normalized = _normalize_lookup_text(prompt)
    if any(marker in normalized for marker in WRITE_OUTPUT_MARKERS):
        return True
    write_verbs = ("写", "填", "放", "输出", "保存", "记录", "插入", "添加", "粘贴")
    destinations = (
        "工作表",
        "子表",
        "新表",
        "新页",
        "sheet",
        "表格",
        "单元格",
        "格子",
    )
    has_cell_address = re.search(
        r"(?<![A-Za-z0-9_])\$?[A-Za-z]{1,3}\$?\d+\b", prompt
    )
    if has_cell_address and any(verb in normalized for verb in write_verbs):
        return True
    return any(verb in normalized for verb in write_verbs) and any(
        destination in normalized for destination in destinations
    )


def _references_previous_result(prompt: str) -> bool:
    normalized = _normalize_lookup_text(prompt)
    return any(marker in normalized for marker in RESULT_REFERENCE_MARKERS)


def _aggregate_scope(
    request: PlanRequest, source_sheets: list[object]
) -> tuple[str, list[object]]:
    prompt = request.prompt
    canonical_prompt = _canonical_text(prompt)
    output_positions = [
        position
        for marker in WRITE_OUTPUT_MARKERS
        if (position := canonical_prompt.find(marker)) >= 0
    ]
    scope_prompt = (
        canonical_prompt[: min(output_positions)]
        if output_positions
        else canonical_prompt
    )
    normalized = _normalize_lookup_text(scope_prompt)
    if any(marker in normalized for marker in ("当前表", "当前工作表", "本表")):
        active_sheets = [
            sheet
            for sheet in source_sheets
            if sheet.name == request.workbook.activeWorksheet
        ]
        return "single", active_sheets or source_sheets
    named_sheets = [
        sheet
        for sheet in source_sheets
        if _normalize_lookup_text(getattr(sheet, "name", "")) in normalized
    ]
    scoped_sheets = named_sheets or source_sheets
    if any(
        marker in normalized
        for marker in ("分别", "各自", "每张", "每个工作表", "逐表", "各表")
    ):
        return "separate", scoped_sheets
    if any(
        marker in normalized
        for marker in ("合并", "一起", "总体", "整体", "所有表", "全部工作表")
    ):
        return "combined", scoped_sheets
    return "implicit", scoped_sheets


def _small_integer(text: str) -> int | None:
    if text.isdigit():
        value = int(text)
        return value if 1 <= value <= 20 else None
    numerals = {
        "一": 1,
        "二": 2,
        "两": 2,
        "三": 3,
        "四": 4,
        "五": 5,
        "六": 6,
        "七": 7,
        "八": 8,
        "九": 9,
        "十": 10,
    }
    if text in numerals:
        return numerals[text]
    if text.startswith("十") and len(text) == 2 and text[1] in numerals:
        return 10 + numerals[text[1]]
    if text.endswith("十") and len(text) == 2 and text[0] in numerals:
        return numerals[text[0]] * 10
    return None


def _extreme_record_intent(prompt: str) -> tuple[str, int, bool] | None:
    normalized = _normalize_lookup_text(prompt)
    direction: str | None = None
    if any(
        marker in normalized
        for marker in ("最高", "最大", "第一名", "排名第一", "并列第一")
    ) or re.search(
        r"前[一二两三四五六七八九十\d]{1,3}(?:个|名|项|条|款|种)?",
        normalized,
    ):
        direction = "max"
    elif any(marker in normalized for marker in ("最低", "最小", "最少")) or re.search(
        r"(?:后|倒数)[一二两三四五六七八九十\d]{1,3}",
        normalized,
    ):
        direction = "min"
    if direction is None:
        return None

    top_n = 1
    count_match = re.search(
        r"(?:前|后|倒数)([一二两三四五六七八九十\d]{1,3})"
        r"|(?:最高|最大|最低|最小|最少)(?:的)?"
        r"([一二两三四五六七八九十\d]{1,3})(?:个|名|项|条|款|种)",
        normalized,
    )
    has_rank_count = count_match is not None
    if count_match:
        parsed = _small_integer(count_match.group(1) or count_match.group(2))
        if parsed is not None:
            top_n = parsed

    record_markers = (
        "谁",
        "哪个",
        "哪一个",
        "哪位",
        "哪项",
        "哪条",
        "哪款",
        "哪种",
        "对应",
        "记录",
        "哪一行",
        "倒数",
        "并列第一",
        "排名第一",
        "第一名",
    )
    record_language = has_rank_count or any(
        marker in normalized for marker in record_markers
    )
    return direction, top_n, record_language


def _identity_columns(
    sheet: object, metric_column: int, prompt: str
) -> tuple[list[int], bool]:
    normalized = _normalize_lookup_text(prompt)
    explicit = [
        index
        for index, header in enumerate(sheet.headers)
        if index != metric_column
        and header not in (None, "")
        and _normalize_lookup_text(header) in normalized
    ]
    if explicit:
        return explicit, True

    for column_index in range(sheet.columnCount):
        if column_index == metric_column:
            continue
        values = [
            row[column_index]
            for row in sheet.dataRows
            if column_index < len(row) and row[column_index] not in (None, "")
        ]
        if values and any(_numeric_value(value) is None for value in values):
            return [column_index], False
    return [], False


def _row_identity(
    sheet: object, row: list[object], columns: list[int], row_number: int
) -> str:
    if not columns:
        return f"第 {row_number} 行"
    if len(columns) == 1:
        column = columns[0]
        return _format_cell_value(row[column]) if column < len(row) else "空白"
    parts = []
    for column in columns:
        header = _column_name(sheet.headers, column)
        value = row[column] if column < len(row) else None
        parts.append(f"{header}={_format_cell_value(value)}")
    return "，".join(parts)


def _ranked_rows(
    entries: list[dict[str, object]], direction: str, top_n: int
) -> list[dict[str, object]]:
    ordered = sorted(
        entries,
        key=lambda item: float(item["value"]),
        reverse=direction == "max",
    )
    if not ordered:
        return []
    cutoff_index = min(top_n, len(ordered)) - 1
    cutoff = float(ordered[cutoff_index]["value"])
    return [
        item
        for item in ordered
        if (
            float(item["value"]) >= cutoff
            if direction == "max"
            else float(item["value"]) <= cutoff
        )
    ]


def _local_extreme_record_response(
    request: PlanRequest, source_sheets: list[object]
) -> AnswerResponse | None:
    intent = _extreme_record_intent(request.prompt)
    if intent is None:
        return None
    direction, top_n, record_language = intent
    scope, scoped_sheets = _aggregate_scope(request, source_sheets)
    candidates: list[dict[str, object]] = []

    for sheet in scoped_sheets:
        numeric_columns = [
            column_index
            for column_index in range(sheet.columnCount)
            if any(
                column_index < len(row)
                and _numeric_value(row[column_index]) is not None
                for row in sheet.dataRows
            )
        ]
        matched_metrics = [
            column_index
            for column_index in numeric_columns
            if column_index < len(sheet.headers)
            and _header_matches_prompt(
                sheet.headers[column_index],
                _normalize_lookup_text(request.prompt),
            )
        ]
        mentions_identity_field = any(
            column_index not in numeric_columns
            and column_index < len(sheet.headers)
            and sheet.headers[column_index] not in (None, "")
            and _normalize_lookup_text(sheet.headers[column_index])
            in _normalize_lookup_text(request.prompt)
            for column_index in range(sheet.columnCount)
        )
        if not matched_metrics and len(numeric_columns) == 1:
            matched_metrics = numeric_columns
        elif (
            not matched_metrics
            and len(numeric_columns) > 1
            and (record_language or mentions_identity_field)
        ):
            fields = "、".join(
                _column_name(sheet.headers, column) for column in numeric_columns
            )
            return AnswerResponse(
                provider="local",
                message=f"「{sheet.name}」中有多个数值字段：{fields}。请明确按哪个字段排名。",
            )
        if len(matched_metrics) > 1:
            fields = "、".join(
                _column_name(sheet.headers, column) for column in matched_metrics
            )
            return AnswerResponse(
                provider="local",
                message=f"「{sheet.name}」中有多个匹配的数值字段：{fields}。请明确按哪个字段排名。",
            )

        for metric_column in matched_metrics:
            metric = _column_name(sheet.headers, metric_column)
            identity_columns, explicit_identity = _identity_columns(
                sheet, metric_column, request.prompt
            )
            entries: list[dict[str, object]] = []
            invalid_count = 0
            for row_index, row in enumerate(sheet.dataRows, start=2):
                if metric_column >= len(row) or row[metric_column] in (None, ""):
                    continue
                numeric = _numeric_value(row[metric_column])
                if numeric is None:
                    invalid_count += 1
                    continue
                entries.append(
                    {
                        "sheet": sheet.name,
                        "metric": metric,
                        "identity": _row_identity(
                            sheet, row, identity_columns, row_index
                        ),
                        "value": numeric,
                    }
                )
            if entries:
                candidates.append(
                    {
                        "sheet": sheet.name,
                        "metric": metric,
                        "entries": entries,
                        "invalidCount": invalid_count,
                        "truncated": sheet.truncated,
                        "explicitIdentity": explicit_identity,
                    }
                )

    if not candidates:
        return None
    if not record_language and not any(
        bool(candidate["explicitIdentity"]) for candidate in candidates
    ):
        return None

    grouped: dict[str, list[dict[str, object]]] = {}
    for candidate in candidates:
        metric_key = _normalize_lookup_text(candidate["metric"])
        key = (
            f"{_normalize_lookup_text(candidate['sheet'])}\0{metric_key}"
            if scope == "separate"
            else metric_key
        )
        grouped.setdefault(key, []).append(candidate)

    if scope == "implicit" and any(len(group) > 1 for group in grouped.values()):
        sheets = sorted(
            {
                str(candidate["sheet"])
                for group in grouped.values()
                if len(group) > 1
                for candidate in group
            }
        )
        return AnswerResponse(
            provider="local",
            message=(
                f"{'、'.join(sheets)} 中包含同名排名字段。"
                "请说明要跨表比较，还是分别查找每张表的记录。"
            ),
        )

    result_entries: list[dict[str, object]] = []
    warnings: list[str] = []
    for group in grouped.values():
        combined_entries = [
            entry
            for candidate in group
            for entry in candidate["entries"]
            if isinstance(entry, dict)
        ]
        result_entries.extend(_ranked_rows(combined_entries, direction, top_n))
        for candidate in group:
            if candidate["invalidCount"]:
                warnings.append(
                    f"「{candidate['sheet']}」的{candidate['metric']}有 "
                    f"{candidate['invalidCount']} 个非数值内容未参与排名。"
                )
            if candidate["truncated"]:
                warnings.append(
                    f"「{candidate['sheet']}」超过读取上限，排名只覆盖前 200 行。"
                )

    if not result_entries:
        return None

    extreme_label = "最高" if direction == "max" else "最低"
    lines: list[str] = []
    if top_n == 1 and len(result_entries) == 1:
        item = result_entries[0]
        sheet_prefix = (
            f"「{item['sheet']}」中，" if len(scoped_sheets) > 1 else ""
        )
        lines.append(
            f"{sheet_prefix}{item['metric']}{extreme_label}的是"
            f"{item['identity']}，{item['metric']}为 "
            f"{_format_cell_value(item['value'])}。"
        )
    elif top_n == 1:
        by_metric: dict[tuple[str, str], list[dict[str, object]]] = {}
        for item in result_entries:
            key = (
                str(item["sheet"]) if scope == "separate" else "",
                str(item["metric"]),
            )
            by_metric.setdefault(key, []).append(item)
        for (sheet_name, metric), items in by_metric.items():
            identities = "、".join(
                (
                    f"{item['identity']}（{item['sheet']}）"
                    if len(scoped_sheets) > 1 and scope != "separate"
                    else str(item["identity"])
                )
                for item in items
            )
            value = _format_cell_value(items[0]["value"])
            prefix = f"「{sheet_name}」中，" if sheet_name else ""
            lines.append(
                f"{prefix}{metric}{extreme_label}的有{identities}，"
                f"{metric}均为 {value}。"
            )
    else:
        lines.append(f"按{extreme_label}顺序找到以下 {len(result_entries)} 项：")
        for index, item in enumerate(result_entries, start=1):
            lines.append(
                f"{index}. {item['identity']}：{item['metric']} "
                f"{_format_cell_value(item['value'])}"
                + (
                    f"（{item['sheet']}）"
                    if len(scoped_sheets) > 1
                    else ""
                )
            )
    lines.extend(f"注意：{warning}" for warning in dict.fromkeys(warnings))

    context_warnings = list(dict.fromkeys(warnings))[:20]
    context_rows = [
        [
            str(item["identity"]),
            str(item["metric"]),
            item["value"],
            str(item["sheet"]),
        ]
        for item in result_entries
    ]
    return AnswerResponse(
        provider="local",
        message=_bounded_lines(lines),
        resultContext=ResultContext(
            title=f"{extreme_label}记录",
            headers=["对象", "指标", "值", "工作表"],
            rows=context_rows,
            primaryValueColumn=None,
            sourceSheets=list(
                dict.fromkeys(str(item["sheet"]) for item in result_entries)
            ),
            warnings=context_warnings,
        ),
    )


def _column_number(name: str) -> int:
    result = 0
    for character in name.upper():
        result = result * 26 + ord(character) - ord("A") + 1
    return result


def _column_letters(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def _autofit_range(start_cell: str, width: int) -> str:
    match = re.fullmatch(r"\$?([A-Za-z]+)\$?\d+", start_cell)
    if not match:
        return "A:D"
    start = _column_number(match.group(1))
    return f"{_column_letters(start)}:{_column_letters(start + width - 1)}"


def _selected_start_cell(address: str | None) -> str | None:
    if not address:
        return None
    cell_part = address.rsplit("!", 1)[-1].split(":", 1)[0]
    match = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?(\d+)", cell_part.strip())
    if not match:
        return None
    return f"{match.group(1).upper()}{match.group(2)}"


def _aggregate_destination(request: PlanRequest) -> tuple[str, str, str]:
    prompt = request.prompt
    canonical_prompt = _canonical_text(prompt)
    marker_end = -1
    for marker in WRITE_OUTPUT_MARKERS:
        position = canonical_prompt.rfind(marker)
        if position >= 0:
            marker_end = max(marker_end, position + len(marker))
    tail = prompt[marker_end:] if marker_end >= 0 else ""
    normalized_tail = _normalize_lookup_text(tail)

    address_match = re.search(
        r"(?<![A-Za-z0-9_])\$?([A-Za-z]{1,3})\$?(\d+)\b",
        tail,
    )
    start_cell = (
        f"{address_match.group(1).upper()}{address_match.group(2)}"
        if address_match
        else "A1"
    )

    for sheet in request.workbook.worksheets:
        if _normalize_lookup_text(sheet.name) in normalized_tail:
            return sheet.name, start_cell, "cell" if address_match else "table"
    if any(
        marker in normalized_tail for marker in ("当前表", "当前工作表", "本表")
    ):
        if any(
            marker in normalized_tail
            for marker in (
                "相应位置",
                "当前位置",
                "选中位置",
                "选择位置",
                "当前单元格",
                "选中单元格",
                "这里",
                "此处",
            )
        ):
            selected_cell = _selected_start_cell(request.workbook.selectedRange)
            if selected_cell:
                return request.workbook.activeWorksheet, selected_cell, "cell"
        return request.workbook.activeWorksheet, start_cell, (
            "cell" if address_match else "table"
        )
    if address_match and not re.search(r"(?:工作表|表)", tail):
        return request.workbook.activeWorksheet, start_cell, "cell"
    if any(
        marker in normalized_tail
        for marker in ("新工作表", "新子表", "新表", "新页", "newsheet")
    ):
        return "AI分析结果", start_cell, "table"

    selection_markers = (
        "相应位置",
        "当前位置",
        "选中位置",
        "选择位置",
        "当前单元格",
        "选中单元格",
        "这里",
        "此处",
    )
    if any(marker in normalized_tail for marker in selection_markers):
        selected_cell = _selected_start_cell(request.workbook.selectedRange)
        if selected_cell:
            return request.workbook.activeWorksheet, selected_cell, "cell"

    new_sheet_match = re.search(
        r'^\s*(?:新建(?:的)?\s*)?[「“"]?'
        r'([^「」“”",，。!！\s]+?)[」”"]?'
        r"(?:工作表|表)(?:里|中|内|的|!|\s|$)",
        tail,
    )
    if new_sheet_match:
        candidate = new_sheet_match.group(1).strip()
        if candidate not in {"新", "新工作", "当前", "这个"} and len(candidate) <= 31:
            return candidate, start_cell, "table"
    return "AI分析结果", start_cell, "table"


def _result_write_actions(
    sheet: str,
    start_cell: str,
    target_kind: str,
    headers: list[str],
    rows: list[list[object]],
    primary_value_column: int | None,
) -> list[dict[str, object]]:
    actions: list[dict[str, object]] = [
        {"type": "createWorksheet", "sheet": sheet}
    ]
    if (
        target_kind == "cell"
        and len(rows) == 1
        and primary_value_column is not None
        and primary_value_column < len(rows[0])
    ):
        actions.append(
            {
                "type": "writeValues",
                "sheet": sheet,
                "range": start_cell,
                "values": [[rows[0][primary_value_column]]],
            }
        )
    else:
        actions.extend(
            [
                {
                    "type": "writeTable",
                    "sheet": sheet,
                    "startCell": start_cell,
                    "headers": headers,
                    "rows": rows,
                },
                {
                    "type": "autofit",
                    "sheet": sheet,
                    "range": _autofit_range(start_cell, len(headers)),
                },
            ]
        )
    actions.append({"type": "activateWorksheet", "sheet": sheet})
    return actions


def _local_result_followup(request: PlanRequest) -> AssistantResponse | None:
    if not _references_previous_result(request.prompt):
        return None
    if not _wants_written_output(request.prompt):
        return None
    if request.lastResult is None:
        return AnswerResponse(
            provider="local",
            message=(
                "我没有收到上一轮的结构化结果，因此无法安全判断“这个结果”指什么。"
                "请重新计算一次，或直接说明要写入的字段和值。"
            ),
        )

    context = request.lastResult
    normalized = _normalize_lookup_text(request.prompt)
    singular_reference = any(
        marker in normalized for marker in ("这个", "那个", "该结果", "此结果")
    )
    plural_reference = any(
        marker in normalized for marker in ("这些", "全部", "所有", "上述")
    )
    if len(context.rows) > 1 and singular_reference and not plural_reference:
        return AnswerResponse(
            provider="local",
            message=(
                f"上一轮包含 {len(context.rows)} 项结果。"
                "请说明要写入哪一项，或说“把这些结果写到新工作表”。"
            ),
            resultContext=context,
        )

    result_sheet, start_cell, target_kind = _aggregate_destination(request)
    plan = AnalysisPlan.model_validate(
        {
            "id": f"local-followup-{uuid.uuid4().hex[:10]}",
            "title": f"写入{context.title}",
            "summary": (
                f"将上一轮的「{context.title}」直接写入"
                f"「{result_sheet}」{start_cell}，不重新扩展为完整统计。"
            ),
            "assumptions": [
                "“这个结果”指向上一轮助手返回的结构化结果。",
            ],
            "warnings": context.warnings,
            "actions": _result_write_actions(
                result_sheet,
                start_cell,
                target_kind,
                context.headers,
                context.rows,
                context.primaryValueColumn,
            ),
        }
    )
    return PlanResponse(plan=plan, provider="local")


def _render_value_sample(values: list[object], limit: int = 8) -> str:
    sample = "、".join(_format_cell_value(value) for value in values[:limit])
    if len(values) <= limit:
        return sample
    return f"{sample}，另有 {len(values) - limit} 个（共 {len(values)} 个）"


def _render_invalid_sample(
    invalid: list[tuple[object, object]], limit: int = 8
) -> str:
    sample = "、".join(
        f"{row_label}={value}" for row_label, value in invalid[:limit]
    )
    if len(invalid) <= limit:
        return sample
    return f"{sample}，另有 {len(invalid) - limit} 项"


def _bounded_lines(lines: list[str], max_length: int = 3800) -> str:
    included: list[str] = []
    for index, line in enumerate(lines):
        omitted = len(lines) - index
        suffix = f"\n另有 {omitted} 项未展开。" if omitted else ""
        candidate = "\n".join([*included, line])
        if len(candidate) + len(suffix) > max_length:
            if omitted:
                included.append(f"另有 {omitted} 项未展开。")
            break
        included.append(line)
    return "\n".join(included)


def _local_aggregate_response(
    request: PlanRequest, source_sheets: list[object]
) -> AssistantResponse | None:
    prompt = _normalize_lookup_text(request.prompt)
    wants_written_output = _wants_written_output(request.prompt)
    asks_for_result = any(
        marker in prompt
        for marker in (
            "多少",
            "是多少",
            "是什么",
            "几",
            "计算",
            "算一下",
            "算出",
            "求出",
            "求平均",
        )
    )
    if not asks_for_result and not wants_written_output:
        return None

    matched = _match_aggregate_operation(
        prompt, supported=frozenset({"sum", "average", "max", "min"})
    )
    if matched is None:
        return None
    operation, op_label, _ = matched

    scope, source_sheets = _aggregate_scope(request, source_sheets)
    groups: dict[str, dict[str, object]] = {}
    for sheet in source_sheets:
        matched_columns = [
            index
            for index, header in enumerate(sheet.headers)
            if header not in (None, "") and _header_matches_prompt(header, prompt)
        ]
        if not matched_columns:
            numeric_columns = []
            for column_index in range(sheet.columnCount):
                if any(
                    column_index < len(row)
                    and _numeric_value(row[column_index]) is not None
                    for row in sheet.dataRows
                ):
                    numeric_columns.append(column_index)
            if len(numeric_columns) == 1:
                matched_columns = numeric_columns

        for column_index in matched_columns:
            header = _column_name(sheet.headers, column_index)
            header_key = _normalize_lookup_text(header)
            key = (
                f"{_normalize_lookup_text(sheet.name)}\0{header_key}"
                if scope == "separate"
                else header_key
            )
            group = groups.setdefault(
                key,
                {
                    "header": header,
                    "values": [],
                    "invalid": [],
                    "sheetNames": [],
                    "truncatedSheets": [],
                },
            )
            values = group["values"]
            invalid = group["invalid"]
            sheet_names = group["sheetNames"]
            truncated_sheets = group["truncatedSheets"]
            assert isinstance(values, list)
            assert isinstance(invalid, list)
            assert isinstance(sheet_names, list)
            assert isinstance(truncated_sheets, list)
            if sheet.name not in sheet_names:
                sheet_names.append(sheet.name)
            if sheet.truncated and sheet.name not in truncated_sheets:
                truncated_sheets.append(sheet.name)
            for row in sheet.dataRows:
                if column_index >= len(row) or row[column_index] in (None, ""):
                    continue
                numeric = _numeric_value(row[column_index])
                if numeric is not None:
                    values.append(numeric)
                    continue
                row_label = next(
                    (
                        str(value)
                        for index, value in enumerate(row)
                        if index != column_index and value not in (None, "")
                    ),
                    f"第 {len(values) + len(invalid) + 2} 行",
                )
                invalid.append((row_label, row[column_index]))

    if scope == "implicit":
        ambiguous_groups = [
            group
            for group in groups.values()
            if isinstance(group["sheetNames"], list)
            and len(group["sheetNames"]) > 1
        ]
        if ambiguous_groups:
            sheet_names = sorted(
                {
                    str(sheet_name)
                    for group in ambiguous_groups
                    for sheet_name in group["sheetNames"]
                }
            )
            return AnswerResponse(
                provider="local",
                message=(
                    f"{'、'.join(sheet_names)} 中包含同名数值字段。"
                    "请说明要“合并计算”还是“分别计算每张表”，"
                    "我不会默认把不同工作表的数据混在一起。"
                ),
            )

    answers: list[str] = []
    result_rows: list[list[object]] = []
    result_warnings: list[str] = []
    for group in groups.values():
        header = str(group["header"])
        values = group["values"]
        invalid = group["invalid"]
        sheet_names = group["sheetNames"]
        truncated_sheets = group["truncatedSheets"]
        assert isinstance(values, list)
        assert isinstance(invalid, list)
        assert isinstance(sheet_names, list)
        assert isinstance(truncated_sheets, list)
        if not values:
            continue
        if operation == "average":
            result = sum(values) / len(values)
        elif operation == "sum":
            result = sum(values)
        elif operation == "max":
            result = max(values)
        else:
            result = min(values)
        label = op_label
        display_header = (
            f"{sheet_names[0]} › {header}"
            if scope == "separate" and sheet_names
            else header
        )
        rendered_values = _render_value_sample(values)
        answer = (
            f"{display_header}的{label}是 {_format_cell_value(result)}"
            f"（按 {len(values)} 个有效数值计算：{rendered_values}）。"
        )
        result_rows.append(
            [display_header, label, round(result, 4), len(values)]
        )
        if invalid:
            rendered_invalid = _render_invalid_sample(invalid)
            answer += f" 未计入非数值内容：{rendered_invalid}。"
            result_warnings.append(
                f"{display_header}未计入非数值内容：{rendered_invalid}。"
            )
        if truncated_sheets:
            truncated_notice = (
                f"{'、'.join(str(name) for name in truncated_sheets)} 已超过读取上限；"
                "结果只基于已读取的前 200 行。"
            )
            answer += f" 注意：{truncated_notice}"
            result_warnings.append(truncated_notice)
        answers.append(answer)

    if not answers:
        return None
    operation_label = str(result_rows[0][1])
    plan_warnings = list(dict.fromkeys(result_warnings))
    if len(plan_warnings) > 20:
        omitted = len(plan_warnings) - 19
        plan_warnings = [
            *plan_warnings[:19],
            f"另有 {omitted} 项数据说明未展开。",
        ]
    result_context = ResultContext(
        title=operation_label,
        headers=["字段", "指标", "结果", "有效数值数"],
        rows=result_rows,
        primaryValueColumn=2,
        sourceSheets=list(dict.fromkeys(sheet.name for sheet in source_sheets)),
        warnings=plan_warnings,
    )
    if wants_written_output:
        result_sheet, start_cell, target_kind = _aggregate_destination(request)
        plan = AnalysisPlan.model_validate(
            {
                "id": f"local-{uuid.uuid4().hex[:10]}",
                "title": f"写入{operation_label}",
                "summary": (
                    f"已计算请求字段的{operation_label}，"
                    "将只写入本次要求的结果，不生成完整字段统计。"
                ),
                "assumptions": [
                    "每张工作表的首行被视为字段名称。",
                ],
                "warnings": plan_warnings,
                "actions": _result_write_actions(
                    result_sheet,
                    start_cell,
                    target_kind,
                    ["字段", "指标", "结果", "有效数值数"],
                    result_rows,
                    2,
                ),
            }
        )
        return PlanResponse(plan=plan, provider="local")
    return AnswerResponse(
        provider="local",
        message=_bounded_lines(answers),
        resultContext=result_context,
    )


_ANALYSIS_INTENT_MARKERS = ("比较", "对比", "汇总", "统计", "分析", "异常", "分布")


def _is_analysis_request(prompt: str) -> bool:
    prompt = _normalize_lookup_text(prompt)
    if any(keyword in prompt for keyword in _ANALYSIS_INTENT_MARKERS):
        return True
    return _match_aggregate_operation(prompt) is not None


def _local_analysis(request: PlanRequest) -> AssistantResponse:
    source_sheets = [
        sheet
        for sheet in request.workbook.worksheets
        if sheet.name not in GENERATED_SHEET_NAMES and _has_data(sheet)
    ]
    followup = _local_result_followup(request)
    if followup is not None:
        return followup
    formula_plan = _local_formula_plan(request, source_sheets)
    if formula_plan is not None:
        return formula_plan
    address_plan = _local_address_plan(request)
    if address_plan is not None:
        return address_plan
    if not source_sheets:
        return AnswerResponse(
            provider="local",
            message=(
                "当前工作簿里没有可供分析的原始数据。"
                "请先把数据粘贴或导入到工作表，再告诉我想比较、汇总或查找什么。"
                "\n\n另外，本地 AI 模型尚未配置，所以普通知识问答目前也不会交给模型。"
            ),
        )

    edit = _local_edit_plan(request, source_sheets)
    if edit is not None:
        return edit

    remove_duplicates = _local_remove_duplicates_plan(request, source_sheets)
    if remove_duplicates is not None:
        return remove_duplicates

    location = _local_location(request, source_sheets)
    if location is not None:
        return location

    lookup = _local_lookup(request, source_sheets)
    if lookup is not None:
        return lookup

    extreme_record = _local_extreme_record_response(request, source_sheets)
    if extreme_record is not None:
        return extreme_record

    aggregate = _local_aggregate_response(request, source_sheets)
    if aggregate is not None:
        return aggregate

    if not _is_analysis_request(request.prompt):
        return AnswerResponse(
            provider="local",
            message=(
                "基础模式没有识别出可确定执行的命令。请直接写出工作表、"
                "字段或单元格地址，或者切换到已配置的模型理解开放式需求。"
            ),
        )

    result_sheet = "AI分析结果"
    stats_rows: list[list[object]] = []
    truncated_sheets: list[str] = []
    for sheet in source_sheets:
        if sheet.truncated:
            truncated_sheets.append(sheet.name)
        width = min(
            sheet.columnCount,
            max(
                len(sheet.headers),
                max((len(row) for row in sheet.dataRows), default=0),
            ),
        )
        for column_index in range(width):
            values = [
                row[column_index]
                for row in sheet.dataRows
                if column_index < len(row) and row[column_index] not in (None, "")
            ]
            numeric_values = [
                float(value)
                for value in values
                if isinstance(value, (int, float)) and not isinstance(value, bool)
            ]
            if not numeric_values:
                continue
            total = sum(numeric_values)
            stats_rows.append(
                [
                    sheet.name,
                    _column_name(sheet.headers, column_index),
                    len(values),
                    len(numeric_values),
                    round(total, 4),
                    round(total / len(numeric_values), 4),
                    min(numeric_values),
                    max(numeric_values),
                ]
            )

    if not stats_rows:
        stats_rows = [
            [
                sheet.name,
                "未发现数值字段",
                sum(
                    1
                    for row in sheet.dataRows
                    if any(value not in (None, "") for value in row)
                ),
                0,
                None,
                None,
                None,
                None,
            ]
            for sheet in source_sheets
        ]

    warnings = [
        "尚未配置 AI 模型；当前执行的是本地通用统计，不会假装理解开放式业务问题。"
    ]
    if truncated_sheets:
        warnings.append(
            f"工作表 {'、'.join(truncated_sheets)} 超过本地快照上限，统计只覆盖前 200 行。"
        )

    plan = AnalysisPlan.model_validate(
        {
            "id": f"local-{uuid.uuid4().hex[:10]}",
            "title": "本地数值字段统计",
            "summary": (
                "已根据实际读取到的工作表数据计算非空数量、数值数量、"
                "合计、平均值、最小值和最大值，并准备写入新的结果表。"
            ),
            "assumptions": [
                f"当前需求：{request.prompt}",
                "每张工作表的首行被视为字段名称。",
            ],
            "warnings": warnings,
            "actions": [
                {"type": "createWorksheet", "sheet": result_sheet},
                {
                    "type": "writeTable",
                    "sheet": result_sheet,
                    "startCell": "A1",
                    "headers": [
                        "工作表",
                        "字段",
                        "非空数量",
                        "数值数量",
                        "合计",
                        "平均值",
                        "最小值",
                        "最大值",
                    ],
                    "rows": stats_rows,
                },
                {"type": "setFont", "sheet": result_sheet, "range": "A1:H1", "bold": True},
                {"type": "autofit", "sheet": result_sheet, "range": "A:H"},
                {"type": "activateWorksheet", "sheet": result_sheet},
            ],
        }
    )
    return PlanResponse(plan=plan, provider="local")


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
