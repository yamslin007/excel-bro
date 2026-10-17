"""Compile a registered local result into a whitelist ``AnalysisPlan``.

The model proposes a *result reference* plus a *target*; this module turns that
proposal into an executable plan — or refuses. It never executes anything and
never calls a host setter: the produced plan still goes through the normal
preview → user confirmation → pre-write review → executor → readback path.

Boundaries kept here:

* Only ``writeValues`` on one already-authorized sheet is generated. No formula
  channel, no structural action, no coordinate/script text may ride along.
* The proposal is validated locally: matrix types and size, target range shape,
  target sheet authorization, host write limits and the result's bound source
  version. A stale result (source changed) is refused, not silently written.
* The compiled matrix stays on the trusted side until it is handed to the local
  preview; it is not sent to the model or a remote plan validator by this code.
"""

from __future__ import annotations

import uuid
from typing import Any

from .capabilities import capability_int
from .local_results import RegisteredResult
from .models import (
    AnalysisPlan,
    CellValue,
    PlanRequest,
    _range_dimensions,
)

SYSTEM_SHEET_NAMES = frozenset({"#eb_rules", "#eb_rules_backup"})

_EXCEL_MAX_ROWS = 1_048_576
_EXCEL_MAX_COLUMNS = 16_384


class ResultPlanError(ValueError):
    """A proposal that cannot be compiled into a safe plan."""

    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code


def _canonical(name: str) -> str:
    return name.strip().casefold()


def _resolve_authorized_sheet(request: PlanRequest, target_sheet: str) -> str:
    wanted = _canonical(target_sheet)
    for sheet in request.workbook.worksheets:
        if _canonical(sheet.name) == wanted:
            return sheet.name
        if sheet.sourceSheet and _canonical(sheet.sourceSheet) == wanted:
            return sheet.name
    raise ResultPlanError(
        "unauthorized_scope",
        "目标工作表不在本次已授权范围内，请先在数据范围中选中它。",
    )


def _assert_safe_values(values: list[list[CellValue]]) -> None:
    for row in values:
        for value in row:
            if not isinstance(value, str):
                continue
            if value.strip().startswith("="):
                # A compiled *values* write must never silently become a formula.
                raise ResultPlanError(
                    "formula_not_allowed",
                    "结果中含疑似公式文本，不能作为固定值写入；请改用公式动作或先处理该字段。",
                )
            if any(ord(character) < 32 and character not in "\t" for character in value):
                raise ResultPlanError(
                    "invalid_value", "结果中含不支持的控制字符。"
                )


def _blank_cell_of_write(value: CellValue) -> CellValue:
    """分析结果里的 ``""`` 折成 ``None``（写入意义上的"空白单元格"）。

    Python 分析算出的矩阵里，"这一格没有内容"常被表达成空字符串。但**单元格的
    空白只有一种**：空单元格。宿主侧的空白语义就是 ``None``/``null`` ——
    计划核心的 ``blankCellMatrix`` 也一律用 ``null`` 造空白。

    不折会在端上变成**双端分叉的真实故障**：WPS 把写入的 ``""`` 当"清空单元格"，
    回读得到 ``null``，而计划值仍是 ``""``，裸严格比较据判"回读与计划不一致"，
    拒绝写回（2026-10-10 实机验收复现）；Excel 端因为自己的值比对本就归一而能过。

    **只在结果计划编译这一处折**：它处理的是"分析产出的占位空白"，不是用户
    显式声明的"写入空字符串"（那是 ``BlankPolicy`` 里独立的一回事：``keep`` 把
    范围排除出写入、``clear`` 走 ``clearRange:contents``、写空串是第三种）。
    执行器层的回读比对因此可以保持严格 —— "写入没生效"仍会被如实拒绝。
    """
    return None if value == "" else value


def result_matrix(
    item: RegisteredResult, *, include_header: bool
) -> list[list[CellValue]]:
    rows = [[_blank_cell_of_write(value) for value in row] for row in item.rows]
    if include_header:
        return [
            [_blank_cell_of_write(header) for header in item.headers],
            *rows,
        ]
    return rows


def compile_result_plan(
    item: RegisteredResult,
    request: PlanRequest,
    *,
    target_sheet: str,
    target_range: str,
    include_header: bool = True,
) -> AnalysisPlan:
    """Validate the proposal and build a single-action ``writeValues`` plan."""
    if item.source_fingerprint is not None:
        current = request.workbook.sourceFingerprint
        if current is not None and current != item.source_fingerprint:
            raise ResultPlanError(
                "source_changed",
                "结果绑定的来源已变化，请重新查询后再写入。",
            )

    sheet = _resolve_authorized_sheet(request, target_sheet)
    if _canonical(sheet) in SYSTEM_SHEET_NAMES:
        raise ResultPlanError("unauthorized_scope", "不能写入内部系统表。")

    dimensions = _range_dimensions(target_range)
    if dimensions is None:
        raise ResultPlanError("invalid_target", "目标区域不是有效的 A1 范围。")
    row_span, column_span = dimensions
    if row_span > _EXCEL_MAX_ROWS or column_span > _EXCEL_MAX_COLUMNS:
        raise ResultPlanError("invalid_target", "目标区域超出工作表边界。")

    matrix = result_matrix(item, include_header=include_header)
    if not matrix or not matrix[0]:
        raise ResultPlanError("empty_result", "结果为空，没有可写入的内容。")
    matrix_rows = len(matrix)
    matrix_columns = len(matrix[0])
    if any(len(row) != matrix_columns for row in matrix):
        raise ResultPlanError("invalid_matrix", "结果不是规则二维数组。")
    if (row_span, column_span) != (matrix_rows, matrix_columns):
        raise ResultPlanError(
            "size_mismatch",
            f"目标区域是 {row_span} 行 × {column_span} 列，"
            f"与结果 {matrix_rows} 行 × {matrix_columns} 列不一致。",
        )

    max_cells = capability_int("agentToolBridge", "maxResultCells")
    if matrix_rows * matrix_columns > max_cells:
        raise ResultPlanError("too_large", "结果超过单次写入上限。")

    if request.hostId == "wps-et":
        wps_max = capability_int("wpsWrite", "maxCells")
        if wps_max and matrix_rows * matrix_columns > wps_max:
            raise ResultPlanError(
                "too_large",
                f"WPS 单次写入上限为 {wps_max} 个单元格，结果超出该上限。",
            )

    _assert_safe_values(matrix)

    # 实际来源依赖必须随计划一起走。少这一步，端上 `attestSourceDependency` 会把
    # 字段删掉（= 依赖不确定），一份来源明确的计划被当成不可恢复。来源取的是
    # **结果据以算出的**那批表（`item.source_sheets` 由可信协调器按实际输入句柄
    # 认定），不是本次授权范围 —— 两者口径不同，不能互相顶替。
    payload: dict[str, Any] = {
        "id": f"result-write-{uuid.uuid4().hex[:10]}",
        "title": f"写入 {sheet}!{target_range.strip().upper()}",
        "summary": (
            f"把本地结果「{item.title}」写入「{sheet}」"
            f"{target_range.strip().upper()}（{matrix_rows} 行 × {matrix_columns} 列）。"
        ),
        "warnings": ["执行后目标区域的现有内容会被覆盖。"],
        # 值只在本机内存；持久侧（规划账本 / 恢复描述）只留这个引用。
        "localResultRef": item.handle,
        "actions": [
            {
                "type": "writeValues",
                "sheet": sheet,
                "range": target_range.strip().upper().replace("$", ""),
                "values": matrix,
            }
        ],
    }
    sources = [name for name in item.source_sheets if isinstance(name, str) and name.strip()]
    if sources:
        payload["sourceDependencySheets"] = sources
    return AnalysisPlan.model_validate(payload)


def result_summary_for_answer(item: RegisteredResult) -> dict[str, Any]:
    """Local result-card payload (values included) for read-only display."""
    warnings = [warning for warning in item.warnings if isinstance(warning, str)]
    return {
        "title": item.title,
        "headers": list(item.headers),
        "rows": [list(row) for row in item.rows],
        "sourceSheets": list(item.source_sheets),
        "warnings": warnings,
    }
