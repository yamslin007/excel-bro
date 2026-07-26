from __future__ import annotations

import base64
import io
import re
import shutil
import uuid
from copy import copy
from datetime import datetime
from pathlib import Path
from typing import Literal

from openpyxl import Workbook, load_workbook
from openpyxl.chart import AreaChart, BarChart, DoughnutChart, LineChart, PieChart, Reference, ScatterChart
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.filters import FilterColumn, Filters
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import absolute_coordinate, get_column_letter, quote_sheetname, range_boundaries
from openpyxl.workbook.defined_name import DefinedName
from pydantic import BaseModel, Field

from .models import (
    ActionExecutionResult,
    AnalysisPlan,
    VerificationCheck,
    VerificationReport,
    WorkbookSnapshot,
    WorksheetSnapshot,
)
from .capabilities import capability_int, capability_text

ROW_LIMIT = capability_int("snapshot", "dataRows")
COLUMN_LIMIT = capability_int("snapshot", "dataColumns")
FILE_LIMIT = capability_int("folder", "maxFiles")
SOURCE_SEPARATOR = " › "
SUPPORTED_SUFFIXES = {".xlsx", ".xlsm"}
OUTPUT_FILE_NAME = capability_text("folder", "outputFileName")


class FolderWorksheetInfo(BaseModel):
    name: str
    rowCount: int
    columnCount: int


class FolderFileInfo(BaseModel):
    id: str
    name: str
    relativePath: str
    worksheets: list[FolderWorksheetInfo] = Field(default_factory=list)
    error: str | None = None


class FolderCatalog(BaseModel):
    sessionId: str
    folderName: str
    folderPath: str
    files: list[FolderFileInfo]


class FolderSelection(BaseModel):
    fileId: str
    sheets: list[str] = Field(min_length=1)


class FolderSnapshotRequest(BaseModel):
    sessionId: str
    selections: list[FolderSelection] = Field(min_length=1)


class FolderExecuteRequest(BaseModel):
    sessionId: str
    plan: AnalysisPlan


class FolderExecuteResponse(BaseModel):
    filesModified: list[str]
    backups: list[str]
    actionResults: list[ActionExecutionResult]
    verification: VerificationReport


class _FolderSession:
    def __init__(self, root: Path, files: dict[str, Path]) -> None:
        self.root = root
        self.files = files
        self.selected_targets: dict[str, tuple[Path, str]] = {}


_sessions: dict[str, _FolderSession] = {}


def choose_folder() -> Path | None:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        selected = filedialog.askdirectory(title="选择包含 Excel 文件的文件夹")
    finally:
        root.destroy()
    return Path(selected).resolve() if selected else None


def _is_supported_file(path: Path) -> bool:
    return (
        path.is_file()
        and path.suffix.lower() in SUPPORTED_SUFFIXES
        and not path.name.startswith("~$")
        and ".excel-bro-backup-" not in path.name
    )


def scan_folder(root: Path) -> FolderCatalog:
    root = root.resolve()
    candidates = sorted(
        (path for path in root.rglob("*") if _is_supported_file(path)),
        key=lambda path: str(path.relative_to(root)).casefold(),
    )[:FILE_LIMIT]
    session_id = uuid.uuid4().hex
    session_files: dict[str, Path] = {}
    files: list[FolderFileInfo] = []

    for path in candidates:
        file_id = uuid.uuid4().hex
        session_files[file_id] = path
        relative_path = str(path.relative_to(root))
        try:
            workbook = load_workbook(
                path,
                read_only=True,
                data_only=True,
                keep_vba=path.suffix.lower() == ".xlsm",
            )
            worksheets = [
                FolderWorksheetInfo(
                    name=sheet.title,
                    rowCount=sheet.max_row,
                    columnCount=sheet.max_column,
                )
                for sheet in workbook.worksheets
            ]
            workbook.close()
            files.append(
                FolderFileInfo(
                    id=file_id,
                    name=path.name,
                    relativePath=relative_path,
                    worksheets=worksheets,
                )
            )
        except Exception as error:
            files.append(
                FolderFileInfo(
                    id=file_id,
                    name=path.name,
                    relativePath=relative_path,
                    error=f"无法读取：{error}",
                )
            )

    _sessions[session_id] = _FolderSession(root, session_files)
    return FolderCatalog(
        sessionId=session_id,
        folderName=root.name,
        folderPath=str(root),
        files=files,
    )


def select_and_scan_folder() -> FolderCatalog | None:
    selected = choose_folder()
    return scan_folder(selected) if selected else None


def _session(session_id: str) -> _FolderSession:
    try:
        return _sessions[session_id]
    except KeyError as error:
        raise ValueError("文件夹会话已失效，请重新选择文件夹") from error


def create_folder_snapshot(request: FolderSnapshotRequest) -> WorkbookSnapshot:
    session = _session(request.sessionId)
    snapshots: list[WorksheetSnapshot] = []
    selected_targets: dict[str, tuple[Path, str]] = {}
    active_name = ""

    for selection in request.selections:
        try:
            path = session.files[selection.fileId]
        except KeyError as error:
            raise ValueError("选择中包含未知文件，请重新扫描文件夹") from error
        workbook = load_workbook(
            path,
            read_only=True,
            data_only=True,
            keep_vba=path.suffix.lower() == ".xlsm",
        )
        available = set(workbook.sheetnames)
        relative_path = str(path.relative_to(session.root))
        for sheet_name in selection.sheets:
            if sheet_name not in available:
                workbook.close()
                raise ValueError(f"{relative_path} 中不存在工作表「{sheet_name}」")
            sheet = workbook[sheet_name]
            row_count = sheet.max_row
            column_count = sheet.max_column
            values = [
                list(row)
                for row in sheet.iter_rows(
                    min_row=1,
                    max_row=min(row_count, ROW_LIMIT + 1),
                    max_col=min(column_count, COLUMN_LIMIT),
                    values_only=True,
                )
            ]
            display_name = f"{relative_path}{SOURCE_SEPARATOR}{sheet_name}"
            selected_targets[display_name] = (path, sheet_name)
            snapshots.append(
                WorksheetSnapshot(
                    name=display_name,
                    sourceFile=relative_path,
                    sourceSheet=sheet_name,
                    usedRange=(
                        f"A1:{get_column_letter(column_count)}{row_count}"
                        if row_count and column_count
                        else None
                    ),
                    rowCount=row_count,
                    columnCount=column_count,
                    headers=list(values[0]) if values else [],
                    dataRows=[list(row) for row in values[1:]],
                    truncated=row_count > ROW_LIMIT + 1
                    or column_count > COLUMN_LIMIT,
                )
            )
            if not active_name:
                active_name = display_name
        workbook.close()

    if not snapshots:
        raise ValueError("请至少选择一个可读取的工作表")
    session.selected_targets = selected_targets
    return WorkbookSnapshot(
        name=f"文件夹：{session.root.name}",
        capturedAt=datetime.now().astimezone().isoformat(),
        activeWorksheet=active_name,
        worksheets=snapshots,
    )


def _target(
    session: _FolderSession, sheet_reference: str
) -> tuple[Path, str, Literal["source", "output"]]:
    if sheet_reference in session.selected_targets:
        path, sheet_name = session.selected_targets[sheet_reference]
        return path, sheet_name, "source"
    if SOURCE_SEPARATOR in sheet_reference:
        raise ValueError(f"计划试图操作未选择的工作表：{sheet_reference}")
    return session.root / OUTPUT_FILE_NAME, sheet_reference, "output"


def _load_for_write(path: Path) -> Workbook:
    if path.exists():
        return load_workbook(path, keep_vba=path.suffix.lower() == ".xlsm")
    return Workbook()


def _get_or_create_sheet(workbook: Workbook, sheet_name: str):
    if sheet_name in workbook.sheetnames:
        return workbook[sheet_name]
    if (
        len(workbook.sheetnames) == 1
        and workbook.active.title == "Sheet"
        and workbook.active.max_row == 1
        and workbook.active.max_column == 1
        and workbook.active["A1"].value is None
    ):
        workbook.active.title = sheet_name
        return workbook.active
    return workbook.create_sheet(sheet_name)


def _values_equal(actual: object, expected: object) -> bool:
    if actual in (None, "") and expected in (None, ""):
        return True
    if (
        isinstance(actual, (int, float))
        and not isinstance(actual, bool)
        and isinstance(expected, (int, float))
        and not isinstance(expected, bool)
    ):
        return float(actual) == float(expected)
    return actual == expected


def _iter_range(sheet, address: str):
    min_column, min_row, max_column, max_row = range_boundaries(address)
    return sheet.iter_rows(
        min_row=min_row,
        max_row=max_row,
        min_col=min_column,
        max_col=max_column,
    )


def _write_matrix(sheet, address: str, values: list[list[object]]) -> None:
    min_column, min_row, max_column, max_row = range_boundaries(address)
    if (
        max_row - min_row + 1 != len(values)
        or not values
        or max_column - min_column + 1 != len(values[0])
    ):
        raise ValueError(f"{address} 与写入数据尺寸不一致")
    for row_offset, row in enumerate(values):
        for column_offset, value in enumerate(row):
            sheet.cell(
                row=min_row + row_offset,
                column=min_column + column_offset,
                value=value,
            )


def _folder_mode_unsupported(action_type: str) -> None:
    raise ValueError(
        f"文件夹模式暂不支持 {action_type}；请打开工作簿后在 Excel 内执行该动作"
    )


def _verify_folder_plan(
    session: _FolderSession, plan: AnalysisPlan
) -> VerificationReport:
    readers: dict[Path, Workbook] = {}
    checks: list[VerificationCheck] = []
    try:
        for criterion in plan.acceptanceCriteria:
            path, sheet_name, _ = _target(session, criterion.sheet)
            if not path.exists():
                checks.append(
                    VerificationCheck(
                        criterion=criterion,
                        passed=False,
                        message=f"文件不存在：{path.name}",
                    )
                )
                continue

            workbook = readers.setdefault(
                path,
                load_workbook(
                    path,
                    data_only=False,
                    read_only=True,
                    keep_vba=path.suffix.lower() == ".xlsm",
                ),
            )
            if criterion.type == "worksheetMissing":
                missing = sheet_name not in workbook.sheetnames
                checks.append(
                    VerificationCheck(
                        criterion=criterion,
                        passed=missing,
                        message=(
                            f"工作表「{sheet_name}」已删除"
                            if missing
                            else f"工作表「{sheet_name}」仍然存在"
                        ),
                    )
                )
                continue
            if sheet_name not in workbook.sheetnames:
                checks.append(
                    VerificationCheck(
                        criterion=criterion,
                        passed=False,
                        message=f"未找到工作表「{sheet_name}」",
                    )
                )
                continue

            if criterion.type == "worksheetExists":
                checks.append(
                    VerificationCheck(
                        criterion=criterion,
                        passed=True,
                        message=f"工作表「{sheet_name}」存在",
                    )
                )
                continue

            sheet = workbook[sheet_name]
            min_column, min_row, max_column, max_row = range_boundaries(
                criterion.range
            )
            actual = [
                [
                    sheet.cell(row=row, column=column).value
                    for column in range(min_column, max_column + 1)
                ]
                for row in range(min_row, max_row + 1)
            ]
            if criterion.type == "rangeEmpty":
                passed = all(
                    value in (None, "") for row in actual for value in row
                )
            else:
                passed = len(actual) == len(criterion.expected) and all(
                    len(actual_row) == len(expected_row)
                    and all(
                        _values_equal(actual_value, expected_value)
                        for actual_value, expected_value in zip(
                            actual_row, expected_row
                        )
                    )
                    for actual_row, expected_row in zip(actual, criterion.expected)
                )
            checks.append(
                VerificationCheck(
                    criterion=criterion,
                    passed=passed,
                    message=(
                        (
                            f"「{sheet_name}」{criterion.range} 已清空"
                            if criterion.type == "rangeEmpty"
                            else (
                                f"「{sheet_name}」{criterion.range} 公式一致"
                                if criterion.type == "formulasEqual"
                                else f"「{sheet_name}」{criterion.range} 写入值一致"
                            )
                        )
                        if passed
                        else f"「{sheet_name}」{criterion.range} 与预期不一致"
                    ),
                    actual=actual,
                )
            )
    finally:
        for workbook in readers.values():
            workbook.close()

    return VerificationReport(
        passed=bool(checks) and all(check.passed for check in checks),
        checks=checks,
    )


def execute_folder_plan(request: FolderExecuteRequest) -> FolderExecuteResponse:
    session = _session(request.sessionId)
    workbooks: dict[Path, Workbook] = {}
    modified_sources: set[Path] = set()
    action_results: list[ActionExecutionResult] = []

    for index, action in enumerate(request.plan.actions):
        path, sheet_name, target_kind = _target(session, action.sheet)
        workbook = workbooks.setdefault(path, _load_for_write(path))
        if target_kind == "source":
            modified_sources.add(path)

        if action.type == "deleteWorksheet":
            if sheet_name not in workbook.sheetnames:
                raise ValueError(f"未找到工作表「{sheet_name}」")
            if len(workbook.sheetnames) == 1:
                raise ValueError("不能删除工作簿中唯一的工作表")
            workbook.remove(workbook[sheet_name])
            action_results.append(
                ActionExecutionResult(
                    index=index,
                    type=action.type,
                    sheet=action.sheet,
                    status="succeeded",
                )
            )
            continue

        if action.type == "splitGroupAggregate":
            if sheet_name not in workbook.sheetnames:
                raise ValueError(f"未找到源工作表「{sheet_name}」")
            source_sheet = workbook[sheet_name]
            if action.sourceRange:
                min_col, min_row, max_col, max_row = range_boundaries(
                    action.sourceRange
                )
            else:
                min_col, min_row = 1, 1
                max_col, max_row = source_sheet.max_column, source_sheet.max_row
            values = [
                [
                    source_sheet.cell(row=row, column=column).value
                    for column in range(min_col, max_col + 1)
                ]
                for row in range(min_row, max_row + 1)
            ]

            def normalized_header(value) -> str:
                return str(value or "").strip().casefold()

            required_fields = [
                action.splitBy,
                *action.groupBy,
                *[
                    metric.field
                    for metric in action.metrics
                    if metric.field is not None
                ],
            ]
            required_normalized = {
                normalized_header(field) for field in required_fields
            }
            header_offset = next(
                (
                    row_index
                    for row_index, row in enumerate(values)
                    if required_normalized.issubset(
                        {
                            normalized_header(value)
                            for value in row
                            if normalized_header(value)
                        }
                    )
                ),
                None,
            )
            if header_offset is None:
                raise ValueError(
                    f"未找到包含这些字段的表头：{'、'.join(required_fields)}"
                )
            header_indexes: dict[str, int] = {}
            for column_index, value in enumerate(values[header_offset]):
                normalized = normalized_header(value)
                if normalized and normalized not in header_indexes:
                    header_indexes[normalized] = column_index

            def field_index(field: str) -> int:
                try:
                    return header_indexes[normalized_header(field)]
                except KeyError as error:
                    raise ValueError(f"未找到字段「{field}」") from error

            split_index = field_index(action.splitBy)
            group_indexes = [field_index(field) for field in action.groupBy]
            metric_indexes = [
                field_index(metric.field) if metric.field else None
                for metric in action.metrics
            ]

            def is_blank(value) -> bool:
                return value is None or (
                    isinstance(value, str) and not value.strip()
                )

            def value_key(value):
                return (type(value).__name__, str(value).strip())

            groups: dict[tuple, dict] = {}
            splits: dict[tuple, dict] = {}
            for row in values[header_offset + 1 :]:
                split_value = row[split_index]
                group_values = [row[column] for column in group_indexes]
                if (
                    not action.includeBlankSplitValues
                    and is_blank(split_value)
                ) or any(is_blank(value) for value in group_values):
                    continue
                group_key = tuple(value_key(value) for value in group_values)
                split_key = (value_key(split_value),)
                group = groups.setdefault(
                    group_key,
                    {
                        "values": group_values,
                        "totals": [0.0 for _ in action.metrics],
                    },
                )
                split = splits.setdefault(
                    split_key, {"value": split_value, "pairs": {}}
                )
                pair = split["pairs"].setdefault(
                    group_key, [0.0 for _ in action.metrics]
                )
                for metric_index, metric in enumerate(action.metrics):
                    cell = (
                        None
                        if metric_indexes[metric_index] is None
                        else row[metric_indexes[metric_index]]
                    )
                    if metric.operation == "countRows":
                        increment = 1.0
                    elif metric.operation == "countNonBlank":
                        increment = 0.0 if is_blank(cell) else 1.0
                    else:
                        try:
                            increment = float(
                                str(cell or "").replace(",", "").strip()
                            )
                        except ValueError:
                            increment = 0.0
                    group["totals"][metric_index] += increment
                    pair[metric_index] += increment

            if not splits:
                raise ValueError("表头下没有可拆分的有效数据行")
            if len(splits) > action.maxOutputSheets:
                raise ValueError(
                    f"将生成 {len(splits)} 个工作表，"
                    f"超过安全上限 {action.maxOutputSheets} 个"
                )

            headers = [*action.groupBy, action.splitBy]
            ratio_columns: list[int] = []
            for metric in action.metrics:
                headers.append(metric.outputName)
                if metric.ratioOutputName:
                    headers.append(metric.ratioOutputName)
                    ratio_columns.append(len(headers))

            existing_names = {
                name.casefold(): name for name in workbook.sheetnames
            }
            generated_names: set[str] = set()

            def safe_name(value) -> str:
                cleaned = re.sub(
                    r"\s+",
                    " ",
                    re.sub(
                        r"[\x00-\x1f:：\\＼/／?？*＊\[\]［］]",
                        " ",
                        str(value or ""),
                    ),
                ).strip().strip("'").strip()
                cleaned = cleaned or "未命名"
                if cleaned.casefold() == "history":
                    cleaned = "History 结果"
                return cleaned[:31]

            def renamed(base: str) -> str:
                if base.casefold() not in existing_names:
                    return base
                suffix = 2
                while True:
                    marker = f" ({suffix})"
                    candidate = f"{base[:31 - len(marker)]}{marker}"
                    if candidate.casefold() not in existing_names:
                        return candidate
                    suffix += 1

            for split in splits.values():
                base_name = safe_name(split["value"])
                existing_name = existing_names.get(base_name.casefold())
                collides_with_current_run = base_name.casefold() in generated_names
                if (
                    existing_name
                    and action.existingSheetPolicy == "skip"
                    and not collides_with_current_run
                ):
                    continue
                if (
                    existing_name
                    and action.existingSheetPolicy == "replace"
                    and not collides_with_current_run
                ):
                    if existing_name.casefold() == sheet_name.casefold():
                        raise ValueError(
                            f"不能用拆分结果覆盖源工作表「{sheet_name}」"
                        )
                    position = workbook.sheetnames.index(existing_name)
                    workbook.remove(workbook[existing_name])
                    target_sheet = workbook.create_sheet(base_name, position)
                    target_name = base_name
                else:
                    target_name = renamed(base_name)
                    target_sheet = workbook.create_sheet(target_name)
                existing_names[target_name.casefold()] = target_name
                generated_names.add(target_name.casefold())

                target_sheet.append(headers)
                for group_key, aggregates in split["pairs"].items():
                    group = groups[group_key]
                    output_row = [*group["values"], split["value"]]
                    for metric_index, metric in enumerate(action.metrics):
                        aggregate = aggregates[metric_index]
                        if metric.operation != "sum" and aggregate.is_integer():
                            aggregate = int(aggregate)
                        output_row.append(aggregate)
                        if metric.ratioOutputName:
                            total = group["totals"][metric_index]
                            output_row.append(
                                None if total == 0 else aggregate / total
                            )
                    target_sheet.append(output_row)
                for cell in target_sheet[1]:
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill("solid", fgColor="DFF3E4")
                for column in ratio_columns:
                    for row in range(2, target_sheet.max_row + 1):
                        target_sheet.cell(row=row, column=column).number_format = (
                            "0.00%"
                        )
                for column in range(1, target_sheet.max_column + 1):
                    letter = get_column_letter(column)
                    width = max(
                        len(str(target_sheet.cell(row=row, column=column).value or ""))
                        for row in range(1, target_sheet.max_row + 1)
                    )
                    target_sheet.column_dimensions[letter].width = min(width + 2, 40)

            action_results.append(
                ActionExecutionResult(
                    index=index,
                    type=action.type,
                    sheet=action.sheet,
                    status="succeeded",
                )
            )
            continue

        sheet = _get_or_create_sheet(workbook, sheet_name)
        if action.type in {"createWorksheet", "activateWorksheet"}:
            if action.type == "activateWorksheet":
                workbook.active = workbook.sheetnames.index(sheet.title)
            action_results.append(
                ActionExecutionResult(
                    index=index,
                    type=action.type,
                    sheet=action.sheet,
                    status="succeeded",
                )
            )
            continue
        if action.type == "writeTable":
            rows = [action.headers, *action.rows]
            start_column, start_row, _, _ = range_boundaries(action.startCell)
            for row_offset, row in enumerate(rows):
                for column_offset, value in enumerate(row):
                    sheet.cell(
                        row=start_row + row_offset,
                        column=start_column + column_offset,
                        value=value,
                    )
        elif action.type == "writeValues":
            _write_matrix(sheet, action.range, action.values)
        elif action.type == "writeFormulas":
            _write_matrix(sheet, action.range, action.formulas)
        elif action.type == "clearRange":
            if action.applyTo in {"all", "contents"}:
                for row in _iter_range(sheet, action.range):
                    for cell in row:
                        cell.value = None
            if action.applyTo in {"all", "formats"}:
                for row in _iter_range(sheet, action.range):
                    for cell in row:
                        cell._style = copy(sheet["A1"]._style)
                        cell.number_format = "General"
                        cell.alignment = Alignment()
            if action.applyTo in {"all", "hyperlinks"}:
                for row in _iter_range(sheet, action.range):
                    for cell in row:
                        cell.hyperlink = None
        elif action.type in {"insertRange", "deleteRange"}:
            _folder_mode_unsupported(action.type)
        elif action.type == "copyRange":
            source_path, source_name, _ = _target(session, action.sourceSheet)
            source_book = workbooks.setdefault(source_path, _load_for_write(source_path))
            if source_name not in source_book.sheetnames:
                raise ValueError(f"未找到源工作表「{source_name}」")
            source_sheet = source_book[source_name]
            source_cells = [list(row) for row in _iter_range(source_sheet, action.sourceRange)]
            target_min_col, target_min_row, _, _ = range_boundaries(action.targetRange)
            if action.transpose:
                source_cells = [list(row) for row in zip(*source_cells)]
            for row_offset, row in enumerate(source_cells):
                for column_offset, source_cell in enumerate(row):
                    target_cell = sheet.cell(
                        row=target_min_row + row_offset,
                        column=target_min_col + column_offset,
                    )
                    if action.skipBlanks and source_cell.value in (None, ""):
                        continue
                    if action.copyType in {"all", "values", "formulas"}:
                        target_cell.value = source_cell.value
                    if action.copyType in {"all", "formats"}:
                        target_cell._style = copy(source_cell._style)
                        target_cell.number_format = source_cell.number_format
                    if action.copyType == "link":
                        target_cell.value = (
                            f"='{source_name}'!{source_cell.coordinate}"
                        )
        elif action.type == "sortRange":
            min_col, min_row, max_col, max_row = range_boundaries(action.range)
            first_data_row = min_row + (1 if action.hasHeaders else 0)
            rows = [
                [sheet.cell(row=row, column=col).value for col in range(min_col, max_col + 1)]
                for row in range(first_data_row, max_row + 1)
            ]
            for key in reversed(action.keys):
                if key.column >= len(rows[0]) if rows else False:
                    raise ValueError("排序列索引超出范围")
                rows.sort(
                    key=lambda row, column=key.column: (
                        row[column] is None,
                        str(row[column]).casefold(),
                    ),
                    reverse=not key.ascending,
                )
            for row_offset, values in enumerate(rows):
                for column_offset, value in enumerate(values):
                    sheet.cell(
                        row=first_data_row + row_offset,
                        column=min_col + column_offset,
                        value=value,
                    )
        elif action.type == "filterRange":
            sheet.auto_filter.ref = action.range
            sheet.auto_filter.filterColumn = [
                FilterColumn(
                    colId=action.column,
                    filters=Filters(filter=[str(value) for value in action.values]),
                )
            ]
        elif action.type == "clearFilter":
            sheet.auto_filter.ref = None
            sheet.auto_filter.filterColumn = []
        elif action.type == "setDataValidation":
            operator_map = {
                "equalTo": "equal",
                "notEqualTo": "notEqual",
                "greaterThanOrEqualTo": "greaterThanOrEqual",
                "lessThanOrEqualTo": "lessThanOrEqual",
            }
            validation_type = {
                "wholeNumber": "whole",
                "textLength": "textLength",
            }.get(action.validationType, action.validationType)
            if validation_type == "list":
                formula1 = '"' + ",".join(str(value) for value in action.values) + '"'
            else:
                formula1 = action.formula1
            validation = DataValidation(
                type=validation_type,
                operator=operator_map.get(action.operator, action.operator),
                formula1=formula1,
                formula2=action.formula2,
                allow_blank=action.allowBlank,
                prompt=action.prompt,
                error=action.errorMessage,
            )
            sheet.add_data_validation(validation)
            validation.add(action.range)
        elif action.type == "setConditionalFormat":
            fill = PatternFill(
                "solid", fgColor=(action.color or "#FFF2CC").removeprefix("#")
            )
            if action.ruleType == "colorScale":
                rule = ColorScaleRule(
                    start_type="min",
                    start_color=(action.minColor or "#F8696B").removeprefix("#"),
                    mid_type="percentile" if action.midColor else None,
                    mid_value=50 if action.midColor else None,
                    mid_color=action.midColor.removeprefix("#") if action.midColor else None,
                    end_type="max",
                    end_color=(action.maxColor or "#63BE7B").removeprefix("#"),
                )
            elif action.ruleType == "custom":
                rule = FormulaRule(formula=[str(action.formula1 or "")], fill=fill)
            else:
                rule = CellIsRule(
                    operator={
                        "equalTo": "equal",
                        "notEqualTo": "notEqual",
                        "greaterThanOrEqualTo": "greaterThanOrEqual",
                        "lessThanOrEqualTo": "lessThanOrEqual",
                    }.get(action.operator or "equalTo", action.operator or "equal"),
                    formula=[
                        str(action.formula1 or ""),
                        *(
                            [str(action.formula2)]
                            if action.formula2 is not None
                            else []
                        ),
                    ],
                    fill=fill,
                )
            sheet.conditional_formatting.add(action.range, rule)
        elif action.type == "setNumberFormat":
            for row in _iter_range(sheet, action.range):
                for cell in row:
                    cell.number_format = action.formatCode
        elif action.type == "setBorders":
            side = Side(
                style={
                    "continuous": "thin",
                    "dash": "dashed",
                    "dashDot": "dashDot",
                    "dot": "dotted",
                    "double": "double",
                    "none": None,
                }[action.style],
                color=action.color.removeprefix("#"),
            )
            for row in _iter_range(sheet, action.range):
                for cell in row:
                    current = cell.border
                    cell.border = Border(
                        left=side if "left" in action.sides else current.left,
                        right=side if "right" in action.sides else current.right,
                        top=side if "top" in action.sides else current.top,
                        bottom=side if "bottom" in action.sides else current.bottom,
                        vertical=side if "insideVertical" in action.sides else current.vertical,
                        horizontal=side if "insideHorizontal" in action.sides else current.horizontal,
                    )
        elif action.type == "setAlignment":
            for row in _iter_range(sheet, action.range):
                for cell in row:
                    cell.alignment = copy(cell.alignment)
                    if action.horizontal is not None:
                        cell.alignment = Alignment(
                            horizontal={
                                "centerAcrossSelection": "centerContinuous"
                            }.get(action.horizontal, action.horizontal),
                            vertical=cell.alignment.vertical,
                            wrap_text=cell.alignment.wrap_text,
                        )
                    if action.vertical is not None or action.wrapText is not None:
                        cell.alignment = Alignment(
                            horizontal=cell.alignment.horizontal,
                            vertical=action.vertical or cell.alignment.vertical,
                            wrap_text=(
                                action.wrapText
                                if action.wrapText is not None
                                else cell.alignment.wrap_text
                            ),
                        )
        elif action.type == "mergeCells":
            sheet.merge_cells(action.range)
        elif action.type == "unmergeCells":
            sheet.unmerge_cells(action.range)
        elif action.type == "resizeRange":
            min_col, min_row, max_col, max_row = range_boundaries(action.range)
            if action.rowHeight is not None:
                for row in range(min_row, max_row + 1):
                    sheet.row_dimensions[row].height = action.rowHeight
            if action.columnWidth is not None:
                for column in range(min_col, max_col + 1):
                    sheet.column_dimensions[get_column_letter(column)].width = action.columnWidth
        elif action.type == "freezePanes":
            if action.rows == 0 and action.columns == 0:
                sheet.freeze_panes = None
            else:
                sheet.freeze_panes = sheet.cell(
                    row=action.rows + 1, column=action.columns + 1
                )
        elif action.type == "setHyperlink":
            for row in _iter_range(sheet, action.range):
                for cell in row:
                    cell.hyperlink = action.address
                    if action.text is not None:
                        cell.value = action.text
        elif action.type in {"addComment", "addNote"}:
            sheet[action.cell].comment = Comment(action.text, "Excel Bro")
        elif action.type == "createTable":
            table = Table(
                displayName=action.name or f"ExcelBroTable{index + 1}",
                ref=action.range,
            )
            table.tableStyleInfo = TableStyleInfo(
                name=action.style or "TableStyleMedium2",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False,
            )
            sheet.add_table(table)
        elif action.type == "createChart":
            chart_class = {
                "ColumnClustered": BarChart,
                "BarClustered": BarChart,
                "Line": LineChart,
                "Pie": PieChart,
                "Doughnut": DoughnutChart,
                "XYScatter": ScatterChart,
                "Area": AreaChart,
            }.get(action.chartType)
            if chart_class is None:
                _folder_mode_unsupported(action.type)
            chart = chart_class()
            if action.chartType == "ColumnClustered":
                chart.type = "col"
            min_col, min_row, max_col, max_row = range_boundaries(action.sourceRange)
            chart.add_data(
                Reference(
                    sheet,
                    min_col=min_col,
                    max_col=max_col,
                    min_row=min_row,
                    max_row=max_row,
                ),
                titles_from_data=True,
            )
            if action.title:
                chart.title = action.title
            sheet.add_chart(chart, action.targetRange or f"{get_column_letter(max_col + 2)}{min_row}")
        elif action.type == "createPivotTable":
            _folder_mode_unsupported(action.type)
        elif action.type == "addNamedRange":
            workbook.defined_names.add(
                DefinedName(
                    action.name,
                    attr_text=(
                        f"{quote_sheetname(sheet.title)}!"
                        f"{absolute_coordinate(action.range)}"
                    ),
                    comment=action.comment,
                )
            )
        elif action.type == "addImage":
            try:
                from openpyxl.drawing.image import Image
                image = Image(io.BytesIO(base64.b64decode(action.base64)))
            except Exception as error:
                raise ValueError(f"无法解析图片数据：{error}") from error
            sheet.add_image(image, action.targetRange.split(":")[0])
        elif action.type == "addShape":
            _folder_mode_unsupported(action.type)
        elif action.type == "setFill":
            min_column, min_row, max_column, max_row = range_boundaries(action.range)
            fill = PatternFill("solid", fgColor=action.color.removeprefix("#"))
            for row in sheet.iter_rows(
                min_row=min_row,
                max_row=max_row,
                min_col=min_column,
                max_col=max_column,
            ):
                for cell in row:
                    cell.fill = fill
        elif action.type == "setFont":
            min_column, min_row, max_column, max_row = range_boundaries(action.range)
            for row in sheet.iter_rows(
                min_row=min_row,
                max_row=max_row,
                min_col=min_column,
                max_col=max_column,
            ):
                for cell in row:
                    cell.font = Font(
                        name=cell.font.name,
                        size=cell.font.size,
                        bold=action.bold
                        if action.bold is not None
                        else cell.font.bold,
                        italic=cell.font.italic,
                        color=action.color.removeprefix("#")
                        if action.color
                        else cell.font.color,
                    )
        elif action.type == "autofit":
            min_column, _, max_column, _ = range_boundaries(action.range)
            for column in range(min_column, max_column + 1):
                letter = get_column_letter(column)
                width = max(
                    (
                        len(str(cell.value))
                        for cell in sheet[letter]
                        if cell.value is not None
                    ),
                    default=8,
                )
                sheet.column_dimensions[letter].width = min(width + 2, 60)
        action_results.append(
            ActionExecutionResult(
                index=index,
                type=action.type,
                sheet=action.sheet,
                status="succeeded",
            )
        )

    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backups: list[str] = []
    modified: list[str] = []
    for path, workbook in workbooks.items():
        if path in modified_sources and path.exists():
            backup = path.with_name(
                f"{path.name}.excel-bro-backup-{timestamp}{path.suffix}"
            )
            shutil.copy2(path, backup)
            backups.append(str(backup.relative_to(session.root)))
        workbook.save(path)
        workbook.close()
        modified.append(str(path.relative_to(session.root)))

    verification = _verify_folder_plan(session, request.plan)
    return FolderExecuteResponse(
        filesModified=modified,
        backups=backups,
        actionResults=action_results,
        verification=verification,
    )
