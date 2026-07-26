from __future__ import annotations

import base64
import binascii
import re
from typing import Annotated, Literal

from pydantic import BaseModel, Field, StringConstraints, field_validator, model_validator

from .capabilities import capability_int

CellValue = str | int | float | bool | None
SheetName = Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=500)]
RangeAddress = Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=80)]
TurnId = Annotated[
    str,
    StringConstraints(
        strip_whitespace=True,
        min_length=8,
        max_length=100,
        pattern=r"^[A-Za-z0-9_-]+$",
    ),
]
MAX_CLARIFICATION_ROUNDS = capability_int(
    "conversation", "maxClarificationRounds"
)
INTENT_MAX_FIELDS = capability_int("intentContext", "maxFieldsPerSheet")
INTENT_MAX_PRIOR_RESULT_ROWS = capability_int(
    "intentContext", "maxPriorResultRows"
)


class WorksheetSnapshot(BaseModel):
    name: SheetName
    sourceFile: str | None = Field(default=None, max_length=500)
    sourceSheet: str | None = Field(default=None, max_length=100)
    usedRange: str | None = None
    rowCount: int = Field(ge=0)
    columnCount: int = Field(ge=0)
    headers: list[CellValue] = Field(default_factory=list, max_length=100)
    dataRows: list[list[CellValue]] = Field(default_factory=list, max_length=200)
    truncated: bool = False


class WorkbookSnapshot(BaseModel):
    name: str = Field(min_length=1, max_length=255)
    capturedAt: str
    activeWorksheet: str
    selectedRange: RangeAddress | None = None
    worksheets: list[WorksheetSnapshot] = Field(min_length=1, max_length=100)


class CreateWorksheetAction(BaseModel):
    type: Literal["createWorksheet"]
    sheet: SheetName


class WriteTableAction(BaseModel):
    type: Literal["writeTable"]
    sheet: SheetName
    startCell: RangeAddress
    headers: list[CellValue] = Field(min_length=1, max_length=50)
    rows: list[list[CellValue]] = Field(default_factory=list, max_length=500)

    @model_validator(mode="after")
    def validate_rows(self) -> "WriteTableAction":
        width = len(self.headers)
        if any(len(row) != width for row in self.rows):
            raise ValueError("writeTable 的每一行必须与 headers 等宽")
        return self


class WriteValuesAction(BaseModel):
    type: Literal["writeValues"]
    sheet: SheetName
    range: RangeAddress
    values: list[list[CellValue]] = Field(min_length=1, max_length=500)

    @model_validator(mode="after")
    def validate_matrix(self) -> "WriteValuesAction":
        width = len(self.values[0])
        if width == 0 or any(len(row) != width for row in self.values):
            raise ValueError("writeValues 必须是规则二维数组")
        return self


class SetFillAction(BaseModel):
    type: Literal["setFill"]
    sheet: SheetName
    range: RangeAddress
    color: Annotated[str, StringConstraints(pattern=r"^#[0-9A-Fa-f]{6}$")]


class SetFontAction(BaseModel):
    type: Literal["setFont"]
    sheet: SheetName
    range: RangeAddress
    bold: bool | None = None
    color: Annotated[str, StringConstraints(pattern=r"^#[0-9A-Fa-f]{6}$")] | None = None


class AutofitAction(BaseModel):
    type: Literal["autofit"]
    sheet: SheetName
    range: RangeAddress


class ActivateWorksheetAction(BaseModel):
    type: Literal["activateWorksheet"]
    sheet: SheetName


class DeleteWorksheetAction(BaseModel):
    type: Literal["deleteWorksheet"]
    sheet: SheetName


class ClearRangeAction(BaseModel):
    type: Literal["clearRange"]
    sheet: SheetName
    range: RangeAddress
    applyTo: Literal["all", "contents", "formats", "hyperlinks"] = "all"


class InsertRangeAction(BaseModel):
    type: Literal["insertRange"]
    sheet: SheetName
    range: RangeAddress
    shift: Literal["down", "right"]


class DeleteRangeAction(BaseModel):
    type: Literal["deleteRange"]
    sheet: SheetName
    range: RangeAddress
    shift: Literal["up", "left"]


class CopyRangeAction(BaseModel):
    type: Literal["copyRange"]
    sheet: SheetName
    sourceSheet: SheetName
    sourceRange: RangeAddress
    targetRange: RangeAddress
    copyType: Literal["all", "values", "formulas", "formats", "link"] = "all"
    skipBlanks: bool = False
    transpose: bool = False


class WriteFormulasAction(BaseModel):
    type: Literal["writeFormulas"]
    sheet: SheetName
    range: RangeAddress
    formulas: list[list[str]] = Field(min_length=1, max_length=500)

    @model_validator(mode="after")
    def validate_matrix(self) -> "WriteFormulasAction":
        width = len(self.formulas[0])
        if width == 0 or any(len(row) != width for row in self.formulas):
            raise ValueError("writeFormulas 必须是规则二维数组")
        return self


class SortKey(BaseModel):
    column: int = Field(ge=0, le=16383)
    ascending: bool = True


class SortRangeAction(BaseModel):
    type: Literal["sortRange"]
    sheet: SheetName
    range: RangeAddress
    keys: list[SortKey] = Field(min_length=1, max_length=20)
    hasHeaders: bool = True


class FilterRangeAction(BaseModel):
    type: Literal["filterRange"]
    sheet: SheetName
    range: RangeAddress
    column: int = Field(ge=0, le=16383)
    values: list[CellValue] = Field(min_length=1, max_length=500)


class ClearFilterAction(BaseModel):
    type: Literal["clearFilter"]
    sheet: SheetName


class SetDataValidationAction(BaseModel):
    type: Literal["setDataValidation"]
    sheet: SheetName
    range: RangeAddress
    validationType: Literal[
        "list", "wholeNumber", "decimal", "date", "textLength", "custom"
    ]
    values: list[CellValue] = Field(default_factory=list, max_length=500)
    formula1: str | int | float | None = None
    formula2: str | int | float | None = None
    operator: Literal[
        "between",
        "notBetween",
        "equalTo",
        "notEqualTo",
        "greaterThan",
        "lessThan",
        "greaterThanOrEqualTo",
        "lessThanOrEqualTo",
    ] = "between"
    allowBlank: bool = True
    prompt: str | None = Field(default=None, max_length=255)
    errorMessage: str | None = Field(default=None, max_length=255)


class SetConditionalFormatAction(BaseModel):
    type: Literal["setConditionalFormat"]
    sheet: SheetName
    range: RangeAddress
    ruleType: Literal["cellValue", "custom", "colorScale"]
    operator: Literal[
        "between",
        "notBetween",
        "equalTo",
        "notEqualTo",
        "greaterThan",
        "lessThan",
        "greaterThanOrEqualTo",
        "lessThanOrEqualTo",
    ] | None = None
    formula1: str | int | float | None = None
    formula2: str | int | float | None = None
    color: Annotated[str, StringConstraints(pattern=r"^#[0-9A-Fa-f]{6}$")] | None = None
    minColor: Annotated[str, StringConstraints(pattern=r"^#[0-9A-Fa-f]{6}$")] | None = None
    midColor: Annotated[str, StringConstraints(pattern=r"^#[0-9A-Fa-f]{6}$")] | None = None
    maxColor: Annotated[str, StringConstraints(pattern=r"^#[0-9A-Fa-f]{6}$")] | None = None


class SetNumberFormatAction(BaseModel):
    type: Literal["setNumberFormat"]
    sheet: SheetName
    range: RangeAddress
    formatCode: str = Field(min_length=1, max_length=100)


class SetBordersAction(BaseModel):
    type: Literal["setBorders"]
    sheet: SheetName
    range: RangeAddress
    sides: list[
        Literal[
            "top",
            "bottom",
            "left",
            "right",
            "insideHorizontal",
            "insideVertical",
        ]
    ] = Field(min_length=1, max_length=6)
    style: Literal[
        "continuous", "dash", "dashDot", "dot", "double", "none"
    ] = "continuous"
    color: Annotated[str, StringConstraints(pattern=r"^#[0-9A-Fa-f]{6}$")] = "#000000"
    weight: Literal["hairline", "thin", "medium", "thick"] = "thin"


class SetAlignmentAction(BaseModel):
    type: Literal["setAlignment"]
    sheet: SheetName
    range: RangeAddress
    horizontal: Literal[
        "general", "left", "center", "right", "fill", "justify", "centerAcrossSelection"
    ] | None = None
    vertical: Literal["top", "center", "bottom", "justify", "distributed"] | None = None
    wrapText: bool | None = None


class MergeCellsAction(BaseModel):
    type: Literal["mergeCells"]
    sheet: SheetName
    range: RangeAddress
    across: bool = False


class UnmergeCellsAction(BaseModel):
    type: Literal["unmergeCells"]
    sheet: SheetName
    range: RangeAddress


class ResizeRangeAction(BaseModel):
    type: Literal["resizeRange"]
    sheet: SheetName
    range: RangeAddress
    rowHeight: float | None = Field(default=None, ge=0, le=409)
    columnWidth: float | None = Field(default=None, ge=0, le=255)


class FreezePanesAction(BaseModel):
    type: Literal["freezePanes"]
    sheet: SheetName
    rows: int = Field(default=0, ge=0, le=1000)
    columns: int = Field(default=0, ge=0, le=1000)


class SetHyperlinkAction(BaseModel):
    type: Literal["setHyperlink"]
    sheet: SheetName
    range: RangeAddress
    address: str = Field(min_length=1, max_length=2048)
    text: str | None = Field(default=None, max_length=500)
    screenTip: str | None = Field(default=None, max_length=500)


class AddCommentAction(BaseModel):
    type: Literal["addComment"]
    sheet: SheetName
    cell: RangeAddress
    text: str = Field(min_length=1, max_length=4000)


class AddNoteAction(BaseModel):
    type: Literal["addNote"]
    sheet: SheetName
    cell: RangeAddress
    text: str = Field(min_length=1, max_length=4000)


class CreateTableAction(BaseModel):
    type: Literal["createTable"]
    sheet: SheetName
    range: RangeAddress
    name: str | None = Field(default=None, max_length=255)
    hasHeaders: bool = True
    style: str | None = Field(default=None, max_length=100)


class CreateChartAction(BaseModel):
    type: Literal["createChart"]
    sheet: SheetName
    sourceRange: RangeAddress
    chartType: Literal[
        "ColumnClustered", "BarClustered", "Line", "Pie", "Doughnut",
        "XYScatter", "Area", "Histogram", "Waterfall", "Funnel"
    ] = "ColumnClustered"
    title: str | None = Field(default=None, max_length=255)
    targetRange: RangeAddress | None = None


class PivotValueField(BaseModel):
    field: str = Field(min_length=1, max_length=255)
    aggregation: Literal["sum", "count", "average", "max", "min"] = "sum"


class CreatePivotTableAction(BaseModel):
    type: Literal["createPivotTable"]
    sheet: SheetName
    sourceSheet: SheetName
    sourceRange: RangeAddress
    name: str = Field(min_length=1, max_length=255)
    destinationCell: RangeAddress
    rowFields: list[str] = Field(default_factory=list, max_length=20)
    columnFields: list[str] = Field(default_factory=list, max_length=20)
    valueFields: list[PivotValueField] = Field(default_factory=list, max_length=20)


class SplitAggregateMetric(BaseModel):
    operation: Literal["countRows", "countNonBlank", "sum"]
    field: str | None = Field(default=None, min_length=1, max_length=255)
    outputName: str = Field(min_length=1, max_length=255)
    ratioOutputName: str | None = Field(default=None, min_length=1, max_length=255)

    @model_validator(mode="after")
    def validate_field(self) -> "SplitAggregateMetric":
        if self.operation != "countRows" and not self.field:
            raise ValueError(f"{self.operation} 聚合必须指定 field")
        return self


class SplitGroupAggregateAction(BaseModel):
    type: Literal["splitGroupAggregate"]
    sheet: SheetName
    sourceRange: RangeAddress | None = None
    splitBy: str = Field(min_length=1, max_length=255)
    groupBy: list[str] = Field(min_length=1, max_length=10)
    metrics: list[SplitAggregateMetric] = Field(min_length=1, max_length=10)
    includeBlankSplitValues: bool = False
    existingSheetPolicy: Literal["rename", "replace", "skip"] = "rename"
    maxOutputSheets: int = Field(default=200, ge=1, le=200)

    @model_validator(mode="after")
    def validate_configuration(self) -> "SplitGroupAggregateAction":
        fields = [field.strip() for field in self.groupBy]
        if any(not field for field in fields):
            raise ValueError("groupBy 不能包含空字段")
        if len(set(fields)) != len(fields):
            raise ValueError("groupBy 不能包含重复字段")
        if self.splitBy.strip() in set(fields):
            raise ValueError("splitBy 不能同时出现在 groupBy")
        output_names = [
            *fields,
            self.splitBy.strip(),
            *[
                name
                for metric in self.metrics
                for name in [metric.outputName, metric.ratioOutputName]
                if name
            ],
        ]
        if len(set(output_names)) != len(output_names):
            raise ValueError("拆分聚合的输出字段名称不能重复")
        return self


class AddNamedRangeAction(BaseModel):
    type: Literal["addNamedRange"]
    sheet: SheetName
    name: str = Field(min_length=1, max_length=255)
    range: RangeAddress
    comment: str | None = Field(default=None, max_length=500)


class AddImageAction(BaseModel):
    type: Literal["addImage"]
    sheet: SheetName
    base64: str = Field(min_length=1, max_length=8_000_000)
    targetRange: RangeAddress
    name: str | None = Field(default=None, max_length=255)


class AddShapeAction(BaseModel):
    type: Literal["addShape"]
    sheet: SheetName
    shapeType: Literal[
        "rectangle", "roundedRectangle", "ellipse", "line", "triangle", "diamond"
    ] = "rectangle"
    targetRange: RangeAddress
    text: str | None = Field(default=None, max_length=1000)
    fillColor: Annotated[str, StringConstraints(pattern=r"^#[0-9A-Fa-f]{6}$")] | None = None


ExcelAction = Annotated[
    CreateWorksheetAction
    | WriteTableAction
    | WriteValuesAction
    | SetFillAction
    | SetFontAction
    | AutofitAction
    | ActivateWorksheetAction
    | DeleteWorksheetAction
    | ClearRangeAction
    | InsertRangeAction
    | DeleteRangeAction
    | CopyRangeAction
    | WriteFormulasAction
    | SortRangeAction
    | FilterRangeAction
    | ClearFilterAction
    | SetDataValidationAction
    | SetConditionalFormatAction
    | SetNumberFormatAction
    | SetBordersAction
    | SetAlignmentAction
    | MergeCellsAction
    | UnmergeCellsAction
    | ResizeRangeAction
    | FreezePanesAction
    | SetHyperlinkAction
    | AddCommentAction
    | AddNoteAction
    | CreateTableAction
    | CreateChartAction
    | CreatePivotTableAction
    | SplitGroupAggregateAction
    | AddNamedRangeAction
    | AddImageAction
    | AddShapeAction,
    Field(discriminator="type"),
]


class WorksheetExistsCriterion(BaseModel):
    type: Literal["worksheetExists"]
    sheet: SheetName


class RangeEqualsCriterion(BaseModel):
    type: Literal["rangeEquals"]
    sheet: SheetName
    range: RangeAddress
    expected: list[list[CellValue]] = Field(min_length=1, max_length=500)

    @model_validator(mode="after")
    def validate_matrix(self) -> "RangeEqualsCriterion":
        width = len(self.expected[0])
        if (
            width == 0
            or width > 50
            or any(len(row) != width for row in self.expected)
        ):
            raise ValueError("rangeEquals.expected 必须是规则二维数组")
        dimensions = _range_dimensions(self.range)
        if dimensions != (len(self.expected), width):
            raise ValueError("rangeEquals.range 必须与 expected 的尺寸一致")
        return self


class WorksheetMissingCriterion(BaseModel):
    type: Literal["worksheetMissing"]
    sheet: SheetName


class RangeEmptyCriterion(BaseModel):
    type: Literal["rangeEmpty"]
    sheet: SheetName
    range: RangeAddress


class FormulasEqualCriterion(BaseModel):
    type: Literal["formulasEqual"]
    sheet: SheetName
    range: RangeAddress
    expected: list[list[str]] = Field(min_length=1, max_length=500)

    @model_validator(mode="after")
    def validate_matrix(self) -> "FormulasEqualCriterion":
        width = len(self.expected[0])
        if width == 0 or any(len(row) != width for row in self.expected):
            raise ValueError("formulasEqual.expected 必须是规则二维数组")
        dimensions = _range_dimensions(self.range)
        if dimensions != (len(self.expected), width):
            raise ValueError("formulasEqual.range 必须与 expected 的尺寸一致")
        return self


VerificationCriterion = Annotated[
    WorksheetExistsCriterion
    | WorksheetMissingCriterion
    | RangeEqualsCriterion
    | RangeEmptyCriterion
    | FormulasEqualCriterion,
    Field(discriminator="type"),
]


def _column_number(name: str) -> int:
    result = 0
    for character in name:
        result = result * 26 + ord(character.upper()) - ord("A") + 1
    return result


def _column_name(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def _matrix_range(start_cell: str, row_count: int, column_count: int) -> str | None:
    match = re.fullmatch(r"\$?([A-Za-z]+)\$?(\d+)", start_cell.strip())
    if not match:
        return None
    start_column = _column_number(match.group(1))
    start_row = int(match.group(2))
    end_column = _column_name(start_column + column_count - 1)
    end_row = start_row + row_count - 1
    return f"{match.group(1).upper()}{start_row}:{end_column}{end_row}"


def _range_dimensions(address: str) -> tuple[int, int] | None:
    match = re.fullmatch(
        r"\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?",
        address.strip(),
    )
    if not match:
        return None
    start_column = _column_number(match.group(1))
    start_row = int(match.group(2))
    end_column = _column_number(match.group(3) or match.group(1))
    end_row = int(match.group(4) or match.group(2))
    if end_column < start_column or end_row < start_row:
        return None
    return end_row - start_row + 1, end_column - start_column + 1


class AnalysisPlan(BaseModel):
    id: str = Field(min_length=1, max_length=100)
    title: str = Field(min_length=1, max_length=100)
    summary: str = Field(min_length=1, max_length=1000)
    assumptions: list[str] = Field(default_factory=list, max_length=20)
    warnings: list[str] = Field(default_factory=list, max_length=20)
    actions: list[ExcelAction] = Field(min_length=1, max_length=50)
    acceptanceCriteria: list[VerificationCriterion] = Field(
        default_factory=list, max_length=100
    )

    @model_validator(mode="after")
    def add_deterministic_acceptance_criteria(self) -> "AnalysisPlan":
        action_sheets = {action.sheet for action in self.actions}
        if self.acceptanceCriteria:
            if any(
                criterion.sheet not in action_sheets
                for criterion in self.acceptanceCriteria
            ):
                raise ValueError("验收条件只能检查计划实际操作的工作表")

        criteria: list[
            WorksheetExistsCriterion
            | WorksheetMissingCriterion
            | RangeEqualsCriterion
            | RangeEmptyCriterion
            | FormulasEqualCriterion
        ] = list(self.acceptanceCriteria)
        criterion_keys = {
            criterion.model_dump_json(exclude_none=True) for criterion in criteria
        }

        def add_criterion(
            criterion: WorksheetExistsCriterion
            | WorksheetMissingCriterion
            | RangeEqualsCriterion
            | RangeEmptyCriterion
            | FormulasEqualCriterion,
        ) -> None:
            key = criterion.model_dump_json(exclude_none=True)
            if key not in criterion_keys:
                criteria.append(criterion)
                criterion_keys.add(key)

        for action in self.actions:
            if action.type == "deleteWorksheet":
                add_criterion(
                    WorksheetMissingCriterion(
                        type="worksheetMissing",
                        sheet=action.sheet,
                    )
                )
                continue
            add_criterion(
                WorksheetExistsCriterion(
                    type="worksheetExists",
                    sheet=action.sheet,
                )
            )
            if action.type == "writeValues":
                add_criterion(
                    RangeEqualsCriterion(
                        type="rangeEquals",
                        sheet=action.sheet,
                        range=action.range,
                        expected=action.values,
                    )
                )
            elif action.type == "writeFormulas":
                add_criterion(
                    FormulasEqualCriterion(
                        type="formulasEqual",
                        sheet=action.sheet,
                        range=action.range,
                        expected=action.formulas,
                    )
                )
            elif action.type == "clearRange" and action.applyTo in {
                "all",
                "contents",
            }:
                add_criterion(
                    RangeEmptyCriterion(
                        type="rangeEmpty",
                        sheet=action.sheet,
                        range=action.range,
                    )
                )
            elif action.type == "writeTable":
                values = [action.headers, *action.rows]
                target_range = _matrix_range(
                    action.startCell, len(values), len(action.headers)
                )
                if target_range:
                    add_criterion(
                        RangeEqualsCriterion(
                            type="rangeEquals",
                            sheet=action.sheet,
                            range=target_range,
                            expected=values,
                        )
                    )
        destructive_types = {"deleteWorksheet", "clearRange", "deleteRange"}
        if any(action.type in destructive_types for action in self.actions) or any(
            action.type == "splitGroupAggregate"
            and action.existingSheetPolicy == "replace"
            for action in self.actions
        ):
            warning = "计划包含删除或清空操作；执行前请确认目标范围，并保留可恢复副本。"
            if warning not in self.warnings:
                self.warnings = [*self.warnings[:19], warning]
        for action in self.actions:
            if action.type != "splitGroupAggregate":
                continue
            warning = (
                f"将按「{action.splitBy}」创建多个结果工作表，"
                f"同名工作表处理方式为 {action.existingSheetPolicy}。"
            )
            if warning not in self.warnings:
                self.warnings = [*self.warnings[:19], warning]
        if len(criteria) > 100:
            raise ValueError("计划及自动补充的验收条件不能超过 100 项")
        self.acceptanceCriteria = criteria
        return self


class ActionExecutionResult(BaseModel):
    index: int = Field(ge=0)
    type: str
    sheet: str
    status: Literal["succeeded"]


class VerificationCheck(BaseModel):
    criterion: VerificationCriterion
    passed: bool
    message: str
    actual: list[list[CellValue]] | None = None


class VerificationReport(BaseModel):
    passed: bool
    checks: list[VerificationCheck]


class ResultContext(BaseModel):
    kind: Literal["table"] = "table"
    title: str = Field(min_length=1, max_length=100)
    headers: list[str] = Field(min_length=1, max_length=20)
    rows: list[list[CellValue]] = Field(min_length=1, max_length=500)
    primaryValueColumn: int | None = Field(default=None, ge=0, lt=20)
    sourceSheets: list[str] = Field(default_factory=list, max_length=100)
    warnings: list[str] = Field(default_factory=list, max_length=20)

    @model_validator(mode="after")
    def validate_rows(self) -> "ResultContext":
        width = len(self.headers)
        if any(len(row) != width for row in self.rows):
            raise ValueError("resultContext 的每一行必须与 headers 等宽")
        if self.primaryValueColumn is not None and self.primaryValueColumn >= width:
            raise ValueError("primaryValueColumn 超出 resultContext 表头范围")
        return self


class ImageAttachment(BaseModel):
    name: str = Field(min_length=1, max_length=255)
    mediaType: Literal["image/png", "image/jpeg", "image/webp"]
    data: str = Field(min_length=4, max_length=5_600_000)

    @field_validator("data")
    @classmethod
    def validate_base64_data(cls, value: str) -> str:
        try:
            decoded = base64.b64decode(value, validate=True)
        except (ValueError, binascii.Error) as error:
            raise ValueError("图片内容不是有效的 Base64 数据") from error
        if len(decoded) > 4 * 1024 * 1024:
            raise ValueError("单张图片不能超过 4 MB")
        return value


class IntentSheetContext(BaseModel):
    name: SheetName
    usedRange: str | None = None
    rowCount: int = Field(ge=0)
    columnCount: int = Field(ge=0)
    headers: list[CellValue] = Field(
        default_factory=list,
        max_length=INTENT_MAX_FIELDS,
    )


class IntentScopeContext(BaseModel):
    workbookName: str = Field(min_length=1, max_length=255)
    sourceMode: Literal["workbook", "folder"]
    selectionMode: Literal["auto", "manual", "folder"]
    activeWorksheet: str = Field(min_length=1, max_length=500)
    selectedRange: RangeAddress | None = None
    totalWorksheetCount: int = Field(ge=1, le=1000)
    sheets: list[IntentSheetContext] = Field(min_length=1, max_length=100)


class IntentConversationMessage(BaseModel):
    role: Literal["user", "assistant"]
    text: str = Field(min_length=1, max_length=1000)


class IntentMemory(BaseModel):
    confirmedPrompt: str = Field(min_length=2, max_length=4500)
    toolRequest: DataToolRequest | None = None


class IntentToolFailure(BaseModel):
    code: str = Field(min_length=1, max_length=100)
    message: str = Field(min_length=1, max_length=1000)
    retryable: bool = True
    availableFields: list[str] = Field(default_factory=list, max_length=100)
    request: DataToolRequest


class IntentCheckRequest(BaseModel):
    turnId: TurnId | None = None
    prompt: str = Field(min_length=2, max_length=4000)
    scope: IntentScopeContext
    imageCount: int = Field(default=0, ge=0, le=3)
    intentConfirmed: bool = False
    clarificationRound: int = Field(
        default=0,
        ge=0,
        le=MAX_CLARIFICATION_ROUNDS,
    )
    conversation: list[IntentConversationMessage] = Field(
        default_factory=list,
        max_length=6,
    )
    priorIntent: IntentMemory | None = None
    priorResult: ResultContext | None = None
    toolFailure: IntentToolFailure | None = None
    modelId: Annotated[
        str,
        StringConstraints(strip_whitespace=True, min_length=1, max_length=200),
    ] | None = None

    @field_validator("priorResult")
    @classmethod
    def validate_prior_result_size(
        cls, value: ResultContext | None
    ) -> ResultContext | None:
        if value and len(value.rows) > INTENT_MAX_PRIOR_RESULT_ROWS:
            raise ValueError(
                "上一轮结果超过意图上下文允许的最大行数"
            )
        return value


class IntentOption(BaseModel):
    id: str = Field(min_length=1, max_length=80)
    label: str = Field(min_length=1, max_length=80)
    description: str = Field(min_length=1, max_length=240)
    resolution: str = Field(min_length=1, max_length=500)
    action: Literal["resolve", "editScope"] = "resolve"


class DataFilter(BaseModel):
    field: str = Field(min_length=1, max_length=200)
    operator: Literal[
        "equals",
        "notEquals",
        "contains",
        "greaterThan",
        "greaterThanOrEqual",
        "lessThan",
        "lessThanOrEqual",
        "isBlank",
        "isNotBlank",
    ]
    value: CellValue = None


class DataMetric(BaseModel):
    operation: Literal[
        "countRows",
        "countDistinct",
        "sum",
        "average",
        "min",
        "max",
    ]
    field: str | None = Field(default=None, max_length=200)
    outputName: str = Field(min_length=1, max_length=200)
    ratioOutputName: str | None = Field(default=None, max_length=200)


class QueryTableArguments(BaseModel):
    mode: Literal["rows", "aggregate", "profile"]
    scope: Literal["selected", "active"] = "selected"
    fields: list[str] = Field(default_factory=list, max_length=30)
    filters: list[DataFilter] = Field(default_factory=list, max_length=20)
    groupBy: list[str] = Field(default_factory=list, max_length=10)
    metrics: list[DataMetric] = Field(default_factory=list, max_length=10)
    profileField: str | None = Field(default=None, max_length=200)
    sortBy: str | None = Field(default=None, max_length=200)
    sortDirection: Literal["asc", "desc"] = "desc"
    limit: int = Field(default=20, ge=1, le=200)

    @model_validator(mode="after")
    def validate_mode_requirements(self) -> "QueryTableArguments":
        if self.mode == "aggregate" and not self.metrics:
            raise ValueError("aggregate 查询至少需要一个指标")
        if self.mode == "profile" and not self.profileField:
            raise ValueError("profile 查询需要 profileField")
        for metric in self.metrics:
            if metric.operation != "countRows" and not metric.field:
                raise ValueError(f"{metric.operation} 指标需要 field")
            if (
                metric.ratioOutputName
                and metric.operation not in ("countRows", "sum")
            ):
                raise ValueError("占比只支持 countRows 或 sum 指标")
        return self


class DataToolRequest(BaseModel):
    id: str = Field(min_length=1, max_length=100)
    tool: Literal["query_table"] = "query_table"
    arguments: QueryTableArguments


class DataToolResult(BaseModel):
    requestId: str = Field(min_length=1, max_length=100)
    tool: Literal["query_table"] = "query_table"
    title: str = Field(min_length=1, max_length=300)
    headers: list[CellValue] = Field(min_length=1, max_length=50)
    rows: list[list[CellValue]] = Field(default_factory=list, max_length=500)
    sourceSheets: list[str] = Field(min_length=1, max_length=100)
    scannedRows: int = Field(ge=0)
    complete: bool
    calculation: str = Field(min_length=1, max_length=1000)
    warnings: list[str] = Field(default_factory=list, max_length=20)

    @model_validator(mode="after")
    def validate_result_rows(self) -> "DataToolResult":
        width = len(self.headers)
        if any(len(row) != width for row in self.rows):
            raise ValueError("数据工具结果的每一行必须与 headers 等宽")
        return self


class IntentClarification(BaseModel):
    id: str = Field(min_length=1, max_length=100)
    summary: str = Field(min_length=1, max_length=300)
    question: str = Field(min_length=1, max_length=300)
    reason: str = Field(min_length=1, max_length=300)
    scopeLabel: str = Field(min_length=1, max_length=300)
    options: list[IntentOption] = Field(min_length=2, max_length=4)


class IntentProceedResponse(BaseModel):
    kind: Literal["proceed"] = "proceed"
    summary: str = Field(min_length=1, max_length=500)
    confirmedPrompt: str = Field(min_length=2, max_length=4500)
    provider: Literal["model", "local"]
    turnId: TurnId | None = None


class IntentClarificationResponse(BaseModel):
    kind: Literal["clarification"] = "clarification"
    clarification: IntentClarification
    provider: Literal["model", "local"]
    turnId: TurnId | None = None


class IntentToolResponse(BaseModel):
    kind: Literal["tool_request"] = "tool_request"
    summary: str = Field(min_length=1, max_length=500)
    confirmedPrompt: str = Field(min_length=2, max_length=4500)
    request: DataToolRequest
    provider: Literal["model", "local"]
    turnId: TurnId | None = None


IntentCheckResponse = Annotated[
    IntentProceedResponse | IntentClarificationResponse | IntentToolResponse,
    Field(discriminator="kind"),
]


class PlanRequest(BaseModel):
    turnId: TurnId | None = None
    prompt: str = Field(min_length=2, max_length=4000)
    workbook: WorkbookSnapshot
    lastResult: ResultContext | None = None
    images: list[ImageAttachment] = Field(default_factory=list, max_length=3)
    dataResults: list[DataToolResult] = Field(default_factory=list, max_length=5)
    modelId: Annotated[
        str,
        StringConstraints(strip_whitespace=True, min_length=1, max_length=200),
    ] | None = None


class PlanResponse(BaseModel):
    kind: Literal["plan"] = "plan"
    plan: AnalysisPlan
    provider: Literal["model", "local"]
    turnId: TurnId | None = None


class AnswerResponse(BaseModel):
    kind: Literal["answer"] = "answer"
    message: str = Field(min_length=1, max_length=4000)
    provider: Literal["model", "local"]
    resultContext: ResultContext | None = None
    turnId: TurnId | None = None


AssistantResponse = Annotated[
    PlanResponse | AnswerResponse,
    Field(discriminator="kind"),
]

TurnRequest = IntentCheckRequest | PlanRequest
TurnResponse = Annotated[
    IntentProceedResponse
    | IntentClarificationResponse
    | IntentToolResponse
    | PlanResponse
    | AnswerResponse,
    Field(discriminator="kind"),
]
