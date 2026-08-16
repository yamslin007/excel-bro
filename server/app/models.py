from __future__ import annotations

import base64
import binascii
import re
from typing import Annotated, Literal

from pydantic import BaseModel, Field, StringConstraints, field_validator, model_validator

from .capabilities import capability_int
from .safety import (
    base64_within_image_limit,
    dangerous_formula,
    dangerous_hyperlink_address,
    max_image_bytes,
)

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


class WorksheetSnapshot(BaseModel):
    name: SheetName
    sourceFileId: str | None = Field(default=None, max_length=100)
    sourceSheetId: str | None = Field(default=None, max_length=100)
    sourceFile: str | None = Field(default=None, max_length=500)
    sourceSheet: str | None = Field(default=None, max_length=100)
    usedRange: str | None = None
    rowCount: int = Field(ge=0)
    columnCount: int = Field(ge=0)
    headers: list[CellValue] = Field(default_factory=list, max_length=100)
    dataRows: list[list[CellValue]] = Field(default_factory=list, max_length=200)
    displayRows: list[list[CellValue]] = Field(default_factory=list, max_length=200)
    truncated: bool = False


class WorkbookSnapshot(BaseModel):
    name: str = Field(min_length=1, max_length=255)
    capturedAt: str
    activeWorksheet: str
    selectedRange: RangeAddress | None = None
    sourceFingerprint: str | None = Field(default=None, min_length=1, max_length=128)
    sourceFingerprintSheets: list[SheetName] = Field(default_factory=list, max_length=100)
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
    filters: list[DataFilter] = Field(default_factory=list, max_length=20)
    hasHeaders: bool = True


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
    filters: list[DataFilter] = Field(default_factory=list, max_length=20)
    hasHeaders: bool = True


class CopyRangeAction(BaseModel):
    type: Literal["copyRange"]
    sheet: SheetName
    sourceSheet: SheetName
    sourceRange: RangeAddress
    targetRange: RangeAddress
    copyType: Literal["all", "values", "formulas", "formats", "link"] = "all"
    skipBlanks: bool = False
    transpose: bool = False
    filters: list[DataFilter] = Field(default_factory=list, max_length=20)
    hasHeaders: bool = True


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
        for row_index, row in enumerate(self.formulas):
            for column_index, formula in enumerate(row):
                matched = dangerous_formula(formula)
                if matched is not None:
                    raise ValueError(
                        f"writeFormulas 第 {row_index + 1} 行第 {column_index + 1}"
                        f" 列的公式包含被禁用的函数：{matched}"
                    )
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


class RemoveDuplicatesAction(BaseModel):
    type: Literal["removeDuplicates"]
    sheet: SheetName
    range: RangeAddress
    columns: list[int] = Field(min_length=1, max_length=100)
    hasHeaders: bool = True
    filters: list[DataFilter] = Field(default_factory=list, max_length=20)

    @field_validator("columns")
    @classmethod
    def validate_columns(cls, columns: list[int]) -> list[int]:
        if any(column < 0 or column > 16383 for column in columns):
            raise ValueError("removeDuplicates 的依据列索引必须在 0 到 16383 之间")
        return columns


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

    @model_validator(mode="after")
    def validate_address(self) -> "SetHyperlinkAction":
        matched = dangerous_hyperlink_address(self.address)
        if matched is not None:
            raise ValueError(f"超链接地址包含被禁用的注入载体：{matched}")
        return self


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

    @model_validator(mode="after")
    def validate_image_size(self) -> "AddImageAction":
        if not base64_within_image_limit(self.base64):
            raise ValueError(
                f"内嵌图片超过 {max_image_bytes() // (1024 * 1024)} MB 上限"
            )
        return self


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
    | RemoveDuplicatesAction
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


class RangeSortedCriterion(BaseModel):
    type: Literal["rangeSorted"]
    sheet: SheetName
    range: RangeAddress
    keys: list[SortKey] = Field(min_length=1, max_length=20)
    hasHeaders: bool = True


class FilterAppliedCriterion(BaseModel):
    type: Literal["filterApplied"]
    sheet: SheetName
    range: RangeAddress
    column: int = Field(ge=0, le=16383)
    values: list[CellValue] = Field(min_length=1, max_length=500)


class FilterClearedCriterion(BaseModel):
    type: Literal["filterCleared"]
    sheet: SheetName


class TableExistsCriterion(BaseModel):
    type: Literal["tableExists"]
    sheet: SheetName
    range: RangeAddress
    name: str | None = Field(default=None, min_length=1, max_length=255)
    hasHeaders: bool = True


class RangeFormatMatchesCriterion(BaseModel):
    type: Literal["rangeFormatMatches"]
    sheet: SheetName
    range: RangeAddress
    fillColor: Annotated[
        str, StringConstraints(pattern=r"^#[0-9A-Fa-f]{6}$")
    ] | None = None
    bold: bool | None = None
    fontColor: Annotated[
        str, StringConstraints(pattern=r"^#[0-9A-Fa-f]{6}$")
    ] | None = None
    numberFormat: str | None = Field(default=None, min_length=1, max_length=100)
    horizontal: str | None = Field(default=None, min_length=1, max_length=50)
    vertical: str | None = Field(default=None, min_length=1, max_length=50)
    wrapText: bool | None = None
    rowHeight: float | None = Field(default=None, ge=0, le=409)
    columnWidth: float | None = Field(default=None, ge=0, le=255)

    @model_validator(mode="after")
    def require_expected_property(self) -> "RangeFormatMatchesCriterion":
        fields = self.model_dump(exclude={"type", "sheet", "range"}, exclude_none=True)
        if not fields:
            raise ValueError("rangeFormatMatches 至少需要一个格式属性")
        return self


class BordersMatchCriterion(BaseModel):
    type: Literal["bordersMatch"]
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
    ]
    color: Annotated[str, StringConstraints(pattern=r"^#[0-9A-Fa-f]{6}$")]
    weight: Literal["hairline", "thin", "medium", "thick"]


class DataValidationMatchesCriterion(BaseModel):
    type: Literal["dataValidationMatches"]
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
    ]
    allowBlank: bool
    prompt: str | None = Field(default=None, max_length=255)
    errorMessage: str | None = Field(default=None, max_length=255)


class FreezePanesMatchesCriterion(BaseModel):
    type: Literal["freezePanesMatches"]
    sheet: SheetName
    rows: int = Field(ge=0, le=1000)
    columns: int = Field(ge=0, le=1000)


class ChartExistsCriterion(BaseModel):
    type: Literal["chartExists"]
    sheet: SheetName
    name: str | None = Field(default=None, min_length=1, max_length=255)
    chartType: str = Field(min_length=1, max_length=100)
    sourceRange: RangeAddress
    title: str | None = Field(default=None, max_length=255)
    targetRange: RangeAddress | None = None


class PivotTableExistsCriterion(BaseModel):
    type: Literal["pivotTableExists"]
    sheet: SheetName
    sourceSheet: SheetName
    sourceRange: RangeAddress
    name: str = Field(min_length=1, max_length=255)
    destinationCell: RangeAddress
    rowFields: list[str] = Field(default_factory=list, max_length=50)
    columnFields: list[str] = Field(default_factory=list, max_length=50)
    valueFields: list[PivotValueField] = Field(
        default_factory=list, max_length=50
    )


VerificationCriterion = Annotated[
    WorksheetExistsCriterion
    | WorksheetMissingCriterion
    | RangeEqualsCriterion
    | RangeEmptyCriterion
    | FormulasEqualCriterion
    | RangeSortedCriterion
    | FilterAppliedCriterion
    | FilterClearedCriterion
    | TableExistsCriterion
    | RangeFormatMatchesCriterion
    | BordersMatchCriterion
    | DataValidationMatchesCriterion
    | FreezePanesMatchesCriterion
    | ChartExistsCriterion
    | PivotTableExistsCriterion,
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
    sourceFingerprint: str | None = Field(default=None, min_length=1, max_length=128)
    sourceFingerprintSheets: list[SheetName] = Field(default_factory=list, max_length=100)
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
            | RangeSortedCriterion
            | FilterAppliedCriterion
            | FilterClearedCriterion
            | TableExistsCriterion
            | RangeFormatMatchesCriterion
            | BordersMatchCriterion
            | DataValidationMatchesCriterion
            | FreezePanesMatchesCriterion
            | ChartExistsCriterion
            | PivotTableExistsCriterion
        ] = list(self.acceptanceCriteria)
        criterion_keys = {
            criterion.model_dump_json(exclude_none=True) for criterion in criteria
        }

        def add_criterion(
            criterion: WorksheetExistsCriterion
            | WorksheetMissingCriterion
            | RangeEqualsCriterion
            | RangeEmptyCriterion
            | FormulasEqualCriterion
            | RangeSortedCriterion
            | FilterAppliedCriterion
            | FilterClearedCriterion
            | TableExistsCriterion
            | RangeFormatMatchesCriterion
            | BordersMatchCriterion
            | DataValidationMatchesCriterion
            | FreezePanesMatchesCriterion
            | ChartExistsCriterion
            | PivotTableExistsCriterion
        ) -> None:
            key = criterion.model_dump_json(exclude_none=True)
            if key not in criterion_keys:
                criteria.append(criterion)
                criterion_keys.add(key)

        for action_index, action in enumerate(self.actions):
            later_actions = self.actions[action_index + 1 :]

            def range_is_sorted_later(address: str) -> bool:
                normalized = address.replace("$", "").strip().upper()
                return any(
                    later.type == "sortRange"
                    and later.sheet.casefold() == action.sheet.casefold()
                    and later.range.replace("$", "").strip().upper() == normalized
                    for later in later_actions
                )

            filter_changes_later = any(
                later.type in {"filterRange", "clearFilter"}
                and later.sheet.casefold() == action.sheet.casefold()
                for later in later_actions
            )

            def later_same_range(action_type: str) -> list[ExcelAction]:
                normalized = getattr(action, "range", "").replace(
                    "$", ""
                ).strip().upper()
                return [
                    later
                    for later in later_actions
                    if later.type == action_type
                    and later.sheet.casefold() == action.sheet.casefold()
                    and getattr(later, "range", "")
                    .replace("$", "")
                    .strip()
                    .upper()
                    == normalized
                ]
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
            if action.type == "writeValues" and not range_is_sorted_later(
                action.range
            ):
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
            elif (
                action.type == "clearRange"
                and action.applyTo in {"all", "contents"}
                and not action.filters
            ):
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
                if target_range and not range_is_sorted_later(target_range):
                    add_criterion(
                        RangeEqualsCriterion(
                            type="rangeEquals",
                            sheet=action.sheet,
                            range=target_range,
                            expected=values,
                        )
                    )
            elif action.type == "sortRange":
                add_criterion(
                    RangeSortedCriterion(
                        type="rangeSorted",
                        sheet=action.sheet,
                        range=action.range,
                        keys=action.keys,
                        hasHeaders=action.hasHeaders,
                    )
                )
            elif action.type == "filterRange" and not filter_changes_later:
                add_criterion(
                    FilterAppliedCriterion(
                        type="filterApplied",
                        sheet=action.sheet,
                        range=action.range,
                        column=action.column,
                        values=action.values,
                    )
                )
            elif action.type == "clearFilter" and not filter_changes_later:
                add_criterion(
                    FilterClearedCriterion(
                        type="filterCleared",
                        sheet=action.sheet,
                    )
                )
            elif action.type == "createTable":
                add_criterion(
                    TableExistsCriterion(
                        type="tableExists",
                        sheet=action.sheet,
                        range=action.range,
                        name=action.name,
                        hasHeaders=action.hasHeaders,
                    )
                )
            elif action.type == "setFill" and not later_same_range("setFill"):
                add_criterion(
                    RangeFormatMatchesCriterion(
                        type="rangeFormatMatches",
                        sheet=action.sheet,
                        range=action.range,
                        fillColor=action.color,
                    )
                )
            elif action.type == "setFont":
                later_fonts = later_same_range("setFont")
                expected_bold = (
                    None
                    if any(later.bold is not None for later in later_fonts)
                    else action.bold
                )
                expected_color = (
                    None
                    if any(later.color is not None for later in later_fonts)
                    else action.color
                )
                if expected_bold is None and expected_color is None:
                    continue
                add_criterion(
                    RangeFormatMatchesCriterion(
                        type="rangeFormatMatches",
                        sheet=action.sheet,
                        range=action.range,
                        bold=expected_bold,
                        fontColor=expected_color,
                    )
                )
            elif action.type == "setNumberFormat" and not later_same_range(
                "setNumberFormat"
            ):
                add_criterion(
                    RangeFormatMatchesCriterion(
                        type="rangeFormatMatches",
                        sheet=action.sheet,
                        range=action.range,
                        numberFormat=action.formatCode,
                    )
                )
            elif action.type == "setAlignment":
                later_alignments = later_same_range("setAlignment")
                expected_horizontal = (
                    None
                    if any(
                        later.horizontal is not None
                        for later in later_alignments
                    )
                    else action.horizontal
                )
                expected_vertical = (
                    None
                    if any(
                        later.vertical is not None
                        for later in later_alignments
                    )
                    else action.vertical
                )
                expected_wrap_text = (
                    None
                    if any(
                        later.wrapText is not None
                        for later in later_alignments
                    )
                    else action.wrapText
                )
                if (
                    expected_horizontal is None
                    and expected_vertical is None
                    and expected_wrap_text is None
                ):
                    continue
                add_criterion(
                    RangeFormatMatchesCriterion(
                        type="rangeFormatMatches",
                        sheet=action.sheet,
                        range=action.range,
                        horizontal=expected_horizontal,
                        vertical=expected_vertical,
                        wrapText=expected_wrap_text,
                    )
                )
            elif action.type == "resizeRange":
                later_resizes = later_same_range("resizeRange")
                expected_row_height = (
                    None
                    if any(
                        later.rowHeight is not None for later in later_resizes
                    )
                    else action.rowHeight
                )
                expected_column_width = (
                    None
                    if any(
                        later.columnWidth is not None
                        for later in later_resizes
                    )
                    else action.columnWidth
                )
                if expected_row_height is None and expected_column_width is None:
                    continue
                add_criterion(
                    RangeFormatMatchesCriterion(
                        type="rangeFormatMatches",
                        sheet=action.sheet,
                        range=action.range,
                        rowHeight=expected_row_height,
                        columnWidth=expected_column_width,
                    )
                )
            elif action.type == "setBorders":
                later_borders = later_same_range("setBorders")
                expected_sides = [
                    side
                    for side in action.sides
                    if not any(
                        side in later.sides for later in later_borders
                    )
                ]
                if not expected_sides:
                    continue
                add_criterion(
                    BordersMatchCriterion(
                        type="bordersMatch",
                        sheet=action.sheet,
                        range=action.range,
                        sides=expected_sides,
                        style=action.style,
                        color=action.color,
                        weight=action.weight,
                    )
                )
            elif action.type == "setDataValidation" and not later_same_range(
                "setDataValidation"
            ):
                add_criterion(
                    DataValidationMatchesCriterion(
                        type="dataValidationMatches",
                        sheet=action.sheet,
                        range=action.range,
                        validationType=action.validationType,
                        values=action.values,
                        formula1=action.formula1,
                        formula2=action.formula2,
                        operator=action.operator,
                        allowBlank=action.allowBlank,
                        prompt=action.prompt,
                        errorMessage=action.errorMessage,
                    )
                )
            elif action.type == "freezePanes" and not any(
                later.type == "freezePanes"
                and later.sheet.casefold() == action.sheet.casefold()
                for later in later_actions
            ):
                add_criterion(
                    FreezePanesMatchesCriterion(
                        type="freezePanesMatches",
                        sheet=action.sheet,
                        rows=action.rows,
                        columns=action.columns,
                    )
                )
            elif action.type == "createPivotTable":
                add_criterion(
                    PivotTableExistsCriterion(
                        type="pivotTableExists",
                        sheet=action.sheet,
                        sourceSheet=action.sourceSheet,
                        sourceRange=action.sourceRange,
                        name=action.name,
                        destinationCell=action.destinationCell,
                        rowFields=action.rowFields,
                        columnFields=action.columnFields,
                        valueFields=action.valueFields,
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
    status: Literal["succeeded", "failed", "not_run"]
    message: str | None = Field(default=None, max_length=1000)


class VerificationCheck(BaseModel):
    criterion: VerificationCriterion
    passed: bool
    message: str
    actual: list[list[CellValue]] | None = None


class UnverifiedAction(BaseModel):
    index: int = Field(ge=0)
    type: str
    sheet: str
    message: str = Field(min_length=1, max_length=1000)


class VerificationReport(BaseModel):
    status: Literal["verified", "executed_unverified", "failed"]
    passed: bool
    checks: list[VerificationCheck]
    unverifiedActions: list[UnverifiedAction] = Field(default_factory=list)

    @model_validator(mode="after")
    def validate_status(self) -> "VerificationReport":
        if self.passed != (self.status == "verified"):
            raise ValueError("passed 必须与 verification status 一致")
        if self.status == "verified" and self.unverifiedActions:
            raise ValueError("verified 结果不能包含未验证动作")
        if self.status == "executed_unverified" and not self.unverifiedActions:
            raise ValueError("executed_unverified 必须说明未验证动作")
        return self


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
    # 全部数据工作表名（不含 #EB_* 系统表）。sheets 只含已选表的数据，
    # 结构操作（删除/重命名等）需引用 worksheetNames 才能规划。
    worksheetNames: list[SheetName] = Field(min_length=1, max_length=1000)
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


class DataCombine(BaseModel):
    mode: Literal["union", "deduplicate", "join"] = "union"
    deduplicateBy: list[str] = Field(default_factory=list, max_length=30)
    leftSourceSheetId: str | None = Field(default=None, max_length=100)
    rightSourceSheetId: str | None = Field(default=None, max_length=100)
    leftKey: str | None = Field(default=None, max_length=200)
    rightKey: str | None = Field(default=None, max_length=200)
    joinHow: Literal["inner", "left", "right", "outer"] = "inner"

    @model_validator(mode="after")
    def validate_combine(self) -> "DataCombine":
        if self.mode == "deduplicate" and not self.deduplicateBy:
            raise ValueError("deduplicate 合并需要 deduplicateBy")
        if self.mode == "join" and not all(
            (
                self.leftSourceSheetId,
                self.rightSourceSheetId,
                self.leftKey,
                self.rightKey,
            )
        ):
            raise ValueError("join 合并需要左右工作表 ID 和关联键")
        return self


class QueryTableArguments(BaseModel):
    mode: Literal["rows", "aggregate", "profile"]
    scope: Literal["selected", "active"] = "selected"
    fields: list[str] = Field(default_factory=list, max_length=30)
    filters: list[DataFilter] = Field(default_factory=list, max_length=20)
    groupBy: list[str] = Field(default_factory=list, max_length=10)
    metrics: list[DataMetric] = Field(default_factory=list, max_length=10)
    combine: DataCombine | None = None
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


class DeterministicQueryTemplate(BaseModel):
    id: str = Field(min_length=1, max_length=100)
    name: str = Field(min_length=1, max_length=100)
    description: str = Field(min_length=1, max_length=500)
    sourceMode: Literal["workbook", "folder"]
    request: DataToolRequest
    sourceSheetNames: list[str] = Field(default_factory=list, max_length=100)
    sourceSheetIds: list[str] = Field(default_factory=list, max_length=100)
    expectedHeaders: list[str] = Field(default_factory=list, max_length=50)


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


class ManagedModelConnectionResponse(BaseModel):
    id: str
    catalogModelId: str
    label: str
    baseUrl: str
    modelId: str
    supportsVision: bool
    apiKeyConfigured: bool
    apiKeyHint: str | None = None


class ModelSettingsResponse(BaseModel):
    baseUrl: str | None = None
    defaultModel: str | None = None
    apiKeyConfigured: bool
    apiKeyHint: str | None = None
    # /function 公式专用模型的 catalog id；空串=跟随全局选择。
    formulaModelId: str = ""
    connections: list[ManagedModelConnectionResponse] = Field(
        default_factory=list,
        max_length=50,
    )


class UpdateModelSettingsRequest(BaseModel):
    apiKey: Annotated[
        str,
        StringConstraints(strip_whitespace=True, min_length=1, max_length=4096),
    ]


class UpsertModelConnectionRequest(BaseModel):
    id: Annotated[
        str,
        StringConstraints(
            strip_whitespace=True,
            min_length=8,
            max_length=80,
            pattern=r"^[A-Za-z0-9_-]+$",
        ),
    ] | None = None
    label: Annotated[
        str,
        StringConstraints(strip_whitespace=True, min_length=1, max_length=80),
    ]
    baseUrl: Annotated[
        str,
        StringConstraints(strip_whitespace=True, min_length=8, max_length=500),
    ]
    modelId: Annotated[
        str,
        StringConstraints(strip_whitespace=True, min_length=1, max_length=200),
    ]
    apiKey: Annotated[
        str,
        StringConstraints(strip_whitespace=True, min_length=1, max_length=4096),
    ] | None = None
    clearApiKey: bool = False
    supportsVision: bool = False

    @field_validator("baseUrl")
    @classmethod
    def validate_base_url(cls, value: str) -> str:
        normalized = value.rstrip("/")
        if not normalized.startswith(("http://", "https://")):
            raise ValueError("服务地址必须以 http:// 或 https:// 开头")
        return normalized


class SetFormulaModelRequest(BaseModel):
    # /function 公式专用模型的 catalog id；空串=跟随全局选择（默认）。
    modelId: Annotated[
        str,
        StringConstraints(strip_whitespace=True, max_length=200),
    ] = ""


class TestModelConnectionResponse(BaseModel):
    ok: bool
    message: str


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


class TurnStepEvent(BaseModel):
    """流式分步进度事件；仅用于 /api/turn/stream 的 step 事件，不参与写入契约。"""

    phase: Literal["planning"] = "planning"
    title: str = Field(min_length=1, max_length=200)
    detail: str | None = Field(default=None, max_length=500)
    completedStep: str | None = Field(default=None, max_length=200)

TurnRequest = IntentCheckRequest | PlanRequest
TurnResponse = Annotated[
    IntentProceedResponse
    | IntentClarificationResponse
    | IntentToolResponse
    | PlanResponse
    | AnswerResponse,
    Field(discriminator="kind"),
]
