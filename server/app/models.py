from __future__ import annotations

import base64
import binascii
import re
from typing import Annotated, Any, Literal

from pydantic import (
    BaseModel,
    ConfigDict,
    Field,
    StringConstraints,
    field_validator,
    model_validator,
)

from .capabilities import capability_int
from .safety import (
    base64_within_image_limit,
    dangerous_formula,
    dangerous_hyperlink_address,
    max_image_bytes,
)

CellValue = str | int | float | bool | None

# Stable service-error codes, shared by JSON/SSE and both host clients.
ModelUsageErrorCode = Literal[
    "MODEL_QUOTA_EXHAUSTED", "MODEL_CONTEXT_LIMIT", "MODEL_OUTPUT_LIMIT",
    "MODEL_BUDGET_LIMIT", "MODEL_TURN_LIMIT",
]

ScopeAdmissionReason = Literal["none", "read_only", "protected", "write_unverified", "host_write_blocked",
    "identity_unverified", "target_changed", "missing", "unreadable", "unsupported", "locked", "directory_not_writable"]


class ScopeAdmissionItem(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)
    id: str
    status: Literal["ready", "read_only", "unavailable"]
    reason: ScopeAdmissionReason


class FolderAdmissionSelection(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)
    fileId: str = Field(min_length=1, max_length=256)
    sheets: list[str] = Field(min_length=1)


class FolderAdmissionRequest(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)
    sessionId: str = Field(min_length=1, max_length=256)
    selections: list[FolderAdmissionSelection] = Field(min_length=1)

    @model_validator(mode="after")
    def validate_selection(self):
        keys = [(selection.fileId, name) for selection in self.selections for name in selection.sheets]
        if len(keys) > capability_int("scopeAdmission", "maxSelections") or len(set(keys)) != len(keys):
            raise ValueError("所选范围过多或重复")
        if any(not name.strip() or len(name) > 31 or "\x00" in name for _, name in keys):
            raise ValueError("工作表名称无效")
        if len({s.fileId for s in self.selections}) != len(self.selections):
            raise ValueError("文件选择重复")
        return self


class FolderAdmissionResponse(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)
    items: list[ScopeAdmissionItem]
    fingerprint: str | None = None


class PythonDataSource(BaseModel):
    """Metadata for one exact authorized range; never a file path."""
    model_config = ConfigDict(extra="forbid", strict=True)
    sheetId: str = Field(min_length=1, max_length=80)
    sheetName: str = Field(min_length=1, max_length=31)
    range: str = Field(pattern=r"^[A-Z]{1,3}[1-9][0-9]{0,6}(?::[A-Z]{1,3}[1-9][0-9]{0,6})?$")


class PythonDataScope(BaseModel):
    """Local-only protocol. Not a field of PlanRequest/history/checkpoints."""
    model_config = ConfigDict(extra="forbid", strict=True)
    version: Literal[1]
    taskId: str = Field(pattern=r"^[A-Za-z0-9_-]{1,80}$")
    hostId: Literal["excel-office", "wps-et"]
    workbookIdentityHash: str = Field(pattern=r"^[0-9a-f]{64}$")
    authorizationVersion: str = Field(pattern=r"^[0-9a-f]{64}$")
    sources: list[PythonDataSource] = Field(min_length=1, max_length=8)


class PythonDataScopeRequest(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)
    scope: PythonDataScope


class PythonDataSnapshotRequest(PythonDataScopeRequest):
    scopeToken: str = Field(pattern=r"^[A-Za-z0-9_-]{43}$")
    sourceIndex: int = Field(ge=0, le=7)
    # Matrix shape/scalars/finite numbers and byte bounds are checked by the
    # trusted registry. FastAPI must not echo validation input in HTTP errors.
    values: list[list[Any]]


class PythonDataRevokeRequest(PythonDataScopeRequest):
    scopeToken: str = Field(pattern=r"^[A-Za-z0-9_-]{43}$")


class PythonDatasetReceipt(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)
    datasetHandle: str = Field(pattern=r"^dataset://[A-Za-z0-9_-]{43}$")
    contentFingerprint: str = Field(pattern=r"^[0-9a-f]{64}$")


class PythonResultReceipt(BaseModel):
    """Local-only reference, never a model tool result or execution permission."""
    model_config = ConfigDict(extra="forbid", strict=True)
    resultHandle: str = Field(pattern=r"^result://[A-Za-z0-9_-]{43}$")
    contentFingerprint: str = Field(pattern=r"^[0-9a-f]{64}$")
    rows: int = Field(ge=1, le=500)
    columns: int = Field(ge=1, le=50)


class PythonResultReadRequest(PythonDataRevokeRequest):
    datasets: list[PythonDatasetReceipt] = Field(min_length=1, max_length=8)
    result: PythonResultReceipt


class PythonMatrixResult(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)
    kind: Literal["matrix"]
    # Trust-boundary validation is in python_results; never coerce nested values.
    values: list[list[Any]]


class ExcelReviewLayoutPrepareRequest(BaseModel):
    """Window metadata only. Never accept paths, plans, values or window handles."""
    model_config = ConfigDict(extra="forbid")
    workbookPathHash: str = Field(pattern=r"^[0-9a-f]{64}$")
    worksheetNameHash: str = Field(pattern=r"^[0-9a-f]{64}$")
    copySheetName: str = Field(pattern=r"^EB_[0-9a-f]{20}$")


class ExcelReviewPaneMeasurement(BaseModel):
    model_config = ConfigDict(extra="forbid")
    # Original window's client width, kept for compatibility/diagnostics.
    clientWidthCssPx: float = Field(gt=0, le=32768, allow_inf_nan=False)
    # The whole monitor work area (taskbar excluded). The AI pane target is computed
    # from THIS, never from a single document window, so both hosts share one meaning.
    workAreaWidthCssPx: float = Field(gt=0, le=32768, allow_inf_nan=False)


class ExcelReviewLayoutTicket(BaseModel):
    model_config = ConfigDict(extra="forbid")
    token: str = Field(pattern=r"^[A-Za-z0-9_-]{43}$")
    layoutVersion: Literal[3]


class ExcelReviewLayoutArrangeRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    token: str = Field(pattern=r"^[A-Za-z0-9_-]{43}$")


class ExcelReviewLayoutResult(BaseModel):
    model_config = ConfigDict(extra="forbid")
    status: Literal["arranged"]
    layoutVersion: Literal[3]
    copySide: Literal["left"]
    originalSide: Literal["right"]
    synchronousScrolling: Literal["enabled", "different_instance", "unavailable"]


class ExcelReviewCopyCloseRequest(BaseModel):
    """Window metadata only. Never accept paths, plans, values or window handles."""
    model_config = ConfigDict(extra="forbid")
    closeVersion: Literal[4]
    workbookPathHash: str = Field(pattern=r"^[0-9a-f]{64}$")
    worksheetNameHash: str = Field(pattern=r"^[0-9a-f]{64}$")
    copySheetName: str = Field(pattern=r"^EB_[0-9a-f]{20}$")


class ExcelReviewCopyCloseResult(BaseModel):
    """Outcome of asking the host to close ONE review copy.

    Deliberately says nothing about the original workbook: whether the writeback
    succeeded is a business fact owned by the add-in, and a close failure must
    never be allowed to overwrite it. `already_closed` is a SUCCESS for cleanup
    (idempotent repeat cleanup, or the user closed the copy themselves).

    `close_unconfirmed` means exactly one thing: the close request went out and the
    window is still there when the poll deadline expired. This bridge has **no dialog
    evidence**, so it never claims the host is asking to save — a slow close or an
    unrelated modal looks identical from here.
    """
    model_config = ConfigDict(extra="forbid")
    status: Literal["closed", "already_closed", "close_unconfirmed", "user_saved",
                    "identity_lost", "ambiguous", "unsupported", "failed"]
    layoutVersion: Literal[4]
    copyClosed: bool
    recoveryCopyPath: str | None = None

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
    catalogOnly: bool = False
    capturedAt: str
    activeWorksheet: str
    selectedRange: RangeAddress | None = None
    sourceFingerprint: str | None = Field(default=None, min_length=1, max_length=128)
    sourceFingerprintSheets: list[SheetName] = Field(default_factory=list, max_length=100)
    # 稳定的工作簿身份，沿用既有 ``url:`` / ``wps-path:`` 命名空间约定。
    # 名字相同但身份不同的两个工作簿必须彼此隔离：桥的归属判定、结果归属
    # 与跨会话拒绝都以此为准，不能只比 ``name``。
    workbookIdentity: str | None = Field(default=None, min_length=5, max_length=4000)
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
    """字体动作。只改显式给出的属性，其余属性与值/公式保持不变。"""

    type: Literal["setFont"]
    sheet: SheetName
    range: RangeAddress
    bold: bool | None = None
    italic: bool | None = None
    name: Annotated[
        str, StringConstraints(min_length=1, max_length=255)
    ] | None = None
    size: float | None = Field(default=None, ge=1, le=409)
    color: Annotated[str, StringConstraints(pattern=r"^#[0-9A-Fa-f]{6}$")] | None = None

    @model_validator(mode="after")
    def require_expected_property(self) -> "SetFontAction":
        if all(
            value is None
            for value in (self.bold, self.italic, self.name, self.size, self.color)
        ):
            raise ValueError("setFont 至少需要一个要修改的字体属性")
        return self


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

    @model_validator(mode="after")
    def reject_duplicate_sides(self) -> "SetBordersAction":
        # 重复的边不会报错，只会让同一位置被画两次 —— 计划里看不出问题，
        # 但期望值与实际落盘的边数对不上。前端 `assertFormatActionFields`
        # 同样拒绝，两端必须一致。
        if len(set(self.sides)) != len(self.sides):
            raise ValueError("setBorders 里有重复的边")
        return self


class SetAlignmentAction(BaseModel):
    """对齐动作。只改显式给出的属性，其余属性与值/公式保持不变。"""

    type: Literal["setAlignment"]
    sheet: SheetName
    range: RangeAddress
    horizontal: Literal[
        "general", "left", "center", "right", "fill", "justify", "centerAcrossSelection"
    ] | None = None
    vertical: Literal["top", "center", "bottom", "justify", "distributed"] | None = None
    wrapText: bool | None = None

    @model_validator(mode="after")
    def require_expected_property(self) -> "SetAlignmentAction":
        if self.horizontal is None and self.vertical is None and self.wrapText is None:
            raise ValueError("setAlignment 至少需要一个要修改的对齐属性")
        return self


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

    @model_validator(mode="after")
    def require_expected_property(self) -> "ResizeRangeAction":
        if self.rowHeight is None and self.columnWidth is None:
            raise ValueError("resizeRange 至少需要一个行高或列宽")
        return self


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
    splitBy: str = Field(
        min_length=1,
        max_length=255,
        description="输出工作表的拆分字段；执行器会自动把它写入结果列。",
    )
    groupBy: list[str] = Field(
        min_length=1,
        max_length=10,
        description="每张结果表内的分组字段；不得重复包含 splitBy。",
    )
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


class ComputeValuesAction(BaseModel):
    """执行时求值的查询动作。

    与 writeValues / writeTable 的根本区别：动作体里**只带查询配置，不带结果**。
    结果由 Excel 端在执行那一刻读取源工作表的实时数据算出来，所以源数据改了之后
    这份计划依然成立、并且会按新数据重算 —— 这正是它存在的理由。

    因此输出的行数在生成计划时**无法确定**，绝不能给它自动补 rangeEquals 验收条件
    （RangeEqualsCriterion 会校验 range 尺寸必须等于 expected 尺寸，见其
    validate_matrix），只补 worksheetExists。参见
    add_deterministic_acceptance_criteria 里的对应分支。
    """

    type: Literal["computeValues"]
    sheet: SheetName
    sourceSheet: SheetName
    sourceRange: RangeAddress | None = None
    hasHeaders: bool = True
    query: QueryTableArguments
    destinationCell: RangeAddress
    includeHeaders: bool = True

    @model_validator(mode="after")
    def validate_configuration(self) -> "ComputeValuesAction":
        if _range_dimensions(self.destinationCell) != (1, 1):
            raise ValueError("computeValues.destinationCell 必须是单个单元格")
        if self.query.combine is not None:
            raise ValueError("computeValues 暂不支持多表合并，请改用文件夹模式")
        referenced = [
            *self.query.fields,
            *self.query.groupBy,
            *[filter_.field for filter_ in self.query.filters],
            *[metric.field for metric in self.query.metrics if metric.field],
        ]
        if any(not str(field).strip() for field in referenced):
            raise ValueError("computeValues 引用的字段名不能为空")
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
    | ComputeValuesAction
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
    fontName: Annotated[
        str, StringConstraints(min_length=1, max_length=255)
    ] | None = None
    fontSize: float | None = Field(default=None, ge=1, le=409)
    fontItalic: bool | None = None
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
    # 指纹覆盖范围：可能宽于计划的实际依赖（端上取的是本次授权范围）。
    sourceFingerprintSheets: list[SheetName] = Field(default_factory=list, max_length=100)
    # 实际来源依赖：计划的值实际从哪些工作表读出来。与上面的指纹范围**分开建模**：
    # 授权范围（authorizedSheetNames）／实际来源依赖（本字段）／写入目标
    # （actions[].sheet）三者互不冒充。本端只从可信记录写入（确定性本地读取、
    # 工具执行记录、已保存工具模板声明），服务端不采信模型自报的来源集合；
    # 缺失即"依赖不确定"，恢复登记退回既有的保守拒绝条件。
    sourceDependencySheets: list[SheetName] = Field(default_factory=list, max_length=100)
    assumptions: list[str] = Field(default_factory=list, max_length=20)
    warnings: list[str] = Field(default_factory=list, max_length=20)
    actions: list[ExcelAction] = Field(min_length=1, max_length=50)
    acceptanceCriteria: list[VerificationCriterion] = Field(
        default_factory=list, max_length=100
    )
    # 本地结果引用：当计划由本机查询结果编译而来时，账本与恢复副本只保留这个
    # 句柄（值留在本机易失登记器），而不是把值矩阵写进持久记录。
    localResultRef: str | None = Field(
        default=None, pattern=r"^res_[0-9a-f]{32}$"
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
                # 后面还有同范围的 setFont 也给了同一字段时，最终值由后者决定 ——
                # 本条不认领它，免得同一区域上出现两条互相矛盾的期望。
                superseded = {
                    field: any(
                        getattr(later, field) is not None for later in later_fonts
                    )
                    for field in ("bold", "italic", "name", "size", "color")
                }
                expected_font = {
                    "bold": None if superseded["bold"] else action.bold,
                    "fontItalic": None if superseded["italic"] else action.italic,
                    "fontName": None if superseded["name"] else action.name,
                    "fontSize": None if superseded["size"] else action.size,
                    "fontColor": None if superseded["color"] else action.color,
                }
                if all(value is None for value in expected_font.values()):
                    continue
                add_criterion(
                    RangeFormatMatchesCriterion(
                        type="rangeFormatMatches",
                        sheet=action.sheet,
                        range=action.range,
                        **expected_font,
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
            elif action.type in ("computeValues", "computeScript"):
                # 故意只保留上面那条无条件补的 worksheetExists，不补任何值类条件。
                # 这两个动作的结果都在执行那一刻才算出来，行数未知 —— 补
                # rangeEquals 既填不出 range，也会被 RangeEqualsCriterion 的
                # 「range 尺寸必须等于 expected 尺寸」校验直接拒掉。
                # 运行时的验收由 Excel 端自建（见 excel.ts 的 splitGroupAggregate 做法）。
                pass
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
            if action.type not in ("splitGroupAggregate", "computeValues"):
                continue
            warning = (
                f"将按「{action.splitBy}」创建多个结果工作表，"
                f"同名工作表处理方式为 {action.existingSheetPolicy}。"
            ) if action.type == "splitGroupAggregate" else (
                f"计划将在执行时从「{action.sourceSheet}」读取实时数据进行计算，"
                f"结果行数由实际数据决定，预览时未固定。"
            )
            if warning not in self.warnings:
                self.warnings = [*self.warnings[:19], warning]
        if len(criteria) > 100:
            raise ValueError("计划及自动补充的验收条件不能超过 100 项")
        self.acceptanceCriteria = criteria
        return self


class PlanValidateRequest(BaseModel):
    plan: AnalysisPlan
    authorized_sheets: list[SheetName] = Field(
        default_factory=list, max_length=100
    )
    existing_sheets: list[SheetName] = Field(default_factory=list, max_length=1000)

    @model_validator(mode="after")
    def enforce_authorized_sheet_boundary(self) -> "PlanValidateRequest":
        authorized = {name.casefold() for name in self.authorized_sheets}
        existing = {name.casefold() for name in self.existing_sheets}
        created: set[str] = set()

        for source_sheet in self.plan.sourceFingerprintSheets:
            if source_sheet.casefold() not in authorized:
                raise ValueError(
                    f"数据来源指纹尝试读取未授权的工作表「{source_sheet}」。"
                )

        # 实际来源依赖同样不许越过本次授权：它是"本机读过的表"，读的时候本来
        # 就该在授权范围内；越界说明端上的记账与授权对不上，宁可拒绝。
        for dependency_sheet in self.plan.sourceDependencySheets:
            if dependency_sheet.casefold() not in authorized:
                raise ValueError(
                    f"计划声明的实际来源依赖超出授权的工作表「{dependency_sheet}」。"
                )

        for action in self.plan.actions:
            source_sheet = getattr(action, "sourceSheet", None)
            if source_sheet:
                source_key = source_sheet.casefold()
                if source_key not in authorized and source_key not in created:
                    raise ValueError(
                        f"动作 {action.type} 尝试读取未授权的工作表"
                        f"「{source_sheet}」。"
                    )

            sheet_key = action.sheet.casefold()
            if action.type == "splitGroupAggregate":
                if sheet_key not in authorized and sheet_key not in created:
                    raise ValueError(
                        f"动作 {action.type} 尝试读取未授权的工作表"
                        f"「{action.sheet}」。"
                    )
                continue

            if (
                (
                    sheet_key in existing
                    or action.type == "deleteWorksheet"
                )
                and sheet_key not in authorized
                and sheet_key not in created
            ):
                raise ValueError(
                    f"动作 {action.type} 尝试写入未授权的工作表"
                    f"「{action.sheet}」。"
                )
            if sheet_key not in existing:
                created.add(sheet_key)

        return self


class ActionExecutionResult(BaseModel):
    index: int = Field(ge=0)
    type: str
    sheet: str
    status: Literal["succeeded", "failed", "not_run"]
    message: str | None = Field(default=None, max_length=1000)
    changed: bool | None = None


class FolderUndoReference(BaseModel):
    undoId: str = Field(min_length=1, max_length=100)
    sessionId: str = Field(min_length=1, max_length=100)


class FolderUndoRequest(BaseModel):
    sessionId: str = Field(min_length=1, max_length=100)
    undoId: str = Field(min_length=1, max_length=100)


class FolderUndoResponse(BaseModel):
    filesRestored: list[str]
    executionsUndone: int = Field(ge=1)
    remainingUndoCount: int = Field(ge=0)


class ManagedFolderRecoveryDirectory(BaseModel):
    directoryId: str = Field(min_length=1, max_length=100)
    folderName: str = Field(min_length=1, max_length=255)
    folderPath: str = Field(min_length=1, max_length=2000)
    status: Literal["active", "unavailable", "unsafe"]
    recoveryCount: int = Field(ge=0)
    attentionCount: int = Field(ge=0)
    operationCount: int = Field(ge=0)
    lastScannedAt: str | None = Field(default=None, max_length=100)
    statusReason: str | None = Field(default=None, max_length=2000)


class ManagedFolderRecoveryDirectoriesResponse(BaseModel):
    directories: list[ManagedFolderRecoveryDirectory] = Field(default_factory=list)


class FolderManagedRecoveryRequest(BaseModel):
    directoryId: str = Field(
        min_length=64, max_length=64, pattern=r"^[0-9a-f]{64}$"
    )
    undoId: str = Field(
        min_length=32, max_length=32, pattern=r"^[0-9a-f]{32}$"
    )
    expectedRevision: int = Field(ge=0)
    operationId: str = Field(
        min_length=8,
        max_length=100,
        pattern=r"^[A-Za-z0-9][A-Za-z0-9_-]{7,99}$",
    )
    confirmed: Literal[True]


class FolderExecutionResumeRequest(BaseModel):
    operationId: str = Field(
        min_length=8,
        max_length=100,
        pattern=r"^[A-Za-z0-9][A-Za-z0-9_.:-]*$",
    )
    expectedRevision: int = Field(ge=1)
    confirmed: Literal[True]
    afterSequence: int = Field(default=0, ge=0)


class FolderManagedRecoveryResponse(BaseModel):
    operationId: str = Field(min_length=8, max_length=100)
    undoId: str = Field(min_length=32, max_length=32)
    status: Literal["completed"]
    filesRestored: list[str]
    alreadyRestored: bool


class FolderBackupFileInfo(BaseModel):
    originalRelativePath: str = Field(min_length=1, max_length=2000)
    backupRelativePath: str | None = Field(default=None, max_length=2000)
    recoveryAction: Literal["restore", "remove"]
    sizeBytes: int = Field(ge=0)


class FolderBackupEntry(BaseModel):
    revision: int = Field(default=0, ge=0)
    undoId: str = Field(min_length=1, max_length=100)
    undoSessionId: str | None = Field(default=None, max_length=100)
    planId: str = Field(min_length=1, max_length=100)
    planTitle: str = Field(min_length=1, max_length=100)
    createdAt: str = Field(min_length=1, max_length=100)
    status: Literal[
        "creating",
        "preparing",
        "prepared",
        "applying",
        "ready",
        "recovering",
        "restored",
        "rolled_back",
        "needs_attention",
    ]
    statusReason: str | None = Field(default=None, max_length=2000)
    undoAvailability: Literal["eligible", "blocked", "view_only", "busy"]
    undoUnavailableReason: str | None = Field(default=None, max_length=2000)
    deleteInvalidatesUndoIds: list[str] = Field(default_factory=list)
    pinned: bool = False
    protectedReasons: list[str] = Field(default_factory=list)
    canUndo: bool
    fileCount: int = Field(ge=0)
    totalBytes: int = Field(ge=0)
    expiresAt: str | None = Field(default=None, max_length=100)
    files: list[FolderBackupFileInfo] = Field(default_factory=list)


class FolderLegacyBackupInfo(BaseModel):
    relativePath: str = Field(min_length=1, max_length=2000)
    sizeBytes: int = Field(ge=0)
    modifiedAt: str = Field(min_length=1, max_length=100)


class FolderBackupStorageSummary(BaseModel):
    automaticBytes: int = Field(ge=0)
    pinnedBytes: int = Field(ge=0)
    abnormalBytes: int = Field(ge=0)
    exportedBytes: int = Field(ge=0)
    legacyBytes: int = Field(ge=0)
    pendingCleanupBytes: int = Field(ge=0)
    totalBytes: int = Field(ge=0)
    softLimitBytes: int = Field(ge=0)
    complete: bool
    calculatedAt: str = Field(min_length=1, max_length=100)


class FolderBackupsResponse(BaseModel):
    folderPath: str = Field(min_length=1, max_length=2000)
    backupDirectory: str = Field(min_length=1, max_length=2000)
    openDirectoryToken: str | None = Field(default=None, min_length=1, max_length=100)
    backupManagementToken: str | None = Field(default=None, min_length=1, max_length=100)
    backups: list[FolderBackupEntry] = Field(default_factory=list)
    legacyCandidates: list[FolderLegacyBackupInfo] = Field(default_factory=list)
    storageSummary: FolderBackupStorageSummary | None = None
    warnings: list[str] = Field(default_factory=list)


class FolderBackupOpenRequest(BaseModel):
    directoryToken: str = Field(min_length=1, max_length=100)


class FolderBackupOpenResponse(BaseModel):
    opened: bool
    directory: str = Field(min_length=1, max_length=2000)


class FolderBackupExportRequest(BaseModel):
    sessionId: str | None = Field(default=None, min_length=1, max_length=100)
    backupManagementToken: str | None = Field(default=None, min_length=1, max_length=100)
    undoId: str = Field(min_length=1, max_length=100)
    operationId: str | None = Field(
        default=None,
        min_length=8,
        max_length=100,
        pattern=r"^[A-Za-z0-9][A-Za-z0-9_-]*$",
    )

    @model_validator(mode="after")
    def validate_management_target(self) -> "FolderBackupExportRequest":
        if not self.sessionId and not self.backupManagementToken:
            raise ValueError("必须提供文件夹会话或备份管理凭证")
        return self


class FolderBackupExportResponse(BaseModel):
    operationId: str = Field(min_length=8, max_length=100)
    exportDirectory: str = Field(min_length=1, max_length=2000)
    files: list[str] = Field(default_factory=list)
    complete: bool = True
    failures: list[str] = Field(default_factory=list)


class FolderBackupDeleteRequest(BaseModel):
    backupManagementToken: str = Field(min_length=1, max_length=100)
    undoId: str = Field(min_length=1, max_length=100)
    confirmed: bool
    expectedRevision: int | None = Field(default=None, ge=0)
    expectedInvalidatedUndoIds: list[str] | None = None


class FolderBackupDeleteResponse(BaseModel):
    deletedUndoId: str = Field(min_length=1, max_length=100)
    discardedUndo: bool
    filesDeleted: int = Field(ge=0)
    bytesFreed: int = Field(ge=0)
    remainingUndoCount: int = Field(ge=0)
    invalidatedUndoIds: list[str] = Field(default_factory=list)
    remainingUndoIds: list[str] = Field(default_factory=list)
    cleanupPending: bool = False
    cleanupWarnings: list[str] = Field(default_factory=list)


class FolderBackupPinRequest(BaseModel):
    backupManagementToken: str = Field(min_length=1, max_length=100)
    undoId: str = Field(min_length=1, max_length=100)
    pinned: bool


class FolderBackupPinResponse(BaseModel):
    undoId: str = Field(min_length=1, max_length=100)
    pinned: bool


class FolderBackupCleanupRequest(BaseModel):
    sessionId: str | None = Field(default=None, min_length=1, max_length=100)
    backupManagementToken: str | None = Field(default=None, min_length=1, max_length=100)

    @model_validator(mode="after")
    def validate_management_target(self) -> "FolderBackupCleanupRequest":
        if not self.sessionId and not self.backupManagementToken:
            raise ValueError("必须提供文件夹会话或备份管理凭证")
        return self


class FolderBackupCleanupResponse(BaseModel):
    deletedExecutions: int = Field(ge=0)
    deletedFiles: int = Field(ge=0)
    bytesFreed: int = Field(ge=0)
    skippedActive: int = Field(ge=0)
    protectedExecutions: int = Field(default=0, ge=0)
    retainedBytes: int = Field(default=0, ge=0)
    overLimitBytes: int = Field(default=0, ge=0)
    failures: list[str] = Field(default_factory=list)
    legacyCandidateCount: int = Field(ge=0)


class FolderLegacyCollectRequest(BaseModel):
    sessionId: str | None = Field(default=None, min_length=1, max_length=100)
    backupManagementToken: str | None = Field(default=None, min_length=1, max_length=100)
    confirmed: bool

    @model_validator(mode="after")
    def validate_management_target(self) -> "FolderLegacyCollectRequest":
        if not self.sessionId and not self.backupManagementToken:
            raise ValueError("必须提供文件夹会话或备份管理凭证")
        return self


class FolderLegacyCollectResponse(BaseModel):
    collectedFiles: list[str] = Field(default_factory=list)
    failures: list[str] = Field(default_factory=list)


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


class FolderExecuteResponse(BaseModel):
    executionId: str | None = Field(default=None, min_length=8, max_length=100)
    filesModified: list[str]
    backups: list[str]
    backupDirectory: str | None = None
    backupWarnings: list[str] = Field(default_factory=list)
    actionResults: list[ActionExecutionResult]
    verification: VerificationReport
    undo: FolderUndoReference | None = None
    executionMs: float = Field(ge=0)
    verificationMs: float = Field(ge=0)


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


class MessageUnderstandingRequest(BaseModel):
    """Message-only routing: no workbook identity, schema or computed results."""

    model_config = ConfigDict(extra="forbid")
    prompt: str = Field(min_length=2, max_length=4000)
    imageCount: int = Field(default=0, ge=0, le=3)
    conversation: list[IntentConversationMessage] = Field(default_factory=list, max_length=6)
    modelId: Annotated[str, StringConstraints(strip_whitespace=True, max_length=200)] | None = None
    # Opaque planning IDs or preplanning-* local request refs; never authorization.
    recoveryTasks: list[Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=200)]] = Field(default_factory=list, max_length=6)


class MessageConversationResponse(BaseModel):
    model_config = ConfigDict(extra="forbid")
    kind: Literal["conversation"] = "conversation"
    message: Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=1500)]
    provider: Literal["model"] = "model"


class MessageContextResponse(BaseModel):
    model_config = ConfigDict(extra="forbid")
    kind: Literal["needs_workbook_context"] = "needs_workbook_context"
    summary: Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=500)]
    provider: Literal["model"] = "model"


class MessageRecoveryResponse(BaseModel):
    model_config = ConfigDict(extra="forbid")
    kind: Literal["task_recovery"] = "task_recovery"
    action: Literal["inspect", "continue"] = "inspect"
    turnId: Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=200)]
    provider: Literal["model"] = "model"


MessageUnderstandingResponse = Annotated[
    MessageConversationResponse | MessageContextResponse | MessageRecoveryResponse, Field(discriminator="kind")
]


class AgentTaskDecision(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)
    questionId: Annotated[str, StringConstraints(strip_whitespace=True, min_length=1,
        max_length=capability_int("taskContext", "maxQuestionIdChars"))] | None = None
    label: Annotated[str, StringConstraints(strip_whitespace=True, min_length=1,
        max_length=capability_int("taskContext", "maxLabelChars"))]
    value: Annotated[str, StringConstraints(strip_whitespace=True, min_length=1,
        max_length=capability_int("taskContext", "maxAnswerChars"))]


class AgentTaskContext(BaseModel):
    """User decisions, never a scope grant, local snapshot or SDK transcript."""
    model_config = ConfigDict(extra="forbid", strict=True)
    schemaVersion: Literal[1] = 1
    originalPrompt: Annotated[str, StringConstraints(strip_whitespace=True, min_length=1,
        max_length=capability_int("taskContext", "maxPromptChars"))]
    confirmations: list[AgentTaskDecision] = Field(default_factory=list,
        max_length=capability_int("taskContext", "maxConfirmations"))

    @field_validator("confirmations")
    @classmethod
    def latest_answer_only(cls, value: list[AgentTaskDecision]) -> list[AgentTaskDecision]:
        result: list[AgentTaskDecision] = []
        positions: dict[str, int] = {}
        for item in value:
            if item.questionId and item.questionId in positions:
                result[positions[item.questionId]] = item
            else:
                if item.questionId:
                    positions[item.questionId] = len(result)
                result.append(item)
        return result


class IntentCheckRequest(BaseModel):
    turnId: TurnId | None = None
    prompt: str = Field(min_length=2, max_length=4000)
    taskContext: AgentTaskContext | None = None
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
    # modelId 缺省/空串都表示「本次没有选择任何模型」——这不是缺参数，
    # 而是一种合法状态：服务端会据此返回 MODEL_NOT_CONFIGURED，让前台
    # 引导用户去「添加模型」，而不是静默回落到基础模式。
    # 旧版加载项缓存的 "local" 会被 config 层归一化成同一状态。
    modelId: Annotated[
        str,
        StringConstraints(strip_whitespace=True, max_length=200),
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
    resolutionMode: Literal["append", "replace"] = "append"
    action: Literal["resolve", "editScope"] = "resolve"


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


# ---------------------------------------------------------------------------
# Task tool bridge (SDK handler ⇄ local host adapter)
#
# The bridge carries local query requests to the host adapter and receipts back.
# These DTOs are the loopback-only wire contract; they never enter a model
# prompt or the durable execution ledger. Positional values (the value matrix)
# travel inside ``BridgeToolReceipt.result`` only on the loopback channel.
# ---------------------------------------------------------------------------

BridgeHostId = Literal["excel-office", "wps-et", "folder"]

# 桥支持的工具名。``query_table`` 是确定性查询；``prepare_python_data`` 让宿主
# 读授权表并把原始值矩阵登记为隔离 Python 的数据集（回执仍是本机回环，值只
# 进服务端易失 registry，不进模型/账本/持久化）。
BridgeToolName = Literal["query_table", "prepare_python_data"]


class BridgeBindingRequest(BaseModel):
    """Binding echoed by the host adapter; matched against the live session.

    ``workbookIdentity`` 是可选的稳定身份（``url:`` / ``wps-path:``）。缺省时
    服务端退回 ``name:<工作簿名>``。名字相同但身份不同的工作簿因此不会互相
    冒充；只比名字不足以隔离两个同名工作簿。
    """

    model_config = ConfigDict(extra="forbid")
    taskId: TurnId
    hostId: BridgeHostId
    workbookName: str = Field(min_length=1, max_length=255)
    workbookIdentity: str | None = Field(default=None, min_length=5, max_length=4000)
    sourceFingerprint: str | None = Field(default=None, max_length=128)


class BridgePendingCall(BaseModel):
    model_config = ConfigDict(extra="forbid")
    operationId: str = Field(pattern=r"^op_[0-9a-f]{32}$")
    tool: BridgeToolName
    arguments: dict[str, Any]
    requestFingerprint: str = Field(pattern=r"^[0-9a-f]{64}$")
    expiresAt: float


class BridgePollResponse(BaseModel):
    model_config = ConfigDict(extra="forbid")
    calls: list[BridgePendingCall] = Field(default_factory=list, max_length=8)
    generation: int = Field(ge=0)


class BridgeToolReceipt(BaseModel):
    """Result of executing one local query call on the host.

    ``query_table`` 走 ``result``（已算好的 ``DataToolResult``）；``prepare_python_data``
    走 ``datasets``（原始值矩阵，直接登记为隔离 Python 的数据集）。值只在本机回环
    通道出现，不进模型/账本/持久化。
    """

    model_config = ConfigDict(extra="forbid")
    result: DataToolResult | None = None
    datasets: list[BridgeDatasetReceipt] | None = Field(default=None, max_length=8)
    errorCode: str | None = Field(default=None, max_length=40)

    @model_validator(mode="after")
    def validate_one_of(self) -> "BridgeToolReceipt":
        if (self.result is None) == (self.errorCode is None) and self.datasets is None:
            # 必须恰好给出 result / datasets / errorCode 之一。
            raise ValueError("回执必须且只能给出 result、datasets 或 errorCode 之一")
        if self.result is not None and (self.datasets is not None or self.errorCode is not None):
            raise ValueError("回执不能同时给出 result 与 datasets/errorCode")
        if self.datasets is not None and self.errorCode is not None:
            raise ValueError("回执不能同时给出 datasets 与 errorCode")
        return self


class BridgeDatasetReceipt(BaseModel):
    """``prepare_python_data`` 宿主回执里的单块原始值矩阵。

    与服务端 :class:`PythonDataSnapshotRequest` 同形：按 ``sourceIndex`` 定位授权
    范围内的某块数据，``values`` 是原始值矩阵。形状/类型/字节上界由可信侧
    :mod:`python_datasets` 在登记时复核，这里只做回环通道的搬运。
    """

    model_config = ConfigDict(extra="forbid")
    sourceIndex: int = Field(ge=0, le=7)
    values: list[list[Any]]


class BridgeReceiptRequest(BridgeBindingRequest):
    operationId: str = Field(pattern=r"^op_[0-9a-f]{32}$")
    requestFingerprint: str = Field(pattern=r"^[0-9a-f]{64}$")
    receipt: BridgeToolReceipt


class BridgeReceiptResponse(BaseModel):
    model_config = ConfigDict(extra="forbid")
    status: Literal["accepted", "duplicate", "rejected"]
    code: str | None = Field(default=None, max_length=40)


class BridgeRegisterRequest(BridgeBindingRequest):
    """Register an already-computed local query result (no re-execution)."""

    result: DataToolResult


class BridgeRegisterResponse(BaseModel):
    model_config = ConfigDict(extra="forbid")
    resultRef: str = Field(pattern=r"^res_[0-9a-f]{32}$")
    structure: dict[str, Any]


class BridgeResultRequest(BridgeBindingRequest):
    resultRef: str = Field(pattern=r"^res_[0-9a-f]{32}$")


class BridgeResultResponse(BaseModel):
    """Loopback-only payload for the local result card (values included)."""

    model_config = ConfigDict(extra="forbid")
    resultRef: str = Field(pattern=r"^res_[0-9a-f]{32}$")
    title: str = Field(min_length=1, max_length=300)
    headers: list[CellValue] = Field(min_length=1, max_length=50)
    rows: list[list[CellValue]] = Field(default_factory=list, max_length=500)
    sourceSheets: list[str] = Field(default_factory=list, max_length=100)
    warnings: list[str] = Field(default_factory=list, max_length=20)
    complete: bool
    calculation: str = Field(min_length=1, max_length=1000)
    scannedRows: int = Field(ge=0)


class DevFaultSeenTarget(BaseModel):
    """服务端见过的一本工作簿（开发入口专用，只在开关打开时记录）。

    给**人**抄身份用：要知道 `workbookIdentity` 该填什么，就从这里取，而不是猜
    `url:` 的归一规则。没有任务 ID、没有数据内容。
    """

    model_config = ConfigDict(extra="forbid")
    hostId: str
    workbookName: str
    workbookIdentity: str


class DevResultRefFaultArmRequest(BaseModel):
    """开发验收专用的结果引用过期入口（A1-E1）。

    **只有"哪本工作簿"**：没有 ``taskId``，也没有 ``resultRef``。
    任务与句柄由服务端在可信路径上绑定（开桥那一刻 + 该任务第一条 Python 结果
    登记那一刻），所以模型既不能指定别的任务，也不能指定别的引用。

    ``workbookIdentity`` 可以省略，但省略时服务端必须"只见过这个名字下的一个
    可核验身份"才肯补全；否则拒绝启用（见 ``dev_fault_target``）。只凭名字认簿
    会把启用落到同名的另一本簿上。
    """

    model_config = ConfigDict(extra="forbid")
    hostId: str = Field(min_length=1, max_length=40)
    workbookName: str = Field(min_length=1, max_length=255)
    workbookIdentity: str = Field(default="", max_length=4000)


class DevResultRefFaultStatus(BaseModel):
    """入口状态；不含任务 ID 与结果句柄（不把归属凭据送回界面）。"""

    model_config = ConfigDict(extra="forbid")
    enabled: bool
    phase: Literal["off", "idle", "armed", "bound", "consumed", "expired"]
    expiresInSeconds: int = Field(ge=0)
    taskBound: bool
    refBound: bool
    identityBound: bool = False
    identitySource: Literal["declared", "observed", "none"] = "none"


class DevSourceChangedFaultArmRequest(BaseModel):
    """开发验收专用的来源变化入口（A1-E3）。

    ``hostId`` / ``workbookName`` **必填**：只按"下一次调用"命中会让注入落到并发
    任务或另一本打开的簿上，被抢的那一轮莫名停住、要验的簿反而看不到停止。
    ``workbookIdentity`` 可以省略，但省略时服务端必须"只见过这个名字下的一个
    可核验身份"才肯补全（``identitySource = observed``）；见过零个或多个 ⇒
    拒绝启用并说明原因。**不能**退回只认名字：同名不同簿在名字上完全一样。

    这里仍然**没有** ``taskId``：任务归属由服务端在开桥那一刻绑定（可信路径），
    调用方（更不用说模型）无从指定"注入给哪一次任务"。
    """

    model_config = ConfigDict(extra="forbid")
    hostId: str = Field(min_length=1, max_length=40)
    workbookName: str = Field(min_length=1, max_length=255)
    workbookIdentity: str = Field(default="", max_length=4000)


class DevSourceChangedFaultStatus(BaseModel):
    """入口状态；不含任务 ID（不把归属凭据送回界面）。"""

    model_config = ConfigDict(extra="forbid")
    enabled: bool
    phase: Literal["off", "idle", "armed", "bound", "consumed", "expired"]
    expiresInSeconds: int = Field(ge=0)
    hostBound: bool
    workbookBound: bool
    taskBound: bool
    identityBound: bool = False
    identitySource: Literal["declared", "observed", "none"] = "none"
    #: 服务端见过的工作簿（最近的在前，最多 16 条）。给验收入抄身份用。
    seenTargets: list[DevFaultSeenTarget] = Field(default_factory=list, max_length=16)


class HostWriteAuditEntry(BaseModel):
    """一次动作的宿主写入事实（A1-E3）。

    **没有**单元格值、公式正文、文件名或工具参数：这里只回答"向宿主发起了几次
    写调用、结果如何、回读有没有过"。计数口径见 ``packages/core/src/hostWriteAudit.ts``
    （两端读同一份）。

    ``outcome`` 是四态而不是布尔：``unknown`` 既不是成功也不是零写入 ——
    宿主抛异常、断线、只写了一半时只能是它，据实记录才谈得上"不宣称"。
    """

    model_config = ConfigDict(extra="forbid")
    actionIndex: int = Field(ge=0, le=10000)
    actionType: str = Field(min_length=1, max_length=40)
    hostWriteCalls: int = Field(ge=0, le=100000)
    mutatingCalls: int = Field(ge=0, le=100000)
    outcome: Literal["written", "failed_before_write", "failed_after_write", "unknown"]
    readback: Literal["verified", "failed", "not_applicable"]
    errorCode: str | None = Field(default=None, max_length=60)


class HostWritesRequest(BaseModel):
    """端上补报一次宿主观测。

    ``ownerToken`` 可选：给了就必须与账本一致（防止另一个执行者往别人的账本里
    写字）；不给只允许**只读身份**（同一 operationId）补报，仍然不能改状态。
    """

    model_config = ConfigDict(extra="forbid")
    ownerToken: str | None = Field(default=None, min_length=32, max_length=100)
    entries: list[HostWriteAuditEntry] = Field(min_length=1, max_length=64)


class HostWritesTotals(BaseModel):
    """汇总只做加法，不做推断。"""

    model_config = ConfigDict(extra="forbid")
    reportedActions: int = Field(ge=0)
    hostWriteCalls: int = Field(ge=0)
    actionsWritten: int = Field(ge=0)
    actionsFailedBeforeWrite: int = Field(ge=0)
    actionsFailedAfterWrite: int = Field(ge=0)
    actionsUnknown: int = Field(ge=0)
    readbackVerified: int = Field(ge=0)
    readbackFailed: int = Field(ge=0)


class HostWritesObservation(BaseModel):
    """**观测覆盖**：这批数字能不能当成"零写入"的证据。

    空数组不等于零写入 —— 上报是 best-effort，可能一条都没到。所以这里把"没有
    观测"和"明确观测到零调用"分开，并且**明确零只能从账本执行边界 + 完整覆盖
    推出**：

    - ``attemptedActions``：账本判定"有可能已经写过"的动作数
      （``next_action_index`` 个已完成 + 在途的那一个）。这是可信执行边界 ——
      没有认领就进不了宿主写 API；
    - ``reportedActions`` / ``missingActions`` / ``unexpectedReports``：覆盖情况；
    - ``observation``：``not_collected``（还没收到）/ ``partial``（收到一部分）/
      ``complete``（边界内每个动作都有观测）；
    - ``callEvidence``：``none``（没有观测）/ ``zero``（**覆盖完整且明确 0 次**）/
      ``nonzero``（至少一次非零，无论覆盖是否完整）/ ``unknown``（覆盖不全或有
      结果未知的条目 ⇒ 不能下"零"的结论）；
    - ``unit``：两端**计量单位不同**，不能互相折算（Excel = 刷新批次，
      WPS = 宿主写 API 调用次数）。
    """

    model_config = ConfigDict(extra="forbid")
    observation: Literal["not_collected", "partial", "complete"]
    callEvidence: Literal["none", "zero", "nonzero", "unknown"]
    attemptedActions: int = Field(ge=0)
    reportedActions: int = Field(ge=0)
    missingActions: int = Field(ge=0)
    unexpectedReports: int = Field(ge=0)
    unit: Literal["excel_sync_batches", "wps_api_calls"]
    lastReportedAt: float | None = None


class HostWritesResponse(BaseModel):
    model_config = ConfigDict(extra="forbid")
    operationId: str
    entries: list[HostWriteAuditEntry]
    totals: HostWritesTotals
    observation: HostWritesObservation


class PythonCodeRequest(BaseModel):
    """Loopback-only request for the generated code body.

    代码正文**不进事件、不进账本**（会落进 SQLite），只在进程内登记表里短期存在；
    界面按 ``codeId`` 经本机回环端点取回。``taskId`` 是必需的：只凭 ``codeId``
    就能取回会让这个端点变成跨任务读取工具。
    """

    model_config = ConfigDict(extra="forbid")
    taskId: str = Field(min_length=1, max_length=64)
    codeId: str = Field(pattern=r"^code_py_[0-9a-f]{12}$")


class PythonCodeResponse(BaseModel):
    """取回后的代码正文。``code`` 是**当时那一版**，不是最新一版。"""

    model_config = ConfigDict(extra="forbid")
    codeId: str = Field(pattern=r"^code_py_[0-9a-f]{12}$")
    attemptId: str = Field(min_length=1, max_length=64)
    code: str
    chars: int = Field(ge=0)


class BridgeToolPointer(BaseModel):
    """Opaque pointer persisted with a progress event.
    Only the operation id and tool name — never the request arguments, which
    stay in the volatile bridge store and are fetched by the host adapter.
    """

    model_config = ConfigDict(extra="forbid")
    operationId: str = Field(pattern=r"^op_[0-9a-f]{32}$")
    tool: BridgeToolName


class ModelOptionResponse(BaseModel):
    id: str
    label: str
    provider: Literal["model"] = "model"
    available: bool
    supportsVision: bool
    agentAvailable: bool
    agentUnavailableReason: str | None = None
    formulaAvailable: bool


class ModelCatalogResponse(BaseModel):
    defaultModelId: str | None = None
    models: list[ModelOptionResponse]


class ManagedModelConnectionResponse(BaseModel):
    protocol: Literal["openai_compatible", "anthropic"] = "openai_compatible"
    agentAvailable: bool = False
    agentUnavailableReason: str | None = None
    formulaAvailable: bool = True
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
    protocol: Literal["openai_compatible", "anthropic"] | None = None
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


class TestModelConnectionRequest(UpsertModelConnectionRequest):
    # Older clients test only chat connectivity; new clients choose explicitly.
    purpose: Literal["agent", "formula"] = "formula"


class TestModelConnectionResponse(BaseModel):
    testedFor: Literal["agent", "formula"] = "formula"
    ok: bool
    message: str


class ClarificationField(BaseModel):
    id: str = Field(min_length=1, max_length=100)
    label: str = Field(min_length=1, max_length=300)
    placeholder: str = Field(default="", max_length=300)


class IntentClarification(BaseModel):
    id: str = Field(min_length=1, max_length=100)
    summary: str = Field(min_length=1, max_length=300)
    question: str = Field(min_length=1, max_length=300)
    reason: str = Field(min_length=1, max_length=300)
    scopeLabel: str = Field(min_length=1, max_length=300)
    options: list[IntentOption] = Field(default_factory=list, max_length=4)
    fields: list[ClarificationField] = Field(default_factory=list, max_length=4)

    @model_validator(mode="after")
    def validate_interaction(self) -> "IntentClarification":
        if not self.fields and len(self.options) < 2:
            raise ValueError("澄清必须提供至少两个选项或至少一个填写字段")
        if len({field.id for field in self.fields}) != len(self.fields):
            raise ValueError("澄清字段 id 不能重复")
        return self


class IntentProceedResponse(BaseModel):
    kind: Literal["proceed"] = "proceed"
    # The established intent route means preview unless the model explicitly
    # identifies a read-only task (e.g. an image or complex Python analysis).
    planningGoal: Literal["read_only", "write_preview"] = "write_preview"
    summary: str = Field(min_length=1, max_length=500)
    confirmedPrompt: str = Field(min_length=2, max_length=4500)
    provider: Literal["model", "local"]
    turnId: TurnId | None = None


class IntentConversationResponse(BaseModel):
    kind: Literal["conversation"] = "conversation"
    message: str = Field(min_length=1, max_length=1500)
    provider: Literal["model", "local"]
    turnId: TurnId | None = None


class IntentClarificationResponse(BaseModel):
    kind: Literal["clarification"] = "clarification"
    clarification: IntentClarification
    provider: Literal["model", "local"]
    turnId: TurnId | None = None


class IntentToolResponse(BaseModel):
    kind: Literal["tool_request"] = "tool_request"
    planningGoal: Literal["read_only", "write_preview"] = "read_only"
    summary: str = Field(min_length=1, max_length=500)
    confirmedPrompt: str = Field(min_length=2, max_length=4500)
    request: DataToolRequest
    provider: Literal["model", "local"]
    turnId: TurnId | None = None


IntentCheckResponse = Annotated[
    IntentProceedResponse | IntentConversationResponse | IntentClarificationResponse | IntentToolResponse,
    Field(discriminator="kind"),
]


class PlanRequest(BaseModel):
    turnId: TurnId | None = None
    # Bound before planning, independent of terminal tool arguments. Not a
    # permission to write; execution still requires the existing confirmation.
    planningGoal: Literal["read_only", "write_preview"] | None = None
    prompt: str = Field(min_length=2, max_length=4000)
    taskContext: AgentTaskContext | None = None
    workbook: WorkbookSnapshot
    lastResult: ResultContext | None = None
    images: list[ImageAttachment] = Field(default_factory=list, max_length=3)
    dataResults: list[DataToolResult] = Field(default_factory=list, max_length=5)
    # 本地已登记结果的不透明引用：只传句柄，值留在本机易失登记器。
    # SDK 后端据此让模型引用本地结果；数据行仍被 dataResults 守卫拒绝。
    localResultRefs: list[str] = Field(default_factory=list, max_length=5)
    planningReadPolicy: Literal["ifNeeded", "none"] = "ifNeeded"
    # 与 IntentCheckRequest 一致：缺省/空串 = 本次没有选择模型，
    # 由 planner 决定是抛 MODEL_NOT_CONFIGURED 还是走纯本地确定性工具。
    modelId: Annotated[
        str,
        StringConstraints(strip_whitespace=True, max_length=200),
    ] | None = None
    # 规划前能力收窄（方案 §4.3）。**这些字段不是授权凭据**：
    # 服务端按 `hostCapabilities` 算出该宿主/版本的动作上界，
    # 只把 `selfReportedActions` 当收窄条件用（交集 `A ∩ B ⊆ A`）。
    # 客户端自报集合不能被当成权威，否则构造过的请求就能让规划器
    # 生成本不该开放的动作。
    hostId: Annotated[
        str,
        StringConstraints(strip_whitespace=True, max_length=40),
    ] | None = None
    hostVersion: Annotated[
        str,
        StringConstraints(strip_whitespace=True, max_length=80),
    ] | None = None
    selfReportedActions: list[
        Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=80)]
    ] = Field(default_factory=list, max_length=200)


class PlanningContinuation(BaseModel):
    """「继续分析」的事实（`/api/turn/{id}/status` 与暂停载荷共用）。

    `available` 只回答"现场是否符合契约且没有已登记的写入"；
    `canContinue` 是它与预算/状态条件的合成，两者不是同一个判断。
    """

    model_config = ConfigDict(extra="forbid")
    available: bool
    code: str | None = None
    message: str = ""
    stage: Literal["before_model", "after_tool"] | None = None
    # 原任务是否依赖来源数据 ⇒ 继续前必须重读并比对来源指纹。
    # **由服务端从检查点判定**，不由客户端猜：猜"需要"会挡住纯结构任务，
    # 猜"不需要"会放过改了数据的现场。客户端只按这个布尔决定要不要真读一次。
    sourceVerificationRequired: bool = False


class PlanningPauseAck(BaseModel):
    turnId: str
    # Receipt of the stop request, not proof that worker cleanup has finished.
    pauseRequested: bool


class PlanningContinueRequest(BaseModel):
    turnId: TurnId
    expectedRevision: int = Field(ge=1)
    requestId: str = Field(min_length=8, max_length=100)
    confirmed: Literal[True]
    afterSequence: int = Field(default=0, ge=0)
    modelId: Annotated[
        str,
        StringConstraints(strip_whitespace=True, min_length=1, max_length=200),
    ] | None = None
    confirmModelSwitch: bool = False
    # 继续之前由客户端**重新核实**的现场证据。三项都是**必填**：
    # 省略、为空、读取失败都拿不到继续权 —— 不允许用"没传就跳过比对"绕过去。
    # 想继续就必须先读一次当前宿主，并把读到的东西交上来。
    workbookIdentity: Annotated[
        str,
        StringConstraints(strip_whitespace=True, min_length=5, max_length=4000),
    ]
    authorizedSheetNames: list[str] = Field(min_length=1, max_length=200)
    # 来源核实的**三态**，不是一个可省略的布尔：
    # - `verified`：本轮重读过依赖来源，`sourceFingerprint` 必填并与记录比对；
    # - `not_collected`：原任务确实没有读过来源（纯结构任务），由服务端判定是否适用；
    # - `failed`：读了但失败。读取失败与"尚未采集"是两件事，必须分开到达服务端，
    #   否则"读不到"会被当成"不需要读"而放行。
    sourceVerification: Literal["verified", "not_collected", "failed"]
    sourceFingerprint: Annotated[
        str,
        StringConstraints(strip_whitespace=True, min_length=1, max_length=400),
    ] | None = None

    @model_validator(mode="after")
    def validate_source_evidence(self) -> "PlanningContinueRequest":
        if self.sourceVerification == "verified" and not (
            self.sourceFingerprint and self.sourceFingerprint.strip()
        ):
            raise ValueError("声明已核实来源时必须提供来源指纹")
        if self.sourceVerification != "verified" and self.sourceFingerprint is not None:
            raise ValueError("未声明已核实来源时不应提供来源指纹")
        if any(not name.strip() for name in self.authorizedSheetNames):
            raise ValueError("授权成员不能包含空表名")
        return self

    @model_validator(mode="after")
    def validate_model_switch_confirmation(self) -> "PlanningContinueRequest":
        if self.modelId is not None and not self.confirmModelSwitch:
            raise ValueError("切换备用模型需要明确确认")
        if self.modelId is None and self.confirmModelSwitch:
            raise ValueError("确认切换时必须选择备用模型")
        return self


class OfficeExecutionRegisterRequest(BaseModel):
    operationId: str = Field(
        min_length=8,
        max_length=100,
        pattern=r"^[A-Za-z0-9][A-Za-z0-9_.:-]*$",
    )
    plan: AnalysisPlan
    authorizedSheetNames: list[str] = Field(min_length=1, max_length=200)
    allowedExternalFiles: list[str] = Field(default_factory=list, max_length=200)
    workbookIdentity: str = Field(min_length=5, max_length=4000)
    # 这个执行是由哪个规划任务产生的。与 WPS 的
    # `WpsExecutionRegisterRequest.taskId` 同一口径：没有它，
    # `has_execution_for_task` 在 Excel 侧永远为假，「继续分析」就发现不了
    # "这个任务其实已经写过原表"，等于把重新规划当成旁路。未知来源（例如
    # `/function` 预览、历史计划重放）允许为空 —— 空只是**不声称**关联，
    # 不能据此放行任何写入。
    taskId: str | None = Field(default=None, min_length=1, max_length=100)
    confirmed: Literal[True]

    @field_validator("workbookIdentity")
    @classmethod
    def validate_workbook_identity(cls, value: str) -> str:
        if not value.startswith("url:"):
            raise ValueError("未保存或无法稳定识别的工作簿不能启用自动断点续跑")
        return value


class OfficeActionStartRequest(BaseModel):
    expectedRevision: int = Field(ge=1)
    ownerToken: str = Field(min_length=32, max_length=100)
    actionIndex: int = Field(ge=0)
    beforeSnapshot: dict[str, Any]


class OfficeActionCompleteRequest(BaseModel):
    expectedRevision: int = Field(ge=1)
    ownerToken: str = Field(min_length=32, max_length=100)
    actionIndex: int = Field(ge=0)
    checkpoint: dict[str, Any]


class OfficeExecutionSuspendRequest(BaseModel):
    expectedRevision: int = Field(ge=1)
    ownerToken: str = Field(min_length=32, max_length=100)
    code: str = Field(min_length=1, max_length=100)
    message: str = Field(min_length=1, max_length=2000)


class OfficeExecutionResumeRequest(BaseModel):
    expectedRevision: int = Field(ge=1)
    confirmed: Literal[True]


class OfficeActionReconcileRequest(BaseModel):
    expectedRevision: int = Field(ge=1)
    actionIndex: int = Field(ge=0)
    outcome: Literal["applied", "not_applied"]
    checkpoint: dict[str, Any] | None = None
    confirmed: Literal[True]

    @model_validator(mode="after")
    def validate_checkpoint(self) -> "OfficeActionReconcileRequest":
        if self.outcome == "applied" and self.checkpoint is None:
            raise ValueError("补记已应用动作必须提供完整检查点")
        return self


class OfficeExecutionFinishRequest(BaseModel):
    expectedRevision: int = Field(ge=1)
    ownerToken: str = Field(min_length=32, max_length=100)
    result: dict[str, Any]


class OfficeExecutionStateResponse(BaseModel):
    operationId: str
    status: Literal["running", "needs_reconciliation", "completed", "failed"]
    revision: int = Field(ge=1)
    workbookIdentity: str
    plan: AnalysisPlan
    authorizedSheetNames: list[str]
    allowedExternalFiles: list[str] = Field(default_factory=list)
    nextActionIndex: int = Field(ge=0)
    inFlightActionIndex: int | None = Field(default=None, ge=0)
    beforeSnapshot: dict[str, Any] | None = None
    checkpoint: dict[str, Any] | None = None
    result: dict[str, Any] | None = None
    error: dict[str, Any] | None = None
    canResume: bool
    requiresReconciliation: bool
    ownerToken: str | None = None


# WPS 格式动作**可以声明**的属性名（= 端上 `declaredFormatFields` 的输出集合，也是
# `apps/wps-addin/src/wpsFormatSnapshot.ts` 的 `ALL_WPS_FORMAT_FIELDS`）。
#
# 尚未实现可恢复执行的属性（`underline` / `columnWidth`）**不在集合里**：
# 端上的投影阶段就产不出来，这里是第二道闸，`main` 里还有按动作类型细分的第三道。
WPS_FORMAT_FIELD_NAMES = frozenset(
    {
        "bold",
        "italic",
        "name",
        "size",
        "color",
        "horizontalAlignment", "verticalAlignment",
        "wrapText",
        "borderTop", "borderBottom", "borderLeft", "borderRight",
        "rowHeight",
        "fillColor", "fillPattern", "numberFormat",
    }
)

# WPS 恢复账本认可的动作类型（= `config/capabilities.json` 的 `wpsWrite.allowedActions`
# 去掉 `clearRange:contents` 的 `:contents` 细分）。`main._WPS_RECOVERABLE_ACTION_TYPES`
# 与它必须一致，`server/tests/test_wps_execution_api.py` 钉住这条同步。
WPS_RECOVERY_ACTION_TYPES = frozenset(
    {
        "writeValues", "sortRange",
        "writeFormulas",
        "clearRange",
        "setFont",
        "setAlignment",
        "resizeRange",
        "setFill", "setNumberFormat", "setBorders",
    }
)


class WpsRecoveryActionDescriptor(BaseModel):
    """WPS 恢复描述里的单个动作。

    **只含元数据，不含 ``values`` / ``formulas`` 矩阵**。``previewFingerprint`` 是
    写前状态的校验指纹；``expectedFingerprint`` 表示写后逻辑内容。
    校验指纹不是加密，不声称能抵御低熵内容的枚举推测。

    值动作的两个指纹都是**值矩阵指纹**；格式动作（``setFont`` / ``setAlignment`` /
    ``resizeRange``）**不改值**，值矩阵指纹在写前写后完全一样、没有区分力，所以那两个
    指纹是**格式证据指纹**（只由声明的属性与目标值加区域形状算出，不含单元格内容），
    并且必须同时声明 ``formatFields``（本次要改的属性名）。``formatFields`` 让"这两个
    指纹是按什么算的"可判定，也让服务端能做属性级校验。

    ``extra="forbid"`` 是出站闸的服务端一侧：任何矩阵字段（``values`` /
    ``formulas`` / ``preview``）或格式快照字段（``cells`` / ``rowHeights``）一旦出现
    在请求里就直接 422，不会"忽略后照常登记"。
    """

    model_config = ConfigDict(extra="forbid")

    actionIndex: int = Field(ge=0)
    type: str = Field(min_length=1, max_length=64)
    sheet: str = Field(min_length=1, max_length=200)
    range: str = Field(min_length=1, max_length=200)
    rowCount: int = Field(ge=1)
    columnCount: int = Field(ge=1)
    cellCount: int = Field(ge=1)
    previewFingerprint: str = Field(min_length=1, max_length=200)
    expectedFingerprint: str | None = Field(default=None, max_length=200)
    applyTo: str | None = None
    formulaR1C1: bool = False
    formatFields: list[str] = Field(default_factory=list, max_length=8)

    @field_validator("formatFields")
    @classmethod
    def validate_format_fields(cls, value: list[str]) -> list[str]:
        """只校验"名字在已实现恢复集合里且不重复"。

        "这个动作类型**该**声明哪些属性"（setFont 只有 bold/italic/name/size/color 等）
        是动作类型相关的判据，放在 ``main._validate_wps_recovery_descriptor``；
        这里只保证**任何**合法记录，读出来后下发给客户端时也不会带出未实现的恢复属性名。
        """
        if len(set(value)) != len(value):
            raise ValueError("格式属性声明不得重复")
        for name in value:
            if name not in WPS_FORMAT_FIELD_NAMES:
                raise ValueError(f"不支持的格式属性「{name}」")
        return value


class WpsResumeSource(BaseModel):
    model_config = ConfigDict(extra="forbid")
    sheetName: str = Field(min_length=1, max_length=31)
    range: str = Field(max_length=40, pattern=r"^\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?$")
    rowCount: int = Field(ge=1, le=1048576)
    columnCount: int = Field(ge=1, le=16384)
    previewFingerprint: str = Field(max_length=80, pattern=r"^v2:\d+:[0-9a-f]{8}$")
    valueFingerprint: str = Field(max_length=80, pattern=r"^v2:\d+:[0-9a-f]{8}$")


class WpsResumeContext(BaseModel):
    model_config = ConfigDict(extra="forbid")
    version: Literal[1] = 1
    hostVersion: str = Field(min_length=1, max_length=80)
    sources: list[WpsResumeSource] = Field(min_length=1, max_length=500)


class WpsRecoveryDescriptor(BaseModel):
    """WPS 写入的**恢复描述**：服务端账本里关于"要写什么"的全部内容。

    早期版本直接上传整个 ``AnalysisPlan``，``writeValues`` 的目标值矩阵因此会出站
    并落进 ``office_executions.plan_json``。这里改成只保存恢复判定必需的元数据：
    动作类型、目标工作表与区域、目标形状、目标指纹，以及动作序列（进度基准）。

    恢复功能**不得**为"能精确续跑"而放宽这条约束 —— 值矩阵不进入服务端账本。
    本机既有撤销记录可被受控续跑复用，来源描述仅包含定位、形状与指纹。
    """

    model_config = ConfigDict(extra="forbid")

    actions: list[WpsRecoveryActionDescriptor] = Field(min_length=1, max_length=500)
    resumeContext: WpsResumeContext | None = None


class WpsExecutionRegisterRequest(BaseModel):
    """WPS 首次原表写入前的持久执行登记。身份用 ``wps-path:`` 前缀。

    与 ``OfficeExecutionRegisterRequest`` 分开：WPS 没有 ``Office.context.document.url``，
    用已保存工作簿的小写磁盘完整路径作为跨关闭身份；两者是不同命名空间，不得互相冒充。

    载荷是 ``descriptor`` 而不是 ``plan``：登记请求不携带单元格值矩阵。
    """

    model_config = ConfigDict(extra="forbid")

    operationId: str = Field(
        min_length=8,
        max_length=100,
        pattern=r"^[A-Za-z0-9][A-Za-z0-9_.:-]*$",
    )
    descriptor: WpsRecoveryDescriptor
    authorizedSheetNames: list[str] = Field(min_length=1, max_length=200)
    workbookIdentity: str = Field(min_length=5, max_length=4000)
    taskId: str = Field(min_length=1, max_length=100)
    confirmed: Literal[True]

    @field_validator("workbookIdentity")
    @classmethod
    def validate_workbook_identity(cls, value: str) -> str:
        if not value.startswith("wps-path:"):
            raise ValueError("WPS 工作簿身份必须以 wps-path: 前缀表示已保存的磁盘路径")
        return value


class WpsActionEvidence(BaseModel):
    """动作级现场证据：只有定位、动作类型与两个状态指纹。

    ``type`` 必须覆盖**已开放的全部可恢复动作**（含格式动作）—— 只认三个值动作时，
    真机 WPS 上格式动作的写前检查点会在登记前就被 422 拒掉。
    """

    model_config = ConfigDict(extra="forbid")
    sheetName: str = Field(min_length=1, max_length=200)
    range: str = Field(min_length=1, max_length=200)
    type: Literal[
        "writeValues",
        "writeFormulas",
        "clearRange",
        "setFont",
        "setAlignment",
        "resizeRange",
    ]
    beforeFingerprint: str = Field(min_length=1, max_length=200)
    afterFingerprint: str | None = Field(default=None, max_length=200)


class WpsActionStartRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    expectedRevision: int = Field(ge=1)
    ownerToken: str = Field(min_length=32, max_length=100)
    actionIndex: int = Field(ge=0)
    beforeSnapshot: WpsActionEvidence


class WpsActionCompleteRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    expectedRevision: int = Field(ge=1)
    ownerToken: str = Field(min_length=32, max_length=100)
    actionIndex: int = Field(ge=0)
    checkpoint: WpsActionEvidence


class WpsExecutionSuspendRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    expectedRevision: int = Field(ge=1)
    ownerToken: str = Field(min_length=32, max_length=100)
    code: str = Field(min_length=1, max_length=100)
    message: str = Field(min_length=1, max_length=2000)


class WpsExecutionResumeRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    expectedRevision: int = Field(ge=1)
    confirmed: Literal[True]


class WpsActionReconcileRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    expectedRevision: int = Field(ge=1)
    actionIndex: int = Field(ge=0)
    outcome: Literal["applied", "not_applied"]
    checkpoint: WpsActionEvidence | None = None
    confirmed: Literal[True]

    @model_validator(mode="after")
    def validate_checkpoint(self) -> "WpsActionReconcileRequest":
        if self.outcome == "applied" and self.checkpoint is None:
            raise ValueError("补记已应用动作必须提供完整检查点")
        return self


class WpsExecutionResult(BaseModel):
    model_config = ConfigDict(extra="forbid")
    actionCount: int = Field(ge=0)
    cellCount: int | None = Field(default=None, ge=0)
    backfilled: bool = False


class WpsExecutionFinishRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    expectedRevision: int = Field(ge=1)
    ownerToken: str = Field(min_length=32, max_length=100)
    result: WpsExecutionResult


class WpsExecutionStateResponse(BaseModel):
    operationId: str
    status: Literal["running", "needs_reconciliation", "completed", "failed"]
    revision: int = Field(ge=1)
    workbookIdentity: str
    taskId: str
    descriptor: WpsRecoveryDescriptor
    authorizedSheetNames: list[str]
    nextActionIndex: int = Field(ge=0)
    totalActions: int = Field(ge=0)
    inFlightActionIndex: int | None = Field(default=None, ge=0)
    beforeSnapshot: WpsActionEvidence | None = None
    checkpoint: WpsActionEvidence | None = None
    result: WpsExecutionResult | None = None
    error: dict[str, Any] | None = None
    canResume: bool
    requiresReconciliation: bool
    ownerActive: bool
    ownerToken: str | None = None


class ColumnMappingPair(BaseModel):
    """映射表的一项。字段名 `source` / `target`，**不设 `from` / `to` 别名**。

    模型侧的工具参数用的是 `{from, to}`（`from` 在 Python 是关键字，只能靠别名），
    但**响应模型不跟着用**：别名只在 `by_alias=True` 时才出现在 JSON 里，而本仓库
    有两条序列化路径 —— `response_model`（FastAPI 默认带 by_alias）与 SSE 里的
    裸 `model_dump(mode="json")`（不带）。同一个模型在两条路径上吐出不同的键名，
    客户端契约校验会在其中一条上静默失败。去掉别名让两条路径必然一致。

    `source` 允许空串：把 "" 写进映射表是"区分空串与真空格"的合法意图
    （`column_mapping.py` 的显式项优先于空白策略）。`target` 不允许 —— 空标签
    写下去与"未写入"在回读侧不可区分。
    """

    source: str = Field(max_length=500)
    target: str = Field(min_length=1, max_length=500)


class ColumnMappingRankEntry(BaseModel):
    """拆分后取最高级时的一条次序声明。`rank` 大者优先。

    等级**由规则数据给出**，服务端与 core 都不认识任何具体等级名 ——
    在意的只有数值大小，以及"没声明时按 `mapping` 声明顺序"这条确定性兜底。
    """

    value: str = Field(min_length=1, max_length=500)
    rank: int


class BlankPolicy(BaseModel):
    """来源格为**空白**时的结构化处理策略。

    与 `packages/core/src/blankPolicy.ts` 同名同义，四种形态：

    - `keep`：保留目标原值 —— 这个位置**根本不进入写入范围**，不是"读回原值
      再写回"（后者是一次真实写入，会改公式与格式）。
    - `clear`：清空目标内容 —— 走 `clearRange:contents`，不是"写空串"。
    - `value`：写入指定值，`value` **允许是空串**（显式写空字符串常量）。
    - `undecided`：还没决定。**不是错误** —— 只有真的读到空白格时才需要决定。

    为什么不再只用一个 `blankLabel: str`：它表达不了 `keep`，也分不清
    `clear` 与"写空串"。用工具能力缺陷去逼用户回答，是把实现问题转嫁给用户。
    """

    model_config = ConfigDict(extra="forbid")

    kind: Literal["keep", "clear", "value", "undecided"]
    # 仅 `kind="value"` 时有意义，且**允许空串**（"显式写空字符串"是明确意图）。
    value: str | None = Field(default=None, max_length=100)

    @model_validator(mode="after")
    def _value_matches_kind(self) -> "BlankPolicy":
        if self.kind == "value" and self.value is None:
            raise ValueError('空白策略 kind="value" 必须带 value')
        if self.kind != "value" and self.value is not None:
            raise ValueError('只有 kind="value" 的空白策略才能带 value')
        return self


class ColumnMappingClarification(BaseModel):
    """待澄清卡片的内容。**不含来源格的原始内容**。

    原始内容留在客户端，由界面就地展示；下发/回传的只有归因与统计
    —— 否则"客户端不上传单元格值"这条会被澄清流程从侧门绕开。
    """

    reason: Literal[
        "matchedDelimiterMissingCategories",
        "unresolvedSeparatorOrMapping",
        # 「读到源区域确实有空白格，但空白怎么处理还没定」。
        # 没读到空白就不该出现 —— 那是凭空提问。
        "blankPolicyNeeded",
    ]
    message: str = Field(min_length=1, max_length=1000)
    # 未覆盖的类别（已去重），供界面列出"要补哪些"。
    missingCategories: list[str] = Field(default_factory=list, max_length=500)
    # 归因不确定时，无法处理的整格**数量**（不回传原文）。
    unresolvedCellCount: int = Field(default=0, ge=0)
    # `blankPolicyNeeded` 时**真实读到**的空白格数量。必须来自一次真实读取。
    blankCellCount: int = Field(default=0, ge=0)
    resolvedBy: list[
        Literal[
            "addCategories",
            "specifyOtherDelimiter",
            "treatWholeCellAsCategory",
            # 空白的四种处理。与上面那组**不混用**：空白卡片上出现
            # 「补充分隔符」只会让用户以为那也能解决空白。
            "keepTarget",
            "clearTarget",
            "writeBlankValue",
            "writeEmptyString",
        ]
    ] = Field(default_factory=list, max_length=4)


class ColumnMappingRule(BaseModel):
    """列映射的**规则**，不含任何单元格值。

    源值不在服务端时（`planningReadPolicy == "none"`），服务端把规则原样下发，
    由客户端调 `packages/core` 的 `buildColumnMappingActions` 就地算出写入计划。
    两侧算出的值必然相同，由 `config/column_mapping_cases.json` 锁住。

    这里刻意不带 `values`：带了就等于把映射结果上传，而单射映射可逆
    —— 标签列加映射表足以还原源列，那与"客户端不上传单元格值"是冲突的。

    多值语义的三组参数（`delimiters` / `valueRank` / `undeclaredSeparators`）
    与 `packages/core` 的 `ColumnMappingInput` **逐字同名**：规则是平的，
    折一层子对象会让"夹具里有、运行时没有"这种错在类型检查下隐形。
    三组都可省略，省略即"一个来源格只当一个值"，与扩展前行为一致。
    """

    kind: Literal["columnMapping"] = "columnMapping"
    sourceSheet: str = Field(min_length=1, max_length=200)
    targetSheet: str = Field(min_length=1, max_length=200)
    sourceRange: str = Field(min_length=2, max_length=50)
    targetRange: str = Field(min_length=2, max_length=50)
    # 保持 [{from,to}] 列表形态而不折成 dict：与模型侧的工具 schema 同形，
    # 且重复键在这一层仍可见（折成 dict 会静默丢掉后者）。
    mapping: list[ColumnMappingPair] = Field(min_length=1, max_length=500)
    # 空白来源格的处理策略。省略 = 未定，只有真的读到空白时才需要决定。
    #
    # 表达「保留目标原值」只有这一种方式：旧字段 `blankLabel` 只能表达
    # "写一个什么标签"，而"保留"意味着这个位置**根本不进入写入范围**，
    # 不是"读回原值再写回"。
    blankPolicy: BlankPolicy | None = None
    # 旧协议的"空白标签"，保留只为兼容旧数据。
    # 非空串等价于 `blankPolicy: {kind:"value", value: 同值}`，语义一点没动；
    # 旧协议里空串本来就被拒绝，所以"空串 = 未定"不会改变任何合法旧计划。
    blankLabel: str | None = Field(default=None, max_length=100)
    # 候选分隔符。拆分时取**出现位置最靠前**的那个。空列表 = 不拆分。
    delimiters: list[str] = Field(default_factory=list, max_length=50)
    # 拆分后取最高级用的次序。空列表 = 按 `mapping` 声明顺序。
    valueRank: list[ColumnMappingRankEntry] = Field(
        default_factory=list, max_length=500
    )
    # 规则明示为分隔符候选、但尚未确认可用来切分的字符。
    undeclaredSeparators: list[str] = Field(default_factory=list, max_length=50)
    # 服务端已判定"缺参数/缺类别、不能算"时带上；带上时客户端只展示澄清卡片。
    clarification: ColumnMappingClarification | None = None


class LocalPlanRuleResponse(BaseModel):
    """"客户端自己算"的响应：服务端只给规则，不给值。

    不伪装成 `PlanResponse`：它没有可执行的 `actions`，塞个占位动作会让
    审批 UI 展示一份并不成立的计划。第三种类型是诚实的形状。
    """

    kind: Literal["local_plan_rule"] = "local_plan_rule"
    rule: ColumnMappingRule
    summary: str = Field(min_length=1, max_length=500)
    warnings: list[str] = Field(default_factory=list, max_length=20)
    assumptions: list[str] = Field(default_factory=list, max_length=20)
    provider: Literal["model", "local"]
    turnId: TurnId | None = None


class PlanResponse(BaseModel):
    # Trusted compiler metadata, never accepted from model-authored JSON.
    planOrigin: Literal["literal_range_edit"] | None = None
    kind: Literal["plan"] = "plan"
    plan: AnalysisPlan
    provider: Literal["model", "local"]
    turnId: TurnId | None = None


class LocalResultCard(BaseModel):
    """Pointer + structure for a local result card. Never carries values.

    Values stay in the trusted volatile registry; the local renderer fetches
    them over the loopback result endpoint using ``resultRef``. This keeps the
    matrix out of the model prompt, the chat history and the durable ledger.
    """

    model_config = ConfigDict(extra="forbid")
    resultRef: str = Field(pattern=r"^res_[0-9a-f]{32}$")
    title: str = Field(min_length=1, max_length=300)
    headers: list[str] = Field(min_length=1, max_length=50)
    rowCount: int = Field(ge=0)
    columnCount: int = Field(ge=1)
    sourceSheets: list[str] = Field(default_factory=list, max_length=100)
    complete: bool = True


class AnswerResponse(BaseModel):
    kind: Literal["answer"] = "answer"
    message: str = Field(min_length=1, max_length=4000)
    provider: Literal["model", "local", "system"]
    outcome: Literal["completed", "failed"] = "completed"
    stopReason: Literal["max_turns", "max_tokens", "timeout", "consecutive_errors", "output_length"] | None = None
    resultContext: ResultContext | None = None
    # SDK 本地结果卡：只带引用与结构，值由本机回环端点按引用取回。
    localResult: LocalResultCard | None = None
    turnId: TurnId | None = None


AssistantResponse = Annotated[
    PlanResponse | AnswerResponse | LocalPlanRuleResponse | IntentClarificationResponse,
    Field(discriminator="kind"),
]


class TurnStepEvent(BaseModel):
    """流式分步进度事件；仅用于 /api/turn/stream 的 step 事件，不参与写入契约。"""

    phase: Literal["planning"] = "planning"
    title: str = Field(min_length=1, max_length=200)
    detail: str | None = Field(default=None, max_length=500)
    completedStep: str | None = Field(default=None, max_length=200)


class TaskProgressTarget(BaseModel):
    file: str | None = Field(default=None, max_length=1000)
    sheet: str | None = Field(default=None, max_length=500)
    range: str | None = Field(default=None, max_length=80)


class TaskProgressFact(BaseModel):
    label: str = Field(min_length=1, max_length=40)
    value: str = Field(min_length=1, max_length=500)
    style: Literal["text", "code"] = "text"


class TaskProgressTrace(BaseModel):
    kind: Literal[
        "understanding",
        "evidence",
        "assumption",
        "decision",
        "conflict",
        "correction",
    ]
    summary: str = Field(min_length=1, max_length=1000)
    basis: str | None = Field(default=None, max_length=1000)
    changedFrom: str | None = Field(default=None, max_length=1000)
    requiresConfirmation: bool = False


class PythonCodeEvent(BaseModel):
    """模型生成的 Python 代码工件的状态（不是单元格数据）。

    代码走独立字段而不是 ``detail``：``detail`` 是"给人看的一句话"，会被截断、
    会被拼接，也会在没有代码时照常出现。这里同时带上工件身份
    （``codeId`` / ``attemptId`` / ``version``），前端才能判定"已运行的就是
    屏幕上这一版"，而不是靠文案猜。
    """

    model_config = ConfigDict(extra="forbid")

    codeId: str = Field(min_length=1, max_length=120)
    attemptId: str = Field(min_length=1, max_length=120)
    stage: Literal[
        "generating",
        "generated",
        "validating",
        "running",
        "analysis_completed",
        "summarizing",
        "paused",
        "not_started",  # 未收到工作程序执行确认，不等同于 Docker CLI 是否启动
        "generation_failed",
        "validation_failed",
        "execution_failed",
    ]
    version: int = Field(default=1, ge=1)
    # 代码正文：上限来自 config/capabilities.json 的 pythonCode.maxCharsPerEvent。
    code: str | None = Field(default=None, max_length=20000)
    delta: str | None = Field(default=None, max_length=20000)
    # 正文字符数：正文不随帧下发时，也要如实告知规模（不是"没有代码"）。
    chars: int | None = Field(default=None, ge=0)
    # 这一帧是否携带正文。账本里的历史帧固定为 false —— 代码不进持久化现场。
    codeRetained: bool | None = None
    truncated: bool = False
    complete: bool | None = None
    executedVersion: int | None = Field(default=None, ge=1)
    sequence: int | None = Field(default=None, ge=0)
    # 未开始执行的具体原因（固定文案，不含路径/请求 ID/原始错误）。
    reason: str | None = Field(default=None, max_length=300)

    @model_validator(mode="after")
    def validate_code_payload(self) -> "PythonCodeEvent":
        # 内容阶段必须带正文；后续的"校验/执行/完成"帧只推进状态，不重复搬运代码
        # （代码已经在同一 stepId 的第一帧里送达，前端按工件身份归并）。
        if self.stage in ("generating", "generated"):
            if self.code is None and self.delta is None and not self.chars:
                raise ValueError("pythonCode 内容阶段必须带 code/delta 或 chars")
        return self


class TaskProgressEvent(BaseModel):
    """由本地代码构造并校验的用户可见任务进度，不承载模型隐藏推理。"""

    schemaVersion: Literal[1] = 1
    operationId: TurnId
    executionId: str | None = Field(
        default=None, pattern=r"^[0-9a-f]{32}$"
    )
    sequence: int = Field(ge=0)
    phase: Literal[
        "intent",
        "planning",
        "preflight",
        "backup",
        "executing",
        "saving",
        "verifying",
        "refreshing",
    ]
    stepId: str = Field(
        min_length=1, max_length=100, pattern=r"^[A-Za-z0-9_.:-]+$"
    )
    status: Literal[
        "started", "progress", "completed", "warning", "failed", "cancelled"
    ]
    title: str = Field(min_length=1, max_length=200)
    summary: str | None = Field(default=None, max_length=1000)
    detail: str | None = Field(default=None, max_length=2000)
    current: int | None = Field(default=None, ge=0)
    total: int | None = Field(default=None, ge=0)
    unit: Literal[
        "rows", "fields", "actions", "sheets", "files", "checks"
    ] | None = None
    actionIndex: int | None = Field(default=None, ge=0)
    actionCount: int | None = Field(default=None, ge=0)
    actionType: str | None = Field(default=None, max_length=100)
    toolName: str | None = Field(default=None, max_length=100)
    facts: list[TaskProgressFact] | None = Field(default=None, max_length=12)
    target: TaskProgressTarget | None = None
    trace: TaskProgressTrace | None = None
    # 本机桥指针：只带操作 ID 与工具名，不含请求参数或任何值。
    bridgeTool: BridgeToolPointer | None = None
    # Python 分析结果卡：包含引用与结构，值由前端另行查询。
    pythonResult: LocalResultCard | None = None
    # Python 代码工件状态（代码正文 + 工件身份 + 阶段），与 pythonResult 分开。
    pythonCode: PythonCodeEvent | None = None
    timestamp: str = Field(min_length=1, max_length=64)

    @model_validator(mode="after")
    def validate_progress_numbers(self) -> "TaskProgressEvent":
        if (
            self.current is not None
            and self.total is not None
            and self.current > self.total
        ):
            raise ValueError("current 不能大于 total")
        return self

TurnRequest = IntentCheckRequest | PlanRequest
TurnResponse = Annotated[
    IntentProceedResponse
    | IntentConversationResponse
    | IntentClarificationResponse
    | IntentToolResponse
    | PlanResponse
    | AnswerResponse
    | LocalPlanRuleResponse,
    Field(discriminator="kind"),
]


class WorksheetDeleteCheckpointRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    operationId: str = Field(pattern=r"^[a-f0-9]{32}$")
    host: Literal["excel", "wps-et"]
    workbookPath: str = Field(min_length=1, max_length=2048)
    hostWindowId: int | None = Field(default=None, gt=0)
    authorizedSheets: list[str] = Field(min_length=1, max_length=256)
    plan: AnalysisPlan
    # Attribution of the originating conversation/request message. Optional so that
    # records written before this field existed stay readable: missing means
    # "unknown", never "the current conversation".
    conversationId: str | None = Field(default=None, min_length=1, max_length=128)
    requestMessageId: str | None = Field(default=None, min_length=1, max_length=128)


class WorksheetCheckpointReference(BaseModel):
    model_config = ConfigDict(extra="forbid")
    operationId: str = Field(pattern=r"^[a-f0-9]{32}$")


class WorksheetCheckpointBackupConsent(WorksheetCheckpointReference):
    backupSheets: list[str] = Field(min_length=1, max_length=256)


class WorksheetRestoreResult(BaseModel):
    """原地撤销的**实际结果**（不是"打开了某个副本"）。

    只报做过的事：哪些表被放回原工作簿、哪些没放回、修了几处被本次删除打断的
    跨表公式与工作簿级名称、哪些类型无法核对。`status` 为 `completed` 才代表
    "名称/顺序/可见性/内容都已核对一致"；`partial` 与 `failed` 都不许显示成已撤销。
    """

    operationId: str
    status: Literal["completed", "partial", "failed", "conflict"]
    restoredSheets: list[str] = Field(default_factory=list)
    failedSheets: list[str] = Field(default_factory=list)
    unverifiedSheets: list[str] = Field(default_factory=list)
    repairedFormulas: int = 0
    repairedNames: int = 0
    limitations: list[str] = Field(default_factory=list)
    message: str


class WorksheetDeleteCheckpointResult(BaseModel):
    operationId: str
    status: Literal["prepared", "needs_backup_consent", "executing", "completed", "partial_or_unknown", "failed"]
    backupSheets: list[str] = Field(default_factory=list)
    deletedSheets: list[str] = Field(default_factory=list)
    remainingSheets: list[str] = Field(default_factory=list)
    recoveryAvailable: bool = False
    message: str
    # 原地撤销的最近一次结果。旧记录没有这一项 —— 缺失只表示"还没撤销过"，
    # 不表示"不能撤销"（能不能撤销由 status / recoveryAvailable / 身份核对决定）。
    restore: WorksheetRestoreResult | None = None


class WorksheetCheckpointOpenResult(BaseModel):
    opened: bool
    recoveryPath: str


class WorksheetCheckpointListRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    workbookPath: str = Field(min_length=1, max_length=2048)
    host: Literal["excel", "wps-et"]


class WorksheetCheckpointSummary(BaseModel):
    operationId: str
    status: Literal["prepared", "needs_backup_consent", "executing", "completed", "partial_or_unknown", "failed"]
    deletedSheets: list[str] = Field(default_factory=list)
    remainingSheets: list[str] = Field(default_factory=list)
    recoveryAvailable: bool = False
    message: str
    createdAt: str | None = None
    planTitle: str | None = None
    # 「备份与恢复」列表要知道这一条是否已经原地撤销过；缺失表示没有撤销记录。
    restoreStatus: Literal["completed", "partial", "failed", "conflict"] | None = None


class WorksheetCheckpointListResult(BaseModel):
    checkpoints: list[WorksheetCheckpointSummary] = Field(default_factory=list)
