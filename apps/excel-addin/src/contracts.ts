// CellValue 的权威定义在 packages/core/src/contracts.ts（L0）。
export interface PlanningPauseAck {
  turnId: string;
  /** 接受停止请求；是否已停下仍须查询任务状态。 */
  pauseRequested: boolean;
}
// 本文件内部大量引用该类型，故 import 后再导出，避免两端各持一份定义后语义漂移。
import type { CellValue } from "@excel-bro/core/contracts";
// 空白策略的权威定义在 packages/core/src/blankPolicy.ts（L2）。协议层只是转发，
// 不另起一份 —— 四种意图一旦在两处各写一遍，就会悄悄漂移。
import type { BlankPolicy } from "@excel-bro/core";
import type {
  PythonCodeEvent,
  PythonCodeLimits,
  PythonCodeStage
} from "@excel-bro/core/pythonCodeArtifact";
// 宿主观测的类型与判据都在 core（唯一口径），协议层只转发。
import type {
  HostWriteAudit,
  HostWriteObservation,
  HostWriteSummary
} from "@excel-bro/core/hostWriteAudit";
// 大小/折叠上限的唯一出处：`config/capabilities.json`（代码只做兜底）。
import capabilities from "../../../config/capabilities.json";

export type { CellValue, BlankPolicy };
export type {
  PythonCodeArtifact,
  PythonCodeEvent,
  PythonCodeStage
} from "@excel-bro/core/pythonCodeArtifact";
export type { ReplySegment } from "@excel-bro/core/replySegments";
export type { ModelUsageErrorCode } from "@excel-bro/core/modelUsageNotice";
export type { PythonResultReceipt, PythonResultReadRequest, PythonMatrixResult } from "@excel-bro/core/contracts";
export type { PythonDataSource, PythonDataScope, PythonDataScopeRequest,
  PythonDataSnapshotRequest, PythonDataRevokeRequest, PythonDatasetReceipt } from "@excel-bro/core/contracts";

// Local Windows layout bridge. No workbook paths, cell data, plans or HWNDs.
export interface ExcelReviewLayoutPrepareRequest {
  workbookPathHash: string;
  worksheetNameHash: string;
  copySheetName: string;
}
export interface ExcelReviewLayoutTicket { token: string; layoutVersion: 3 }
export interface ExcelReviewPaneMeasurement {
  /** 原表窗口内容区宽度（Win32 client rect；诊断/兼容用，不用于计算 AI 比例）。 */
  clientWidthCssPx: number;
  /** 原表所在显示器的整个工作区宽度 —— AI 窗格比例的唯一参照。 */
  workAreaWidthCssPx: number;
}
export interface ExcelReviewLayoutResult {
  status: "arranged";
  layoutVersion: 3;
  copySide: "left";
  originalSide: "right";
  synchronousScrolling: "enabled" | "different_instance" | "unavailable";
}
/**
 * 关闭本次审阅副本的结果。**不包含**原表写回结果 —— 那是业务事实，
 * 由前端持有；副本关闭失败绝不允许把"原表写回成功"改写成"执行失败"。
 * `already_closed` 是**成功**收尾（重复清理或用户自己已关掉）。
 */
export interface ExcelReviewCopyCloseRequest {
  closeVersion: 4;
  workbookPathHash: string;
  worksheetNameHash: string;
  copySheetName: string;
}
export type ExcelReviewCopyCloseStatus =
  | "closed" | "already_closed" | "close_unconfirmed" | "user_saved"
  | "identity_lost" | "ambiguous" | "unsupported" | "failed";
export interface ExcelReviewCopyCloseResult {
  status: ExcelReviewCopyCloseStatus;
  layoutVersion: 4;
  copyClosed: boolean;
  recoveryCopyPath: string | null;
}

// /function 短链：AI 生成原生 Excel 公式（主路径）
export interface FormulaDictionarySheet {
  name: string;
  rows: string[][];
}

// /function 多工作簿：用户勾选的外部工作簿工作表上下文（供生成跨文件公式）。
export interface FormulaExtraSheet {
  sourceFile: string;
  sourcePath: string;
  sheetName: string;
  headers: string[];
  columns: string[];
  sampleRows: string[][];
  /** 外部表总行数（含表头行），本地拼跨表公式时用于确定引用区域末行。 */
  rowCount: number;
}

export interface GenerateFormulaRequest {
  description: string;
  activeCell: string;
  headers?: string[];
  columns?: string[];
  sampleRows?: string[][];
  dictionary?: FormulaDictionarySheet | null;
  extraSheets?: FormulaExtraSheet[];
  modelId?: string | null;
}

export interface GenerateFormulaResponse {
  modernFormula: string;
  modernExplanation: string;
  compatFormula: string;
  compatExplanation: string;
}

export interface WorksheetSnapshot {
  name: string;
  sourceFileId?: string | null;
  sourceSheetId?: string | null;
  sourceFile?: string | null;
  sourceSheet?: string | null;
  usedRange: string | null;
  rowCount: number;
  columnCount: number;
  headers: CellValue[];
  previewRows?: CellValue[][];
  dataRows: CellValue[][];
  displayRows?: CellValue[][];
  truncated: boolean;
}

export interface WorkbookSnapshot {
  name: string;
  /** Metadata directory only; dimensions/fields have not been captured. */
  catalogOnly?: boolean;
  capturedAt: string;
  activeWorksheet: string;
  selectedRange?: string | null;
  sourceFingerprint?: string | null;
  sourceFingerprintSheets?: string[];
  /**
   * 稳定工作簿身份，沿用既有 ``url:`` / ``wps-path:`` 命名空间。
   * 名字相同但身份不同的两个工作簿必须彼此隔离，桥归属按此判定。
   */
  workbookIdentity?: string | null;
  worksheets: WorksheetSnapshot[];
}

/**
 * 格式动作的封闭取值域 —— **与 `server/app/models.py` 里的 `Literal` 逐字一致**。
 *
 * 抽成具名别名的两个原因：
 *   1. 同一个枚举在 action 与 criterion 两处出现，字面量写两遍迟早会漂；
 *   2. 跨语言一致性测试（`server/tests/test_protocol_consistency.py`）要能从
 *      TS 源码里把这几组取值抠出来跟 Python 的 `Literal` 逐项比对，
 *      具名别名让"抠哪一段"是确定的，不用去猜。
 *
 * 改这里**必须**同步改 `models.py`，否则一致性测试会红。
 */
export type BorderSide =
  | "top"
  | "bottom"
  | "left"
  | "right"
  | "insideHorizontal"
  | "insideVertical";
export type BorderLineStyle =
  | "continuous"
  | "dash"
  | "dashDot"
  | "dot"
  | "double"
  | "none";
export type BorderWeight = "hairline" | "thin" | "medium" | "thick";
export type HorizontalAlignmentRule =
  | "general"
  | "left"
  | "center"
  | "right"
  | "fill"
  | "justify"
  | "centerAcrossSelection";
export type VerticalAlignmentRule =
  | "top"
  | "center"
  | "bottom"
  | "justify"
  | "distributed";

export const BORDER_SIDES: readonly BorderSide[] = [
  "top",
  "bottom",
  "left",
  "right",
  "insideHorizontal",
  "insideVertical"
];
export const BORDER_LINE_STYLES: readonly BorderLineStyle[] = [
  "continuous",
  "dash",
  "dashDot",
  "dot",
  "double",
  "none"
];
export const BORDER_WEIGHTS: readonly BorderWeight[] = [
  "hairline",
  "thin",
  "medium",
  "thick"
];
export const HORIZONTAL_ALIGNMENTS: readonly HorizontalAlignmentRule[] = [
  "general",
  "left",
  "center",
  "right",
  "fill",
  "justify",
  "centerAcrossSelection"
];
export const VERTICAL_ALIGNMENTS: readonly VerticalAlignmentRule[] = [
  "top",
  "center",
  "bottom",
  "justify",
  "distributed"
];

/**
 * 颜色一律 `#RRGGBB`。Python 侧用 `StringConstraints(pattern=...)` 强制，
 * TS 侧在 `assertEditablePlan` 里用同一条 pattern 做运行时校验 ——
 * 只在类型上写 `string` 等于放行 `red`、`#fff`、`rgb(...)`，
 * 那些值到了宿主才失败，会把失败点推到手已经动过之后。
 */
export const HEX_COLOR_PATTERN = /^#[0-9A-Fa-f]{6}$/;

/** 行高/列宽与宿主一致的上下界（Excel 与 WPS 都用磅 / 字符宽，取 Office.js 的值域）。 */
export const ROW_HEIGHT_RANGE = { min: 0, max: 409 } as const;
export const COLUMN_WIDTH_RANGE = { min: 0, max: 255 } as const;
/** 字体名的长度上限，与 Python `Field(min_length=1, max_length=255)` 对齐。 */
export const FONT_NAME_MAX_LENGTH = 255;
/** 字号值域，与 Python `Field(ge=1, le=409)` 对齐。 */
export const FONT_SIZE_RANGE = { min: 1, max: 409 } as const;
/** 数字格式码长度上限，与 Python `Field(min_length=1, max_length=100)` 对齐。 */
export const NUMBER_FORMAT_MAX_LENGTH = 100;

export type ExcelAction =
  | { type: "createWorksheet"; sheet: string }
  | {
      type: "writeTable";
      sheet: string;
      startCell: string;
      headers: CellValue[];
      rows: CellValue[][];
    }
  | {
      type: "writeValues";
      sheet: string;
      range: string;
      values: CellValue[][];
    }
  | {
      type: "setFill";
      sheet: string;
      range: string;
      color: string;
    }
  | {
      /**
       * 字体动作。**只改这里显式给出的属性**，其余属性与值/公式保持不变。
       *
       * 五个属性各自可选且互相独立：`bold`/`italic` 只动字形，`name`/`size` 只动
       * 字族与字号，`color` 只动字色。模型可以只给其中一项。
       *
       * 每个字段都允许 `null`：服务端 `response_model_exclude_none=True` 会把未设置
       * 的字段整个省略，但计划编辑器与手工构造的计划可能显式给 `null`，
       * 两者语义相同 —— 都是"这一项不要动"。执行侧一律用 `!= null` 判空。
       */
      type: "setFont";
      sheet: string;
      range: string;
      bold?: boolean | null;
      italic?: boolean | null;
      /** 字体名（如 `"微软雅黑"`）。长度上限见 `FONT_NAME_MAX_LENGTH`。 */
      name?: string | null;
      /** 字号（磅）。值域见 `FONT_SIZE_RANGE`。 */
      size?: number | null;
      color?: string | null;
    }
  | { type: "autofit"; sheet: string; range: string }
  | { type: "activateWorksheet"; sheet: string }
  | { type: "deleteWorksheet"; sheet: string }
  | {
      type: "clearRange";
      sheet: string;
      range: string;
      applyTo: "all" | "contents" | "formats" | "hyperlinks";
      filters?: DataFilter[];
      hasHeaders?: boolean;
    }
  | {
      type: "insertRange";
      sheet: string;
      range: string;
      shift: "down" | "right";
    }
  | {
      type: "deleteRange";
      sheet: string;
      range: string;
      shift: "up" | "left";
      filters?: DataFilter[];
      hasHeaders?: boolean;
    }
  | {
      type: "copyRange";
      sheet: string;
      sourceSheet: string;
      sourceRange: string;
      targetRange: string;
      copyType: "all" | "values" | "formulas" | "formats" | "link";
      skipBlanks: boolean;
      transpose: boolean;
      filters?: DataFilter[];
      hasHeaders?: boolean;
    }
  | {
      type: "writeFormulas";
      sheet: string;
      range: string;
      formulas: string[][];
      /**
       * 锚点无关的 R1C1 形式；存在时按区域形状铺满，相对引用逐格自动平移。
       *
       * 口径（2026-10-30 核对真实传输与消费后统一；三件事分开说）：
       * - **字段名一致**：两端都叫 `formulaR1C1`；
       * - **类型一致**：可省略、可为 `null`（服务端 `str | None = None`），
       *   但**不能**是别的类型 —— 本文件只声明 `string`/`null`；
       * - **运行时约束**：服务端 `max_length=512`（513 字符 ⇒ 422 `string_too_long`），
       *   本文件没有长度预检；**空串等同"没有 R1C1"**，两端与两个宿主都按真值判
       *   （服务端 `if action.formulaR1C1`、Excel `if (action.formulaR1C1)`、
       *   WPS 同一写法），所以 `""`/`null`/省略 三者行为相同。
       *
       * 真实客户端（`App.tsx`）只会"非空才带上"，不会发 `null`；`null` 之所以在
       * 类型里，是因为服务端本来就接受它 —— 类型与线上一旦不一致，读代码的人只会
       * 得出错误结论。
       */
      formulaR1C1?: string | null;
    }
  | {
      type: "sortRange";
      sheet: string;
      range: string;
      keys: Array<{ column: number; ascending: boolean }>;
      hasHeaders: boolean;
    }
  | {
      type: "removeDuplicates";
      sheet: string;
      range: string;
      columns: number[];
      hasHeaders: boolean;
      filters?: DataFilter[];
    }
  | {
      type: "filterRange";
      sheet: string;
      range: string;
      column: number;
      values: CellValue[];
    }
  | { type: "clearFilter"; sheet: string }
  | {
      type: "setDataValidation";
      sheet: string;
      range: string;
      validationType:
        | "list"
        | "wholeNumber"
        | "decimal"
        | "date"
        | "textLength"
        | "custom";
      values: CellValue[];
      formula1?: string | number | null;
      formula2?: string | number | null;
      operator: string;
      allowBlank: boolean;
      prompt?: string | null;
      errorMessage?: string | null;
    }
  | {
      type: "setConditionalFormat";
      sheet: string;
      range: string;
      ruleType: "cellValue" | "custom" | "colorScale";
      operator?: string | null;
      formula1?: string | number | null;
      formula2?: string | number | null;
      color?: string | null;
      minColor?: string | null;
      midColor?: string | null;
      maxColor?: string | null;
    }
  | { type: "setNumberFormat"; sheet: string; range: string; formatCode: string }
  | {
      /**
       * 边框。`style`/`color`/`weight` 与 Python 侧一样**可以省略**（服务端有默认值
       * `continuous` / `#000000` / `thin`），执行侧用 `normalizeBorderAction` 补齐 ——
       * 类型层写必填而实现层要兼容缺省，中间那段空档就是"模型少给一个字段、
       * 计划在校验期通过、执行期才炸"的来源。
       */
      type: "setBorders";
      sheet: string;
      range: string;
      sides: BorderSide[];
      style?: BorderLineStyle | null;
      color?: string | null;
      weight?: BorderWeight | null;
    }
  | {
      type: "setAlignment";
      sheet: string;
      range: string;
      horizontal?: HorizontalAlignmentRule | null;
      vertical?: VerticalAlignmentRule | null;
      wrapText?: boolean | null;
    }
  | { type: "mergeCells"; sheet: string; range: string; across: boolean }
  | { type: "unmergeCells"; sheet: string; range: string }
  | {
      type: "resizeRange";
      sheet: string;
      range: string;
      rowHeight?: number | null;
      columnWidth?: number | null;
    }
  | { type: "freezePanes"; sheet: string; rows: number; columns: number }
  | {
      type: "setHyperlink";
      sheet: string;
      range: string;
      address: string;
      text?: string | null;
      screenTip?: string | null;
    }
  | { type: "addComment"; sheet: string; cell: string; text: string }
  | { type: "addNote"; sheet: string; cell: string; text: string }
  | {
      type: "createTable";
      sheet: string;
      range: string;
      name?: string | null;
      hasHeaders: boolean;
      style?: string | null;
    }
  | {
      type: "createChart";
      sheet: string;
      sourceRange: string;
      chartType: string;
      title?: string | null;
      targetRange?: string | null;
    }
  | {
      type: "createPivotTable";
      sheet: string;
      sourceSheet: string;
      sourceRange: string;
      name: string;
      destinationCell: string;
      rowFields: string[];
      columnFields: string[];
      valueFields: Array<{
        field: string;
        aggregation: "sum" | "count" | "average" | "max" | "min";
      }>;
    }
  | {
      type: "splitGroupAggregate";
      sheet: string;
      sourceRange?: string | null;
      splitBy: string;
      groupBy: string[];
      metrics: Array<{
        operation: "countRows" | "countNonBlank" | "sum";
        field?: string | null;
        outputName: string;
        ratioOutputName?: string | null;
      }>;
      includeBlankSplitValues: boolean;
      existingSheetPolicy: "rename" | "replace" | "skip";
      maxOutputSheets: number;
    }
  | {
      /**
       * 执行时求值：动作体只带查询配置，不带算好的结果。执行那一刻才读
       * sourceSheet 的实时数据算出输出，所以源数据改了计划依然成立。
       */
      type: "computeValues";
      sheet: string;
      sourceSheet: string;
      sourceRange?: string | null;
      hasHeaders: boolean;
      query: QueryTableArguments;
      destinationCell: string;
      includeHeaders: boolean;
    }
  | {
      type: "addNamedRange";
      sheet: string;
      name: string;
      range: string;
      comment?: string | null;
    }
  | {
      type: "addImage";
      sheet: string;
      base64: string;
      targetRange: string;
      name?: string | null;
    }
  | {
      type: "addShape";
      sheet: string;
      shapeType: string;
      targetRange: string;
      text?: string | null;
      fillColor?: string | null;
    };

export type VerificationCriterion =
  | { type: "worksheetExists"; sheet: string }
  | { type: "worksheetMissing"; sheet: string }
  | {
      type: "rangeEquals";
      sheet: string;
      range: string;
      expected: CellValue[][];
    }
  | { type: "rangeEmpty"; sheet: string; range: string }
  | {
      type: "formulasEqual";
      sheet: string;
      range: string;
      expected: string[][];
    }
  | {
      type: "formulasR1C1Equal";
      sheet: string;
      range: string;
      expected: string;
    }
  | {
      type: "rangeSorted";
      sheet: string;
      range: string;
      keys: Array<{ column: number; ascending: boolean }>;
      hasHeaders: boolean;
    }
  | {
      type: "filterApplied";
      sheet: string;
      range: string;
      column: number;
      values: CellValue[];
    }
  | { type: "filterCleared"; sheet: string }
  | {
      type: "tableExists";
      sheet: string;
      range: string;
      name?: string | null;
      hasHeaders: boolean;
    }
  | {
      /**
       * 格式回读核对。**至少给一项**，给到的每一项都必须与真实格式属性一致
       * —— 判据是"读宿主当前格式"，不是"值没变"。
       */
      type: "rangeFormatMatches";
      sheet: string;
      range: string;
      fillColor?: string | null;
      bold?: boolean | null;
      fontColor?: string | null;
      fontName?: string | null;
      fontSize?: number | null;
      fontItalic?: boolean | null;
      numberFormat?: string | null;
      horizontal?: HorizontalAlignmentRule | null;
      vertical?: VerticalAlignmentRule | null;
      wrapText?: boolean | null;
      rowHeight?: number | null;
      columnWidth?: number | null;
    }
  | {
      type: "bordersMatch";
      sheet: string;
      range: string;
      sides: BorderSide[];
      style: BorderLineStyle;
      color: string;
      weight: BorderWeight;
    }
  | {
      type: "dataValidationMatches";
      sheet: string;
      range: string;
      validationType:
        | "list"
        | "wholeNumber"
        | "decimal"
        | "date"
        | "textLength"
        | "custom";
      values: CellValue[];
      formula1?: string | number | null;
      formula2?: string | number | null;
      operator: string;
      allowBlank: boolean;
      prompt?: string | null;
      errorMessage?: string | null;
    }
  | {
      type: "freezePanesMatches";
      sheet: string;
      rows: number;
      columns: number;
    }
  | {
      type: "chartExists";
      sheet: string;
      name?: string | null;
      chartType: string;
      sourceRange: string;
      title?: string | null;
      targetRange?: string | null;
    }
  | {
      type: "pivotTableExists";
      sheet: string;
      sourceSheet: string;
      sourceRange: string;
      name: string;
      destinationCell: string;
      rowFields: string[];
      columnFields: string[];
      valueFields: Array<{
        field: string;
        aggregation: "sum" | "count" | "average" | "max" | "min";
      }>;
    };

export interface AnalysisPlan {
  id: string;
  title: string;
  summary: string;
  sourceFingerprint?: string | null;
  /**
   * **指纹覆盖范围**：生成这份计划时对它做了来源指纹的工作表。
   *
   * 它可能**宽于**计划的实际依赖（App 取的是本次授权范围）—— 变更检测宁可宽，
   * 但"宽"的代价是它不能用来回答"这份计划到底依赖了谁"。
   */
  sourceFingerprintSheets?: string[];
  /**
   * **实际来源依赖**：这份计划的值实际是从哪些工作表读出来的。
   *
   * 与 `sourceFingerprintSheets` 刻意分开建模，三者口径互不冒充：
   * 授权范围（`authorizedSheetNames`，边界）／实际来源依赖（本字段）／写入目标
   * （`actions[].sheet`）。
   *
   * 只由本端写入，且只来自**可信记录**：确定性本地读取、工具执行记录
   * （`DataToolResult.sourceSheets`）、已保存工具模板的声明来源。服务端或模型
   * 在计划里带回来的值一律由 `attestSourceDependency` 覆盖或剥离 —— 依赖是
   * "本机读了什么"的记账，不是可以被声明出来的东西。也**不**由"授权范围减去
   * 写入目标"推出来。
   *
   * 缺失 = 依赖不确定 = 不登记断点恢复（fail-closed，退回既有的保守口径）。
   */
  sourceDependencySheets?: string[];
  /**
   * **公式来源证据**：登记这份执行时，公式依赖的那几张表（
   * `formulaSourceEvidenceRequirement` 算出的集合）的值指纹。
   *
   * 口径只有一条，两端一致：**指纹覆盖的正是公式文本读的那几张表**，不是本次
   * 授权范围的整张快照。授权范围宽于依赖时，拿整张快照比对会在无关表被改时误拒；
   * 更糟的是它证明不了公式的输入没变（"只检查一个与公式依赖无关的快照"）。
   *
   * 写入时机：**登记断点恢复那一刻**（`register_office_execution` 之前由宿主现场读值），
   * 不是计划生成那一刻 —— 续跑要对照的是"开始写之前现场是什么"，中间夹着的预览期
   * 改动本来就该由首次执行按新值算。
   *
   * 缺失 = 没有登记来源证据 = 公式计划**不能**精确续跑（fail-closed）；不拿
   * "续跑时现读一次"补一个基线（那正是被反复禁止的写法）。
   */
  sourceDependencyFingerprint?: string | null;
  assumptions: string[];
  warnings: string[];
  actions: ExcelAction[];
  acceptanceCriteria?: VerificationCriterion[];
  /**
   * 本地结果引用：计划由本机查询结果编译而来时只带这个句柄。
   * 持久侧（规划账本 / 恢复副本）保存的是带引用、不带值的投影。
   */
  localResultRef?: string | null;
}

export interface PlanValidateRequest {
  plan: AnalysisPlan;
  authorized_sheets: string[];
  existing_sheets: string[];
}

export interface ResultContext {
  kind: "table";
  title: string;
  headers: string[];
  rows: CellValue[][];
  primaryValueColumn?: number | null;
  sourceSheets: string[];
  warnings: string[];
}

export interface ImageAttachment {
  name: string;
  mediaType: "image/png" | "image/jpeg" | "image/webp";
  data: string;
}

export interface IntentSheetContext {
  name: string;
  usedRange: string | null;
  rowCount: number;
  columnCount: number;
  headers: CellValue[];
}

export interface IntentScopeContext {
  workbookName: string;
  sourceMode: "workbook" | "folder";
  selectionMode: "auto" | "manual" | "folder";
  activeWorksheet: string;
  selectedRange?: string | null;
  totalWorksheetCount: number;
  // 全部数据工作表名（不含 #EB_* 系统表）。sheets 只含已选表的数据，结构操作
  // （删除/重命名等用户明确要求的跨表改动）需引用 worksheetNames 才能规划。
  worksheetNames: string[];
  sheets: IntentSheetContext[];
}

/** First semantic stage: intentionally has no workbook/schema/result fields. */
export interface MessageUnderstandingRequest {
  prompt: string;
  imageCount?: number;
  conversation?: Array<{ role: "user" | "assistant"; text: string }>;
  modelId?: string | null;
  /** Opaque task IDs or preplanning-* local request refs; no workbook data or authorization. */
  recoveryTasks?: string[];
}

export type MessageUnderstandingResponse =
  | { kind: "clarification"; clarification: IntentClarification; provider: "model" }
  | { kind: "conversation"; message: string; provider: "model" }
  | { kind: "needs_workbook_context"; summary: string; provider: "model" }
  | { kind: "task_recovery"; turnId: string; action?: "inspect" | "continue"; provider: "model" };

export function assertMessageUnderstandingResponse(value: unknown): asserts value is MessageUnderstandingResponse {
  if (!value || typeof value !== "object") throw new Error("消息理解响应格式无效");
  const item = value as Record<string, unknown>;
  if (item.kind === "clarification") {
    if (item.provider !== "model" || Object.keys(item).some((key) => !["kind", "provider", "clarification"].includes(key))) {
      throw new Error("消息理解澄清格式无效");
    }
    assertIntentCheckResponse(item);
    return;
  }
  const key = item.kind === "conversation" ? "message" : item.kind === "needs_workbook_context" ? "summary" : item.kind === "task_recovery" ? "turnId" : null;
  if (!key || item.provider !== "model" || typeof item[key] !== "string" || !(item[key] as string).trim()
      || (item[key] as string).length > (key === "message" ? 1500 : key === "turnId" ? 200 : 500)
      || (item.kind === "task_recovery" && item.action !== undefined && !["inspect", "continue"].includes(String(item.action)))
      || Object.keys(item).some((field) => !["kind", key, "provider", ...(item.kind === "task_recovery" ? ["action"] : [])].includes(field))) {
    throw new Error("消息理解响应格式无效；未读取工作表结构");
  }
}

export interface IntentCheckRequest {
  turnId?: string | null;
  prompt: string;
  /** Explicit semantic decisions, never authorization or local source values. */
  taskContext?: import("@excel-bro/core/agentTaskContext").AgentTaskContext | null;
  scope: IntentScopeContext;
  imageCount?: number;
  intentConfirmed?: boolean;
  clarificationRound?: number;
  conversation?: Array<{
    role: "user" | "assistant";
    text: string;
  }>;
  priorIntent?: IntentMemory | null;
  priorResult?: ResultContext | null;
  toolFailure?: {
    code: string;
    message: string;
    retryable: boolean;
    availableFields?: string[];
    request: DataToolRequest;
  } | null;
  modelId?: string | null;
}

export interface IntentOption {
  id: string;
  label: string;
  description: string;
  resolution: string;
  resolutionMode?: "append" | "replace";
  action?: "resolve" | "editScope";
}

export type DataFilterOperator =
  | "equals"
  | "notEquals"
  | "contains"
  | "greaterThan"
  | "greaterThanOrEqual"
  | "lessThan"
  | "lessThanOrEqual"
  | "isBlank"
  | "isNotBlank";

export interface DataFilter {
  field: string;
  operator: DataFilterOperator;
  value?: CellValue;
}

export interface DataMetric {
  operation:
    | "countRows"
    | "countDistinct"
    | "sum"
    | "average"
    | "min"
    | "max";
  field?: string | null;
  outputName: string;
  ratioOutputName?: string | null;
}

export interface QueryTableArguments {
  mode: "rows" | "aggregate" | "profile";
  scope?: "selected" | "active";
  fields?: string[];
  filters?: DataFilter[];
  groupBy?: string[];
  metrics?: DataMetric[];
  combine?: {
    mode: "union" | "deduplicate" | "join";
    deduplicateBy?: string[];
    leftSourceSheetId?: string | null;
    rightSourceSheetId?: string | null;
    leftKey?: string | null;
    rightKey?: string | null;
    joinHow?: "inner" | "left" | "right" | "outer";
  } | null;
  profileField?: string | null;
  sortBy?: string | null;
  sortDirection?: "asc" | "desc";
  limit?: number;
}

export interface DataToolRequest {
  id: string;
  tool: "query_table";
  arguments: QueryTableArguments;
}

export interface DeterministicQueryTemplate {
  id: string;
  name: string;
  description: string;
  sourceMode: "workbook" | "folder";
  request: DataToolRequest;
  sourceSheetNames: string[];
  sourceSheetIds: string[];
  expectedHeaders: string[];
}

export interface IntentMemory {
  confirmedPrompt: string;
  toolRequest?: DataToolRequest | null;
}

export interface ModelOption {
  agentAvailable?: boolean;
  agentUnavailableReason?: string | null;
  formulaAvailable?: boolean;
  id: string;
  label: string;
  provider: "model";
  available: boolean;
  supportsVision: boolean;
}

export interface ModelCatalog {
  defaultModelId: string | null;
  models: ModelOption[];
}

export interface ModelSettings {
  baseUrl: string | null;
  defaultModel: string | null;
  apiKeyConfigured: boolean;
  apiKeyHint: string | null;
  // /function 公式专用模型的 catalog id；空串=跟随全局选择。
  formulaModelId: string;
  connections: ManagedModelConnection[];
}

export type ModelConnectionProtocol = "openai_compatible" | "anthropic";
export type ModelConnectionPurpose = "agent" | "formula";

export interface ManagedModelConnection {
  protocol?: ModelConnectionProtocol;
  agentAvailable?: boolean;
  agentUnavailableReason?: string | null;
  formulaAvailable?: boolean;
  id: string;
  catalogModelId: string;
  label: string;
  baseUrl: string;
  modelId: string;
  supportsVision: boolean;
  apiKeyConfigured: boolean;
  apiKeyHint: string | null;
}

export interface UpdateModelSettingsRequest {
  apiKey: string;
}

export interface UpsertModelConnectionRequest {
  protocol?: ModelConnectionProtocol;
  id?: string | null;
  label: string;
  baseUrl: string;
  modelId: string;
  apiKey?: string | null;
  clearApiKey?: boolean;
  supportsVision: boolean;
}

export interface SetFormulaModelRequest {
  modelId: string;
}

export interface TestModelConnectionRequest extends UpsertModelConnectionRequest {
  purpose?: ModelConnectionPurpose;
}

export interface TestModelConnectionResponse {
  testedFor?: ModelConnectionPurpose;
  ok: boolean;
  message: string;
}

export function assertTestModelConnectionResponse(
  value: unknown
): asserts value is TestModelConnectionResponse {
  if (!value || typeof value !== "object") {
    throw new Error("服务返回的连接测试结果不是有效对象");
  }
  const result = value as Partial<TestModelConnectionResponse>;
  if (result.ok !== true || typeof result.message !== "string" ||
      (result.testedFor !== undefined && result.testedFor !== "agent" && result.testedFor !== "formula")) {
    throw new Error("服务返回的连接测试结果缺少必要字段");
  }
}

export function assertModelSettings(
  value: unknown
): asserts value is ModelSettings {
  if (!value || typeof value !== "object") {
    throw new Error("服务返回的模型设置不是有效对象");
  }
  const settings = value as Partial<ModelSettings>;
  if (
    (settings.baseUrl !== null && typeof settings.baseUrl !== "string") ||
    (settings.defaultModel !== null &&
      typeof settings.defaultModel !== "string") ||
    typeof settings.apiKeyConfigured !== "boolean" ||
    (settings.apiKeyHint !== null &&
      typeof settings.apiKeyHint !== "string") ||
    (settings.formulaModelId !== undefined &&
      typeof settings.formulaModelId !== "string") ||
    !Array.isArray(settings.connections) ||
    settings.connections.some(
      (connection) =>
        !connection ||
        typeof connection.id !== "string" ||
        typeof connection.catalogModelId !== "string" ||
        typeof connection.label !== "string" ||
        typeof connection.baseUrl !== "string" ||
        typeof connection.modelId !== "string" ||
        (connection.protocol !== undefined && connection.protocol !== "openai_compatible" && connection.protocol !== "anthropic") ||
        (connection.agentAvailable !== undefined && typeof connection.agentAvailable !== "boolean") ||
        (connection.formulaAvailable !== undefined && typeof connection.formulaAvailable !== "boolean") ||
        (connection.agentUnavailableReason != null && typeof connection.agentUnavailableReason !== "string") ||
        typeof connection.supportsVision !== "boolean" ||
        typeof connection.apiKeyConfigured !== "boolean" ||
        (connection.apiKeyHint !== null &&
          typeof connection.apiKeyHint !== "string")
    )
  ) {
    throw new Error("服务返回的模型设置无效");
  }
}

export interface DataToolResult {
  requestId: string;
  tool: "query_table";
  title: string;
  headers: CellValue[];
  rows: CellValue[][];
  sourceSheets: string[];
  scannedRows: number;
  complete: boolean;
  calculation: string;
  warnings: string[];
}

export function assertDataToolResult(
  value: unknown
): asserts value is DataToolResult {
  const item = value as Partial<DataToolResult>;
  if (
    !item ||
    typeof item !== "object" ||
    typeof item.requestId !== "string" ||
    item.tool !== "query_table" ||
    typeof item.title !== "string" ||
    !Array.isArray(item.headers) ||
    !item.headers.every((header) => typeof header === "string") ||
    !Array.isArray(item.rows) ||
    !item.rows.every((row) => Array.isArray(row)) ||
    !Array.isArray(item.sourceSheets) ||
    !item.sourceSheets.every((sheet) => typeof sheet === "string") ||
    typeof item.scannedRows !== "number" ||
    typeof item.complete !== "boolean" ||
    typeof item.calculation !== "string" ||
    !Array.isArray(item.warnings) ||
    !item.warnings.every((warning) => typeof warning === "string")
  ) {
    throw new Error("本地服务返回了无效的数据工具结果");
  }
}

// ---------------------------------------------------------------------------
// 任务工具桥协议（本机回环）：SDK handler ⇄ 本地宿主适配器。
// 与 server/app/models.py 的 Bridge* 模型一一对应。请求参数只经本回环通道，
// 不进模型提示、不进持久进度账本；值矩阵只在本通道与本地结果端点出现。
// ---------------------------------------------------------------------------

export type SdkBridgeHostId = "excel-office" | "wps-et" | "folder";
export type SdkBridgeToolName = "query_table" | "prepare_python_data";

/**
 * 进度事件里携带的桥指针：**只有操作 ID 与工具名**。
 * 请求参数留在本机易失桥里由宿主适配器拉取，不进模型提示、不进持久账本。
 */
export interface SdkBridgeToolPointer {
  operationId: string;
  tool: SdkBridgeToolName;
}

export interface SdkBridgeBindingRequest {
  taskId: string;
  hostId: SdkBridgeHostId;
  workbookName: string;
  /** 稳定工作簿身份；服务端缺省时退回 ``name:<工作簿名>``。 */
  workbookIdentity?: string | null;
  sourceFingerprint?: string | null;
}

export interface SdkBridgePendingCall {
  operationId: string;
  tool: SdkBridgeToolName;
  arguments: QueryTableArguments;
  requestFingerprint: string;
  expiresAt: number;
}

export interface SdkBridgePollResponse {
  calls: SdkBridgePendingCall[];
  generation: number;
}

/** ``prepare_python_data`` 宿主回执里的单块原始值矩阵（对应服务端 BridgeDatasetReceipt）。 */
export interface SdkBridgeDatasetReceipt {
  sourceIndex: number;
  values: CellValue[][];
}

export interface SdkBridgeToolReceipt {
  result?: DataToolResult | null;
  datasets?: SdkBridgeDatasetReceipt[] | null;
  errorCode?: string | null;
}

export interface SdkBridgeReceiptRequest extends SdkBridgeBindingRequest {
  operationId: string;
  requestFingerprint: string;
  receipt: SdkBridgeToolReceipt;
}

export interface SdkBridgeReceiptResponse {
  status: "accepted" | "duplicate" | "rejected";
  code?: string | null;
}

export interface SdkBridgeRegisterRequest extends SdkBridgeBindingRequest {
  result: DataToolResult;
}

export interface SdkBridgeRegisterResponse {
  resultRef: string;
  structure: {
    headers: CellValue[];
    rowCount: number;
    columnCount: number;
    complete: boolean;
  };
}

export interface SdkBridgeResultRequest extends SdkBridgeBindingRequest {
  resultRef: string;
}

/** 本地结果卡的完整载荷（含值），仅本机回环端点返回，不进模型/历史/账本。 */
export interface SdkBridgeResultResponse {
  resultRef: string;
  title: string;
  headers: CellValue[];
  rows: CellValue[][];
  sourceSheets: string[];
  warnings: string[];
  complete: boolean;
  calculation: string;
  scannedRows: number;
}

/**
 * 开发验收专用的「结果引用过期」入口（A1-E1）。
 *
 * 请求里**只有工作簿身份**：没有 `taskId`、没有 `resultRef`。任务与句柄由服务端在
 * 可信路径上绑定（开桥那一刻 = 哪次任务；该任务第一条 Python 结果登记那一刻 = 哪个
 * 引用），模型无从指定别的任务或别的引用。服务端未开启开发模式时这些端点表现为 404。
 */
export interface DevResultRefFaultArmRequest {
  hostId: "excel-office" | "wps-et";
  workbookName: string;
  workbookIdentity?: string;
}

export interface DevResultRefFaultStatus {
  enabled: boolean;
  phase: "off" | "idle" | "armed" | "bound" | "consumed" | "expired";
  expiresInSeconds: number;
  taskBound: boolean;
  refBound: boolean;
}

/**
 * 结果卡指针：只带引用与结构，**不含任何值**。
 * 真实值由本机回环端点按 `resultRef` 取回后由本地渲染器展示。
 */
export interface LocalResultCard {
  resultRef: string;
  title: string;
  headers: string[];
  rowCount: number;
  columnCount: number;
  sourceSheets: string[];
  complete: boolean;
}

export interface ClarificationField {
  id: string;
  label: string;
  placeholder?: string;
}

export interface IntentClarification {
  id: string;
  summary: string;
  question: string;
  reason: string;
  scopeLabel: string;
  options: IntentOption[];
  fields?: ClarificationField[];
}

export type IntentCheckResponse =
  | {
      kind: "conversation";
      message: string;
      provider: "model" | "local";
      turnId?: string | null;
    }
  | {
      kind: "proceed";
      planningGoal?: "read_only" | "write_preview";
      summary: string;
      confirmedPrompt: string;
      provider: "model" | "local";
      turnId?: string | null;
    }
  | {
      kind: "clarification";
      clarification: IntentClarification;
      provider: "model" | "local";
      turnId?: string | null;
    }
  | {
      kind: "tool_request";
      planningGoal?: "read_only" | "write_preview";
      summary: string;
      confirmedPrompt: string;
      request: DataToolRequest;
      provider: "model" | "local";
      turnId?: string | null;
    };

export interface PlanRequest {
  planningGoal?: "read_only" | "write_preview" | null;
  turnId?: string | null;
  prompt: string;
  taskContext?: import("@excel-bro/core/agentTaskContext").AgentTaskContext | null;
  workbook: WorkbookSnapshot;
  lastResult?: ResultContext | null;
  images?: ImageAttachment[];
  dataResults?: DataToolResult[];
  /**
   * 本地已登记结果的不透明引用。SDK 后端据此让模型引用本地结果，
   * 而值始终留在本机易失登记器；`dataResults` 的数据行守卫仍然有效。
   */
  localResultRefs?: string[];
  planningReadPolicy?: "ifNeeded" | "none";
  modelId?: string | null;
  /**
   * 规划前能力收窄（方案 §4.3）。**这些字段不是授权凭据**：
   * 服务端按 `hostCapabilities` 算该宿主/版本的动作上界，
   * 只把 `selfReportedActions` 当收窄条件用（交集 `A ∩ B ⊆ A`）。
   * 客户端多报不会放开任何动作，少报只会减少自己的功能。
   *
   * 未声明 `hostId` 的调用方保持原行为，动作边界由各通道后续校验负责。
   */
  hostId?: "excel-office" | "wps-et" | "folder" | null;
  hostVersion?: string | null;
  selfReportedActions?: string[];
}

export function assertIntentCheckResponse(
  value: unknown
): asserts value is IntentCheckResponse {
  if (!value || typeof value !== "object") {
    throw new Error("服务返回的意图判断不是有效对象");
  }
  const response = value as Partial<IntentCheckResponse>;
  if (response.kind === "conversation") {
    if (typeof response.message !== "string" || !response.message.trim()) {
      throw new Error("服务返回的对话回复无效");
    }
    return;
  }
  if (response.kind === "proceed") {
    if (
      !("confirmedPrompt" in response) ||
      typeof response.confirmedPrompt !== "string"
    ) {
      throw new Error("服务返回的明确需求无效");
    }
    return;
  }
  if (response.kind === "tool_request") {
    const request = response.request;
    if (
      !request ||
      request.tool !== "query_table" ||
      typeof request.id !== "string" ||
      !request.arguments ||
      !["rows", "aggregate", "profile"].includes(request.arguments.mode)
    ) {
      throw new Error("服务返回的数据工具请求无效");
    }
    return;
  }
  if (response.kind !== "clarification") {
    throw new Error("服务返回了未知的意图判断类型");
  }
  const clarification = response.clarification;
  if (
    !clarification ||
    typeof clarification.question !== "string" ||
    !Array.isArray(clarification.options) ||
    clarification.options.length > capabilities.clarification.maxOptions ||
    (clarification.options.length < 2 && !clarification.fields?.length) ||
    (clarification.fields !== undefined && (
      !Array.isArray(clarification.fields) || clarification.fields.length > 4 ||
      clarification.fields.some((field) => !field || typeof field.id !== "string" || !field.id.trim() ||
        typeof field.label !== "string" || !field.label.trim() ||
        (field.placeholder !== undefined && typeof field.placeholder !== "string")) ||
      new Set(clarification.fields.map((field) => field.id)).size !== clarification.fields.length
    )) ||
    clarification.options.some(
      (option) =>
        !option ||
        typeof option.id !== "string" ||
        typeof option.label !== "string" ||
        typeof option.description !== "string" ||
        typeof option.resolution !== "string" ||
        (option.resolutionMode !== undefined &&
          option.resolutionMode !== "append" &&
          option.resolutionMode !== "replace")
    )
  ) {
    throw new Error("服务返回的确认问题无效");
  }
}

export interface FolderWorksheetInfo {
  name: string;
  rowCount: number;
  columnCount: number;
}

export interface FolderFileInfo {
  id: string;
  name: string;
  relativePath: string;
  worksheets: FolderWorksheetInfo[];
  error?: string | null;
}

export interface FolderCatalog {
  sessionId: string;
  folderName: string;
  folderPath: string;
  files: FolderFileInfo[];
  totalFiles: number;
  truncated: boolean;
  expiresAt: string;
}

export interface FolderSelection {
  fileId: string;
  sheets: string[];
}

/** Local selection preflight, not data-read consent or a write approval. */
export interface FolderAdmissionRequest {
  sessionId: string;
  selections: FolderSelection[];
}
export interface FolderAdmissionResponse {
  items: import("@excel-bro/core/scopeAdmission").ScopeAdmissionItem[];
  fingerprint: string | null;
}

export interface FolderExecuteResult {
  executionId?: string | null;
  filesModified: string[];
  backups: string[];
  backupDirectory?: string | null;
  backupWarnings?: string[];
  actionResults: ActionExecutionResult[];
  verification: VerificationReport;
  undo?: FolderUndoReference | null;
  executionMs: number;
  verificationMs: number;
}

export interface FolderUndoReference {
  undoId: string;
  sessionId: string;
}

export interface FolderUndoResult {
  filesRestored: string[];
  executionsUndone: number;
  remainingUndoCount: number;
}

export interface FolderExecutionStatusResult {
  operationId: string;
  status: "running" | "completed" | "failed" | "needs_reconciliation";
  lastSequence: number;
  result?: FolderExecuteResult | null;
  error?: {
    status?: number;
    code?: string;
    message?: string;
    retryable?: boolean;
  } | null;
  revision?: number | null;
  canResume?: boolean;
  checkpointStage?: "publish_prepared" | "publishing" | "published" | null;
  resumeReason?: string | null;
  resumeFilesApplied?: number | null;
  resumeFilesPending?: number | null;
}

export interface ManagedFolderRecoveryDirectory {
  directoryId: string;
  folderName: string;
  folderPath: string;
  status: "active" | "unavailable" | "unsafe";
  recoveryCount: number;
  attentionCount: number;
  operationCount: number;
  lastScannedAt?: string | null;
  statusReason?: string | null;
}

export interface ManagedFolderRecoveryDirectoriesResult {
  directories: ManagedFolderRecoveryDirectory[];
}

export interface FolderManagedRecoveryResult {
  operationId: string;
  undoId: string;
  status: "completed";
  filesRestored: string[];
  alreadyRestored: boolean;
}

export interface FolderBackupFileInfo {
  originalRelativePath: string;
  backupRelativePath?: string | null;
  recoveryAction: "restore" | "remove";
  sizeBytes: number;
}

export interface FolderBackupEntry {
  revision?: number;
  undoId: string;
  undoSessionId?: string | null;
  planId: string;
  planTitle: string;
  createdAt: string;
  status:
    | "creating"
    | "preparing"
    | "prepared"
    | "applying"
    | "ready"
    | "recovering"
    | "restored"
    | "rolled_back"
    | "needs_attention";
  statusReason?: string | null;
  undoAvailability: "eligible" | "blocked" | "view_only" | "busy";
  undoUnavailableReason?: string | null;
  deleteInvalidatesUndoIds: string[];
  pinned: boolean;
  protectedReasons: string[];
  canUndo: boolean;
  fileCount: number;
  totalBytes: number;
  expiresAt?: string | null;
  files: FolderBackupFileInfo[];
}

export interface FolderLegacyBackupInfo {
  relativePath: string;
  sizeBytes: number;
  modifiedAt: string;
}

export interface FolderBackupStorageSummary {
  automaticBytes: number;
  pinnedBytes: number;
  abnormalBytes: number;
  exportedBytes: number;
  legacyBytes: number;
  pendingCleanupBytes: number;
  totalBytes: number;
  softLimitBytes: number;
  complete: boolean;
  calculatedAt: string;
}

export interface FolderBackupsResult {
  folderPath: string;
  backupDirectory: string;
  openDirectoryToken?: string | null;
  backupManagementToken?: string | null;
  backups: FolderBackupEntry[];
  legacyCandidates: FolderLegacyBackupInfo[];
  storageSummary?: FolderBackupStorageSummary | null;
  warnings: string[];
}

export interface FolderBackupOpenResult {
  opened: boolean;
  directory: string;
}

export interface FolderBackupExportResult {
  operationId: string;
  exportDirectory: string;
  files: string[];
  complete?: boolean;
  failures?: string[];
}

export interface FolderBackupDeleteResult {
  deletedUndoId: string;
  discardedUndo: boolean;
  filesDeleted: number;
  bytesFreed: number;
  remainingUndoCount: number;
  invalidatedUndoIds: string[];
  remainingUndoIds: string[];
  cleanupPending: boolean;
  cleanupWarnings: string[];
}

export interface FolderBackupPinResult {
  undoId: string;
  pinned: boolean;
}

export interface FolderBackupCleanupResult {
  deletedExecutions: number;
  deletedFiles: number;
  bytesFreed: number;
  skippedActive: number;
  protectedExecutions?: number;
  retainedBytes?: number;
  overLimitBytes?: number;
  failures: string[];
  legacyCandidateCount: number;
}

export interface FolderLegacyCollectResult {
  collectedFiles: string[];
  failures: string[];
}

export interface ActionExecutionResult {
  index: number;
  type: ExcelAction["type"];
  sheet: string;
  status: "succeeded" | "failed" | "not_run";
  message?: string | null;
  changed?: boolean | null;
}

export interface VerificationCheck {
  criterion: VerificationCriterion;
  passed: boolean;
  message: string;
  actual?: CellValue[][] | null;
}

export interface UnverifiedAction {
  index: number;
  type: ExcelAction["type"];
  sheet: string;
  message: string;
}

export interface VerificationReport {
  status: "verified" | "executed_unverified" | "failed";
  passed: boolean;
  checks: VerificationCheck[];
  unverifiedActions: UnverifiedAction[];
}

export interface PlanExecutionResult {
  actionResults: ActionExecutionResult[];
  verification: VerificationReport;
  undoSnapshot?: ExecutionUndoSnapshot | null;
  executionMs: number;
  verificationMs: number;
}

/**
 * 一次撤销动作**实际要恢复**的字段集合，由动作的**实际影响**派生（纯函数 `undoRestoreScope`）。
 *
 * 存在的理由：旧实现无差别地"恢复值 + 恢复填充 + 恢复字体 + 恢复全部格式 + 恢复行高列宽"，
 * 于是一个只改行高的动作撤销时会顺带把单元格格式写一遍。用户对不被本动作改过的属性的后续编辑
 * 就被静默覆盖了。按范围恢复后：
 *  - 只有 `formulas` 为真的动作才写回值/公式；
 *  - 只有 `cellFormat` 为真的动作才碰逐格格式；
 *  - 只有 `rowHeight`/`columnWidth` 为真的动作才碰行高/列宽。
 *
 * **`cellFormat` ⇒ 隐含 `numberFormat`**：兼容路径（宿主无 ExcelApi 1.19）恢复逐格格式用的
 * 是 `clear(Formats)`，而"清除格式"**连数字格式一起清**。所以只要动作会走那条路径，
 * 数字格式就必须一并捕获与写回，否则恢复完数字格式反而丢了。这是机制约束，不是动作影响，
 * 在 `undoRestoreScope` 里显式注明。
 */
export interface UndoRestoreScope {
  /** 是否写回值/公式（`writeValues`/`writeFormulas`/`writeTable`/`clearRange(contents)`）。 */
  formulas: boolean;
  /** 是否写回数字格式（`setNumberFormat`，或被 `cellFormat` 的 `clear(Formats)` 连带）。 */
  numberFormat: boolean;
  /** 是否恢复逐格格式（`setFill`/`setFont`/`setBorders`/`setAlignment`）。 */
  cellFormat: boolean;
  /** 是否恢复行高（`resizeRange` 且该动作确实给了 `rowHeight`）。 */
  rowHeight: boolean;
  /** 是否恢复列宽（`resizeRange` 且该动作确实给了 `columnWidth`）。 */
  columnWidth: boolean;
}

/**
 * 单个区域的撤销快照。
 *
 * 格式分两层：value/formula 与常用格式总是记录；`format` 里的完整格式（边框、
 * 对齐、合并）只在区域规模不超过 undo.formatCaptureMaxCells 时捕获 —— 整表操作
 * 时逐格读边框会让执行明显变慢，超限则降级为只恢复常用格式。
 */
export interface UndoRangeSnapshot {
  actionIndex: number;
  sheet: string;
  range: string;
  formulas: CellValue[][];
  numberFormat: string[][];
  fillColor: string;
  fontBold: boolean;
  fontColor: string;
  /**
   * 本动作**实际恢复范围**。缺失 = 旧恢复点，按"全范围恢复"（旧行为）处理，
   * 以免升级本身把已落盘的恢复点判成范围不符。
   */
  scope?: UndoRestoreScope;
  /**
   * 本动作实际修改的格式属性名（`excelFormatActionFields`）。缺失 = 未知（旧记录）。
   * 冲突检测用它定位"到底哪个属性不是我们最后写的那个值"。
   */
  formatFields?: string[];
  /** 本次执行完成后该范围的确定性指纹；撤销前用于阻止覆盖后续人工修改。 */
  postFingerprint?: string;
  /**
   * 本次执行完成后该范围的**扩展格式摘要**（第二代判据，覆盖对齐/自动换行/行高列宽/
   * 字体名・字号・斜体・下划线/六条边框）。
   *
   * `postFingerprint` 是**第一代**判据，只看 值/公式/数字格式/填充色/加粗/字色；
   * 执行完一个格式动作之后用户手动改了"对齐"或"边框"，第一代一点都察觉不到。
   * 所以新记录同时写这一项，撤销前优先比它。
   *
   * 字段保持可选：已经落盘的旧恢复点是按第一代载荷算的，改载荷会让它们一律对不上、
   * 撤销被误判成"执行后又被改过"而全部拒绝。缺失 ⇒ 退回第一代判据（安全性略弱，
   * 但不会把旧记录变成不可撤销）。
   */
  extendedFormatDigest?: string;
  /**
   * **按恢复范围、逐格/逐行**计算的执行后状态摘要（第三代判据）。
   *
   * 载荷恰好是 `scope` 为真的那几项 —— 只比"这次撤销真正会写回去的字段"。这样：
   *  - 用户在**不会被写回**的属性上的后续编辑 ⇒ 摘要不变 ⇒ 撤销照常执行且**保留**该编辑；
   *  - 用户在**会被写回**的属性上的后续编辑 ⇒ 摘要变化 ⇒ 在**任何写入之前**报冲突。
   *
   * 与 `extendedFormatDigest` 的区别是**覆盖范围**而非粒度：后者是固定的全属性集合，
   * 前者跟着动作走。只改行高的动作不会因为用户改了字色而拒绝撤销。
   */
  scopeDigest?: string;
  /**
   * 第三代判据的按字段拆分（`formulas`/`numberFormat`/`cellFormat`/`rowHeight`/`columnWidth`），
   * 只用于在冲突时**指名到底哪一项对不上**，不参与是否冲突的判定。
   */
  scopeFieldDigests?: Record<string, string>;
  /**
   * 第三代判据里 `cellFormat` 那一项的**按属性再拆分**（键为 `ExcelFormatField`，
   * 如 `fontBold`/`fill`/`borders`）。
   *
   * `scopeFieldDigests` 只能说到"单元格格式对不上"，用户在一大片格式里仍然不知道该看哪个；
   * 这份拆分让冲突文案能指名**具体属性**（如"单元格格式（加粗）"）。同样只用于呈报，
   * 不参与是否冲突的判定（判定仍只比 `scopeDigest`）。
   */
  cellFormatFieldDigests?: Record<string, string>;
  /** 完整格式快照；区域过大而降级捕获时为 null。 */
  format?: UndoRangeFormat | null;
}

/**
 * 超出常用格式之外、按**属性实际粒度**恢复的格式快照。
 *
 * 粒度分三类（缺一就会把"混合格式"恢复成单一值，属于不可接受的静默降级）：
 *  - 逐格：字体（名/号/粗/斜/字色/下划线）、对齐（水平/垂直/换行/缩进）、填充色、六条边框；
 *  - 逐行：行高（行级属性，落在 B3:D5 上真实改的是第 3–5 整行）；
 *  - 逐列：列宽（列级属性）。
 *
 * 旧字段（`horizontalAlignment`/`rowHeight`/`borders`…）是**第一版区域级单值**，
 * 保留只为兼容已落盘的旧恢复点；新捕获一律写 `cells`/`rowHeights`/`columnWidths`，
 * 恢复时优先走按粒度字段，缺失才退回旧单值。
 */
export interface UndoRangeFormat {
  horizontalAlignment: string;
  verticalAlignment: string;
  wrapText: boolean | null;
  rowHeight: number;
  columnWidth: number;
  indentLevel: number;
  fontName: string;
  fontSize: number;
  fontItalic: boolean;
  fontUnderline: string;
  borders: UndoBorderState[];
  /** 逐格格式矩阵（与目标区域同形）。新捕获时必填；区域过大降级捕获时为 null。 */
  cells?: UndoCellFormat[][] | null;
  /**
   * 逐行行高（按目标区域的行序）。元素三态：
   *  - `number`：该行行高**直接设置**为这个值（撤销写回该值）；
   *  - `"standard"`：该行未直接设置行高（`useStandardHeight === true`），撤销走标准行高恢复；
   *  - `null`：该行行高状态**无法确定**（宿主异常）。含 `null` 且动作会改行高时应执行前拒绝。
   * 缺失数组表示未逐行采集，退回区域级 `rowHeight`（仅旧恢复点兼容）。
   */
  rowHeights?: Array<number | "standard" | null>;
  /**
   * 逐列列宽（按目标区域的列序）。元素三态语义同 `rowHeights`：
   *  - `number`：直接设置；`"standard"`：未直接设置（`useStandardWidth === true`）；`null`：无法确定。
   */
  columnWidths?: Array<number | "standard" | null>;
  /**
   * **兼容路径**（宿主 < ExcelApi 1.19，批量许可永久版最高到 1.17）的逐格
   * **直接格式**矩阵，与 `cells` 互斥：`cells` 是"显示值"，本字段是"直接设置集合"。
   *
   * 恢复语义（`restoreRangeFormat` 走这条分支时）：先 `range.clear(Formats)` 清掉
   * 目标区域的直接格式，再 `setCellProperties` 写回这里的直接格式集合。于是
   * **直接设置过的属性逐格精确还原**，而**未直接设置的属性回到继承样式/默认**——
   * 恢复的是"格式的来源语义"，不是"把显示效果写成显式值"。
   *
   * 为空（null/缺失）表示未走兼容路径。
   */
  directCells?: UndoCellDirectFormat[][] | null;
  /**
   * 兼容路径下，本次动作实际修改的属性里**未直接设置**（继承样式/条件格式/默认）
   * 的那部分属性名并集（去重排序）。只用于如实呈报来源，不参与写入。
   */
  inheritedFields?: string[];
  /** 标记该快照由兼容路径（无 ExcelApi 1.19）采集。 */
  compatPath?: boolean;
}

/**
 * 兼容路径下单个单元格的**直接格式**快照（来自 `Range.getCellProperties`，ExcelApi 1.9）。
 *
 * 语义与 `UndoCellFormat`（显示值）刻意不同：
 *  - `UndoCellFormat` 每个字段都是一个**确定的显示值**；
 *  - 本类型只装**直接设置在单元格上**的属性。未直接设置的属性**不出现**在 `direct` 里 ——
 *    这是"继承样式/条件格式/默认"这一**确定状态**，不是"未知"，因此**不允许**
 *    填一个默认值冒充"采集到了"（那会把继承语义写成显式值）。
 *
 * 真正的"未知"只发生在矩阵形状不符/元素非对象时，那种情况在写入前直接拒绝，
 * 不会生成这个结构。
 */
export interface UndoCellDirectFormat {
  /** 该格直接设置的属性子集；未直接设置的属性不出现。 */
  direct: DirectCellFormatValues;
  /** 本次动作实际修改的属性里，该格**未直接设置**的那些（供如实呈报来源）。 */
  inherited: string[];
}

/**
 * 直接格式属性子集。字段缺席 = 该属性在该格上未直接设置。
 *
 * **字段清单不是"我们想到哪算哪"**，而是与 `Excel.CellPropertiesLoadOptions` /
 * `CellPropertiesFormatLoadOptions` / `CellPropertiesFontLoadOptions` /
 * `CellPropertiesFillLoadOptions` / `CellPropertiesBorderLoadOptions`
 * 逐项对齐 —— 因为兼容路径的恢复要先 `clear(Formats)`，**凡是能读的都要读**，
 * 否则"清空再写回"会把读不到的属性静默丢掉。
 *
 * **已知无法完整保存的场景（如实限制，不冒充覆盖）**：单元格**内部富文本**
 * （`CellPropertiesLoadOptions.textRuns`，ExcelApi 1.18）在 1.17 宿主上**读不到也写不回**，
 * 因此本机既无法判断某格是否含富文本，也无法在富文本被改变时还原。
 *
 * 关于 `clear(Formats)`：官方 Excel 语义说它**不清除**单元格内字符级格式
 * （"Clear Formats 不改单元格内单个字符的格式"）。这只是**那一步操作的语义依据**，
 * **不构成"含富文本的区域可以安全执行并撤销"的整体结论** —— 完整的写入与撤销链路
 * **未经真实宿主验证**，且一旦宿主实际行为与官方语义不符，本机既发现不了也修复不了。
 * 不在这里补默认值、也不声称"已覆盖全部格式"。
 */
export interface DirectCellFormatValues {
  /**
   * 单元格**命名样式**（`CellPropertiesLoadOptions.style`，ExcelApi 1.9）。
   *
   * 它不是 `format` 的子字段，而是 `SettableCellProperties.style`（与 `format` 同级）。
   * `clear(Formats)` 会把命名样式退回 `Normal`，所以必须捕获并在恢复时写回 ——
   * 这是"非 Normal 单元格样式 + 直接格式覆盖"能被正确还原的前提。
   * 缺席 = 未采到（旧快照），恢复时不写。
   */
  style?: string;
  /** 直接设置的填充样式（如 `"Solid"`）；缺席 = 未直接设置填充。 */
  fillPattern?: string;
  fillColor?: string;
  /** 图案色（`CellPropertiesFill.patternColor`），如斜线图案的线条色。 */
  fillPatternColor?: string;
  fillPatternTintAndShade?: number;
  fillTintAndShade?: number;
  fontName?: string;
  fontSize?: number;
  fontBold?: boolean;
  fontItalic?: boolean;
  fontColor?: string;
  fontUnderline?: string;
  /** 字体删除线（`CellPropertiesFont.strikethrough`）。 */
  fontStrikethrough?: boolean;
  /** 下标 / 上标（`CellPropertiesFont`）。两者互斥，但都如实捕获。 */
  fontSubscript?: boolean;
  fontSuperscript?: boolean;
  fontTintAndShade?: number;
  horizontalAlignment?: string;
  verticalAlignment?: string;
  wrapText?: boolean;
  indentLevel?: number;
  shrinkToFit?: boolean;
  textOrientation?: number;
  readingOrder?: string;
  /** 自动缩进（`CellPropertiesFormat.autoIndent`）。 */
  autoIndent?: boolean;
  /** 单元格保护（`locked` / `formulaHidden`）。`clear(Formats)` 会重置它，必须一并捕获。 */
  locked?: boolean;
  formulaHidden?: boolean;
  /**
   * 逐边直接格式。**八条边**（`Excel.CellBorderCollection`：top/bottom/left/right/
   * horizontal/vertical/diagonalDown/diagonalUp），不是旧实现里的四条 —— 斜线边框
   * 正是靠 `diagonalDown`/`diagonalUp` 表达，漏采就会被 `clear(Formats)` 清掉。
   */
  borders?: Partial<Record<DirectBorderSide, CellBorderState>>;
}

/**
 * 单个单元格的格式快照。所有字段都在**捕获那一刻**被逐格读取，混合值不会被折叠成区域级单值。
 *
 * 语义（撤销恢复的是**执行前那一刻的实际显示值**，不是"直接设置状态"）：
 * 逐格属性用 `getDisplayedCellProperties`（ExcelApi 1.19）捕获，它返回单元格**实际显示**
 * 的格式 —— 无论该格式是直接设置在单元格上，还是来自表格样式 / 命名样式 / 条件格式的
 * 间接来源。因此这里的每个字段都是一个**确定的显示值**：
 *
 *  - `fillPattern: "None"` = 明确无填充（显示上没有填充）；`"Solid"` 等 = 显示上有填充，
 *    无论其来源是直接设置还是样式/条件格式（此时 `fillColor` 是显示色）。
 *  - `fontBold: false` = 显示为不加粗（可能是直接设 false，也可能继承样式默认不加粗），
 *    撤销写回 `bold: false` 以恢复"执行前显示为不加粗"。
 *
 * 之所以存"显示值"而非"直接设置值"：`getCellProperties` 对未直接设置的属性返回 `undefined`，
 * 它**不告诉你**该属性当前显示是什么（可能正被样式/条件格式影响）。用 `undefined`（存 null）
 * 并在撤销时跳过，会把"执行动作临时写上的显式值"留在单元格上，无法恢复执行前的显示。
 * 唯一可靠、可恢复的依据是**显示值**。
 *
 * `fillColor` 是唯一的 `null` 合法值：当 `fillPattern === "None"` 时显示上没有填充色，存
 * `null` 表示"无填充色"（这是确定状态，不是"未采集"）。
 */
export interface UndoCellFormat {
  /** 显示填充样式：`"None"` = 明确无填充；`"Solid"`/`"LinearGradient"` 等 = 显示有填充。 */
  fillPattern: string;
  /** 显示填充色。`null` = 无填充色（`fillPattern === "None"`）。非 `None` 时必有值。 */
  fillColor: string | null;
  fontName: string;
  fontSize: number;
  fontBold: boolean;
  fontItalic: boolean;
  fontColor: string;
  fontUnderline: string;
  horizontalAlignment: string;
  verticalAlignment: string;
  wrapText: boolean;
  indentLevel: number;
  /** 逐格边框：单元格四条边的显示状态。每条边都是确定的显示值。 */
  borders: Record<CellBorderSide, CellBorderState>;
}

/** 逐格边框的四条边名（与 `Excel.CellBorderCollection` 的 top/bottom/left/right 对应）。 */
export type CellBorderSide = "top" | "bottom" | "left" | "right";

/**
 * 直接格式路径的逐边边名：`Excel.CellBorderCollection` 的**八条**边。
 *
 * 与显示值路径的 `CellBorderSide`（四条）刻意分开：显示值路径只读四边，
 * 直接格式路径要连同内部边（`horizontal`/`vertical`）与斜线（`diagonalDown`/`diagonalUp`）
 * 一起捕获 —— 漏任何一条，`clear(Formats)` 都会把它清掉且还原不回来。
 */
export type DirectBorderSide =
  | "top"
  | "bottom"
  | "left"
  | "right"
  | "horizontal"
  | "vertical"
  | "diagonalDown"
  | "diagonalUp";

/** 逐格边框单边的显示状态。字段都是确定的显示值（来自 `getDisplayedCellProperties`）。 */
export interface CellBorderState {
  style: string;
  color: string;
  weight: string;
  /** 底纹深浅（`CellBorder.tintAndShade`）；仅直接格式路径可能带上。 */
  tintAndShade?: number;
}

export interface UndoBorderState {
  sideIndex: Excel.BorderIndex;
  style: string;
  color: string;
  weight: string;
}

// 被删除工作表的撤销快照：删前把整张表内容拍下来，撤销时新建并原样恢复。
export interface UndoWorksheetSnapshot {
  sheet: string;
  position: number;
  usedRange: string;
  formulas: CellValue[][];
  numberFormat: string[][];
}

export interface UndoCreatedWorksheetState {
  sheet: string;
  /** 新建表执行完成后的内容指纹；缺失时不能安全自动删除该表。 */
  postFingerprint?: string;
}

export interface ExecutionUndoSnapshot {
  planId: string;
  /** 同一计划重复执行时仍保持唯一的本次执行标识。 */
  executionId?: string;
  /** 工作簿本地标识；可取得时用于阻止把快照应用到另一个工作簿。 */
  workbookIdentity?: string;
  capturedAt: string;
  /** partial 表示失败前只捕获并执行了部分动作。 */
  coverage?: "full" | "partial";
  completedActionCount?: number;
  uncertainActionIndexes?: number[];
  /** complete 才允许直接撤销；partial 表示写后状态未能全部核对。 */
  postStateCoverage?: "complete" | "partial";
  ranges: UndoRangeSnapshot[];
  // 本次执行新建的工作表名，撤销时移除。
  createdWorksheets: string[];
  createdWorksheetStates?: UndoCreatedWorksheetState[];
  // 本次执行删除的工作表，撤销时恢复。
  deletedWorksheets: UndoWorksheetSnapshot[];
  // 删除但因过大未记录撤销快照的工作表名（无法恢复）。
  skippedWorksheets: string[];
}

/**
 * 服务端持久化的 Office.js 写入现场。只有受控区域动作且工作簿具有稳定 URL
 * 身份时才会创建；ownerToken 只在登记/领取写入权后返回。
 */
/**
 * 「补记完成」写进 `result` 的最小事实。
 *
 * 服务端 `complete` 只要求一个 dict（不重算计划结果），所以补记时**不伪造**
 * 一份 `PlanExecutionResult`：这里只记"这次没有执行任何动作、只是把已完成的账补上"。
 * 补记路径不调用 `executePlan`，一个单元格都不写。
 */
export interface OfficeExecutionBackfillResult {
  backfilled: true;
  actionCount: number;
}

/**
 * 区分"真的执行过"与"只补记"。
 *
 * 补记结果没有 `undoSnapshot` / `verification`：拿它去读撤销快照会把"没执行"说成
 * "执行了但没有可撤销数据"。判据用 `backfilled` 标记，不按字段有没有来猜。
 */
export function isPlanExecutionResult(
  value: PlanExecutionResult | OfficeExecutionBackfillResult | null | undefined
): value is PlanExecutionResult {
  return Boolean(value) && !("backfilled" in (value as OfficeExecutionBackfillResult));
}

/**
 * 这次执行被用户终态处置的**首次时刻**（服务端只读投影，`null` = 这条留痕不存在）。
 *
 * 两个取值不互冒充：`undoneAt` = 确认之后撤销本次执行，`abandonedAt` = 审阅阶段
 * 放弃变更。重开窗格靠它判断"撤销入口是不是已经被用掉了"。
 */
export interface OfficeExecutionDispositionState {
  undoneAt?: number | null;
  abandonedAt?: number | null;
}

export interface OfficeExecutionState {
  operationId: string;
  status: "running" | "needs_reconciliation" | "completed" | "failed";
  revision: number;
  workbookIdentity: string;
  plan: AnalysisPlan;
  authorizedSheetNames: string[];
  allowedExternalFiles: string[];
  nextActionIndex: number;
  inFlightActionIndex?: number | null;
  beforeSnapshot?: UndoRangeSnapshot | null;
  checkpoint?: ExecutionUndoSnapshot | null;
  result?: PlanExecutionResult | OfficeExecutionBackfillResult | null;
  error?: { code?: string; message?: string } | null;
  canResume: boolean;
  requiresReconciliation: boolean;
  ownerToken?: string | null;
  disposition?: OfficeExecutionDispositionState | null;
}

export function assertOfficeExecutionState(
  value: unknown
): asserts value is OfficeExecutionState {
  if (!value || typeof value !== "object") {
    throw new Error("服务返回的 Office 恢复状态不是有效对象");
  }
  const state = value as Partial<OfficeExecutionState>;
  if (
    typeof state.operationId !== "string" ||
    !["running", "needs_reconciliation", "completed", "failed"].includes(
      String(state.status)
    ) ||
    typeof state.revision !== "number" ||
    typeof state.workbookIdentity !== "string" ||
    !state.plan ||
    !Array.isArray(state.plan.actions) ||
    !Array.isArray(state.authorizedSheetNames) ||
    !Array.isArray(state.allowedExternalFiles) ||
    typeof state.nextActionIndex !== "number" ||
    typeof state.canResume !== "boolean" ||
    typeof state.requiresReconciliation !== "boolean"
  ) {
    throw new Error("服务返回的 Office 恢复状态缺少必要字段");
  }
  if (
    state.inFlightActionIndex !== undefined &&
    state.inFlightActionIndex !== null &&
    typeof state.inFlightActionIndex !== "number"
  ) {
    throw new Error("服务返回的 Office 在途动作位置无效");
  }
}

/**
 * WPS 恢复描述里的**单个动作**。与 `server/app/models.py: WpsRecoveryActionDescriptor` 同形。
 *
 * 只带恢复判定必需的元数据：动作类型、目标工作表与区域、目标形状、目标指纹，
 * 以及 `clearRange` 的 `applyTo` / `writeFormulas` 的 `formulaR1C1`（服务端据此
 * 拒绝不支持的变体）。**不含 `values` / `formulas` 矩阵** —— 目标值只停在本机内存。
 *
 * 格式动作（`setFont` / `setAlignment` / `resizeRange`）多一个 `formatFields`：
 * 本次声明要改的**属性名**（`bold`/`size`/`horizontalAlignment`/`rowHeight` 等）。
 * 它让服务端能做属性级校验，也让"这两个指纹是按格式状态算的"可判定 —— 值动作的
 * 指纹是值矩阵指纹，格式动作的是格式证据指纹（格式动作不改值，值矩阵指纹在写前
 * 写后完全一样，没有区分力）。`formatFields` 里**不含单元格内容**，它只是属性名。
 */
export interface WpsRecoveryActionDescriptor {
  actionIndex: number;
  type: string;
  sheet: string;
  range: string;
  rowCount: number;
  columnCount: number;
  cellCount: number;
  /** 写前逻辑指纹（值动作：写前值；格式动作：写前格式证据）。校验不是加密。 */
  previewFingerprint: string;
  applyTo?: string | null;
  expectedFingerprint?: string | null;
  formulaR1C1?: boolean;
  formatFields?: string[];
}

/**
 * WPS 写入的**恢复描述**。与 `server/app/models.py: WpsRecoveryDescriptor` 同形。
 *
 * 这是持久账本里关于"这次要写什么"的**全部**内容。它替代了早期直接上传整个
 * `AnalysisPlan` 的做法 —— 那个做法会把 `writeValues` 的目标值矩阵带出站并落盘。
 * 恢复功能不得为"能续跑"而放宽这条约束：宁可续跑能力受限，也不让单元格内容出站。
 */
export interface WpsResumeSource {
  sheetName: string;
  range: string;
  rowCount: number;
  columnCount: number;
  previewFingerprint: string;
  valueFingerprint: string;
}
export interface WpsResumeContext {
  /** 1: nonempty source evidence; 2: trusted literal-only plan with no source dependency. */
  version: 1 | 2;
  hostVersion: string;
  sources: WpsResumeSource[];
}
export interface WpsRecoveryDescriptor {
  resumeContext?: WpsResumeContext | null;
  actions: WpsRecoveryActionDescriptor[];
}

/**
 * WPS 恢复账本认可的**动作类型**（唯一来源）。
 *
 * 与 `config/capabilities.json` 的 `wpsWrite.allowedActions` 同口径（`clearRange`
 * 在那里按 `clearRange:contents` 细分），也与执行器 `WPS_SUPPORTED_ACTION_KEYS`
 * 一致。曾经这里硬编码成"只认三个值动作"，于是**格式动作在 fetch 之前就被拒**，
 * 表现成真实 WPS 上"确认写回"时提示「恢复描述的动作定位或形状无效」。
 */
export const WPS_RECOVERABLE_ACTION_TYPES = [
  "writeValues", "sortRange",
  "writeFormulas",
  "clearRange",
  "setFont",
  "setAlignment",
  "setFill",
  "setNumberFormat",
  "resizeRange",
  "setBorders"
] as const;

/**
 * 每个格式动作**可以声明**的属性名（= 执行器 `declaredFormatFields` 的输出集合，
 * 也与 `wpsFormatSnapshot.ALL_WPS_FORMAT_FIELDS` 的字段名同源）。
 *
 * 尚未实现可恢复执行的属性（`underline` / `columnWidth`）**不在任何一行里**：
 * 它们在投影阶段就产不出来，这里是第二道闸 —— 服务端还有第三道。
 */
export const WPS_FORMAT_ACTION_FIELDS: Record<string, readonly string[]> = {
  setFont: ["bold", "italic", "name", "size", "color"],
  setAlignment: ["horizontalAlignment", "verticalAlignment", "wrapText"],
  setFill: ["fillColor", "fillPattern"],
  setNumberFormat: ["numberFormat"],
  setBorders: ["borderTop", "borderBottom", "borderLeft", "borderRight"],
  resizeRange: ["rowHeight"]
};

export function assertWpsEvidence(value: unknown): void {
  if (!value || typeof value !== "object" ||
      Object.keys(value).some((key) => !["sheetName", "range", "type", "beforeFingerprint", "afterFingerprint"].includes(key)))
    throw new Error("WPS 动作证据含未知字段或单元格矩阵");
  const evidence = value as Record<string, unknown>;
  if (["sheetName", "range", "type", "beforeFingerprint"].some((key) => typeof evidence[key] !== "string") ||
      (evidence.afterFingerprint != null && typeof evidence.afterFingerprint !== "string"))
    throw new Error("WPS 动作证据必须只含定位和指纹文本");
  if (!(WPS_RECOVERABLE_ACTION_TYPES as readonly string[]).includes(String(evidence.type)))
    throw new Error("WPS 动作证据的动作类型不在可恢复范围内");
}

export function assertWpsExecutionResult(value: unknown): void {
  if (!value || typeof value !== "object") throw new Error("WPS 完成回执无效");
  const result = value as Record<string, unknown>;
  if (Object.keys(result).some((key) => !["actionCount", "cellCount", "backfilled"].includes(key)) ||
      !Number.isInteger(result.actionCount) || Number(result.actionCount) < 0 ||
      (result.cellCount != null && (!Number.isInteger(result.cellCount) || Number(result.cellCount) < 0)) ||
      (result.backfilled != null && typeof result.backfilled !== "boolean"))
    throw new Error("WPS 完成回执只允许进度计数，不允许单元格内容");
}

export function assertWpsRecoveryDescriptor(value: unknown): asserts value is WpsRecoveryDescriptor {
  if (!value || typeof value !== "object" || Object.keys(value).some((key) => !["actions", "resumeContext"].includes(key)) ||
      !Array.isArray((value as WpsRecoveryDescriptor).actions) || !(value as WpsRecoveryDescriptor).actions.length)
    throw new Error("WPS 恢复描述无效");
  const context = (value as WpsRecoveryDescriptor).resumeContext;
  if (context != null) {
    if (![1,2].includes(context.version) || typeof context.hostVersion !== "string" || !context.hostVersion ||
        context.hostVersion.length > 80 || Object.keys(context).some(key => !["version", "hostVersion", "sources"].includes(key)) ||
        !Array.isArray(context.sources) || context.sources.length > 500 ||
        (context.version === 1 ? !context.sources.length : context.sources.length !== 0) ||
        (context.version === 2 && (value as WpsRecoveryDescriptor).actions.some(action=>action.type!=="writeValues")))
      throw new Error("WPS 续跑来源证据无效");
    for (const source of context.sources) {
      if (!source || Object.keys(source).some(key => !["sheetName", "range", "rowCount", "columnCount", "previewFingerprint", "valueFingerprint"].includes(key)) ||
          typeof source.sheetName !== "string" || !source.sheetName || source.sheetName.length > 31 ||
          typeof source.range !== "string" || !/^\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?$/.test(source.range) ||
          !Number.isInteger(source.rowCount) || source.rowCount < 1 || source.rowCount > 1048576 ||
          !Number.isInteger(source.columnCount) || source.columnCount < 1 || source.columnCount > 16384 ||
          typeof source.previewFingerprint !== "string" || typeof source.valueFingerprint !== "string" ||
          !/^v2:\d+:[0-9a-f]{8}$/.test(source.previewFingerprint) || !/^v2:\d+:[0-9a-f]{8}$/.test(source.valueFingerprint))
        throw new Error("WPS 续跑来源含未知字段或无效指纹");
    }
  }
  const keys = ["actionIndex", "type", "sheet", "range", "rowCount", "columnCount", "cellCount",
    "previewFingerprint", "expectedFingerprint", "applyTo", "formulaR1C1", "formatFields"];
  for (const [index, action] of (value as WpsRecoveryDescriptor).actions.entries()) {
    if (!action || Object.keys(action).some((key) => !keys.includes(key)))
      throw new Error("WPS 恢复描述含未知字段或单元格矩阵，已拒绝接收");
    if (action.actionIndex !== index || !(WPS_RECOVERABLE_ACTION_TYPES as readonly string[]).includes(action.type) ||
        typeof action.sheet !== "string" || typeof action.range !== "string" ||
        !Number.isInteger(action.rowCount) || action.rowCount < 1 ||
        !Number.isInteger(action.columnCount) || action.columnCount < 1 ||
        action.cellCount !== action.rowCount * action.columnCount || typeof action.previewFingerprint !== "string" ||
        (action.expectedFingerprint != null && typeof action.expectedFingerprint !== "string") ||
        (action.applyTo != null && typeof action.applyTo !== "string") ||
        (action.formulaR1C1 != null && typeof action.formulaR1C1 !== "boolean"))
      throw new Error("WPS 恢复描述的动作定位或形状无效");
    assertWpsFormatFields(action.type, (action as { formatFields?: unknown }).formatFields);
  }
}

/**
 * 格式属性声明的形状与白名单校验。
 *
 * 值动作**不得**声明格式属性（否则"这个指纹是值指纹还是格式指纹"会变得不可判定）；
 * 格式动作**必须**至少声明一个，且只能取已验证的属性名。要求"至少一个"不是形式主义：
 * 空的格式动作本来就被协议层拒绝，这里再兜一次，避免登记出一份"目标指纹等于写前指纹"
 * 的记录 —— 那样的记录对恢复判定没有区分力。
 */
function assertWpsFormatFields(type: string, declared: unknown): void {
  const allowed = WPS_FORMAT_ACTION_FIELDS[type];
  if (!allowed) {
    if (declared != null && (!Array.isArray(declared) || declared.length))
      throw new Error("WPS 恢复描述的值动作不允许声明格式属性");
    return;
  }
  if (!Array.isArray(declared) || !declared.length ||
      declared.some((field) => typeof field !== "string" || !allowed.includes(field)) ||
      new Set(declared).size !== declared.length)
    throw new Error("WPS 恢复描述的格式属性声明无效");
}

/**
 * WPS 执行恢复状态，与 `server/app/models.py: WpsExecutionStateResponse` 同形。
 *
 * 与 Office 的区别：没有 `allowedExternalFiles`（WPS 不支持外部文件写入），
 * 多一个 `taskId`（加载项侧任务稳定身份，重载后继续时用来定位同一任务）；
 * 载荷是 `descriptor` 而不是 `plan` —— 服务端账本不存单元格值矩阵，
 * 所以下发的状态里也不会带回它们。
 * 身份用 `wps-path:` 前缀（已保存工作簿的小写磁盘完整路径），不与 Office 的
 * `url:` 前缀共享命名空间。
 */
export interface WpsExecutionState {
  operationId: string;
  status: "running" | "needs_reconciliation" | "completed" | "failed";
  revision: number;
  workbookIdentity: string;
  taskId: string;
  descriptor: WpsRecoveryDescriptor;
  authorizedSheetNames: string[];
  nextActionIndex: number;
  /** 动作总数。客户端据此算"剩余动作是否为空"，不能拿矩阵长度或内存进度代替。 */
  totalActions: number;
  inFlightActionIndex?: number | null;
  beforeSnapshot?: Record<string, unknown> | null;
  checkpoint?: Record<string, unknown> | null;
  result?: Record<string, unknown> | null;
  error?: { code?: string; message?: string } | null;
  canResume: boolean;
  /** True while another executor's durable lease is active; recovery cannot steal it. */
  ownerActive: boolean;
  requiresReconciliation: boolean;
  ownerToken?: string | null;
  /**
   * 处置留痕只读投影（2026-10-31）：与 `OfficeExecutionState.disposition` 同口径。
   * 撤销/放弃之后账本 `status` 与游标都不变，跨重启续跑判据必须靠它识别终态
   * （否则 `undoWpsExecution` 只在宿主缓存里把 `record.state` 改成 `restored`，
   * 重开后缓存没了，续跑会被重新开放）。
   */
  disposition?: OfficeExecutionDispositionState | null;
}

export function assertWpsExecutionState(
  value: unknown
): asserts value is WpsExecutionState {
  if (!value || typeof value !== "object") {
    throw new Error("服务返回的 WPS 恢复状态不是有效对象");
  }
  const state = value as Partial<WpsExecutionState>;
  if (
    typeof state.operationId !== "string" ||
    !["running", "needs_reconciliation", "completed", "failed"].includes(
      String(state.status)
    ) ||
    typeof state.revision !== "number" ||
    typeof state.workbookIdentity !== "string" ||
    !state.workbookIdentity.startsWith("wps-path:") ||
    typeof state.taskId !== "string" ||
    !state.descriptor ||
    !Array.isArray(state.descriptor.actions) ||
    !Array.isArray(state.authorizedSheetNames) ||
    typeof state.nextActionIndex !== "number" ||
    typeof state.totalActions !== "number" ||
    typeof state.canResume !== "boolean" ||
    typeof state.ownerActive !== "boolean" ||
    typeof state.requiresReconciliation !== "boolean"
  ) {
    throw new Error("服务返回的 WPS 恢复状态缺少必要字段");
  }
  if (
    state.inFlightActionIndex !== undefined &&
    state.inFlightActionIndex !== null &&
    typeof state.inFlightActionIndex !== "number"
  ) {
    throw new Error("服务返回的 WPS 在途动作位置无效");
  }
  /*
   * 出站闸的第二道：服务端若（因版本错配或改动失误）把单元格矩阵塞回状态里，
   * 客户端**拒绝接受**而不是照单收下再渲染。恢复功能只允许指纹级信息往返。
   */
  assertWpsRecoveryDescriptor(state.descriptor);
  if (state.totalActions !== state.descriptor.actions.length || !Number.isInteger(state.revision) || state.revision! < 1 ||
      !Number.isInteger(state.nextActionIndex) || state.nextActionIndex! < 0 || state.nextActionIndex! > state.totalActions!)
    throw new Error("WPS 恢复进度无效");
  if ("plan" in value) throw new Error("WPS 恢复状态含完整计划，已拒绝接收");
  if (state.beforeSnapshot) assertWpsEvidence(state.beforeSnapshot);
  if (state.checkpoint) assertWpsEvidence(state.checkpoint);
  if (state.result) assertWpsExecutionResult(state.result);
  for (const action of state.descriptor.actions) {
    if (!action || typeof action !== "object") {
      throw new Error("服务返回的 WPS 恢复描述含无效动作");
    }
    const candidate = action as unknown as Record<string, unknown>;
    if ("values" in candidate || "formulas" in candidate || "preview" in candidate) {
      throw new Error("服务返回的 WPS 恢复描述含单元格矩阵，已拒绝接收");
    }
  }
}

/**
 * **未完成任务恢复包**的类型权威在 `packages/core/src/executionRecoveryPackage.ts`。
 *
 * 协议层只转发，不另写一份结构：恢复包同时被宿主端（构建/水合）、共享层（白名单
 * 校验）与服务端（Pydantic 镜像）读取，多一份定义就多一处漂移。校验函数也用
 * core 那一份（`assertExecutionRecoveryPackage`），这里只补"传输信封"的形状。
 */
export type {
  ExecutionRecoveryActionType,
  ExecutionRecoveryPackage,
  ExecutionRecoveryPackageAction,
  ExecutionRecoverySource,
  ExecutionRecoverySourceEvidence
} from "@excel-bro/core/executionRecoveryPackage";

/**
 * 恢复包的本机读取结果（与 `server/app/models.py: ExecutionRecoveryPackageMaterial` 同形）。
 *
 * `packageJson` 是**原样文本**：客户端保存时写下哪一串字节，读回就必须是哪一串。
 * `expiresAt` 由服务端按保留期算出（默认 7 天），客户端不得自行延长。
 */
export interface ExecutionRecoveryPackageMaterial {
  packageJson: string;
  digest: string;
  size: number;
  createdAt: string;
  expiresAt: string;
  savedAt: string;
}

export function assertExecutionRecoveryMaterial(
  value: unknown
): asserts value is ExecutionRecoveryPackageMaterial {
  if (!value || typeof value !== "object") throw new Error("服务返回的恢复包材料无效");
  const material = value as Record<string, unknown>;
  if (typeof material.packageJson !== "string" || !material.packageJson ||
      typeof material.digest !== "string" || !/^[0-9a-f]{64}$/.test(material.digest) ||
      !Number.isInteger(material.size) || (material.size as number) < 1 ||
      typeof material.createdAt !== "string" || !Number.isFinite(Date.parse(material.createdAt)) ||
      typeof material.expiresAt !== "string" || !Number.isFinite(Date.parse(material.expiresAt)) ||
      typeof material.savedAt !== "string" || !Number.isFinite(Date.parse(material.savedAt)))
    throw new Error("服务返回的恢复包材料缺少必要字段");
}

/**
 * **未完成执行发现**的类型与准入闸权威在 `packages/core/src/pendingExecutionDiscovery.ts`。
 *
 * 协议层只转发，不另写一份结构：发现结果同时被两端客户端、共享判据与服务端
 * （`server/app/models.py: PendingExecutionsResponse`）读写，多一份定义就多一处漂移。
 * 本轮新增的两个端点：
 *
 * - `GET /api/{wps|office}/executions/pending?workbookIdentity=…`（只读，按完整身份精确查找）
 * - `POST /api/wps/executions/{id}/dismiss-discovery`（只写发现处置元数据）
 *
 * 边界：发现入口**不等于**写入资格 —— 响应里没有执行权令牌、单元格内容、恢复包正文或
 * 完整值计划；认领与续跑仍走既有状态查询与材料水合。
 */
export type {
  PendingExecutionNamespace,
  PendingExecutionStatus,
  PendingExecutionCandidate,
  PendingExecutionDiscoveryResponse,
  PendingExecutionSelection
} from "@excel-bro/core/pendingExecutionDiscovery";

export {
  PENDING_DISCOVERY_SCHEMA_VERSION,
  PENDING_DISCOVERY_LIMIT,
  assertPendingExecutionDiscovery,
  pendingExecutionSelection
} from "@excel-bro/core/pendingExecutionDiscovery";

/**
 * **用户处置留痕**（放弃变更 / 撤销本次执行）的类型权威在
 * `packages/core/src/executionDisposition.ts`。
 *
 * 与"未完成执行发现"同一口径：协议层只转发，不另写一份结构。同一份取值被两端客户端、
 * 共享判据与服务端（`server/app/models.py` 的 `OfficeExecutionDispositionRequest` /
 * `WpsExecutionDispositionRequest`，两处都是 `Literal["abandoned", "undone"]`）读写，
 * 多一份定义就多一处漂移。
 *
 * 端点与语义：
 *
 * - `POST /api/{office|wps}/executions/{id}/disposition`，体 `{ disposition }`，成功 `204`；
 * - 记录不存在 / 身份前缀不属于本宿主 ⇒ `404 EXECUTION_DISPOSITION_NOT_FOUND`，**不补建账本**；
 * - 只写一列时间戳：不推进 `revision`、不改 `status`、不动动作游标、不删恢复材料；
 * - **首次写赢** ⇒ 重复上报幂等（重试只重发这一条 POST，绝不回到宿主再撤一次）。
 */
export type {
  ExecutionDisposition,
  ExecutionDispositionOutcome
} from "@excel-bro/core/executionDisposition";

export {
  EXECUTION_DISPOSITION_LABEL,
  normalizeExecutionDisposition,
  recordExecutionDisposition
} from "@excel-bro/core/executionDisposition";

/**
 * 与 `server/app/models.py: ColumnMappingPair` 同形。
 *
 * 用 `source` / `target` 而不是模型工具参数里的 `from` / `to`：服务端那侧若给
 * `from` 设别名，别名只在 `by_alias=True` 的序列化路径上出现，而 SSE 路径用的是
 * 裸 `model_dump`，两条路径会吐出不同的键名，其中一条上的契约校验会静默失败。
 */
export interface ColumnMappingPair {
  source: string;
  target: string;
}

/**
 * 拆分后取最高级时的一条次序声明。与 `@excel-bro/core` 的
 * `ColumnMappingRankEntry` 同形。`rank` 大者优先。
 */
export interface ColumnMappingRankEntry {
  value: string;
  rank: number;
}

/**
 * 待澄清的**归因**。用户界面据此给出**具体选项**，而不是只弹一句
 * "请补充信息" —— 归因不同，能做的补救动作不同。
 */
export type ColumnMappingClarificationReason =
  | "matchedDelimiterMissingCategories"
  | "unresolvedSeparatorOrMapping"
  /** 读到源区域**确实有空白格**，而空白怎么处理还没定。没读到就不该出现。 */
  | "blankPolicyNeeded";

/**
 * 澄清卡片上可以给用户选的动作。
 *
 * - `addCategories`：补类别映射
 * - `specifyOtherDelimiter`：指定其他分隔符
 * - `treatWholeCellAsCategory`：把整格视为一个类别
 * - `keepTarget` / `clearTarget` / `writeBlankValue` / `writeEmptyString`：
 *   空白的四种处理。**与上面那组不混用** —— 空白卡片上出现"补充分隔符"
 *   只会让用户以为那也能解决空白。
 *
 * 归因决定哪些动作有意义：分隔符已经对上的情形给出"补充分隔符"
 * 只会误导用户去改一个没问题的设置。
 */
export type ColumnMappingResolution =
  | "addCategories"
  | "specifyOtherDelimiter"
  | "treatWholeCellAsCategory"
  | "keepTarget"
  | "clearTarget"
  | "writeBlankValue"
  | "writeEmptyString";

/**
 * 待澄清卡片的内容。**不含来源格的原始内容** —— 原始内容留在客户端，
 * 由 UI 就地展示；这里只有归因与统计。
 */
export interface ColumnMappingClarification {
  reason: ColumnMappingClarificationReason;
  message: string;
  /** 未覆盖的类别（已去重），供界面列出"要补哪些"。 */
  missingCategories: string[];
  /** 归因不确定时，无法处理的整格数量（不回传原文）。 */
  unresolvedCellCount: number;
  /**
   * `blankPolicyNeeded` 时**真实读到**的空白格数量。
   *
   * 必须来自一次真实读取。没有读证据就不许声称"源区域有空白" ——
   * 那是凭空提问。
   */
  blankCellCount?: number;
  /** 界面应当提供的动作。顺序即推荐顺序。 */
  resolvedBy: ColumnMappingResolution[];
}

/**
 * 列映射的**规则**，不含任何单元格值。
 *
 * 源值不在服务端时（`planningReadPolicy: "none"`），服务端下发规则，
 * 客户端调 `@excel-bro/core` 的 `buildColumnMappingActions` 就地算出写入计划。
 * 刻意不带 `values`：单射映射可逆，上传标签列等于上传源列。
 *
 * 多值语义的三组参数（`delimiters` / `valueRank` / `undeclaredSeparators`）
 * 与 `@excel-bro/core` 的 `ColumnMappingInput` **逐字同名**：规则是平的，
 * 折一层子对象会让"夹具里有、运行时没有"这种错在类型检查下隐形。
 * 三组都可省略，省略即"一个来源格只当一个值"，与扩展前行为一致。
 */
export interface ColumnMappingRule {
  kind: "columnMapping";
  sourceSheet: string;
  targetSheet: string;
  sourceRange: string;
  targetRange: string;
  mapping: ColumnMappingPair[];
  /**
   * 空白来源格的处理策略。**省略 = 未定**，只有真的读到空白时才需要决定。
   *
   * 表达"保留目标原值"只有这一种方式：旧字段 `blankLabel` 只能表达
   * "写一个什么标签"，而"保留"意味着这个位置**根本不进入写入范围**，
   * 不是"读回原值再写回"。
   */
  blankPolicy?: BlankPolicy;
  /**
   * 旧协议的"空白标签"。**保留只为兼容旧数据**，新规则一律用 `blankPolicy`。
   *
   * 兼容口径：非空串等价于 `blankPolicy: {kind:"value", value: 同值}`，
   * 语义一点没动。旧协议里空串本来就被拒绝，所以"空串 = 未定"这条折法
   * 不会改变任何一条曾经合法的旧计划。
   */
  blankLabel?: string;
  /**
   * 候选分隔符。拆分时取**出现位置最靠前**的那个。
   * 空数组（或省略）= 不拆分。
   */
  delimiters?: string[];
  /** 拆分后取最高级用的次序，`rank` 大者优先。省略 = 按 `mapping` 声明顺序。 */
  valueRank?: ColumnMappingRankEntry[];
  /**
   * 规则**明示为分隔符候选、但尚未确认可用来切分**的字符。
   * 整格没命中却含有这里的字符时，归因不确定，走澄清。
   */
  undeclaredSeparators?: string[];
  /**
   * 服务端已判定"这条需求缺参数/缺类别、不能算"时带上的澄清内容。
   * 带上时客户端只展示澄清卡片，不算、不写。
   */
  clarification?: ColumnMappingClarification;
}

export type AssistantResponse =
  | Extract<IntentCheckResponse, { kind: "clarification" }>
  | {
      kind: "answer";
      message: string;
      provider: "model" | "local" | "system";
      outcome?: "completed" | "failed";
      stopReason?: "max_turns" | "max_tokens" | "timeout" | "consecutive_errors" | "output_length" | null;
      resultContext?: ResultContext | null;
      /** SDK 本地结果卡指针：只带引用与结构，值由本机回环端点取回。 */
      localResult?: LocalResultCard | null;
      turnId?: string | null;
    }
  | {
      kind: "plan";
      /** Trusted single/batch literal compiler metadata; raw model JSON cannot attest it. */
      planOrigin?: "literal_range_edit" | null;
      plan: AnalysisPlan;
      provider: "model" | "local";
      turnId?: string | null;
    }
  | {
      kind: "local_plan_rule";
      rule: ColumnMappingRule;
      summary: string;
      warnings: string[];
      assumptions: string[];
      provider: "model" | "local";
      turnId?: string | null;
    };

export interface PlanningPauseResult {
  kind: "planning_pause";
  turnId: string;
  status: "waiting_user" | "interrupted";
  revision: number;
  lastSequence: number;
  reason: string;
  activeElapsedSeconds: number;
  approvedWallTimeSeconds: number;
  extensionSeconds: number;
  extensionCount: number;
  maximumExtensions: number;
  canContinue: boolean;
  requiresExtension: boolean;
  /**
   * 暂停原因。`budget_exhausted` 是 2026-10-10 新增：累计活跃执行时间已达
   * 总时限，恢复能力受引用是否仍然有效限制，**不能**显示成"必然可以继续"。
   */
  pauseType: "budget" | "budget_exhausted" | "model_error" | "interrupted";
  currentModelId?: string | null;
  modelSwitchCount: number;
  canSwitchModel: boolean;
  availableModelConnections: Array<{
    catalogModelId: string;
    label: string;
    modelId: string;
    supportsVision: boolean;
  }>;
  sourceFingerprint?: string | null;
  sourceFingerprintSheets?: string[];
  /**
   * 「继续分析」的事实：能不能继续、为什么不能、停在哪个阶段。
   *
   * 由服务端用共享契约 `planningContext` + 执行账本算出，**客户端不得自行推断**
   * ——`canContinue` 是它的合成结果，不是它的替代品：预算或状态不允许时
   * `available` 仍可能是 true。
   */
  continuation: {
    available: boolean;
    /** 固定拒绝码（`PLANNING_CONTEXT_*`）；available 时为 null。 */
    code: string | null;
    message: string;
    stage: "before_model" | "after_tool" | null;
    /**
     * 原任务依赖来源数据 ⇒ 继续前必须重读并比对来源指纹。
     * 服务端判定，客户端只据此决定要不要真读一次。
     */
    sourceVerificationRequired?: boolean;
  };
}

export type PlanningStreamResult = AssistantResponse | PlanningPauseResult;

/**
 * 持久账本里的**版本化恢复描述**。
 *
 * 落盘的规划记录**不是**一个 `AssistantResponse`：值矩阵被省略过，回放出来会是
 * 一个"矩阵全空"的假计划。所以要靠 `valuesOmitted` 分辨这条记录能不能当完整
 * 响应回放，靠 `resultRef` 指出要取真值该走哪个本地结果引用。
 */
export interface PlanningRecoveryDescriptor {
  schemaVersion: number;
  kind: string | null;
  /** true = 有值被省略，`response` 只能当结构参考；false = 可当完整响应回放。 */
  valuesOmitted: boolean;
  resultRef: string | null;
  response: Record<string, unknown>;
}

export type PlanningTaskStatusResult =
  | (Omit<PlanningPauseResult, "status"> & {
      status: "running" | "waiting_user" | "interrupted";
    })
  | {
      turnId: string;
      status: "completed";
      revision: number;
      lastSequence: number;
      /**
       * **没有值被省略**时才有：可直接回放的完整响应。
       * 两个字段互斥 —— 有 `recovery` 就不会有 `result`。
       */
      result?: AssistantResponse;
      /** 值被省略过：只给版本化恢复描述，绝不冒充完整响应。 */
      recovery?: PlanningRecoveryDescriptor | null;
      canContinue: false;
    }
  | {
      turnId: string;
      status: "failed" | "cancelled";
      revision: number;
      lastSequence: number;
      error?: {
        status?: number;
        code?: string;
        message?: string;
        retryable?: boolean;
      } | null;
      canContinue: false;
    };

export interface TurnStepEvent {
  phase: "planning";
  title: string;
  detail?: string | null;
  completedStep?: string | null;
}

export type TaskProgressPhase =
  | "intent"
  | "planning"
  | "preflight"
  | "backup"
  | "executing"
  | "saving"
  | "verifying"
  | "refreshing";

export type TaskProgressStatus =
  | "started"
  | "progress"
  | "completed"
  | "warning"
  | "failed"
  | "cancelled";

export type DecisionTraceKind =
  | "understanding"
  | "evidence"
  | "assumption"
  | "decision"
  | "conflict"
  | "correction";

export interface TaskProgressTarget {
  file?: string | null;
  sheet?: string | null;
  range?: string | null;
}

export interface TaskProgressFact {
  label: string;
  value: string;
  style?: "text" | "code";
}

export interface TaskProgressTrace {
  kind: DecisionTraceKind;
  summary: string;
  basis?: string | null;
  changedFrom?: string | null;
  requiresConfirmation?: boolean;
}

export interface TaskProgressEvent {
  schemaVersion: 1;
  operationId: string;
  executionId?: string | null;
  sequence: number;
  phase: TaskProgressPhase;
  stepId: string;
  status: TaskProgressStatus;
  title: string;
  summary?: string | null;
  detail?: string | null;
  current?: number | null;
  total?: number | null;
  unit?: "rows" | "fields" | "actions" | "sheets" | "files" | "checks" | null;
  actionIndex?: number | null;
  actionCount?: number | null;
  actionType?: string | null;
  toolName?: string | null;
  facts?: TaskProgressFact[] | null;
  target?: TaskProgressTarget | null;
  trace?: TaskProgressTrace | null;
  /** 桥指针：收到即表示服务端在等本机执行一次本地查询并回执。 */
  bridgeTool?: SdkBridgeToolPointer | null;
  /** Python 分析结果卡：包含引用与结构，值由前端另行查询。 */
  pythonResult?: LocalResultCard | null;
  /**
   * Python 代码工件状态。
   *
   * 代码**不塞进** `detail` 之类的进度文案字段：那些字段是给人看的一句话，
   * 会被截断、会被拼接，也会在没有代码时照常出现。代码走这个专门字段，
   * 携带工件身份（`codeId` / `attemptId` / `version`）与阶段，前端据此判定
   * "已运行的就是这一版"，而不是猜。
   */
  pythonCode?: PythonCodeEvent | null;
  timestamp: string;
}

const allowedActionTypes = new Set<ExcelAction["type"]>([
  "createWorksheet",
  "writeTable",
  "writeValues",
  "setFill",
  "setFont",
  "autofit",
  "activateWorksheet",
  "deleteWorksheet",
  "clearRange",
  "insertRange",
  "deleteRange",
  "copyRange",
  "writeFormulas",
  "sortRange",
  "removeDuplicates",
  "filterRange",
  "clearFilter",
  "setDataValidation",
  "setConditionalFormat",
  "setNumberFormat",
  "setBorders",
  "setAlignment",
  "mergeCells",
  "unmergeCells",
  "resizeRange",
  "freezePanes",
  "setHyperlink",
  "addComment",
  "addNote",
  "createTable",
  "createChart",
  "createPivotTable",
  "splitGroupAggregate",
  "computeValues",
  "addNamedRange",
  "addImage",
  "addShape"
]);
const allowedCriterionTypes = new Set<VerificationCriterion["type"]>([
  "worksheetExists",
  "worksheetMissing",
  "rangeEquals",
  "rangeEmpty",
  "formulasEqual",
  "formulasR1C1Equal",
  "rangeSorted",
  "filterApplied",
  "filterCleared",
  "tableExists",
  "rangeFormatMatches",
  "bordersMatch",
  "dataValidationMatches",
  "freezePanesMatches",
  "chartExists",
  "pivotTableExists"
]);

function rangeDimensions(address: string): [number, number] | null {
  const match =
    /^\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?$/.exec(
      address.trim()
    );
  if (!match) return null;
  const columnNumber = (name: string) =>
    [...name.toUpperCase()].reduce(
      (result, character) => result * 26 + character.charCodeAt(0) - 64,
      0
    );
  const startColumn = columnNumber(match[1]);
  const startRow = Number(match[2]);
  const endColumn = columnNumber(match[3] ?? match[1]);
  const endRow = Number(match[4] ?? match[2]);
  if (endColumn < startColumn || endRow < startRow) return null;
  return [endRow - startRow + 1, endColumn - startColumn + 1];
}

export function assertAssistantResponse(
  value: unknown
): asserts value is AssistantResponse {
  if (!value || typeof value !== "object") {
    throw new Error("服务返回的内容不是有效对象");
  }

  const response = value as Partial<AssistantResponse>;
  if (response.kind === "clarification") {
    assertIntentCheckResponse(value);
    return;
  }
  if (response.kind === "answer") {
    if (response.outcome !== undefined && response.outcome !== "completed" && response.outcome !== "failed") {
      throw new Error("服务返回的规划终态无效");
    }
    if (response.stopReason != null && !["max_turns", "max_tokens", "timeout", "consecutive_errors", "output_length"].includes(response.stopReason)) {
      throw new Error("服务返回的规划停止原因无效");
    }
    if (!("message" in response) || typeof response.message !== "string") {
      throw new Error("服务返回的回答缺少 message");
    }
    if (
      response.resultContext !== undefined &&
      response.resultContext !== null &&
      (response.resultContext.kind !== "table" ||
        !Array.isArray(response.resultContext.headers) ||
        !Array.isArray(response.resultContext.rows))
    ) {
      throw new Error("服务返回的 resultContext 无效");
    }
    return;
  }

  if (response.kind === "local_plan_rule") {
    assertColumnMappingRule(response.rule);
    if (typeof response.summary !== "string" || !response.summary.trim()) {
      throw new Error("服务返回的本地计算规则缺少 summary");
    }
    return;
  }

  if (response.kind !== "plan") {
    throw new Error("服务返回了未知响应类型");
  }
  assertEditablePlan(response.plan);
}

/**
 * 校验可选的空白策略。
 *
 * `undefined` 合法（未定）；给了就必须是四种已知形态之一，且 `kind:"value"`
 * 必须带字符串 `value` —— 空串是合法的（"显式写空字符串"是一种明确意图），
 * 缺字段或未知 kind 都要拒。
 */
function assertOptionalBlankPolicy(value: unknown): void {
  if (value === undefined || value === null) return;
  if (typeof value !== "object" || Array.isArray(value)) {
    throw new Error("服务返回的空白策略不是有效对象");
  }
  const policy = value as Record<string, unknown>;
  switch (policy.kind) {
    case "keep":
    case "clear":
    case "undecided":
      return;
    case "value":
      if (typeof policy.value !== "string") {
        throw new Error("服务返回的空白策略缺少字符串类型的 value");
      }
      return;
    default:
      throw new Error(`服务返回了无法识别的空白策略：${String(policy.kind)}`);
  }
}

/**
 * 校验列映射规则。**不校验区域是否在已授权范围内** —— 那要读宿主，
 * 属于客户端执行前的授权闸，不是契约层的事。
 */
export function assertColumnMappingRule(
  value: unknown
): asserts value is ColumnMappingRule {
  if (!value || typeof value !== "object") {
    throw new Error("服务返回的映射规则不是有效对象");
  }
  const rule = value as Partial<ColumnMappingRule>;
  if (rule.kind !== "columnMapping") {
    throw new Error("服务返回了未知的本地计算规则类型");
  }
  for (const field of [
    "sourceSheet",
    "targetSheet",
    "sourceRange",
    "targetRange"
  ] as const) {
    if (typeof rule[field] !== "string" || !rule[field]!.trim()) {
      throw new Error(`服务返回的映射规则缺少 ${field}`);
    }
  }
  /*
   * 空白策略：**给了就必须认得出来，没给就不算错**。
   *
   * 没给 = 未定，只在真的读到空白时才升级为一次澄清 —— 若这里直接拒，
   * 等于把"还没读源数据就先问空白怎么办"这条老路又开回来。
   * 给了却认不出来则必须拒：静默折成"未定"会让用户答过的问题被重问一遍。
   */
  assertOptionalBlankPolicy(rule.blankPolicy);
  if (!Array.isArray(rule.mapping) || rule.mapping.length === 0) {
    throw new Error("服务返回的映射规则没有任何映射项");
  }
  for (const pair of rule.mapping) {
    // source 可以是空串（"区分空串与真空格"的合法意图）；target 不行。
    if (
      !pair ||
      typeof pair.source !== "string" ||
      typeof pair.target !== "string" ||
      !pair.target.trim()
    ) {
      throw new Error("服务返回的映射项不是一对非空字符串");
    }
  }
  // 多值语义参数都是可选的，但**给了就必须是字符串数组**：
  // 静默忽略一个类型不对的 delimiters 会让客户端按"不拆分"算，
  // 与服务端（若它算得出）不一致，而这种不一致在单元格层面看不出来。
  for (const field of ["delimiters", "undeclaredSeparators"] as const) {
    const items = rule[field];
    if (items === undefined) continue;
    if (
      !Array.isArray(items) ||
      items.some((item) => typeof item !== "string" || item.length === 0)
    ) {
      throw new Error(`服务返回的映射规则里 ${field} 不是非空字符串数组`);
    }
  }
  if (rule.valueRank !== undefined) {
    if (!Array.isArray(rule.valueRank)) {
      throw new Error("服务返回的映射规则里 valueRank 不是数组");
    }
    for (const entry of rule.valueRank) {
      if (
        !entry ||
        typeof entry.value !== "string" ||
        !entry.value.trim() ||
        typeof entry.rank !== "number"
      ) {
        throw new Error("服务返回的映射规则里 valueRank 项无效");
      }
    }
  }
  if (rule.clarification !== undefined && rule.clarification !== null) {
    assertColumnMappingClarification(rule.clarification);
  }
}

/**
 * 校验澄清卡片。**不校验原文** —— 原文根本不在协议里（留在客户端），
 * 这里能查的只有归因、计数与候选动作。
 */
export function assertColumnMappingClarification(
  value: unknown
): asserts value is ColumnMappingClarification {
  if (!value || typeof value !== "object") {
    throw new Error("服务返回的澄清内容不是有效对象");
  }
  const clarification = value as Partial<ColumnMappingClarification>;
  if (
    clarification.reason !== "matchedDelimiterMissingCategories" &&
    clarification.reason !== "unresolvedSeparatorOrMapping"
  ) {
    throw new Error("服务返回的澄清内容缺少有效归因");
  }
  if (
    typeof clarification.message !== "string" ||
    !clarification.message.trim()
  ) {
    throw new Error("服务返回的澄清内容缺少说明");
  }
  if (
    !Array.isArray(clarification.missingCategories) ||
    clarification.missingCategories.some((item) => typeof item !== "string")
  ) {
    throw new Error("服务返回的澄清内容里 missingCategories 无效");
  }
  if (
    typeof clarification.unresolvedCellCount !== "number" ||
    clarification.unresolvedCellCount < 0
  ) {
    throw new Error("服务返回的澄清内容里 unresolvedCellCount 无效");
  }
  if (!Array.isArray(clarification.resolvedBy)) {
    throw new Error("服务返回的澄清内容里 resolvedBy 无效");
  }
  for (const option of clarification.resolvedBy) {
    if (
      option !== "addCategories" &&
      option !== "specifyOtherDelimiter" &&
      option !== "treatWholeCellAsCategory"
    ) {
      throw new Error("服务返回的澄清内容里有未知的可选动作");
    }
  }
}

export function assertPlanningPauseResult(
  value: unknown
): asserts value is PlanningPauseResult {
  if (!value || typeof value !== "object") {
    throw new Error("服务返回的暂停状态不是有效对象");
  }
  const result = value as Partial<PlanningPauseResult>;
  if (
    result.kind !== "planning_pause" ||
    typeof result.turnId !== "string" ||
    (result.status !== "waiting_user" && result.status !== "interrupted") ||
    typeof result.revision !== "number" ||
    typeof result.lastSequence !== "number" ||
    typeof result.reason !== "string" ||
    typeof result.activeElapsedSeconds !== "number" ||
    typeof result.approvedWallTimeSeconds !== "number" ||
    typeof result.extensionSeconds !== "number" ||
    typeof result.extensionCount !== "number" ||
    typeof result.maximumExtensions !== "number" ||
    typeof result.canContinue !== "boolean" ||
    typeof result.requiresExtension !== "boolean" ||
    (result.pauseType !== "budget" &&
      result.pauseType !== "budget_exhausted" &&
      result.pauseType !== "model_error" &&
      result.pauseType !== "interrupted") ||
    typeof result.modelSwitchCount !== "number" ||
    typeof result.canSwitchModel !== "boolean" ||
    !Array.isArray(result.availableModelConnections)
  ) {
    throw new Error("服务返回的规划暂停状态缺少必要字段");
  }
  if (
    result.sourceFingerprintSheets !== undefined &&
    !Array.isArray(result.sourceFingerprintSheets)
  ) {
    throw new Error("服务返回的规划来源工作表无效");
  }
  if (
    result.availableModelConnections.some(
      (connection) =>
        !connection ||
        typeof connection !== "object" ||
        typeof connection.catalogModelId !== "string" ||
        typeof connection.label !== "string" ||
        typeof connection.modelId !== "string" ||
        typeof connection.supportsVision !== "boolean"
    )
  ) {
    throw new Error("服务返回的备用模型连接无效");
  }
}

/**
 * 六个**格式动作**的联合类型。
 *
 * 抽出来是因为它们共享两件事：`sheet` + `range` 定位，以及"只改指定属性"的语义。
 * 字段级校验、预览的属性清单、影响范围说明都只对这一段做，用 `ExcelAction` 全量
 * 去 switch 会连 `createWorksheet`（根本没有 `range`）一起卷进来。
 */
export type FormatAction = Extract<
  ExcelAction,
  {
    type:
      | "setFill"
      | "setFont"
      | "setBorders"
      | "setAlignment"
      | "setNumberFormat"
      | "resizeRange";
  }
>;

const FORMAT_ACTION_TYPES = new Set<string>([
  "setFill",
  "setFont",
  "setBorders",
  "setAlignment",
  "setNumberFormat",
  "resizeRange"
]);

export function isFormatAction(action: ExcelAction): action is FormatAction {
  return FORMAT_ACTION_TYPES.has(action.type);
}

/** `setBorders` 的缺省值。服务端 `models.py` 给的是同一组，执行器与回读判据都从这里取。 */
export const BORDER_DEFAULTS = {
  style: "continuous",
  color: "#000000",
  weight: "thin"
} as const satisfies { style: BorderLineStyle; color: string; weight: BorderWeight };

/**
 * 补齐 `setBorders` 的缺省值。
 *
 * 类型层把 `style`/`color`/`weight` 标成可选（对齐 Python 的默认值），所以**不能**
 * 在执行器里直接 `action.style` —— 那会得到 `undefined` 并写成非法值。执行与回读
 * 必须走同一条补齐逻辑，否则"写入用的 style"和"核对用的 style"会在缺省时对不上。
 */
export function normalizeBorderAction(
  action: Extract<ExcelAction, { type: "setBorders" }>
): { sides: BorderSide[]; style: BorderLineStyle; color: string; weight: BorderWeight } {
  return {
    sides: action.sides,
    style: action.style ?? BORDER_DEFAULTS.style,
    color: action.color ?? BORDER_DEFAULTS.color,
    weight: action.weight ?? BORDER_DEFAULTS.weight
  };
}

function assertHexColor(value: unknown, label: string): void {
  if (typeof value !== "string" || !HEX_COLOR_PATTERN.test(value)) {
    throw new Error(`${label}必须是 #RRGGBB 形式的颜色，收到 ${JSON.stringify(value)}`);
  }
}

function assertInClosedSet<T extends string>(
  value: unknown,
  allowed: readonly T[],
  label: string
): void {
  if (typeof value !== "string" || !(allowed as readonly string[]).includes(value)) {
    throw new Error(`${label}只能是 ${allowed.join(" / ")}，收到 ${JSON.stringify(value)}`);
  }
}

function assertOptionalBoolean(value: unknown, label: string): void {
  if (value != null && typeof value !== "boolean") {
    throw new Error(`${label}必须是布尔值`);
  }
}

/**
 * 格式动作的**字段级**校验（与 `server/app/models.py` 的 `Literal` / `Field` 等价）。
 *
 * 动作名白名单只保证"这个动作被允许"，不保证"参数是宿主能接受的"。颜色写成 `red`、
 * 对齐写成 `middle`、列宽写成 900，白名单一条都不拦 —— 但它们会一路走到宿主才失败，
 * 而那一刻前面的动作可能已经落盘了。所以格式参数必须在**动任何单元格之前**查完。
 *
 * 这条校验对**模型生成的计划**和**用户手改的 JSON** 一视同仁：手改的不该更宽松。
 */
export function assertFormatActionFields(action: ExcelAction): void {
  // 非格式动作没有 `sheet`/`range` 这两个定位字段，直接跳过。
  if (!isFormatAction(action)) return;
  const where = `「${action.sheet}!${action.range}」`;
  switch (action.type) {
    case "setFill":
      assertHexColor(action.color, `${where}的填充色`);
      return;
    case "setFont": {
      const font = action;
      assertOptionalBoolean(font.bold, `${where}的加粗标记`);
      assertOptionalBoolean(font.italic, `${where}的斜体标记`);
      if (font.color != null) assertHexColor(font.color, `${where}的字色`);
      if (font.name != null) {
        if (
          typeof font.name !== "string" ||
          !font.name.trim() ||
          font.name.length > FONT_NAME_MAX_LENGTH
        ) {
          throw new Error(`${where}的字体名无效`);
        }
      }
      if (font.size != null) {
        if (
          typeof font.size !== "number" ||
          !Number.isFinite(font.size) ||
          font.size < FONT_SIZE_RANGE.min ||
          font.size > FONT_SIZE_RANGE.max
        ) {
          throw new Error(
            `${where}的字号必须在 ${FONT_SIZE_RANGE.min}–${FONT_SIZE_RANGE.max} 之间`
          );
        }
      }
      // 一个属性都没给 = 这个动作什么也不做。放过去会让计划"看起来成功、实际没改"，
      // 与"只有指定属性被改变"的承诺正好相反，所以当场拒绝。
      if (
        font.bold == null &&
        font.italic == null &&
        font.name == null &&
        font.size == null &&
        font.color == null
      ) {
        throw new Error(`${where}的 setFont 没有指定任何要修改的字体属性`);
      }
      return;
    }
    case "setBorders": {
      const borders = action;
      if (!Array.isArray(borders.sides) || borders.sides.length === 0) {
        throw new Error(`${where}的 setBorders 至少需要一条边`);
      }
      if (new Set(borders.sides).size !== borders.sides.length) {
        throw new Error(`${where}的 setBorders 里有重复的边`);
      }
      borders.sides.forEach((side) =>
        assertInClosedSet(side, BORDER_SIDES, `${where}的边框位置`)
      );
      const normalized = normalizeBorderAction(borders);
      assertInClosedSet(normalized.style, BORDER_LINE_STYLES, `${where}的线型`);
      assertInClosedSet(normalized.weight, BORDER_WEIGHTS, `${where}的线宽`);
      assertHexColor(normalized.color, `${where}的边框颜色`);
      return;
    }
    case "setAlignment": {
      const alignment = action;
      if (alignment.horizontal != null) {
        assertInClosedSet(alignment.horizontal, HORIZONTAL_ALIGNMENTS, `${where}的水平对齐`);
      }
      if (alignment.vertical != null) {
        assertInClosedSet(alignment.vertical, VERTICAL_ALIGNMENTS, `${where}的垂直对齐`);
      }
      assertOptionalBoolean(alignment.wrapText, `${where}的自动换行标记`);
      if (
        alignment.horizontal == null &&
        alignment.vertical == null &&
        alignment.wrapText == null
      ) {
        throw new Error(`${where}的 setAlignment 没有指定任何要修改的对齐属性`);
      }
      return;
    }
    case "setNumberFormat": {
      const numberFormat = action;
      if (
        typeof numberFormat.formatCode !== "string" ||
        numberFormat.formatCode.length === 0 ||
        numberFormat.formatCode.length > NUMBER_FORMAT_MAX_LENGTH
      ) {
        throw new Error(
          `${where}的数字格式码必须是 1–${NUMBER_FORMAT_MAX_LENGTH} 个字符`
        );
      }
      return;
    }
    case "resizeRange": {
      const resize = action;
      const checkRange = (
        value: unknown,
        range: { min: number; max: number },
        label: string
      ) => {
        if (value == null) return;
        if (
          typeof value !== "number" ||
          !Number.isFinite(value) ||
          value < range.min ||
          value > range.max
        ) {
          throw new Error(`${label}必须在 ${range.min}–${range.max} 之间`);
        }
      };
      checkRange(resize.rowHeight, ROW_HEIGHT_RANGE, `${where}的行高`);
      checkRange(resize.columnWidth, COLUMN_WIDTH_RANGE, `${where}的列宽`);
      if (resize.rowHeight == null && resize.columnWidth == null) {
        throw new Error(`${where}的 resizeRange 没有指定行高或列宽`);
      }
      return;
    }
  }
}

/**
 * 校验一份计划的动作白名单与验收条件结构。
 *
 * 模型返回的计划和用户手改的 JSON 走同一道门 —— 手改的不该比模型生成的宽松，
 * 否则「能编辑」就变成了绕过白名单的后门。
 *
 * 这里只做结构、白名单与**格式动作的字段级**校验，不查 title/summary 之类的必填
 * 字段：那些交给服务端 Pydantic 规范化，前端多加一层只会和既有调用方打架。
 */
export function assertEditablePlan(
  value: unknown
): asserts value is AnalysisPlan {
  if (!value || typeof value !== "object") {
    throw new Error("计划不是有效对象");
  }
  const response = { plan: value as AnalysisPlan };
  if (!Array.isArray(response.plan.actions)) {
    throw new Error("计划缺少 actions");
  }

  for (const action of response.plan.actions as Array<{ type?: string }>) {
    if (!action.type || !allowedActionTypes.has(action.type as ExcelAction["type"])) {
      throw new Error(`计划包含未授权动作：${action.type ?? "unknown"}`);
    }
    assertFormatActionFields(action as ExcelAction);
  }

  const criteria = response.plan.acceptanceCriteria;
  if (criteria !== undefined && !Array.isArray(criteria)) {
    throw new Error("计划的验收条件不是有效数组");
  }
  const actionSheets = new Set(response.plan.actions.map((action) => action.sheet));
  for (const criterion of criteria ?? []) {
    if (
      !criterion.type ||
      !allowedCriterionTypes.has(criterion.type as VerificationCriterion["type"])
    ) {
      throw new Error(`计划包含未知验收条件：${criterion.type ?? "unknown"}`);
    }
    if (!actionSheets.has(criterion.sheet)) {
      throw new Error("验收条件只能检查计划实际操作的工作表");
    }
    if (
      [
        "rangeSorted",
        "filterApplied",
        "tableExists",
        "rangeFormatMatches",
        "bordersMatch",
        "dataValidationMatches"
      ].includes(
        criterion.type
      ) &&
      typeof (criterion as { range?: unknown }).range !== "string"
    ) {
      throw new Error(`${criterion.type} 缺少有效范围`);
    }
    if (
      criterion.type === "rangeSorted" &&
      (!Array.isArray(criterion.keys) ||
        criterion.keys.length === 0 ||
        typeof criterion.hasHeaders !== "boolean" ||
        criterion.keys.some(
          (key) =>
            !Number.isInteger(key?.column) ||
            key.column < 0 ||
            typeof key.ascending !== "boolean"
        ))
    ) {
      throw new Error("rangeSorted 的排序键无效");
    }
    if (
      criterion.type === "filterApplied" &&
      (!Number.isInteger(criterion.column) ||
        criterion.column < 0 ||
        !Array.isArray(criterion.values) ||
        criterion.values.length === 0)
    ) {
      throw new Error("filterApplied 的筛选条件无效");
    }
    if (
      criterion.type === "tableExists" &&
      (typeof criterion.hasHeaders !== "boolean" ||
        (criterion.name !== undefined &&
          criterion.name !== null &&
          typeof criterion.name !== "string"))
    ) {
      throw new Error("tableExists 的表格条件无效");
    }
    if (criterion.type === "rangeFormatMatches") {
      const expectedProperties = [
        criterion.fillColor,
        criterion.bold,
        criterion.fontColor,
        criterion.fontName,
        criterion.fontSize,
        criterion.fontItalic,
        criterion.numberFormat,
        criterion.horizontal,
        criterion.vertical,
        criterion.wrapText,
        criterion.rowHeight,
        criterion.columnWidth
      ];
      if (expectedProperties.every((property) => property == null)) {
        throw new Error("rangeFormatMatches 缺少格式属性");
      }
    }
    if (
      criterion.type === "bordersMatch" &&
      (!Array.isArray(criterion.sides) ||
        criterion.sides.length === 0 ||
        typeof criterion.style !== "string" ||
        typeof criterion.color !== "string" ||
        typeof criterion.weight !== "string")
    ) {
      throw new Error("bordersMatch 的边框条件无效");
    }
    if (
      criterion.type === "dataValidationMatches" &&
      (typeof criterion.validationType !== "string" ||
        !Array.isArray(criterion.values) ||
        typeof criterion.operator !== "string" ||
        typeof criterion.allowBlank !== "boolean")
    ) {
      throw new Error("dataValidationMatches 的验证条件无效");
    }
    if (
      criterion.type === "freezePanesMatches" &&
      (!Number.isInteger(criterion.rows) ||
        criterion.rows < 0 ||
        !Number.isInteger(criterion.columns) ||
        criterion.columns < 0)
    ) {
      throw new Error("freezePanesMatches 的冻结位置无效");
    }
    if (
      criterion.type === "chartExists" &&
      (typeof criterion.chartType !== "string" ||
        typeof criterion.sourceRange !== "string" ||
        (criterion.name != null && typeof criterion.name !== "string") ||
        (criterion.targetRange != null &&
          typeof criterion.targetRange !== "string"))
    ) {
      throw new Error("chartExists 的图表条件无效");
    }
    if (
      criterion.type === "pivotTableExists" &&
      (typeof criterion.name !== "string" ||
        typeof criterion.sourceSheet !== "string" ||
        typeof criterion.sourceRange !== "string" ||
        typeof criterion.destinationCell !== "string" ||
        !Array.isArray(criterion.rowFields) ||
        !Array.isArray(criterion.columnFields) ||
        !Array.isArray(criterion.valueFields))
    ) {
      throw new Error("pivotTableExists 的数据透视表条件无效");
    }
    if (
      criterion.type === "rangeEquals" ||
      criterion.type === "formulasEqual"
    ) {
      const expected = criterion.expected;
      const width = Array.isArray(expected?.[0]) ? expected[0].length : 0;
      const rectangular =
        Array.isArray(expected) &&
        expected.length > 0 &&
        expected.length <= 500 &&
        width > 0 &&
        width <= 50 &&
        expected.every((row) => Array.isArray(row) && row.length === width);
      const dimensions =
        typeof criterion.range === "string"
          ? rangeDimensions(criterion.range)
          : null;
      if (
        !rectangular ||
        !dimensions ||
        dimensions[0] !== expected.length ||
        dimensions[1] !== width
      ) {
        throw new Error(`${criterion.type} 的范围与 expected 尺寸不一致`);
      }
    }
    if (criterion.type === "formulasR1C1Equal") {
      const dimensions =
        typeof criterion.range === "string"
          ? rangeDimensions(criterion.range)
          : null;
      if (
        typeof criterion.expected !== "string" ||
        criterion.expected.length === 0 ||
        !dimensions
      ) {
        throw new Error("formulasR1C1Equal 的范围或 expected 无效");
      }
    }
  }
}

export function parseTurnStepEvent(value: unknown): TurnStepEvent | null {
  if (!value || typeof value !== "object") {
    return null;
  }
  const event = value as Partial<TurnStepEvent>;
  if (event.phase !== "planning" || typeof event.title !== "string" || !event.title) {
    return null;
  }
  return {
    phase: "planning",
    title: event.title,
    detail: typeof event.detail === "string" ? event.detail : null,
    completedStep:
      typeof event.completedStep === "string" ? event.completedStep : null
  };
}

const taskProgressPhases = new Set<TaskProgressPhase>([
  "intent",
  "planning",
  "preflight",
  "backup",
  "executing",
  "saving",
  "verifying",
  "refreshing"
]);

const taskProgressStatuses = new Set<TaskProgressStatus>([
  "started",
  "progress",
  "completed",
  "warning",
  "failed",
  "cancelled"
]);

const decisionTraceKinds = new Set<DecisionTraceKind>([
  "understanding",
  "evidence",
  "assumption",
  "decision",
  "conflict",
  "correction"
]);

function optionalNonNegativeNumber(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) && value >= 0
    ? value
    : null;
}

const TASK_PROGRESS_FACT_LIMIT = 12;
const TASK_PROGRESS_FACT_VALUE_LIMIT = 500;
const TASK_PROGRESS_TRUNCATION_MARKER = "…（内容过长，已截断）";
const SENSITIVE_PROGRESS_LABEL =
  /token|authorization|cookie|api[ _-]?key|secret|password|credential|密钥|口令/i;
const SAFE_PROGRESS_FACT_LABELS = new Set([
  "工具",
  "动作",
  "动作数",
  "参数",
  "目标",
  "结果",
  "依据",
  "查询",
  "范围",
  "工作表",
  "工作表数",
  "文件",
  "字段数",
  "匹配数",
  "进度",
  "授权范围",
  "读取方式",
  "快照"
]);

function truncateProgressText(value: string, limit: number): string {
  if (value.length <= limit) return value;
  const keep = Math.max(0, limit - TASK_PROGRESS_TRUNCATION_MARKER.length);
  return value.slice(0, keep) + TASK_PROGRESS_TRUNCATION_MARKER;
}

function redactProgressValue(value: string): string {
  return value
    .replace(/\bBearer\s+[A-Za-z0-9._~+/=-]{8,}/gi, "Bearer 已隐藏")
    .replace(
      /([A-Za-z][A-Za-z0-9+.-]*:\/\/)([^/\s:@]+):([^@\s/]+)@/g,
      "$1已隐藏@"
    )
    .replace(/\bsk-[A-Za-z0-9_-]{12,}\b/g, "已隐藏")
    .replace(
      /\b(api[ _-]?key|token|secret|password)\s*[:=]\s*[^\s,;]+/gi,
      "$1=已隐藏"
    );
}

function parseProgressFacts(value: unknown): TaskProgressFact[] | null {
  if (!Array.isArray(value)) return null;
  const parsed: TaskProgressFact[] = [];
  for (const item of value.slice(0, TASK_PROGRESS_FACT_LIMIT)) {
    if (!item || typeof item !== "object") continue;
    const candidate = item as Record<string, unknown>;
    if (
      typeof candidate.label !== "string" ||
      !candidate.label.trim() ||
      typeof candidate.value !== "string" ||
      !candidate.value.trim()
    ) {
      continue;
    }
    const label = truncateProgressText(candidate.label.trim(), 40);
    if (
      !SENSITIVE_PROGRESS_LABEL.test(label) &&
      !SAFE_PROGRESS_FACT_LABELS.has(label)
    ) {
      continue;
    }
    const rawValue = candidate.value.trim();
    const safeValue = SENSITIVE_PROGRESS_LABEL.test(label)
      ? "已隐藏"
      : redactProgressValue(rawValue);
    parsed.push({
      label,
      value: truncateProgressText(
        safeValue,
        TASK_PROGRESS_FACT_VALUE_LIMIT
      ),
      style: candidate.style === "code" ? "code" : "text"
    });
  }
  return parsed.length > 0 ? parsed : null;
}

function parseProgressTarget(value: unknown): TaskProgressTarget | null {
  if (!value || typeof value !== "object") return null;
  const candidate = value as Record<string, unknown>;
  const rawFile =
    typeof candidate.file === "string" ? candidate.file.trim() : "";
  const file = rawFile
    ? /^[A-Za-z]:[\\/]|^\\\\|^\//.test(rawFile)
      ? (() => {
          const parts = rawFile.split(/[\\/]/).filter(Boolean);
          return parts.length > 0 ? parts[parts.length - 1] : null;
        })()
      : truncateProgressText(rawFile, 1000)
    : null;
  const sheet =
    typeof candidate.sheet === "string" && candidate.sheet.trim()
      ? truncateProgressText(candidate.sheet.trim(), 500)
      : null;
  const range =
    typeof candidate.range === "string" && candidate.range.trim()
      ? truncateProgressText(candidate.range.trim(), 80)
      : null;
  return file || sheet || range ? { file, sheet, range } : null;
}

/**
 * 单帧代码正文的接收上限。
 *
 * 上限本身放在 `config/capabilities.json` 的 `pythonCode.maxCharsPerEvent`（配置
 * 优先，代码只是兜底）：这是**协议边界**的上限（一帧能带多少），不是界面显示
 * 上限。超出即截断并置 `truncated`，界面据此不再提供"复制全部"。
 */
export const PYTHON_CODE_MAX_CHARS =
  Number(capabilities.pythonCode?.maxCharsPerEvent) > 0
    ? Number(capabilities.pythonCode.maxCharsPerEvent)
    : 20000;

function positiveInt(value: unknown, fallback: number): number {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed > 0 ? Math.floor(parsed) : fallback;
}

/** 展示上限同样来自配置：预览几行、超过多少行折叠。 */
export const PYTHON_CODE_LIMITS: PythonCodeLimits = {
  maxChars: PYTHON_CODE_MAX_CHARS,
  previewLines: positiveInt(capabilities.pythonCode?.previewLines, 6),
  collapseAfterLines: positiveInt(capabilities.pythonCode?.collapseAfterLines, 24)
};

// not_started 表示未收到工作程序执行确认；不把 Docker CLI 启动当作已执行。
const PYTHON_CODE_STAGES = new Set([  "generating",
  "generated",
  "validating",
  "running",
  "analysis_completed",
  "summarizing",
  "paused",
  "not_started",
  "generation_failed",
  "validation_failed",
  "execution_failed"
]);

/**
 * 严格解析 Python 代码工件状态。
 *
 * 形状不符一律当"没有"：半成品工件会让界面把一段不完整、来源不明的代码
 * 当成"已生成"，从而出现"看起来能跑"的误导。代码正文按上限截断，并在
 * `truncated` 上留痕 —— 前端据此关掉"复制全部"。
 */
export function parsePythonCodeEvent(value: unknown): PythonCodeEvent | null {
  if (!value || typeof value !== "object") return null;
  const candidate = value as Record<string, unknown>;
  const codeId = typeof candidate.codeId === "string" ? candidate.codeId.trim() : "";
  const attemptId =
    typeof candidate.attemptId === "string" ? candidate.attemptId.trim() : "";
  if (!codeId || codeId.length > 120 || !attemptId || attemptId.length > 120) {
    return null;
  }
  if (
    typeof candidate.stage !== "string" ||
    !PYTHON_CODE_STAGES.has(candidate.stage)
  ) {
    return null;
  }
  const code = typeof candidate.code === "string" ? candidate.code : undefined;
  const delta = typeof candidate.delta === "string" ? candidate.delta : undefined;
  // 内容阶段（生成中 / 已生成）必须带正文；验证、执行、完成这些帧只推进状态。
  if (
    (candidate.stage === "generating" || candidate.stage === "generated") &&
    code === undefined &&
    delta === undefined
  ) {
    return null;
  }
  const limit = PYTHON_CODE_MAX_CHARS;
  const bounded = (text: string | undefined) =>
    text === undefined
      ? undefined
      : text.length > limit
        ? text.slice(0, limit)
        : text;
  const clippedCode = bounded(code);
  const clippedDelta = bounded(delta);
  const droppedByLimit =
    (code !== undefined && clippedCode !== code) ||
    (delta !== undefined && clippedDelta !== delta);
  return {
    codeId,
    attemptId,
    stage: candidate.stage as PythonCodeStage,
    ...(typeof candidate.version === "number" && Number.isFinite(candidate.version)
      ? { version: Math.max(1, Math.floor(candidate.version)) }
      : {}),
    ...(clippedCode !== undefined ? { code: clippedCode } : {}),
    ...(clippedDelta !== undefined ? { delta: clippedDelta } : {}),
    truncated: Boolean(candidate.truncated) || droppedByLimit,
    ...(typeof candidate.complete === "boolean" ? { complete: candidate.complete } : {}),
    // 正文字符数：正文没随事件下发时，界面靠它说清规模（"812 字符"）。
    ...(typeof candidate.chars === "number" &&
    Number.isFinite(candidate.chars) &&
    candidate.chars >= 0
      ? { chars: Math.floor(candidate.chars) }
      : {}),
    // 服务端显式声明的"正文不随事件保留"。缺省不等于 false。
    ...(typeof candidate.codeRetained === "boolean"
      ? { codeRetained: candidate.codeRetained }
      : {}),
    // 未开始执行的具体原因：只允许一句固定文案（服务端生成，长度受限）。
    ...(typeof candidate.reason === "string" && candidate.reason.trim()
      ? { reason: candidate.reason.trim().slice(0, 300) }
      : {}),
    ...(typeof candidate.executedVersion === "number" &&
    Number.isFinite(candidate.executedVersion)
      ? { executedVersion: Math.max(1, Math.floor(candidate.executedVersion)) }
      : {}),
    ...(typeof candidate.sequence === "number" && Number.isFinite(candidate.sequence)
      ? { sequence: Math.max(0, Math.floor(candidate.sequence)) }
      : {})
  };
}

export function parseTaskProgressEvent(
  value: unknown,
  legacyOperationId = "legacy-turn",
  legacySequence = 0
): TaskProgressEvent | null {  if (!value || typeof value !== "object") return null;
  const candidate = value as Partial<TaskProgressEvent>;
  if (candidate.schemaVersion !== 1) {
    const legacy = parseTurnStepEvent(value);
    if (!legacy) return null;
    return {
      schemaVersion: 1,
      operationId: legacyOperationId,
      sequence: legacySequence,
      phase: "planning",
      stepId: `legacy-${legacySequence}`,
      status: "progress",
      title: legacy.title,
      detail: legacy.detail ?? null,
      summary: null,
      timestamp: new Date().toISOString()
    };
  }
  if (
    typeof candidate.operationId !== "string" ||
    !candidate.operationId.trim() ||
    typeof candidate.sequence !== "number" ||
    !Number.isInteger(candidate.sequence) ||
    candidate.sequence < 0 ||
    !taskProgressPhases.has(candidate.phase as TaskProgressPhase) ||
    !taskProgressStatuses.has(candidate.status as TaskProgressStatus) ||
    typeof candidate.stepId !== "string" ||
    !candidate.stepId.trim() ||
    typeof candidate.title !== "string" ||
    !candidate.title.trim() ||
    typeof candidate.timestamp !== "string" ||
    !candidate.timestamp.trim()
  ) {
    return null;
  }
  const current = optionalNonNegativeNumber(candidate.current);
  const total = optionalNonNegativeNumber(candidate.total);
  if (current !== null && total !== null && current > total) return null;
  const traceValue = candidate.trace;
  const trace =
    traceValue &&
    typeof traceValue === "object" &&
    decisionTraceKinds.has(traceValue.kind as DecisionTraceKind) &&
    typeof traceValue.summary === "string" &&
    traceValue.summary.trim()
      ? {
          kind: traceValue.kind as DecisionTraceKind,
          summary: traceValue.summary,
          basis: typeof traceValue.basis === "string" ? traceValue.basis : null,
          changedFrom:
            typeof traceValue.changedFrom === "string"
              ? traceValue.changedFrom
              : null,
          requiresConfirmation: Boolean(traceValue.requiresConfirmation)
        }
      : null;
  return {
    schemaVersion: 1,
    operationId: candidate.operationId,
    executionId:
      typeof candidate.executionId === "string" ? candidate.executionId : null,
    sequence: candidate.sequence,
    phase: candidate.phase as TaskProgressPhase,
    stepId: candidate.stepId,
    status: candidate.status as TaskProgressStatus,
    title: candidate.title,
    summary:
      typeof candidate.summary === "string"
        ? truncateProgressText(candidate.summary, 1000)
        : null,
    detail:
      typeof candidate.detail === "string"
        ? truncateProgressText(candidate.detail, 2000)
        : null,
    current,
    total,
    unit: candidate.unit ?? null,
    actionIndex: optionalNonNegativeNumber(candidate.actionIndex),
    actionCount: optionalNonNegativeNumber(candidate.actionCount),
    actionType:
      typeof candidate.actionType === "string"
        ? truncateProgressText(candidate.actionType, 100)
        : null,
    toolName:
      typeof candidate.toolName === "string"
        ? truncateProgressText(candidate.toolName, 100)
        : null,
    facts: parseProgressFacts(
      (candidate as TaskProgressEvent & { facts?: unknown }).facts
    ),
    target: parseProgressTarget(candidate.target),
    trace,
    bridgeTool: parseBridgeToolPointer(
      (candidate as TaskProgressEvent & { bridgeTool?: unknown }).bridgeTool
    ),
    pythonResult: parseLocalResultCard(candidate.pythonResult),
    pythonCode: parsePythonCodeEvent(
      (candidate as TaskProgressEvent & { pythonCode?: unknown }).pythonCode
    ),
    timestamp: candidate.timestamp
  };
}

const BRIDGE_OPERATION_ID = /^op_[0-9a-f]{32}$/;

/** Rebuild metadata only; never persist arbitrary fields from a progress frame. */
export function parseLocalResultCard(value: unknown): LocalResultCard | null {
  if (!value || typeof value !== "object") return null;
  const card = value as Partial<LocalResultCard>;
  if (typeof card.resultRef !== "string" || !/^res_[0-9a-f]{32}$/.test(card.resultRef) ||
      typeof card.title !== "string" || !card.title || card.title.length > 300 ||
      !Array.isArray(card.headers) || !card.headers.length || card.headers.length > 50 ||
      !card.headers.every((header) => typeof header === "string") ||
      !Number.isInteger(card.rowCount) || card.rowCount! < 0 ||
      !Number.isInteger(card.columnCount) || card.columnCount! < 1 ||
      !Array.isArray(card.sourceSheets) || card.sourceSheets.length > 100 ||
      !card.sourceSheets.every((sheet) => typeof sheet === "string") ||
      typeof card.complete !== "boolean") return null;
  return { resultRef: card.resultRef, title: card.title, headers: [...card.headers],
    rowCount: card.rowCount!, columnCount: card.columnCount!,
    sourceSheets: [...card.sourceSheets], complete: card.complete };
}

/** 严格解析桥指针：形状不符就当没有，绝不把半成品指针交给宿主。 */
export function parseBridgeToolPointer(value: unknown): SdkBridgeToolPointer | null {
  if (value === null || typeof value !== "object") return null;
  const pointer = value as Partial<SdkBridgeToolPointer>;
  if (
    typeof pointer.operationId !== "string" ||
    !BRIDGE_OPERATION_ID.test(pointer.operationId) ||
    (pointer.tool !== "query_table" && pointer.tool !== "prepare_python_data")
  ) {
    return null;
  }
  return { operationId: pointer.operationId, tool: pointer.tool };
}


export interface WorksheetDeleteCheckpointRequest {
  operationId: string;
  host: "excel" | "wps-et";
  workbookPath: string;
  hostWindowId?: number | null;
  authorizedSheets: string[];
  plan: AnalysisPlan;
  /**
   * 发起本次操作的原会话与请求消息。
   *
   * 旧记录没有这两个字段：缺失即"归属未知"，服务端不会把它算进当前会话，
   * 也不会因此拒绝恢复。
   */
  conversationId?: string;
  requestMessageId?: string;
}
export interface WorksheetCheckpointReference { operationId: string; }
export interface WorksheetCheckpointBackupConsent extends WorksheetCheckpointReference { backupSheets: string[]; }
/**
 * 原地撤销的实际结果。
 *
 * 与 `WorksheetCheckpointOpenResult` 是两件事：那边只是"打开了执行前副本"，
 * 这边是"把表放回了当前打开的原工作簿"。`status === "completed"` 才代表
 * 名称、顺序、可见性与内容都核对过；`partial` / `failed` 都不许说成已撤销。
 */
export interface WorksheetRestoreResult {
  operationId: string;
  status: "completed" | "partial" | "failed" | "conflict";
  restoredSheets: string[];
  failedSheets: string[];
  unverifiedSheets: string[];
  repairedFormulas: number;
  repairedNames: number;
  limitations: string[];
  message: string;
}
export interface WorksheetDeleteCheckpointResult {
  operationId: string;
  status: "prepared" | "needs_backup_consent" | "executing" | "completed" | "partial_or_unknown" | "failed";
  backupSheets?: string[];
  deletedSheets: string[];
  remainingSheets: string[];
  recoveryAvailable: boolean;
  message: string;
  /** 最近一次原地撤销结果；缺失只表示"还没撤销过"。 */
  restore?: WorksheetRestoreResult | null;
}
export interface WorksheetCheckpointOpenResult { opened: boolean; recoveryPath: string; }
export interface WorksheetCheckpointListRequest { workbookPath: string; host: "excel" | "wps-et"; }
/** 服务端按工作簿身份列出的恢复点摘要；不含单元格内容。 */
export interface WorksheetCheckpointSummary {
  operationId: string;
  status: "prepared" | "needs_backup_consent" | "executing" | "completed" | "partial_or_unknown" | "failed";
  deletedSheets: string[];
  remainingSheets: string[];
  recoveryAvailable: boolean;
  message: string;
  createdAt: string | null;
  planTitle: string | null;
  restoreStatus?: WorksheetRestoreResult["status"] | null;
}
export interface WorksheetCheckpointListResult { checkpoints: WorksheetCheckpointSummary[]; }


/** Reject incomplete/stale service receipts; a successful HTTP status is not deletion evidence. */
export function assertWorksheetCheckpointResult(value: unknown, operationId: string): asserts value is WorksheetDeleteCheckpointResult {
  const item = value as Partial<WorksheetDeleteCheckpointResult> | null;
  const names = (list: unknown): list is string[] => Array.isArray(list) && list.every(name => typeof name === "string" && name.length > 0) && new Set(list).size === list.length;
  if (!item || item.operationId !== operationId ||
      !["prepared", "needs_backup_consent", "executing", "completed", "partial_or_unknown", "failed"].includes(String(item.status)) ||
      !names(item.deletedSheets) || !names(item.remainingSheets) || typeof item.recoveryAvailable !== "boolean" || typeof item.message !== "string" ||
      (item.backupSheets !== undefined && !names(item.backupSheets)) ||
      (item.status === "needs_backup_consent" && (!names(item.backupSheets) || !item.backupSheets.length || item.recoveryAvailable || item.deletedSheets.length > 0)) ||
      (item.status === "completed" && (!item.recoveryAvailable || !item.deletedSheets.length || !item.remainingSheets.length || item.deletedSheets.some(name => item.remainingSheets!.includes(name))))) {
    throw new Error("工作表检查点回执不完整，执行状态尚未确认；请查询状态，不要重复删除。");
  }
}

/**
 * 撤销回执同样要**先自证**，不能拿 HTTP 200 当证据。
 *
 * 三条硬约束：
 * - 归属必须是这一次执行；
 * - 声称 `completed` 时，已放回的表必须**覆盖**本次删除的全部表，且没有失败/未核对的；
 * - `partial` 必须真的"只完成了一部分"（有放回的、也有没放回的），否则措辞会骗人。
 */
export function assertWorksheetRestoreResult(
  value: unknown,
  operationId: string,
  deletedSheets: readonly string[]
): asserts value is WorksheetRestoreResult {
  const item = value as Partial<WorksheetRestoreResult> | null;
  const names = (list: unknown): list is string[] =>
    Array.isArray(list) && list.every(name => typeof name === "string" && name.length > 0) && new Set(list).size === list.length;
  const statuses = ["completed", "partial", "failed", "conflict"];
  if (!item || item.operationId !== operationId || !statuses.includes(String(item.status)) ||
      !names(item.restoredSheets) || !names(item.failedSheets) || !names(item.unverifiedSheets) ||
      !names(item.limitations) || typeof item.message !== "string" ||
      typeof item.repairedFormulas !== "number" || typeof item.repairedNames !== "number") {
    throw new Error("撤销回执不完整，工作簿状态尚未确认；请核对工作表列表。");
  }
  const covered = deletedSheets.every(name => item.restoredSheets!.includes(name));
  if (item.status === "completed" && (!covered || item.failedSheets!.length || item.unverifiedSheets!.length)) {
    throw new Error("撤销回执声称全部完成，但覆盖范围对不上；请核对工作表列表。");
  }
  if (item.status === "partial" && (!item.restoredSheets!.length || !item.failedSheets!.length)) {
    throw new Error("撤销回执声称部分完成，但清单与状态不一致；请核对工作表列表。");
  }
}

/**
 * 宿主观测的读回形状（A1-E3）。**只读诊断**，不参与任何写入判断。
 *
 * `observation` 是"这批数字能不能当证据"的判据，由服务端给出（只有服务端知道
 * 账本里的动作认领边界）。空数组不等于零写入 —— 判读一律走
 * `describeHostWriteObservation()`，不要自己看 `totals.hostWriteCalls === 0`。
 */
export interface HostWritesSummaryResponse {
  operationId: string;
  entries: HostWriteAudit[];
  totals: HostWriteSummary;
  observation: HostWriteObservation;
}
