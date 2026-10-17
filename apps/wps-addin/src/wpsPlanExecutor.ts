import capabilities from "../../../config/capabilities.json";
import { readWpsRangeValues } from "./wpsRangeValues";
import { writeWpsMatrix } from "./wpsWriteMatrix";
import type { WorkbookLocator } from "@excel-bro/core/contracts";
import {
  buildRangeChangeReview,
  type ChangeReviewStatus,
  type ChangeSnapshot,
  type RangeChangeReview
} from "@excel-bro/core/changeReview";
import type { AnalysisPlan, CellValue, ExcelAction } from "../../excel-addin/src/contracts";
import {
  assertRectangularValues,
  blankCellMatrix,
  parseExplicitA1Rect,
  writeMatrixShapeIssue
} from "../../excel-addin/src/planCore";
import { assertSafeFormula } from "../../excel-addin/src/formulaSafety";
import type { WpsApplication, WpsRange, WpsWorksheet } from "./wpsHost";
import {
  assertWpsAuthorization,
  authorizeWpsSheets,
  normalizeWpsMatrix,
  type WpsReadAuthorization
} from "./wpsWorkbookReader";
import { wpsWriteVersionStatus } from "./wpsWriteVersionGate";
import {
  assertWpsVerifiedSetAlignment,
  assertWpsVerifiedResizeRange,
  wpsColorFromHex,
  wpsHorizontalAlignment
} from "./wpsFormatCapability";
import {
  captureWpsFormatSnapshot,
  restoreWpsFormatSnapshot,
  sameWpsFormatSnapshot,
  type WpsFormatField,
  type WpsFormatSnapshot
} from "./wpsFormatSnapshot";

const STORAGE_KEY = "excel-bro.wps.undo.v1";

// 双栏对照的规模上限只影响"展示与采集"，不影响写入与撤销的正确性。
// WPS 自有 wpsWrite.maxCells（当前 1000，见 prepareWpsPlan）远小于此处，
// 这里是配置被调大之后的第二道闸：超限就不再采集显示值，只留公式判定改动。
const CHANGE_REVIEW_CAPTURE_MAX_CELLS = capabilities.changeReview.captureMaxCells;
const CHANGE_REVIEW_LIMIT = {
  maxRows: capabilities.changeReview.maxRows,
  maxTotalCells: capabilities.changeReview.maxTotalCells
};

type SupportedAction =
  | Extract<ExcelAction, { type: "writeValues" }>
  | Extract<ExcelAction, { type: "writeFormulas" }>
  | Extract<ExcelAction, { type: "clearRange" }>
  | Extract<ExcelAction, { type: "setFont" }>
  | Extract<ExcelAction, { type: "setAlignment" }>
  | Extract<ExcelAction, { type: "resizeRange" }>;

/** 值动作（writeValues / writeFormulas / clearRange）—— 走值矩阵路径。 */
type ValueAction = Extract<SupportedAction, { type: "writeValues" | "writeFormulas" | "clearRange" }>;

/** 格式动作（setFont / setAlignment / resizeRange）—— 走格式快照路径。 */
type FormatAction = Extract<SupportedAction, { type: "setFont" | "setAlignment" | "resizeRange" }>;

/** 格式动作在 StoredAction 里可用的字段形状（不含 sheet/range 定位字段，那些单独传）。 */
export type FormatActionFields =
  | { type: "setFont"; bold?: boolean | null; italic?: boolean | null; name?: string | null; size?: number | null; color?: string | null }
  | { type: "setAlignment"; horizontal?: string | null; vertical?: string | null; wrapText?: boolean | null }
  | { type: "resizeRange"; rowHeight?: number | null; columnWidth?: number | null };

/** 从 ExcelAction 格式动作提取 FormatActionFields（prepareFormatAction 用）。 */
function asFormatFieldsFromExcelAction(
  action: Extract<ExcelAction, { type: "setFont" | "setAlignment" | "resizeRange" }>
): FormatActionFields {
  switch (action.type) {
    case "setFont":
      return { type: "setFont", bold: action.bold, italic: action.italic, name: action.name, size: action.size, color: action.color };
    case "setAlignment":
      return { type: "setAlignment", horizontal: action.horizontal, vertical: action.vertical, wrapText: action.wrapText };
    case "resizeRange":
      return { type: "resizeRange", rowHeight: action.rowHeight, columnWidth: action.columnWidth };
  }
}

/**
 * 把 prepared/stored 动作收窄成格式动作字段形状（供 apply/回读/恢复/副本审阅使用）。
 * 优先从 `formatAction` 字段读取（`prepareFormatAction` 已把具体字段存进去）。
 */
export function asFormatFields(action: { formatAction?: FormatActionFields } | StoredAction): FormatActionFields {
  return action.formatAction ?? asFormatFieldsFromExcelAction(action as unknown as Extract<ExcelAction, { type: "setFont" | "setAlignment" | "resizeRange" }>);
}

/**
 * 执行器**真正**能跑完闭环（预览快照 / 执行 / 回读 / 撤销）的动作键。
 *
 * 之所以要有这份运行期清单，而不只靠类型：`capabilities.wpsWrite.allowedActions`
 * 是**配置**，`supportedAction()` 是**类型**，两者之间没有任何东西保证一致。
 * 曾经的风险形态就是"配置里放开了某个动作、执行器其实不认"——它在运行期
 * 表现成"计划被静默拒绝"，而配置读起来却像是已接通。
 *
 * `assertAllowedAction` 与漂移检测测试都认这一份，`clearRange` 按 `applyTo` 细分。
 *
 * 格式动作（setFont/setAlignment/resizeRange）按**已实机验证的属性子集**开放：
 * 动作名进了这份清单只说明"执行器认这个动作类型"，属性级守卫在
 * `assertWpsVerifiedSetFont`/`assertWpsVerifiedSetAlignment`/`assertWpsVerifiedResizeRange`
 * 与 `prepareWpsPlan` 预检里兜住（underline/vertical/columnWidth 等未验证属性明确拒绝）。
 */
export const WPS_SUPPORTED_ACTION_KEYS: readonly string[] = [
  "writeValues",
  "writeFormulas",
  "clearRange:contents",
  "setFont",
  "setAlignment",
  "resizeRange"
];

type Rect = {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
};

export type WpsPreparedAction = {
  actionIndex: number;
  type: "writeValues" | "writeFormulas" | "clearRange" | "setFont" | "setAlignment" | "resizeRange";
  sheetIndex: number;
  sheetName: string;
  range: string;
  rowCount: number;
  columnCount: number;
  cellCount: number;
  values: CellValue[][];
  preview: CellValue[][];
  /**
   * 预览生成那一刻、目标区域在宿主里的**真实原值**（与 `previewFingerprint` 同源同一次读取）。
   *
   * 只活在内存里，供确认卡首屏的「修改前」列使用 —— 让用户在写入前看到"到底改了什么"。
   * 出站闸 `wpsRecoveryDescriptor` 是**白名单投影**，不会取这两个字段，所以它
   * 没有路径进入登记请求或落盘。
   */
  beforeValues: CellValue[][];
  /** false = 宿主这次没交出 `Value2`，确认卡必须标"原值未取得"，不得拿公式顶替。 */
  beforeValuesCaptured: boolean;
  /** Target state when the user was shown this preview. */
  previewFingerprint: string;
  rect: Rect;
  /**
   * 格式动作（setFont/setAlignment/resizeRange）专属：预览时采集的**逐格/逐行格式快照**。
   * 值动作（writeValues 等）没有此字段。
   */
  formatSnapshot?: WpsFormatSnapshot;
  /** 格式动作的具体字段（bold/horizontal/rowHeight 等），供执行/回读/恢复使用。 */
  formatAction?: FormatActionFields;
};

/** Kept in memory only: source values are never sent to the model or persisted as confirmation metadata. */
export type WpsSourceSnapshot = {
  sheetName: string;
  range: string;
  values: CellValue[][];
};

type WpsPreparedSource = {
  sheetIndex: number;
  sheetName: string;
  range: string;
  rowCount: number;
  columnCount: number;
  previewFingerprint: string;
};

export type WpsPreparedPlan = {
  plan: AnalysisPlan;
  /** In-memory only; needed to recheck the original calculation input after a preview edit. */
  sourceSnapshots: WpsSourceSnapshot[];
  authorization: WpsReadAuthorization;
  workbookIdentity: string;
  hostVersion: string | null;
  actions: WpsPreparedAction[];
  sources: WpsPreparedSource[];
  totalCells: number;
  previewRevision: string;
  confirmationSnapshot: string;
};

export type StoredAction = WpsPreparedAction & {
  beforeValues: CellValue[][];
  beforeFormulas: CellValue[][];
  beforeFingerprint: string;
  afterFingerprint: string | null;
  appliedRows: number;
  /** 格式动作的写前格式快照（执行前采集，撤销用）。值动作无此字段。 */
  beforeFormatSnapshot?: WpsFormatSnapshot;
  /**
   * 格式动作**整份计划执行后**的最终格式快照（按本动作区域采集）。值动作无此字段。
   *
   * 撤销冲突检测按**最终恢复影响范围**比对：本动作声明过的属性，现场是否仍等于
   * 计划整体产出的最终格式。**不能**用"本动作声明的目标值"—— 同属性后续覆盖
   * （字号先 12→14 再 14→16）时第一步声明的 14 早被覆盖成 16，按目标值判会误报冲突。
   */
  afterFormatSnapshot?: WpsFormatSnapshot;
};

/** 一次动作写完之后回读到的整块结果。只活在内存里 —— 不进 `WpsUndoRecord`（那个会落 PluginStorage）。 */
export type WpsActionAfter = {
  values: CellValue[][];
  formulas: CellValue[][];
  /** false = 宿主这次没把 Value2 暴露出来，双栏只能显示"显示值未采集"。 */
  valuesCaptured: boolean;
};

export type WpsUndoRecord = {
  version: 1 | 2;
  executionId: string;
  planId: string;
  planTitle: string;
  createdAt: string;
  workbookIdentity: string;
  authorizedSheetNames: string[];
  state: "ready" | "applying" | "verified" | "partial_write_recoverable" | "undoing" | "restored" | "undo_conflict";
  actions: StoredAction[];
  error: string | null;
};

export type WpsExecutionProgress = {
  stage: "snapshot" | "writing" | "verifying" | "undoing";
  message: string;
  actionIndex?: number;
  actionType?: WpsPreparedAction["type"];
  sheetName?: string;
  range?: string;
  completedRows?: number;
  totalRows?: number;
  actionCount?: number;
  cellCount?: number;
};

function columnName(index: number): string {
  let value = index;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function parseA1Range(address: string): Rect {
  const rect = parseExplicitA1Rect(address);
  if (!rect) throw new Error(`WPS 首版只支持明确的 A1 矩形区域，无法执行「${address}」。`);
  return rect;
}

function rangeAddress(rect: Rect, offsetRows = 0, rowCount?: number): string {
  const startRow = rect.startRow + offsetRows;
  const endRow = rowCount == null ? rect.endRow : startRow + rowCount - 1;
  const first = `${columnName(rect.startColumn)}${startRow}`;
  const last = `${columnName(rect.endColumn)}${endRow}`;
  return first === last ? first : `${first}:${last}`;
}

function rangeFor(sheet: WpsWorksheet, address: string): WpsRange {
  const range = sheet.Range?.(address);
  if (!range) throw new Error(`WPS 无法读取区域「${sheet.Name}!${address}」。`);
  return range;
}

function canonicalCell(value: unknown): CellValue {
  if (value == null || value === "") return null;
  if (typeof value === "string" || typeof value === "boolean") return value;
  if (typeof value === "number") return Number.isFinite(value) ? value : String(value);
  return String(value);
}

function canonicalMatrix(values: readonly (readonly unknown[])[]): CellValue[][] {
  return values.map((row) => row.map(canonicalCell));
}

function legacyMatrixFingerprint(values: CellValue[][], formulas: CellValue[][]): string {
  const normalizedValues = canonicalMatrix(values);
  const normalizedFormulas = canonicalMatrix(formulas);
  const logicalCells = normalizedValues.map((row, rowIndex) =>
    row.map((value, columnIndex) => {
      const formula = normalizedFormulas[rowIndex]?.[columnIndex];
      return typeof formula === "string" && formula.startsWith("=")
        ? { formula }
        : { value };
    })
  );
  const text = JSON.stringify(logicalCells);
  let hash = 0x811c9dc5;
  for (let index = 0; index < text.length; index += 1) {
    hash ^= text.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return `${text.length}:${(hash >>> 0).toString(16).padStart(8, "0")}`;
}

function typedCell(value: CellValue): readonly [string, CellValue?] {
  if (value === null) return ["null"];
  return [typeof value, value];
}

/** v2 keeps null and empty string distinct and binds a formula ahead of its displayed value. */
export function matrixFingerprintV2(values: CellValue[][], formulas: CellValue[][]): string {
  const logicalCells = values.map((row, rowIndex) =>
    row.map((value, columnIndex) => {
      const formula = formulas[rowIndex]?.[columnIndex];
      return typeof formula === "string" && formula.startsWith("=")
        ? ["formula", formula]
        : typedCell(value);
    })
  );
  const text = JSON.stringify(logicalCells);
  let hash = 0x811c9dc5;
  for (let index = 0; index < text.length; index += 1) {
    hash ^= text.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return `v2:${text.length}:${(hash >>> 0).toString(16).padStart(8, "0")}`;
}

function fingerprintForVersion(values: CellValue[][], formulas: CellValue[][], version: 1 | 2): string {
  return version === 1
    ? legacyMatrixFingerprint(values, formulas)
    : matrixFingerprintV2(values, formulas);
}

function readMatrices(range: WpsRange, rows: number, columns: number, version: 1 | 2 = 2) {
  const { values, valuesCaptured } = readWpsRangeValues(range, rows, columns);
  const formulas = normalizeWpsMatrix(range.Formula, rows, columns);
  return {
    values,
    formulas,
    fingerprint: fingerprintForVersion(values, formulas, version),
    // 宿主没把 Value2 交出来时不能拿公式顶替：`normalizeWpsMatrix` 会把 undefined
    // 归一成 null，看上去跟"空单元格"一模一样。这里如实记一笔，交给双栏显示。
    valuesCaptured
  };
}

/**
 * `range.Formula` 对常量单元格返回的是常量本身（"123"、"abc"），只有 `=` 开头才是
 * 真公式 —— 与 `legacyMatrixFingerprint` / `matrixFingerprintV2` 的口径一致，
 * 双栏不得另立一套，否则同一格在两处会有不同身份。
 */
export function toChangeFormulas(matrix: CellValue[][]): (string | null)[][] {
  return matrix.map((row) =>
    row.map((cell) => (typeof cell === "string" && cell.startsWith("=") ? cell : null))
  );
}

/**
 * 由「本次写前快照 + 本次写后回读」生成一份双栏对照。纯函数，不碰宿主。
 *
 * `matched` 三态：true = 回读与计划一致；false = 不一致；null = 根本没走到比对
 * （写了但中途抛错）。三者都不能算"已核对通过"，只有 true 能得到 `applied`。
 */
export function wpsActionChangeReview(
  action: StoredAction,
  after: WpsActionAfter | null,
  matched: boolean | null
): RangeChangeReview {
  const status: ChangeReviewStatus =
    action.appliedRows === 0
      ? "not_executed"
      : matched === true
        ? "applied"
        : matched === false
          ? "readback_mismatch"
          : "readback_failed";
  const capture = action.rowCount * action.columnCount <= CHANGE_REVIEW_CAPTURE_MAX_CELLS;
  const before: ChangeSnapshot = {
    values: capture
      ? action.beforeValues
      : action.beforeValues.map((row) => row.map(() => null)),
    formulas: toChangeFormulas(action.beforeFormulas),
    valuesCaptured: capture
  };
  // 没有 after（压根没写、或没读成）时**不**用空矩阵冒充"写成了空"：
  // 整块标成未采集，界面只会看到"显示值未采集"。
  const afterSnapshot: ChangeSnapshot =
    after && capture
      ? {
          values: after.values,
          formulas: toChangeFormulas(after.formulas),
          valuesCaptured: after.valuesCaptured
        }
      : { values: [], formulas: [], valuesCaptured: false };
  return buildRangeChangeReview({
    sheet: action.sheetName,
    range: action.range,
    phase: "actual",
    status,
    before,
    after: afterSnapshot,
    appliedRows: action.appliedRows,
    limit: CHANGE_REVIEW_LIMIT
  });
}

function hasAmbiguousLogicalCell(values: CellValue[][], formulas: CellValue[][]): boolean {
  return values.some((row, rowIndex) => row.some((value, columnIndex) => {
    const formula = formulas[rowIndex]?.[columnIndex];
    return !(typeof formula === "string" && formula.startsWith("=")) && (value === null || value === "");
  }));
}

function confirmationSnapshot(prepared: WpsPreparedPlan): string {
  return JSON.stringify({
    plan: prepared.plan,
    sourceSnapshots: prepared.sourceSnapshots,
    sheets: prepared.authorization.sheets.map(({ index, name }) => ({ index, name })),
    workbookIdentity: prepared.workbookIdentity,
    hostVersion: prepared.hostVersion,
    actions: prepared.actions,
    sources: prepared.sources,
    totalCells: prepared.totalCells,
    previewRevision: prepared.previewRevision
  });
}

/** A staged native comparison must not bypass the preview's immutable confirmation snapshot. */
export function assertWpsPreparedPlanUnchanged(prepared: WpsPreparedPlan): void {
  if (confirmationSnapshot(prepared) !== prepared.confirmationSnapshot) {
    throw new Error("写入计划或预览已变化，请重新预览并确认后再执行。");
  }
}

function matricesEqual(actual: CellValue[][], expected: CellValue[][]): boolean {
  // Readback must preserve the plan's cell values. The legacy fingerprint above
  // intentionally folds empty strings for existing undo records, but that
  // compatibility rule must not make a failed write look verified.
  return JSON.stringify(actual) === JSON.stringify(expected);
}

function overlaps(left: WpsPreparedAction, right: WpsPreparedAction): boolean {
  if (left.sheetName !== right.sheetName) return false;
  const geometric = !(
    left.rect.endRow < right.rect.startRow ||
    right.rect.endRow < left.rect.startRow ||
    left.rect.endColumn < right.rect.startColumn ||
    right.rect.endColumn < left.rect.startColumn
  );
  if (!geometric) return false;
  // 两个都是格式动作时允许同区域重叠：它们改的是（可能不同的）格式字段，
  // 属性级恢复（`declaredFormatFields`）各自恢复各自的字段，不会互相覆盖。
  // 值动作重叠、值动作 × 格式动作重叠仍拒绝（值动作会整块覆盖区域）。
  if (isFormatActionType(left.type) && isFormatActionType(right.type)) return false;
  return true;
}

function workbookIdentity(authorization: WpsReadAuthorization): string {
  const identity = authorization.workbookFullName?.trim();
  if (!identity || !/^[A-Za-z]:\\[^\r\n]+\.xlsx$/i.test(identity)) {
    throw new Error("WPS 写入只对已保存到本地磁盘的 .xlsx 工作簿开放。请先保存文件。");
  }
  return identity.toLowerCase();
}

/**
 * WPS 身份的共享层内部表示（方案 §2.3 第 1 条）。
 *
 * 值沿用 `workbookIdentity` 的既有归一（仅小写，**不**把反斜杠转正斜杠）——
 * 已落盘的 `WpsUndoRecord.workbookIdentity` 就是这个形态，改动会让现有撤销记录
 * 对不上。core 的 `sameLocator` 比对 wps-path 时对两种分隔符都归一，故不受影响。
 *
 * 这个 locator **不会**出网：WPS 恢复按方案 §2.3 第 2 条保持客户端权威，
 * 服务端不参与，也不存在折叠成 `url:` 字符串的路径。
 */
export function wpsWorkbookLocator(
  authorization: WpsReadAuthorization
): WorkbookLocator {
  return { kind: "wps-path", value: workbookIdentity(authorization) };
}

function selectedSheet(
  application: WpsApplication,
  authorization: WpsReadAuthorization,
  sheetName: string
): { index: number; sheet: WpsWorksheet } {
  const authorized = authorization.sheets.find((sheet) => sheet.name === sheetName);
  if (!authorized) throw new Error(`计划引用了未授权工作表「${sheetName}」，已拒绝整份计划。`);
  return {
    index: authorized.index,
    sheet: assertWpsAuthorization(application, authorization, authorized)
  };
}

function valuesForAction(action: ValueAction, rows: number, columns: number): CellValue[][] {
  if (action.type === "clearRange") {
    if (action.applyTo !== "contents" || (action.filters?.length ?? 0) > 0) {
      throw new Error("WPS 当前只支持清空明确区域的内容，不支持清除格式、超链接或带筛选条件的删除。");
    }
    return blankCellMatrix(rows, columns);
  }
  if (action.type === "writeFormulas") {
    if (action.formulaR1C1 !== undefined) {
      throw new Error("WPS 的 R1C1 铺满公式尚未完成独立回读验收；请使用与目标区域同形的 A1 公式矩阵。");
    }
    if (!Array.isArray(action.formulas) || action.formulas.some((row) => !Array.isArray(row))) {
      throw new Error(`区域「${action.sheet}!${action.range}」的公式不是二维数组。`);
    }
    assertRectangularValues(action.formulas);
    const shapeIssue = writeMatrixShapeIssue(action.range, action.formulas);
    if (shapeIssue) throw new Error(shapeIssue);
    action.formulas.forEach((row, rowIndex) => row.forEach((formula, columnIndex) => {
      if (typeof formula !== "string" || !formula.startsWith("=")) {
        throw new Error(`区域「${action.sheet}!${action.range}」第 ${rowIndex + 1} 行第 ${columnIndex + 1} 列不是 A1 公式。`);
      }
      assertSafeFormula(formula, `${action.sheet}!${action.range} 第 ${rowIndex + 1} 行第 ${columnIndex + 1} 列`);
    }));
    return action.formulas.map((row) => [...row]);
  }
  if (!Array.isArray(action.values) || action.values.some((row) => !Array.isArray(row))) {
    throw new Error(`区域「${action.sheet}!${action.range}」的写入数据不是二维数组。`);
  }
  assertRectangularValues(action.values);
  const shapeIssue = writeMatrixShapeIssue(action.range, action.values);
  if (shapeIssue) throw new Error(shapeIssue);
  const values = action.values.map((row) => row.map((value) =>
    value === "" ? "" : canonicalCell(value)
  ));
  if (values.some((row) => row.some((value) => typeof value === "string" && value.startsWith("=")))) {
    throw new Error("WPS 不允许通过 writeValues 注入公式；请使用 writeFormulas 动作并预览确认。");
  }
  return values;
}

function supportedAction(action: ExcelAction): action is SupportedAction {
  return (
    action.type === "writeValues" ||
    action.type === "writeFormulas" ||
    action.type === "clearRange" ||
    action.type === "setFont" ||
    action.type === "setAlignment" ||
    action.type === "resizeRange"
  );
}

function assertAllowedAction(action: SupportedAction): void {
  const key = action.type === "clearRange" ? `${action.type}:${action.applyTo}` : action.type;
  // 两道闸，缺一不可：① 执行器真的实现了这个动作；② 配置真的开放了它。
  // 只查配置会让"配置放开了但执行器不认"溜过去，只查实现会让"实现好了但还没
  // 验收就对外放开"溜过去。
  if (!WPS_SUPPORTED_ACTION_KEYS.includes(key)) {
    throw new Error(`WPS 当前未实现「${key}」，已拒绝整份计划。`);
  }
  if (!(capabilities.wpsWrite.allowedActions as readonly string[]).includes(key)) {
    throw new Error(`WPS 当前未开放「${key}」，已拒绝整份计划。`);
  }
}

/** 是否为格式动作（走格式快照路径，而非值矩阵路径）。 */
export function isFormatActionType(type: string): type is "setFont" | "setAlignment" | "resizeRange" {
  return type === "setFont" || type === "setAlignment" || type === "resizeRange";
}

/**
 * 格式动作的预检 + 逐格/逐行快照采集。
 *
 * 与值动作不同：格式动作**不改值/公式**，但为了走通副本审阅（`stageWpsNativeComparison`
 * 会读值矩阵与 `previewFingerprint` 比对"值没被外部改过"），这里仍采集目标区域的
 * **真实值矩阵**作为 `values`/`beforeValues`/`previewFingerprint`。值在格式动作里
 * 保持不变（`values === beforeValues`），回读与撤销的**格式维度**依据是
 * `formatSnapshot`（逐格字体/对齐/换行 + 逐行行高）。
 */
function prepareFormatAction(
  action: Extract<ExcelAction, { type: "setFont" | "setAlignment" | "resizeRange" }>,
  actionIndex: number,
  sheetIndex: number,
  sheet: WpsWorksheet,
  hostRange: WpsRange,
  targetRange: string,
  rect: Rect,
  rowCount: number,
  columnCount: number
): WpsPreparedAction {
  // 属性级守卫：未验证属性明确拒绝（不静默忽略）。
  if (action.type === "setAlignment") {
    assertWpsVerifiedSetAlignment(action);
  } else if (action.type === "resizeRange") {
    assertWpsVerifiedResizeRange(action);
  }

  // 逐格/逐行格式快照。采集不完整（任一目标属性读不到）→ 拒绝，不产生不可撤销的修改。
  const formatSnapshot = captureWpsFormatSnapshot(hostRange, rowCount, columnCount);
  if (!formatSnapshot.captured) {
    throw new Error(
      `区域「${action.sheet}!${targetRange}」的格式无法完整逐格采集（宿主未返回全部属性），已拒绝整份计划。`
    );
  }

  // 值矩阵：格式动作不改值，但需要真实值作为副本审阅的"值没变"基线。
  if (hostRange.Formula === undefined) {
    throw new Error(`区域「${action.sheet}!${targetRange}」无法生成公式快照。`);
  }
  const beforePreview = readMatrices(hostRange, rowCount, columnCount);
  const formatAction = asFormatFieldsFromExcelAction(action);
  return {
    actionIndex,
    type: action.type,
    sheetIndex,
    sheetName: action.sheet,
    range: targetRange,
    rowCount,
    columnCount,
    cellCount: rowCount * columnCount,
    values: beforePreview.values,
    preview: beforePreview.values.slice(0, 5).map((row) => row.slice(0, 5)),
    beforeValues: beforePreview.values,
    beforeValuesCaptured: beforePreview.valuesCaptured,
    previewFingerprint: beforePreview.fingerprint,
    rect,
    formatSnapshot,
    formatAction
  };
}

export function prepareWpsPlan(
  application: WpsApplication,
  selectedSheetIndexes: readonly number[],
  plan: AnalysisPlan,
  sourceSnapshots: readonly WpsSourceSnapshot[] = []
): WpsPreparedPlan {
  if (!capabilities.wpsWrite.enabled) throw new Error("WPS 写入功能当前已由能力开关关闭。");
  const hostVersion = String(application.Version ?? "").trim();
  if (wpsWriteVersionStatus(hostVersion) === "blocked") {
    throw new Error(`当前 WPS JSAPI 版本「${hostVersion || "未知"}」尚未开放写入。`);
  }
  const workbook = application.ActiveWorkbook;
  if (!workbook || workbook.ReadOnly !== false) throw new Error("当前 WPS 工作簿没有明确处于可编辑状态。");
  if (!plan.actions.length) throw new Error("计划中没有可执行操作。");
  const authorization = authorizeWpsSheets(application, selectedSheetIndexes);
  const identity = workbookIdentity(authorization);
  const actions: WpsPreparedAction[] = plan.actions.map((action, actionIndex) => {
    if (!supportedAction(action)) {
      throw new Error(`WPS 当前不执行「${action.type}」。请使用当前已开放的写入动作。`);
    }
    assertAllowedAction(action);
    const targetRange = action.range;
    const rect = parseA1Range(targetRange);
    const rowCount = rect.endRow - rect.startRow + 1;
    const columnCount = rect.endColumn - rect.startColumn + 1;
    const { index, sheet } = selectedSheet(application, authorization, action.sheet);
    if (sheet.ProtectContents !== false) throw new Error(`工作表「${action.sheet}」受保护或保护状态未知。`);
    const hostRange = rangeFor(sheet, targetRange);
    if (hostRange.MergeCells !== false) throw new Error(`区域「${action.sheet}!${targetRange}」包含合并单元格或状态未知。`);

    // 格式动作：属性级守卫 + 逐格/逐行格式快照，不采集值矩阵。
    if (action.type === "setFont" || action.type === "setAlignment" || action.type === "resizeRange") {
      return prepareFormatAction(action, actionIndex, index, sheet, hostRange, targetRange, rect, rowCount, columnCount);
    }

    if (hostRange.Formula === undefined) throw new Error(`区域「${action.sheet}!${targetRange}」无法生成公式快照。`);
    const beforePreview = readMatrices(hostRange, rowCount, columnCount);
    const values = valuesForAction(action, rowCount, columnCount);
    return {
      actionIndex,
      type: action.type,
      sheetIndex: index,
      sheetName: action.sheet,
      range: targetRange,
      rowCount,
      columnCount,
      cellCount: rowCount * columnCount,
      values,
      preview: values.slice(0, 5).map((row) => row.slice(0, 5)),
      // 与 `previewFingerprint` 来自同一次 `readMatrices`（零额外宿主读取）。
      beforeValues: beforePreview.values,
      beforeValuesCaptured: beforePreview.valuesCaptured,
      previewFingerprint: beforePreview.fingerprint,
      rect
    };
  });
  const sources: WpsPreparedSource[] = sourceSnapshots.map((snapshot) => {
    const rect = parseA1Range(snapshot.range);
    const rowCount = rect.endRow - rect.startRow + 1;
    const columnCount = rect.endColumn - rect.startColumn + 1;
    if (snapshot.values.length !== rowCount || snapshot.values.some((row) => row.length !== columnCount)) {
      throw new Error(`源区域「${snapshot.sheetName}!${snapshot.range}」的读取矩阵尺寸不符。`);
    }
    const { index, sheet } = selectedSheet(application, authorization, snapshot.sheetName);
    const current = readMatrices(rangeFor(sheet, snapshot.range), rowCount, columnCount);
    const emptyFormulas = blankCellMatrix(rowCount, columnCount);
    if (matrixFingerprintV2(current.values, emptyFormulas) !== matrixFingerprintV2(snapshot.values, emptyFormulas)) {
      throw new Error(`源区域「${snapshot.sheetName}!${snapshot.range}」在计算后发生变化，请重新生成计划。`);
    }
    return {
      sheetIndex: index,
      sheetName: snapshot.sheetName,
      range: snapshot.range,
      rowCount,
      columnCount,
      previewFingerprint: current.fingerprint
    };
  });
  for (let left = 0; left < actions.length; left += 1) {
    for (let right = left + 1; right < actions.length; right += 1) {
      if (overlaps(actions[left], actions[right])) {
        throw new Error("同一计划中存在重叠的 WPS 写入区域，已拒绝整份计划。");
      }
    }
  }
  const totalCells = actions.reduce((sum, action) => sum + action.cellCount, 0);
  if (totalCells > capabilities.wpsWrite.maxCells) {
    throw new Error(`本次计划涉及 ${totalCells} 个单元格，超过 WPS 单次 ${capabilities.wpsWrite.maxCells} 个单元格的上限。`);
  }
  const prepared: WpsPreparedPlan = {
    plan,
    sourceSnapshots: sourceSnapshots.map((snapshot) => ({
      sheetName: snapshot.sheetName,
      range: snapshot.range,
      values: snapshot.values.map((row) => [...row])
    })),
    authorization,
    workbookIdentity: identity,
    hostVersion,
    actions,
    sources,
    totalCells,
    previewRevision: randomId(),
    confirmationSnapshot: ""
  };
  prepared.confirmationSnapshot = confirmationSnapshot(prepared);
  return prepared;
}

function randomId(): string {
  return globalThis.crypto?.randomUUID?.() ?? `wps-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function readRecords(application: WpsApplication): WpsUndoRecord[] {
  const raw = application.PluginStorage.getItem(STORAGE_KEY);
  if (!raw) return [];
  try {
    const value = typeof raw === "string" ? JSON.parse(raw) : raw;
    return Array.isArray(value) ? value as WpsUndoRecord[] : [];
  } catch {
    return [];
  }
}

function writeRecords(application: WpsApplication, records: WpsUndoRecord[]): void {
  application.PluginStorage.setItem(STORAGE_KEY, JSON.stringify(records.slice(0, capabilities.undo.maxEntries)));
}

function upsertRecord(application: WpsApplication, record: WpsUndoRecord): void {
  const records = readRecords(application).filter((item) => item.executionId !== record.executionId);
  writeRecords(application, [record, ...records]);
}

function getStoredSheet(
  application: WpsApplication,
  record: WpsUndoRecord,
  action: StoredAction
): WpsWorksheet {
  const workbook = application.ActiveWorkbook;
  const fullName = String(workbook?.FullName ?? "").trim().toLowerCase();
  if (!workbook || fullName !== record.workbookIdentity) throw new Error("当前工作簿与恢复点不一致，已停止操作。");
  const sheet = workbook.Worksheets.Item(action.sheetIndex);
  if (!sheet || String(sheet.Name) !== action.sheetName) throw new Error(`工作表「${action.sheetName}」身份已变化，已停止操作。`);
  return sheet;
}

function setRangeValues(range: WpsRange, values: CellValue[][]): void {
  writeWpsMatrix(range, values);
}

function setRangeFormulas(range: WpsRange, formulas: CellValue[][]): void {
  writeWpsMatrix(range, formulas, "Formula");
}

/* ------------------------------------------------------------------------- *
 * 格式动作：执行 / 回读判定 / 撤销恢复（Task #53）
 * ------------------------------------------------------------------------- */

/** 把格式动作写到整块区域（逐格写字体/对齐/换行；resizeRange 逐行写行高）。 */
export function applyFormatAction(
  range: WpsRange,
  action: FormatActionFields
): void {
  if (action.type === "setFont") {
    for (let r = 0; r < (range.Rows?.Count ?? 0); r += 1) {
      for (let c = 0; c < (range.Columns?.Count ?? 0); c += 1) {
        const cell = range.Item ? range.Item(r + 1, c + 1) : range;
        const font = cell.Font;
        if (!font) continue;
        if (action.bold != null) font.Bold = action.bold;
        if (action.italic != null) font.Italic = action.italic;
        if (action.name != null) font.Name = action.name;
        if (action.size != null) font.Size = action.size;
        if (action.color != null) font.Color = wpsColorFromHex(action.color);
      }
    }
    return;
  }
  if (action.type === "setAlignment") {
    for (let r = 0; r < (range.Rows?.Count ?? 0); r += 1) {
      for (let c = 0; c < (range.Columns?.Count ?? 0); c += 1) {
        const cell = range.Item ? range.Item(r + 1, c + 1) : range;
        if (action.horizontal != null) cell.HorizontalAlignment = wpsHorizontalAlignment(action.horizontal);
        if (action.wrapText != null) cell.WrapText = action.wrapText;
      }
    }
    return;
  }
  // resizeRange：只开放 rowHeight（columnWidth 已在属性级守卫拒绝）。
  if (action.rowHeight != null) {
    for (let r = 0; r < (range.Rows?.Count ?? 0); r += 1) {
      const cell = range.Item ? range.Item(r + 1, 1) : range;
      cell.RowHeight = action.rowHeight;
    }
  }
}

/**
 * 提取格式动作**实际改动**的属性名（非 null 字段）—— 撤销时只恢复这些属性，
 * 从而**保留用户对其他属性的后续修改**（Task #53 阻断点 2：属性级恢复）。
 *
 * 这也是**可恢复属性白名单**的唯一来源：恢复账本（`wpsExecutionLedger`）投影出的
 * `formatFields` 就是它，服务端按同一份名单校验。多一份平行名单就会漂移。
 */
export function declaredFormatFields(action: FormatActionFields): WpsFormatField[] {
  const fields: WpsFormatField[] = [];
  if (action.type === "setFont") {
    if (action.bold != null) fields.push("bold");
    if (action.italic != null) fields.push("italic");
    if (action.name != null) fields.push("name");
    if (action.size != null) fields.push("size");
    if (action.color != null) fields.push("color");
  } else if (action.type === "setAlignment") {
    if (action.horizontal != null) fields.push("horizontalAlignment");
    if (action.wrapText != null) fields.push("wrapText");
  } else if (action.type === "resizeRange") {
    if (action.rowHeight != null) fields.push("rowHeight");
  }
  return fields;
}

/* ------------------------------------------------------------------------- *
 * 格式动作的**状态证据指纹**（恢复账本用；Task #61/#62）
 * ------------------------------------------------------------------------- */

/** 格式属性的规范化取值：布尔 / 数字 / 字符串 / null。**不含任何单元格内容**。 */
type WpsFormatScalar = boolean | number | string | null;

function formatFingerprintOf(payload: unknown[]): string {
  const text = JSON.stringify(payload);
  let hash = 0x811c9dc5;
  for (let index = 0; index < text.length; index += 1) {
    hash ^= text.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return `f1:${text.length}:${(hash >>> 0).toString(16).padStart(8, "0")}`;
}

/**
 * 一份格式快照在**已声明属性**上的确定性指纹（`f1:` 前缀）。
 *
 * 为什么不能沿用值矩阵指纹：格式动作**不改值**，`matrixFingerprintV2` 在写前与写后
 * 完全一样，对"到底改没改到"没有区分力；而恢复判定必须能区分
 * **写前 / 写后 / 都不是（= 被外部改过）**三种现场。
 *
 * 这里只取 `fields` 覆盖到的属性，逐格（行高逐行）取值后用**同一个**序列化口径出
 * 哈希，所以同一份属性声明在三种现场上给出三个可区分的结果。指纹**不是加密**，
 * 也不承诺抵御低熵枚举 —— 与 `matrixFingerprintV2` 同一条口径。
 *
 * 出站的只有这个哈希：快照本身含逐格属性值，属于本机现场，不上传、不落账本。
 */
export function wpsFormatEvidenceFingerprint(
  type: string,
  fields: readonly WpsFormatField[],
  snapshot: WpsFormatSnapshot
): string {
  const payload: unknown[] = [type, [...fields]];
  for (const field of fields) {
    if (field === "rowHeight") {
      payload.push(snapshot.rowHeights.map((height) => height));
      continue;
    }
    payload.push(snapshot.cells.map((row) => row.map((cell) => cell[field])));
  }
  return formatFingerprintOf(payload);
}

/** `wpsFormatEvidenceFingerprint` 的动作重载（声明字段由 `declaredFormatFields` 给出）。 */
export function formatEvidenceFingerprint(
  action: FormatActionFields,
  snapshot: WpsFormatSnapshot
): string {
  return wpsFormatEvidenceFingerprint(action.type, declaredFormatFields(action), snapshot);
}

/** 声明属性在**执行后**的取值。`applyFormatAction` 把声明值写到区域内的每一格/每一行。 */
function formatTargetValue(action: FormatActionFields, field: WpsFormatField): WpsFormatScalar {
  switch (field) {
    case "rowHeight":
      return action.type === "resizeRange" ? action.rowHeight ?? null : null;
    case "bold":
      return action.type === "setFont" ? action.bold ?? null : null;
    case "italic":
      return action.type === "setFont" ? action.italic ?? null : null;
    case "name":
      return action.type === "setFont" ? action.name ?? null : null;
    case "size":
      return action.type === "setFont" ? action.size ?? null : null;
    case "color":
      return action.type === "setFont" && action.color != null ? wpsColorFromHex(action.color) : null;
    case "horizontalAlignment":
      return action.type === "setAlignment" && action.horizontal != null
        ? wpsHorizontalAlignment(action.horizontal)
        : null;
    case "wrapText":
      return action.type === "setAlignment" ? action.wrapText ?? null : null;
  }
}

/**
 * 写后**目标**指纹：由"动作声明的属性 + 目标值 + 区域形状"唯一决定，与任何单元格内容无关。
 *
 * 之所以能在**登记那一刻**就把它算出来：`applyFormatAction` 会把声明值写到区域内
 * 每一格（行高写到每一行），所以目标现场的声明属性取值是常量。这里显式构造同一形状
 * 的取值，再用与回读侧**同一个**序列化口径出哈希 —— 两边共用一套归一化，才不会出现
 * "回读判定通过、指纹却对不上"的自相矛盾。
 */
export function formatExpectedFingerprint(
  action: FormatActionFields,
  rowCount: number,
  columnCount: number
): string {
  const fields = declaredFormatFields(action);
  const payload: unknown[] = [action.type, [...fields]];
  for (const field of fields) {
    const value = formatTargetValue(action, field);
    payload.push(
      field === "rowHeight"
        ? Array.from({ length: rowCount }, () => value)
        : Array.from({ length: rowCount }, () => Array.from({ length: columnCount }, () => value))
    );
  }
  return formatFingerprintOf(payload);
}

/**
 * 一个格式动作在**执行那一刻**的写前格式快照 = 预览快照 + 同一份计划里更早的、
 * 区域真正相交的格式动作的**声明值**。
 *
 * 为什么必须有这一步：同区域多个格式动作按顺序执行时，后一个动作的"写前"是前一个
 * 动作执行后的状态。拿**预览时的原始快照**当基准，会同时坏掉两件事：
 * 1. 写前复核会把一条完全正常的链路误报成"格式在预览后发生变化"；
 * 2. 账本会登记一个**永远达不到**的 `previewFingerprint`，服务端随后以
 *    "动作证据与登记的写前状态不一致"拒掉写前检查点，恢复链整条断在登记上。
 *
 * 合成是**属性级**的：只看前序动作**声明过**的属性，且只落在两个矩形真正相交的
 * 格/行上 —— 与 `applyFormatAction` 的实际写入范围逐格一致，也就是"属性级恢复"
 * 的同一套语义。前序动作没声明的属性一律不动（保持预览快照里的原值）。
 *
 * 纯函数、只读：不碰宿主，也不重新读一次宿主（快照是本机现场，一律不出站）。
 * 没有任何格式快照时返回 `null`（调用方据此拒绝登记，而不是拿空快照冒充）。
 */
export function composedFormatSnapshot(
  actions: readonly WpsPreparedAction[],
  index: number
): WpsFormatSnapshot | null {
  const target = actions[index];
  if (!target?.formatSnapshot) return null;
  const fields = declaredFormatFields(asFormatFields(target));
  const composed: WpsFormatSnapshot = {
    cells: target.formatSnapshot.cells.map((row) => row.map((cell) => ({ ...cell }))),
    rowHeights: [...target.formatSnapshot.rowHeights],
    captured: target.formatSnapshot.captured
  };
  for (let earlier = 0; earlier < index; earlier += 1) {
    const source = actions[earlier];
    if (!source || !isFormatActionType(source.type)) continue;
    if (source.sheetName !== target.sheetName) continue;
    /*
     * 影响范围按**字段粒度**判：
     * - 行高是**行级**属性，只要行相交就算（A2:A3 与 C3:C4 在第 3 行上确实互相影响，
     *   哪怕列完全不相交）；
     * - 其余属性落在格上，必须矩形真正相交。
     * 拿"矩形相交"一刀切会把行高的整行连带影响漏掉。
     */
    const rowsIntersect =
      source.rect.startRow <= target.rect.endRow && target.rect.startRow <= source.rect.endRow;
    if (!rowsIntersect) continue;
    const cellsIntersect =
      source.rect.startColumn <= target.rect.endColumn && target.rect.startColumn <= source.rect.endColumn;
    const sourceFields = declaredFormatFields(asFormatFields(source));
    const firstRow = Math.max(source.rect.startRow, target.rect.startRow);
    const lastRow = Math.min(source.rect.endRow, target.rect.endRow);
    const firstColumn = Math.max(source.rect.startColumn, target.rect.startColumn);
    const lastColumn = Math.min(source.rect.endColumn, target.rect.endColumn);
    for (const field of fields) {
      if (!sourceFields.includes(field)) continue;
      const value = formatTargetValue(asFormatFields(source), field);
      if (field === "rowHeight") {
        // 行高是**行级**属性：相交的整行都归前序动作。
        for (let row = firstRow; row <= lastRow; row += 1) {
          composed.rowHeights[row - target.rect.startRow] = value as number | null;
        }
        continue;
      }
      if (!cellsIntersect) continue;
      for (let row = firstRow; row <= lastRow; row += 1) {
        for (let column = firstColumn; column <= lastColumn; column += 1) {
          const cell = composed.cells[row - target.rect.startRow]?.[column - target.rect.startColumn];
          if (cell) (cell as Record<string, WpsFormatScalar>)[field] = value;
        }
      }
    }
  }
  return composed;
}

/** 逐格比对「声明属性都改到目标值」。任一格不符即 false。 */
function declaredFieldsMatch(action: FormatActionFields, cells: WpsFormatSnapshot["cells"]): boolean {
  for (let r = 0; r < cells.length; r += 1) {
    for (let c = 0; c < cells[r].length; c += 1) {
      const cell = cells[r][c];
      if (action.type === "setFont") {
        if (action.bold != null && cell.bold !== action.bold) return false;
        if (action.italic != null && cell.italic !== action.italic) return false;
        if (action.name != null && cell.name !== action.name) return false;
        if (action.size != null && cell.size !== action.size) return false;
        if (action.color != null && cell.color !== wpsColorFromHex(action.color)) return false;
      } else if (action.type === "setAlignment") {
        if (action.horizontal != null && cell.horizontalAlignment !== wpsHorizontalAlignment(action.horizontal)) return false;
        if (action.wrapText != null && cell.wrapText !== action.wrapText) return false;
      }
    }
  }
  if (action.type === "resizeRange" && action.rowHeight != null) {
    // 行高在 rowHeights 里单独比（cells 里不含 rowHeight）。
    return true;
  }
  return true;
}

/**
 * 执行后回读判定：逐格/逐行核对「声明属性都改到了目标值，**未声明属性保持写前值**」。
 * 任一不符即 `false`（= 回读不一致）。这里要**全字段**比对 —— 保证格式动作精确只改
 * 它声明的那几个属性，没有连带改动别的属性。
 */
export function formatActionApplied(
  action: FormatActionFields,
  before: WpsFormatSnapshot,
  after: WpsFormatSnapshot
): boolean {
  // 声明属性必须都改到目标值。
  if (!declaredFieldsMatch(action, after.cells)) return false;
  // 未声明的属性必须保持写前值（确保没有连带改动）—— 覆盖**所有**字段，
  // 不只 bold/italic（阻断点 2：检查范围必须等于恢复范围）。
  const declared = new Set<WpsFormatField>(declaredFormatFields(action));
  for (let r = 0; r < after.cells.length; r += 1) {
    for (let c = 0; c < after.cells[r].length; c += 1) {
      const beforeCell = before.cells[r][c];
      const afterCell = after.cells[r][c];
      if (!declared.has("bold") && afterCell.bold !== beforeCell.bold) return false;
      if (!declared.has("italic") && afterCell.italic !== beforeCell.italic) return false;
      if (!declared.has("name") && afterCell.name !== beforeCell.name) return false;
      if (!declared.has("size") && afterCell.size !== beforeCell.size) return false;
      if (!declared.has("color") && afterCell.color !== beforeCell.color) return false;
      if (!declared.has("horizontalAlignment") && afterCell.horizontalAlignment !== beforeCell.horizontalAlignment) return false;
      if (!declared.has("wrapText") && afterCell.wrapText !== beforeCell.wrapText) return false;
    }
  }
  if (action.type === "resizeRange" && action.rowHeight != null) {
    for (let r = 0; r < after.rowHeights.length; r += 1) {
      if (after.rowHeights[r] !== action.rowHeight) return false;
    }
  } else {
    // 非行高动作：行高必须保持写前值（未声明）。
    for (let r = 0; r < after.rowHeights.length; r += 1) {
      if (after.rowHeights[r] !== before.rowHeights[r]) return false;
    }
  }
  return true;
}

/**
 * 撤销冲突检测：当前现场是否**仍是「执行后」状态**（只比对声明属性）。
 *
 * 与 `formatActionApplied` 不同：撤销只恢复声明属性，用户对其他属性的后续修改
 * 不该被覆盖、也不该触发冲突 —— 所以这里**只**看"声明属性是否仍等于执行后的目标值"。
 * 声明属性被用户改掉 ⇒ 恢复会覆盖用户的这部分修改 ⇒ 返回 false（拒绝自动覆盖）。
 */
function formatActionStillApplied(
  action: FormatActionFields,
  current: WpsFormatSnapshot
): boolean {
  if (!declaredFieldsMatch(action, current.cells)) return false;
  if (action.type === "resizeRange" && action.rowHeight != null) {
    for (let r = 0; r < current.rowHeights.length; r += 1) {
      if (current.rowHeights[r] !== action.rowHeight) return false;
    }
  }
  return true;
}

/**
 * 撤销冲突检测（Task #53 收口）：按**最终恢复影响范围**比对 —— 本动作声明过的
 * 属性，现场是否仍等于"**整份计划执行后的最终格式**"。
 *
 * 为什么不是"本动作声明的目标值"（`formatActionStillApplied`）：同属性后续覆盖
 * （如字号先 12→14 再 14→16）时，第一步声明的 14 早已被第二步覆盖成 16，按目标值
 * 判会把正常的组合误报成"又被修改"。最终状态才是计划整体产出的样子，现场与它
 * 不一致才说明有**外部改动**。
 *
 * 只比对声明字段：用户对其他属性的后续修改不参与判定（撤销也不会恢复它们）。
 */
function formatActionUnchangedSinceFinal(
  action: FormatActionFields,
  finalSnapshot: WpsFormatSnapshot,
  current: WpsFormatSnapshot
): boolean {
  const declared = new Set<WpsFormatField>(declaredFormatFields(action));
  for (let r = 0; r < current.cells.length; r += 1) {
    for (let c = 0; c < current.cells[r].length; c += 1) {
      const finalCell = finalSnapshot.cells[r]?.[c];
      if (!finalCell) return false;
      const currentCell = current.cells[r][c];
      if (declared.has("bold") && currentCell.bold !== finalCell.bold) return false;
      if (declared.has("italic") && currentCell.italic !== finalCell.italic) return false;
      if (declared.has("name") && currentCell.name !== finalCell.name) return false;
      if (declared.has("size") && currentCell.size !== finalCell.size) return false;
      if (declared.has("color") && currentCell.color !== finalCell.color) return false;
      if (declared.has("horizontalAlignment") && currentCell.horizontalAlignment !== finalCell.horizontalAlignment) return false;
      if (declared.has("wrapText") && currentCell.wrapText !== finalCell.wrapText) return false;
    }
  }
  if (declared.has("rowHeight")) {
    for (let r = 0; r < current.rowHeights.length; r += 1) {
      if (current.rowHeights[r] !== finalSnapshot.rowHeights[r]) return false;
    }
  }
  return true;
}

/**
 * 撤销冲突检测的统一入口：有最终快照就按最终状态比；历史/部分写入记录缺最终快照时
 * 退回"声明目标值"旧口径（尽力而为，不静默放行到覆盖用户改动）。
 */
function formatActionSafeToUndo(
  action: FormatActionFields,
  afterFinal: WpsFormatSnapshot | undefined,
  current: WpsFormatSnapshot
): boolean {
  return afterFinal
    ? formatActionUnchangedSinceFinal(action, afterFinal, current)
    : formatActionStillApplied(action, current);
}

/**
 * 恢复后核对：本次恢复的属性是否回到写前值（只比对声明字段）。
 *
 * 属性级恢复只动声明字段，所以恢复成功 = "声明字段逐格/逐行回到写前快照"。
 * 用户对其他属性的修改仍在现场，不参与这里判定。
 */
function formatActionRestored(
  action: FormatActionFields,
  before: WpsFormatSnapshot,
  restored: WpsFormatSnapshot,
  declared: ReadonlySet<WpsFormatField>
): boolean {
  for (let r = 0; r < restored.cells.length; r += 1) {
    for (let c = 0; c < restored.cells[r].length; c += 1) {
      const beforeCell = before.cells[r][c];
      const restoredCell = restored.cells[r][c];
      if (declared.has("bold") && restoredCell.bold !== beforeCell.bold) return false;
      if (declared.has("italic") && restoredCell.italic !== beforeCell.italic) return false;
      if (declared.has("name") && restoredCell.name !== beforeCell.name) return false;
      if (declared.has("size") && restoredCell.size !== beforeCell.size) return false;
      if (declared.has("color") && restoredCell.color !== beforeCell.color) return false;
      if (declared.has("horizontalAlignment") && restoredCell.horizontalAlignment !== beforeCell.horizontalAlignment) return false;
      if (declared.has("wrapText") && restoredCell.wrapText !== beforeCell.wrapText) return false;
    }
  }
  if (action.type === "resizeRange" && action.rowHeight != null) {
    for (let r = 0; r < restored.rowHeights.length; r += 1) {
      if (restored.rowHeights[r] !== before.rowHeights[r]) return false;
    }
  }
  return true;
}

function restorationMatrix(action: StoredAction): CellValue[][] {
  return action.beforeValues.map((row, rowIndex) =>
    row.map((value, columnIndex) => {
      const formula = action.beforeFormulas[rowIndex]?.[columnIndex];
      return typeof formula === "string" && formula.startsWith("=") ? formula : value;
    })
  );
}

function yieldControl(): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, 0));
}

/**
 * 动作级检查点 —— 交给调用方（面板）提交到服务端持久账本。
 *
 * **只含指纹与动作元数据，不含单元格值矩阵**：`beforeFingerprint`/`afterFingerprint`
 * 是服务端恢复时比对现场的证据，重放矩阵既不必要也不允许出站。恢复时用指纹判断
 * "现场是写前、写后还是两者都不是"，而不是把值矩阵存进账本。
 */
export type WpsActionCheckpoint = {
  actionIndex: number;
  sheetName: string;
  range: string;
  type: "writeValues" | "writeFormulas" | "clearRange" | "setFont" | "setAlignment" | "resizeRange";
  beforeFingerprint: string;
  afterFingerprint: string | null;
};

export type WpsActionCheckpointHook = (
  phase: "before" | "after",
  checkpoint: WpsActionCheckpoint
) => Promise<void>;

export async function executeWpsPlan(
  application: WpsApplication,
  prepared: WpsPreparedPlan,
  onProgress?: (progress: WpsExecutionProgress) => void,
  /**
   * 双栏变更对照只走回调交出去，**不写进 `WpsUndoRecord`** —— record 会落 PluginStorage，
   * 而前后矩阵按 §7 只许停在会话内内存。
   *
   * 一并回传 `executionId`：失败路径同样会回调一次（失败后必须先给对照，再由用户决定
   * 撤销还是重做），那时调用方拿不到 record。
   */
  onChangeReview?: (reviews: RangeChangeReview[], executionId: string) => void,
  /**
   * 动作级持久检查点。写前/写后各触发一次，供调用方提交服务端账本（CAP-03）。
   * 回调抛错会**中止后续写入** —— 拿不到持久动作回执就不继续写，宁可不写也不留
   * "写了一半却没有恢复证据"的现场。
   */
  onActionCheckpoint?: WpsActionCheckpointHook,
  /**
   * 覆盖默认随机生成的 `executionId`。CAP-03 用它把 PluginStorage 恢复点与服务端
   * 账本的 `operationId` 对齐，使撤销、核对、续跑都指向同一份执行身份。
   */
  executionIdOverride?: string
): Promise<WpsUndoRecord> {
  if (!capabilities.wpsWrite.enabled) throw new Error("WPS 写入功能当前已由能力开关关闭。");
  const hostVersion = String(application.Version ?? "").trim();
  if (hostVersion !== prepared.hostVersion || wpsWriteVersionStatus(hostVersion) === "blocked") {
    throw new Error("WPS 宿主版本已变化或不在当前写入范围，未执行写入。");
  }
  assertWpsPreparedPlanUnchanged(prepared);
  prepared.plan.actions.forEach((action) => {
    if (!supportedAction(action)) throw new Error(`WPS 当前不执行「${action.type}」。`);
    assertAllowedAction(action);
  });
  const workbook = application.ActiveWorkbook;
  if (!workbook || workbook.ReadOnly !== false) throw new Error("当前 WPS 工作簿已不再处于可编辑状态。");
  for (const source of prepared.sources) {
    const authorized = prepared.authorization.sheets.find((sheet) => sheet.index === source.sheetIndex);
    if (!authorized || authorized.name !== source.sheetName) {
      throw new Error(`源工作表「${source.sheetName}」不在预览时的授权范围，请重新预览。`);
    }
    const sheet = assertWpsAuthorization(application, prepared.authorization, authorized);
    const current = readMatrices(rangeFor(sheet, source.range), source.rowCount, source.columnCount);
    if (current.fingerprint !== source.previewFingerprint) {
      throw new Error(`源区域「${source.sheetName}!${source.range}」在预览后发生变化，请重新预览并确认。`);
    }
  }
  onProgress?.({
    stage: "snapshot",
    message: "正在保存执行前恢复点…",
    actionCount: prepared.actions.length,
    cellCount: prepared.totalCells
  });
  const storedActions: StoredAction[] = prepared.actions.map((action) => {
    const authorized = prepared.authorization.sheets.find((sheet) => sheet.index === action.sheetIndex);
    if (!authorized) throw new Error(`工作表「${action.sheetName}」已不在授权范围内。`);
    const sheet = assertWpsAuthorization(application, prepared.authorization, authorized);
    const range = rangeFor(sheet, action.range);
    if (sheet.ProtectContents !== false) throw new Error(`工作表「${action.sheetName}」已被保护，未执行写入。`);
    if (range.MergeCells !== false) throw new Error(`区域「${action.sheetName}!${action.range}」的合并状态已变化，未执行写入。`);

    // 格式动作：`beforeFormatSnapshot` 在执行循环里**现场采集**（见下），
    // 因为同区域多个格式动作按顺序执行时，后一个动作的"写前格式"是前一个动作
    // 执行后的状态，不能在开头统一采集。
    if (isFormatActionType(action.type)) {
      return {
        ...action,
        beforeValues: action.beforeValues,
        beforeFormulas: action.beforeValues,
        beforeFingerprint: action.previewFingerprint,
        afterFingerprint: null,
        appliedRows: 0,
        beforeFormatSnapshot: undefined
      };
    }

    if (range.Formula === undefined) throw new Error(`区域「${action.sheetName}!${action.range}」无法保存恢复快照。`);
    const before = readMatrices(range, action.rowCount, action.columnCount);
    if (before.fingerprint !== action.previewFingerprint) {
      throw new Error(`区域「${action.sheetName}!${action.range}」在预览后发生变化，请重新预览并确认。`);
    }
    return {
      ...action,
      beforeValues: before.values,
      beforeFormulas: before.formulas,
      beforeFingerprint: before.fingerprint,
      afterFingerprint: null,
      appliedRows: 0
    };
  });
  const record: WpsUndoRecord = {
    version: 2,
    executionId: executionIdOverride ?? randomId(),
    planId: prepared.plan.id,
    planTitle: prepared.plan.title,
    createdAt: new Date().toISOString(),
    workbookIdentity: prepared.workbookIdentity,
    authorizedSheetNames: prepared.authorization.sheets.map((sheet) => sheet.name),
    state: "ready",
    actions: storedActions,
    error: null
  };
  upsertRecord(application, record);

  // 回读矩阵按 actionIndex 存，只活在这次调用的闭包里。
  const afterByIndex = new Map<number, WpsActionAfter>();
  const matchByIndex = new Map<number, boolean>();
  const emitReviews = () => {
    onChangeReview?.(
      record.actions.map((action) =>
        wpsActionChangeReview(
          action,
          afterByIndex.get(action.actionIndex) ?? null,
          matchByIndex.get(action.actionIndex) ?? null
        )
      ),
      record.executionId
    );
  };

  try {
    record.state = "applying";
    upsertRecord(application, record);
    for (const [position, action] of record.actions.entries()) {
      const sheet = getStoredSheet(application, record, action);

      // 格式动作：执行前现场采集逐格/逐行格式快照（撤销依据 + 回读基线）。
      // 同区域多个格式动作按顺序执行，后一个动作的"写前格式"是前一个执行后的状态，
      // 必须在**这里**（本动作执行前）采集，不能在开头统一采集。
      if (isFormatActionType(action.type)) {
        const formatAction = asFormatFields(action);
        const formatRange = rangeFor(sheet, action.range);
        const beforeFormat = captureWpsFormatSnapshot(formatRange, action.rowCount, action.columnCount);
        if (!beforeFormat.captured) {
          throw new Error(`区域「${action.sheetName}!${action.range}」的格式无法完整逐格采集，未执行写入。`);
        }
        action.beforeFormatSnapshot = beforeFormat;
        /*
         * 写前**格式证据**：与预览时登记的必须是同一份。
         *
         * 值动作在下面用 `current.fingerprint !== action.beforeFingerprint` 复核"确认后没被改过"；
         * 格式动作的值不会变，那种复核对格式改动**没有区分力**，所以这里改用格式指纹复核 ——
         * 少了这一步，"预览后有人改了格式"会一路走到写入。
         *
         * 基准取**合成后的**写前快照（前序同区域格式动作的声明值已叠加）：直接拿预览
         * 原始快照比，会把"同区域按序叠加"这条正常链路误报成外部修改。
         */
        const beforeFingerprint = formatEvidenceFingerprint(formatAction, beforeFormat);
        const expectedBefore = composedFormatSnapshot(record.actions, position);
        if (expectedBefore) {
          const previewFingerprint = formatEvidenceFingerprint(formatAction, expectedBefore);
          if (beforeFingerprint !== previewFingerprint) {
            throw new Error(`区域「${action.sheetName}!${action.range}」的格式在预览后发生变化，请重新预览并确认。`);
          }
        }
        if (onActionCheckpoint) {
          await onActionCheckpoint("before", {
            actionIndex: action.actionIndex,
            sheetName: action.sheetName,
            range: action.range,
            type: action.type as WpsActionCheckpoint["type"],
            beforeFingerprint,
            afterFingerprint: null
          });
        }
        onProgress?.({
          stage: "writing",
          message: `正在写入格式「${action.sheetName}!${action.range}」…`,
          actionIndex: action.actionIndex,
          actionType: action.type,
          sheetName: action.sheetName,
          range: action.range,
          actionCount: record.actions.length,
          cellCount: action.cellCount
        });
        applyFormatAction(formatRange, formatAction);
        action.appliedRows = action.rowCount;
        // 回读格式快照，比对目标属性是否改到、未声明属性是否保持写前值。
        const afterFormat = captureWpsFormatSnapshot(formatRange, action.rowCount, action.columnCount);
        const afterFingerprint = formatEvidenceFingerprint(formatAction, afterFormat);
        action.afterFingerprint = afterFingerprint;
        const matched = afterFormat.captured && formatActionApplied(
          formatAction,
          beforeFormat,
          afterFormat
        );
        matchByIndex.set(action.actionIndex, matched);
        if (!matched) {
          throw new Error(`区域「${action.sheetName}!${action.range}」的格式回读结果与计划不一致。`);
        }
        /*
         * 回读指纹必须**精确等于**登记时算出的目标指纹。两条判据读的是同一份声明，
         * 但一条走现场取值、一条走声明取值；任何一边的归一化漂移都会在这里暴露，
         * 而不是变成"账本里存着一个永远达不到的目标指纹"。
         */
        if (afterFingerprint !== formatExpectedFingerprint(formatAction, action.rowCount, action.columnCount)) {
          throw new Error(`区域「${action.sheetName}!${action.range}」的格式写后证据与登记的目标指纹不一致。`);
        }
        if (onActionCheckpoint) {
          await onActionCheckpoint("after", {
            actionIndex: action.actionIndex,
            sheetName: action.sheetName,
            range: action.range,
            type: action.type as WpsActionCheckpoint["type"],
            beforeFingerprint,
            afterFingerprint
          });
        }
        upsertRecord(application, record);
        await yieldControl();
        continue;
      }

      const current = readMatrices(rangeFor(sheet, action.range), action.rowCount, action.columnCount);
      if (current.fingerprint !== action.beforeFingerprint) {
        throw new Error(`区域「${action.sheetName}!${action.range}」在确认后发生变化，未执行写入。`);
      }
      // 动作写前检查点：把写前指纹登记到服务端账本，之后才允许动这块区域。
      if (onActionCheckpoint) {
        await onActionCheckpoint("before", {
          actionIndex: action.actionIndex,
          sheetName: action.sheetName,
          range: action.range,
          type: action.type,
          beforeFingerprint: action.beforeFingerprint,
          afterFingerprint: null
        });
      }
      for (let offset = 0; offset < action.rowCount; offset += capabilities.wpsWrite.chunkRows) {
        const rows = Math.min(capabilities.wpsWrite.chunkRows, action.rowCount - offset);
        onProgress?.({
          stage: "writing",
          message: `正在写入「${action.sheetName}!${rangeAddress(action.rect, offset, rows)}」…`,
          actionIndex: action.actionIndex,
          actionType: action.type,
          sheetName: action.sheetName,
          range: action.range,
          completedRows: offset,
          totalRows: action.rowCount,
          actionCount: record.actions.length,
          cellCount: action.cellCount
        });
        const chunk = rangeFor(sheet, rangeAddress(action.rect, offset, rows));
        const chunkValues = action.values.slice(offset, offset + rows);
        if (action.type === "writeFormulas") setRangeFormulas(chunk, chunkValues);
        else setRangeValues(chunk, chunkValues);
        action.appliedRows = offset + rows;
        const afterChunk = readMatrices(rangeFor(sheet, action.range), action.rowCount, action.columnCount);
        action.afterFingerprint = afterChunk.fingerprint;
        // 这一读本来就有（分块进度 + 指纹），矩阵随即丢弃；这里顺手留一份，
        // 分块中途失败时它就是最后一块真实的现场。不新增任何宿主调用。
        afterByIndex.set(action.actionIndex, {
          values: afterChunk.values,
          formulas: afterChunk.formulas,
          valuesCaptured: afterChunk.valuesCaptured
        });
        upsertRecord(application, record);
        onProgress?.({
          stage: "writing",
          message: `已写入「${action.sheetName}!${rangeAddress(action.rect, offset, rows)}」`,
          actionIndex: action.actionIndex,
          actionType: action.type,
          sheetName: action.sheetName,
          range: action.range,
          completedRows: action.appliedRows,
          totalRows: action.rowCount,
          actionCount: record.actions.length,
          cellCount: action.cellCount
        });
        await yieldControl();
      }
      onProgress?.({
        stage: "verifying",
        message: `正在回读核对「${action.sheetName}!${action.range}」…`,
        actionIndex: action.actionIndex,
        actionType: action.type,
        sheetName: action.sheetName,
        range: action.range,
        actionCount: record.actions.length,
        cellCount: action.cellCount
      });
      const after = readMatrices(rangeFor(sheet, action.range), action.rowCount, action.columnCount);
      action.afterFingerprint = after.fingerprint;
      afterByIndex.set(action.actionIndex, {
        values: after.values,
        formulas: after.formulas,
        valuesCaptured: after.valuesCaptured
      });
      const observed = action.type === "writeFormulas" ? after.formulas : after.values;
      const matched = matricesEqual(observed, action.values);
      matchByIndex.set(action.actionIndex, matched);
      if (!matched) {
        throw new Error(`区域「${action.sheetName}!${action.range}」回读结果与计划不一致。`);
      }
      // 动作写后检查点：回读一致后才提交完成回执，账本据此推进 nextActionIndex。
      if (onActionCheckpoint) {
        await onActionCheckpoint("after", {
          actionIndex: action.actionIndex,
          sheetName: action.sheetName,
          range: action.range,
          type: action.type,
          beforeFingerprint: action.beforeFingerprint,
          afterFingerprint: action.afterFingerprint
        });
      }
      upsertRecord(application, record);
    }
    // 整份计划全部执行完后再采集格式动作的**最终**格式快照（每个格式动作按其区域）。
    // 撤销冲突检测按最终状态比 —— 同属性后续覆盖时，中间态代表不了计划产出。
    for (const action of record.actions) {
      if (!isFormatActionType(action.type) || action.appliedRows === 0) continue;
      try {
        const sheet = getStoredSheet(application, record, action);
        const finalFormat = captureWpsFormatSnapshot(rangeFor(sheet, action.range), action.rowCount, action.columnCount);
        if (finalFormat.captured) action.afterFormatSnapshot = finalFormat;
      } catch {
        // 读不到就不强求：撤销时退回"声明目标值"旧口径（见 `formatActionSafeToUndo`）。
      }
    }
    record.state = "verified";
    upsertRecord(application, record);
    emitReviews();
    return record;
  } catch (error) {
    for (const action of record.actions) {
      try {
        const sheet = getStoredSheet(application, record, action);
        const current = readMatrices(rangeFor(sheet, action.range), action.rowCount, action.columnCount);
        // 只为"已经写过但没留下回读矩阵"的动作补一份现场（例如第一块写入就抛错）。
        // 没写过的动作不补：它的"修改后"就该是未采集，不能拿"跟修改前一样"糊过去。
        if (action.appliedRows > 0 && !afterByIndex.has(action.actionIndex)) {
          afterByIndex.set(action.actionIndex, {
            values: current.values,
            formulas: current.formulas,
            valuesCaptured: current.valuesCaptured
          });
        }
        if (current.fingerprint !== action.beforeFingerprint) {
          action.appliedRows = action.rowCount;
          action.afterFingerprint = current.fingerprint;
        }
        // 格式动作：部分写入时也留一份**当前**格式快照当最终状态，撤销冲突检测才
        // 按最终状态比（读不到就留空，撤销退回"声明目标值"旧口径）。
        if (isFormatActionType(action.type) && action.appliedRows > 0) {
          const partialFormat = captureWpsFormatSnapshot(rangeFor(sheet, action.range), action.rowCount, action.columnCount);
          if (partialFormat.captured) action.afterFormatSnapshot = partialFormat;
        }
      } catch {
        // Keep the durable before-snapshot even when the host can no longer be inspected.
      }
    }
    record.state = record.actions.some((action) => action.appliedRows > 0)
      ? "partial_write_recoverable"
      : "restored";
    record.error = error instanceof Error ? error.message : "WPS 写入失败。";
    upsertRecord(application, record);
    // §6.5：失败后先给实际变更对照（哪些步写了、哪些没写），再让用户决定撤销还是重做。
    emitReviews();
    throw error;
  }
}

export function latestWpsUndo(application: WpsApplication): WpsUndoRecord | null {
  const workbook = application.ActiveWorkbook;
  const identity = String(workbook?.FullName ?? "").trim().toLowerCase();
  return readRecords(application).find((record) =>
    record.workbookIdentity === identity &&
    (record.state === "verified" || record.state === "partial_write_recoverable" || record.state === "undo_conflict")
  ) ?? null;
}

/**
 * 按 `executionId` 读**原始**记录 —— **不**过滤可撤销性。
 *
 * 判断"这次写入到底有没有发生"时不能只看 `latestWpsUndo`：它只返回当前**还能撤销**
 * 的那一条，会漏掉 `restored`（执行器已把区域放回写前）与 `ready`/`applying` 等中间态。
 * 核对写前/写后状态需要看到记录的真实 state，所以单独开这个读口。
 *
 * 读不到时返回 `null`；**调用方不得把 `null` 解读成"没写过"** ——
 * 存储被清空、键损坏、宿主拒绝读取都会得到 `null`，那时只能说"未知"。
 */
export function wpsUndoRecordFor(
  application: WpsApplication,
  executionId: string
): WpsUndoRecord | null {
  return readRecords(application).find((record) => record.executionId === executionId) ?? null;
}

/**
 * 一次写入尝试的**只读**核对结果。
 *
 * `conflict` 同时覆盖"被人工改过"和"只写了一部分"：两者的下一步都是"先核对、
 * 不覆盖"，`detail` 里说清是哪一种。
 */
export type WpsUndoInspection =
  | { kind: "completed" }
  | { kind: "not_started" }
  | { kind: "conflict"; detail: string }
  | { kind: "unknown"; detail: string };

/**
 * 只读核对：现场到底是"已写完"、"仍是写前"还是"两者都不是"。
 *
 * **不写任何单元格**，也不改恢复记录 —— 只读目标区域并与记录里的写前/写后指纹比。
 * `undoWpsExecution` 也会做同样的比对，但它比对通过之后就开始写回；核对阶段
 * 需要的是"先看清楚再决定"，所以单独开这个只读入口。
 *
 * 判据只认正面证据：
 * - 每个**报过写入**的动作都与 `afterFingerprint` 一致 → 已完成；
 * - 都与 `beforeFingerprint` 一致 → 确实没开始；
 * - 混合、或与两者都不一致 → 冲突（需要人工看）；
 * - 没有动作报过写入、记录缺失、读取抛错 → **未知**（不得当成"没写过"）。
 */
export function inspectWpsUndoRecord(
  application: WpsApplication,
  executionId: string
): WpsUndoInspection {
  const record = readRecords(application).find((item) => item.executionId === executionId);
  if (!record) {
    return { kind: "unknown", detail: "找不到这次执行的恢复记录。" };
  }
  const applicable = record.actions.filter((action) => action.appliedRows > 0);
  if (applicable.length === 0) {
    // 一个动作都没报过写入：执行器若明确说区域已回到写前，才敢说"未开始"。
    if (record.state === "restored") return { kind: "not_started" };
    return { kind: "unknown", detail: "执行记录里没有任何已写入的动作，无法判断现场。" };
  }
  let atAfter = 0;
  let atBefore = 0;
  for (const action of applicable) {
    const sheet = getStoredSheet(application, record, action);

    // 格式动作：核对当前格式快照是"写后"还是"写前"。
    if (isFormatActionType(action.type)) {
      let currentFormat: WpsFormatSnapshot;
      try {
        currentFormat = captureWpsFormatSnapshot(rangeFor(sheet, action.range), action.rowCount, action.columnCount);
      } catch (error) {
        return { kind: "unknown", detail: error instanceof Error ? error.message : "读取格式现场失败。" };
      }
      if (!currentFormat.captured) {
        return { kind: "unknown", detail: `区域「${action.sheetName}!${action.range}」的格式无法逐格采集。` };
      }
      const atAfterState = formatActionSafeToUndo(
        asFormatFields(action),
        action.afterFormatSnapshot,
        currentFormat
      );
      if (atAfterState) {
        atAfter += 1;
        continue;
      }
      if (sameWpsFormatSnapshot(currentFormat, action.beforeFormatSnapshot!)) {
        atBefore += 1;
        continue;
      }
      return {
        kind: "conflict",
        detail: `区域「${action.sheetName}!${action.range}」的格式与写前、写后都不一致。`
      };
    }

    let fingerprint: string;
    try {
      const current = readMatrices(
        rangeFor(sheet, action.range),
        action.rowCount,
        action.columnCount,
        record.version
      );
      fingerprint = current.fingerprint;
    } catch (error) {
      return {
        kind: "unknown",
        detail: error instanceof Error ? error.message : "读取原表现场失败。"
      };
    }
    if (action.afterFingerprint && fingerprint === action.afterFingerprint) {
      atAfter += 1;
      continue;
    }
    if (fingerprint === action.beforeFingerprint) {
      atBefore += 1;
      continue;
    }
    return {
      kind: "conflict",
      detail: `区域「${action.sheetName}!${action.range}」与写前、写后都不一致。`
    };
  }
  if (atAfter === applicable.length) return { kind: "completed" };
  if (atBefore === applicable.length) return { kind: "not_started" };
  return {
    kind: "conflict",
    detail: "一部分区域与写后一致、一部分仍是写前，属于只写了一部分的现场。"
  };
}

export async function undoWpsExecution(  application: WpsApplication,
  executionId: string,
  onProgress?: (progress: WpsExecutionProgress) => void
): Promise<WpsUndoRecord> {
  const record = readRecords(application).find((item) => item.executionId === executionId);
  if (!record) throw new Error("找不到这次 WPS 执行的恢复点。");
  if (!["verified", "partial_write_recoverable", "undo_conflict"].includes(record.state)) {
    throw new Error("这个恢复点当前不可撤销。");
  }
  const workbook = application.ActiveWorkbook;
  if (!workbook || workbook.ReadOnly !== false) throw new Error("当前 WPS 工作簿不可编辑，无法撤销。");
  for (const action of record.actions) {
    if (action.appliedRows === 0) continue;
    const sheet = getStoredSheet(application, record, action);
    if (sheet.ProtectContents !== false) throw new Error(`工作表「${action.sheetName}」已被保护，无法撤销。`);
    const target = rangeFor(sheet, action.range);
    if (target.MergeCells !== false) throw new Error(`区域「${action.sheetName}!${action.range}」的合并状态已变化，无法撤销。`);
    // 格式动作：撤销前核对当前格式快照是否仍是"执行后"状态（未被人改）。
    if (isFormatActionType(action.type)) {
      const currentFormat = captureWpsFormatSnapshot(target, action.rowCount, action.columnCount);
      if (!currentFormat.captured) {
        record.state = "undo_conflict";
        record.error = `区域「${action.sheetName}!${action.range}」的格式无法逐格采集，不能自动覆盖。`;
        upsertRecord(application, record);
        throw new Error(record.error);
      }
      const expectedAfter = formatActionSafeToUndo(
        asFormatFields(action),
        action.afterFormatSnapshot,
        currentFormat
      );
      if (!expectedAfter) {
        record.state = "undo_conflict";
        record.error = `区域「${action.sheetName}!${action.range}」的格式在执行后又被修改，不能自动覆盖。`;
        upsertRecord(application, record);
        throw new Error(record.error);
      }
      continue;
    }
    const current = readMatrices(target, action.rowCount, action.columnCount, record.version);
    if (record.version === 1 && (
      hasAmbiguousLogicalCell(action.beforeValues, action.beforeFormulas) ||
      hasAmbiguousLogicalCell(current.values, current.formulas)
    )) {
      record.state = "undo_conflict";
      record.error = `旧版恢复点无法区分区域「${action.sheetName}!${action.range}」的空字符串与空单元格，已保留恢复资料并拒绝自动覆盖。`;
      upsertRecord(application, record);
      throw new Error(record.error);
    }
    if (!action.afterFingerprint || current.fingerprint !== action.afterFingerprint) {
      record.state = "undo_conflict";
      record.error = `区域「${action.sheetName}!${action.range}」在执行后又被修改，不能自动覆盖。`;
      upsertRecord(application, record);
      throw new Error(record.error);
    }
  }
  record.state = "undoing";
  record.error = null;
  upsertRecord(application, record);
  for (const action of [...record.actions].reverse()) {
    if (action.appliedRows === 0) continue;
    onProgress?.({
      stage: "undoing",
      message: `正在恢复「${action.sheetName}!${action.range}」…`,
      actionIndex: action.actionIndex,
      actionType: action.type,
      sheetName: action.sheetName,
      range: action.range,
      actionCount: record.actions.length,
      cellCount: action.cellCount
    });
    const sheet = getStoredSheet(application, record, action);
    // 格式动作：恢复逐格/逐行格式快照，然后回读核对是否回到写前。
    if (isFormatActionType(action.type)) {
      const formatTarget = rangeFor(sheet, action.range);
      const formatFields = asFormatFields(action);
      // 属性级恢复：只恢复本次动作实际改动过的属性，保留用户对其他属性的后续修改。
      restoreWpsFormatSnapshot(formatTarget, action.beforeFormatSnapshot!, declaredFormatFields(formatFields));
      const restoredFormat = captureWpsFormatSnapshot(formatTarget, action.rowCount, action.columnCount);
      // 恢复后核对只比对「本次恢复的属性」是否回到写前值（不是整块全字段），
      // 因为用户对其他属性的修改本来就该被保留、不该拿来判定恢复失败。
      const declared = new Set<WpsFormatField>(declaredFormatFields(formatFields));
      const restoredOk = restoredFormat.captured && formatActionRestored(
        formatFields,
        action.beforeFormatSnapshot!,
        restoredFormat,
        declared
      );
      if (!restoredOk) {
        record.state = "undo_conflict";
        record.error = `区域「${action.sheetName}!${action.range}」的格式恢复后校验失败。`;
        upsertRecord(application, record);
        throw new Error(record.error);
      }
      await yieldControl();
      continue;
    }
    setRangeFormulas(rangeFor(sheet, action.range), restorationMatrix(action));
    const restored = readMatrices(rangeFor(sheet, action.range), action.rowCount, action.columnCount, record.version);
    if (restored.fingerprint !== action.beforeFingerprint) {
      record.state = "undo_conflict";
      record.error = `区域「${action.sheetName}!${action.range}」恢复后校验失败。`;
      upsertRecord(application, record);
      throw new Error(record.error);
    }
    await yieldControl();
  }
  record.state = "restored";
  upsertRecord(application, record);
  return record;
}
