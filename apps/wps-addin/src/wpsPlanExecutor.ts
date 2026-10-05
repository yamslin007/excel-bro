import capabilities from "../../../config/capabilities.json";
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
  | Extract<ExcelAction, { type: "clearRange" }>;

type Rect = {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
};

export type WpsPreparedAction = {
  actionIndex: number;
  type: "writeValues" | "writeFormulas" | "clearRange";
  sheetIndex: number;
  sheetName: string;
  range: string;
  rowCount: number;
  columnCount: number;
  cellCount: number;
  values: CellValue[][];
  preview: CellValue[][];
  /** Target state when the user was shown this preview. */
  previewFingerprint: string;
  rect: Rect;
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
  const values = normalizeWpsMatrix(range.Value2, rows, columns);
  const formulas = normalizeWpsMatrix(range.Formula, rows, columns);
  return {
    values,
    formulas,
    fingerprint: fingerprintForVersion(values, formulas, version),
    // 宿主没把 Value2 交出来时不能拿公式顶替：`normalizeWpsMatrix` 会把 undefined
    // 归一成 null，看上去跟"空单元格"一模一样。这里如实记一笔，交给双栏显示。
    valuesCaptured: range.Value2 !== undefined && range.Value2 !== null
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
  return !(
    left.rect.endRow < right.rect.startRow ||
    right.rect.endRow < left.rect.startRow ||
    left.rect.endColumn < right.rect.startColumn ||
    right.rect.endColumn < left.rect.startColumn
  );
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

function valuesForAction(action: SupportedAction, rows: number, columns: number): CellValue[][] {
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
  return action.type === "writeValues" || action.type === "writeFormulas" || action.type === "clearRange";
}

function assertAllowedAction(action: SupportedAction): void {
  const key = action.type === "clearRange" ? `${action.type}:${action.applyTo}` : action.type;
  if (!(capabilities.wpsWrite.allowedActions as readonly string[]).includes(key)) {
    throw new Error(`WPS 当前未开放「${key}」，已拒绝整份计划。`);
  }
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
  range.Value2 = values.length === 1 && values[0].length === 1 ? values[0][0] : values;
}

function setRangeFormulas(range: WpsRange, formulas: CellValue[][]): void {
  range.Formula = formulas.length === 1 && formulas[0].length === 1 ? formulas[0][0] : formulas;
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
  type: "writeValues" | "writeFormulas" | "clearRange";
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
    for (const action of record.actions) {
      const sheet = getStoredSheet(application, record, action);
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
    let fingerprint: string;
    try {
      const sheet = getStoredSheet(application, record, action);
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
