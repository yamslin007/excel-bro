import capabilities from "../../../config/capabilities.json";
import type { WorkbookLocator } from "@excel-bro/core/contracts";
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

type StoredAction = WpsPreparedAction & {
  beforeValues: CellValue[][];
  beforeFormulas: CellValue[][];
  beforeFingerprint: string;
  afterFingerprint: string | null;
  appliedRows: number;
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
  return { values, formulas, fingerprint: fingerprintForVersion(values, formulas, version) };
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

export async function executeWpsPlan(
  application: WpsApplication,
  prepared: WpsPreparedPlan,
  onProgress?: (progress: WpsExecutionProgress) => void
): Promise<WpsUndoRecord> {
  if (!capabilities.wpsWrite.enabled) throw new Error("WPS 写入功能当前已由能力开关关闭。");
  const hostVersion = String(application.Version ?? "").trim();
  if (hostVersion !== prepared.hostVersion || wpsWriteVersionStatus(hostVersion) === "blocked") {
    throw new Error("WPS 宿主版本已变化或不在当前写入范围，未执行写入。");
  }
  if (confirmationSnapshot(prepared) !== prepared.confirmationSnapshot) {
    throw new Error("写入计划或预览已变化，请重新预览并确认后再执行。");
  }
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
    executionId: randomId(),
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

  try {
    record.state = "applying";
    upsertRecord(application, record);
    for (const action of record.actions) {
      const sheet = getStoredSheet(application, record, action);
      const current = readMatrices(rangeFor(sheet, action.range), action.rowCount, action.columnCount);
      if (current.fingerprint !== action.beforeFingerprint) {
        throw new Error(`区域「${action.sheetName}!${action.range}」在确认后发生变化，未执行写入。`);
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
      const observed = action.type === "writeFormulas" ? after.formulas : after.values;
      if (!matricesEqual(observed, action.values)) {
        throw new Error(`区域「${action.sheetName}!${action.range}」回读结果与计划不一致。`);
      }
      upsertRecord(application, record);
    }
    record.state = "verified";
    upsertRecord(application, record);
    return record;
  } catch (error) {
    for (const action of record.actions) {
      try {
        const sheet = getStoredSheet(application, record, action);
        const current = readMatrices(rangeFor(sheet, action.range), action.rowCount, action.columnCount);
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

export async function undoWpsExecution(
  application: WpsApplication,
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
