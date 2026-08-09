import type {
  AnalysisPlan,
  ActionExecutionResult,
  CellValue,
  ExecutionUndoSnapshot,
  ExcelAction,
  PlanExecutionResult,
  VerificationCheck,
  VerificationCriterion,
  VerificationReport,
  UnverifiedAction,
  WorkbookSnapshot,
  WorksheetSnapshot
} from "./contracts";
import capabilities from "../../../config/capabilities.json";
import { detectSheetFields } from "./tableSchema";
import {
  batchSplitAggregateOutputs,
  buildSplitAggregateOutputs,
  safeWorksheetBaseName
} from "./splitAggregate";
import { workbookNameFromDocumentUrl } from "./workbookIdentity";
import {
  displayCellValue,
  normalizeCellValue
} from "./cellNormalization";

const DATA_ROW_LIMIT = capabilities.snapshot.dataRows;
const DATA_COLUMN_LIMIT = capabilities.snapshot.dataColumns;
const STRUCTURE_ROW_LIMIT = capabilities.snapshot.structureRows;
const STRUCTURE_COLUMN_LIMIT = capabilities.snapshot.structureColumns;
const FINGERPRINT_CHUNK_ROWS = capabilities.queryTable.chunkRows;
const FINGERPRINT_CHUNK_CELLS =
  capabilities.queryTable.chunkRows * capabilities.queryTable.maxColumns;
const SPLIT_AGGREGATE_BATCH_SHEETS =
  capabilities.excelExecution.splitAggregateBatchSheets;
// EB 规则系统早期把规则存进这两张隐藏工作表（见 docs/EB_FUNCTIONS.md）。
// 现已改为纯内置规则，代码不再读写它们，但用户的旧文件里可能残留。
// 用精确名单过滤，避免误伤用户自己以 # 开头命名的正常表。
const EB_SYSTEM_SHEETS = new Set(["#EB_RULES", "#EB_RULES_BACKUP"]);

function isEBSystemSheet(name: string): boolean {
  return EB_SYSTEM_SHEETS.has(name);
}

let structureCache:
  | { key: string; snapshot: WorkbookSnapshot }
  | null = null;
let structureCacheEnabled = true;

// 事件驱动的数据版本号：每次监听到工作表内容变化就自增，作为"数据没变"的安全闸。
// 命中缓存的路径只读取这些计数器，零全表扫描（避免"用全量指纹当闸=重跑一遍扫描"）。
const GLOBAL_DATA_EPOCH_KEY = "__global__";
const dataEpochBySheet = new Map<string, number>();
let globalDataEpoch = 0;

function bumpSheetDataEpoch(sheetName: string | null): void {
  if (sheetName) {
    dataEpochBySheet.set(sheetName, (dataEpochBySheet.get(sheetName) ?? 0) + 1);
    return;
  }
  // 解析不到受影响的具体表名时全局自增：宁可整体 miss，不可误命中。
  globalDataEpoch += 1;
}

/**
 * 给定来源表集合，快照当前各表 epoch 与全局 epoch。写缓存时调用。
 */
export function snapshotDataEpochs(
  sheets: string[]
): Record<string, number> {
  const snapshot: Record<string, number> = {
    [GLOBAL_DATA_EPOCH_KEY]: globalDataEpoch
  };
  for (const name of sheets) {
    snapshot[name] = dataEpochBySheet.get(name) ?? 0;
  }
  return snapshot;
}

/**
 * 判断自快照以来数据是否变化：全局 epoch 变了，或任一记录的表 epoch 变了，返回 true。
 */
export function dataEpochsChanged(
  snapshot: Record<string, number>
): boolean {
  if ((snapshot[GLOBAL_DATA_EPOCH_KEY] ?? 0) !== globalDataEpoch) {
    return true;
  }
  for (const [name, epoch] of Object.entries(snapshot)) {
    if (name === GLOBAL_DATA_EPOCH_KEY) continue;
    if ((dataEpochBySheet.get(name) ?? 0) !== epoch) {
      return true;
    }
  }
  return false;
}

export function invalidateWorkbookStructureCache(): void {
  structureCache = null;
}

interface SheetChangeEventArgs {
  worksheetId?: string;
}

export async function watchWorkbookStructureChanges(
  onInvalidated?: () => void
): Promise<() => void> {
  // 监听到内容变化：失效结构缓存 + 按受影响表自增 dataEpoch。
  // idToName 在监听建立时载入，用来把 change 事件里的 worksheetId 映射回表名；
  // 映射不到（如删表、批量改只报部分地址）就交给 bumpSheetDataEpoch(null) 全局自增。
  const idToName = new Map<string, string>();
  const handler = async (args?: SheetChangeEventArgs) => {
    invalidateWorkbookStructureCache();
    const sheetName =
      args?.worksheetId && idToName.has(args.worksheetId)
        ? idToName.get(args.worksheetId) ?? null
        : null;
    bumpSheetDataEpoch(sheetName);
    onInvalidated?.();
  };
  const supportsCollectionChanges =
    Office.context.requirements.isSetSupported("ExcelApi", "1.9");
  structureCacheEnabled = supportsCollectionChanges;
  const supportsCollectionLifecycle =
    Office.context.requirements.isSetSupported("ExcelApi", "1.7");
  const watched = await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    if (supportsCollectionChanges) {
      worksheets.onChanged.add(handler);
    }
    if (supportsCollectionLifecycle) {
      worksheets.onAdded.add(handler);
      worksheets.onDeleted.add(handler);
    }
    worksheets.load("items/name,items/id");
    await context.sync();
    idToName.clear();
    for (const sheet of worksheets.items) {
      idToName.set(sheet.id, sheet.name);
    }
    if (!supportsCollectionChanges) {
      for (const sheet of worksheets.items) {
        sheet.onChanged.add(handler);
      }
      await context.sync();
    }
    return worksheets.items.map((sheet) => sheet.name);
  });
  return () => {
    void Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;
      if (supportsCollectionChanges) {
        worksheets.onChanged.remove(handler);
      } else {
        for (const name of watched) {
          const sheet = worksheets.getItemOrNullObject(name);
          sheet.load("isNullObject");
          await context.sync();
          if (!sheet.isNullObject) sheet.onChanged.remove(handler);
        }
      }
      if (supportsCollectionLifecycle) {
        worksheets.onAdded.remove(handler);
        worksheets.onDeleted.remove(handler);
      }
      await context.sync();
    });
  };
}

export class PlanExecutionError extends Error {
  constructor(
    message: string,
    public readonly actionResults: ActionExecutionResult[]
  ) {
    super(message);
    this.name = "PlanExecutionError";
  }
}

export interface PlanPreflightIssue {
  index: number;
  message: string;
}

export interface PlanPreflightCatalog {
  tableNames?: string[];
  pivotTableNames?: string[];
  namedRangeNames?: string[];
  shapeNamesBySheet?: Record<string, string[]>;
}

function normalizeValue(value: unknown): CellValue {
  return normalizeCellValue(value);
}

export function sourceFingerprintForSnapshot(
  snapshot: Pick<WorkbookSnapshot, "name" | "worksheets">,
  sourceSheets: string[]
): string {
  const selected = new Set(sourceSheets);
  const serialized = JSON.stringify({
    name: snapshot.name,
    worksheets: snapshot.worksheets.map((sheet) => ({
      name: sheet.name,
      usedRange: sheet.usedRange,
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount,
      headers: selected.has(sheet.name) ? sheet.headers : [],
      dataRows: selected.has(sheet.name) ? sheet.dataRows : [],
      truncated: sheet.truncated
    }))
  });
  let hash = 0x811c9dc5;
  for (let index = 0; index < serialized.length; index += 1) {
    hash ^= serialized.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return `fnv1a32:${(hash >>> 0).toString(16).padStart(8, "0")}`;
}

const TOOL_SCHEMA_FINGERPRINT_PREFIX = "tool-schema-fnv1a32:";

function normalizeSchemaName(value: string): string {
  return value.trim().toLowerCase();
}

function requiredToolFieldsBySheet(
  plan: AnalysisPlan
): Map<string, Set<string>> {
  const required = new Map<string, Set<string>>();
  const add = (sheet: string, fields: Array<string | null | undefined>) => {
    const key = normalizeSchemaName(sheet);
    const current = required.get(key) ?? new Set<string>();
    fields.forEach((field) => {
      if (field?.trim()) current.add(normalizeSchemaName(field));
    });
    required.set(key, current);
  };
  for (const action of plan.actions) {
    if (action.type === "splitGroupAggregate") {
      add(action.sheet, [
        action.splitBy,
        ...action.groupBy,
        ...action.metrics.map((metric) => metric.field)
      ]);
    }
    if (action.type === "createPivotTable") {
      add(action.sourceSheet, [
        ...action.rowFields,
        ...action.columnFields,
        ...action.valueFields.map((value) => value.field)
      ]);
    }
  }
  return required;
}

/**
 * Saved tools are reusable recipes, so their safety gate follows the schema
 * they need instead of pinning every source cell. Row counts, values, used
 * ranges and column order may change; missing source sheets or required fields
 * still invalidate the preview.
 */
export function toolSchemaFingerprintForSnapshot(
  plan: AnalysisPlan,
  snapshot: Pick<WorkbookSnapshot, "worksheets">
): string {
  const requiredBySheet = requiredToolFieldsBySheet(plan);
  const requested = [...new Set(plan.sourceFingerprintSheets ?? [])]
    .map(normalizeSchemaName)
    .sort();
  const available = new Map(
    snapshot.worksheets.map((sheet) => [
      normalizeSchemaName(sheet.name),
      new Set(
        sheet.headers
          .map((header) => normalizeSchemaName(String(header ?? "")))
          .filter(Boolean)
      )
    ])
  );
  const serialized = JSON.stringify({
    version: 1,
    sheets: requested.map((sheet) => {
      const fields = available.get(sheet);
      const requiredFields = [...(requiredBySheet.get(sheet) ?? [])].sort();
      return {
        sheet,
        exists: Boolean(fields),
        requiredFields: requiredFields.map((field) => ({
          field,
          exists: fields?.has(field) === true
        }))
      };
    })
  });
  let hash = 0x811c9dc5;
  for (let index = 0; index < serialized.length; index += 1) {
    hash ^= serialized.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return `${TOOL_SCHEMA_FINGERPRINT_PREFIX}${(hash >>> 0)
    .toString(16)
    .padStart(8, "0")}`;
}

function isToolSchemaFingerprint(value: string): boolean {
  return value.startsWith(TOOL_SCHEMA_FINGERPRINT_PREFIX);
}

function updateFingerprint(seed: number, text: string): number {
  let hash = seed;
  for (let index = 0; index < text.length; index += 1) {
    hash ^= text.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return hash >>> 0;
}

export async function captureWorkbookSourceFingerprint(
  sourceSheetNames: string[]
): Promise<string> {
  const requested = [...new Set(sourceSheetNames)].sort((left, right) =>
    left.localeCompare(right)
  );
  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();
    const available = new Set(worksheets.items.map((sheet) => sheet.name));
    const missingName = requested.find((name) => !available.has(name));
    if (missingName) {
      throw new Error(`来源工作表「${missingName}」不存在，请重新生成预览`);
    }
    const pending = requested.map((name) => {
      const worksheet = worksheets.getItem(name);
      const usedRange = worksheet.getUsedRangeOrNullObject(true);
      worksheet.load("name");
      usedRange.load(
        "isNullObject,address,rowIndex,columnIndex,rowCount,columnCount"
      );
      return { name, worksheet, usedRange };
    });
    await context.sync();
    let primary = 0x811c9dc5;
    let secondary = 0x9e3779b9;
    const include = (value: unknown) => {
      const text = `${JSON.stringify(value)}\u001f`;
      primary = updateFingerprint(primary, text);
      secondary = updateFingerprint(secondary ^ 0x85ebca6b, text);
    };
    include(workbookNameFromDocumentUrl(Office.context.document.url));
    for (const { worksheet, usedRange } of pending) {
      include(worksheet.name);
      if (usedRange.isNullObject) {
        include(null);
        continue;
      }
      include([
        usedRange.address,
        usedRange.rowCount,
        usedRange.columnCount
      ]);
      const fingerprintRowsPerChunk = Math.max(
        1,
        Math.min(
          FINGERPRINT_CHUNK_ROWS,
          Math.floor(FINGERPRINT_CHUNK_CELLS / usedRange.columnCount)
        )
      );
      for (
        let offset = 0;
        offset < usedRange.rowCount;
        offset += fingerprintRowsPerChunk
      ) {
        const rowCount = Math.min(
          fingerprintRowsPerChunk,
          usedRange.rowCount - offset
        );
        const range = worksheet.getRangeByIndexes(
          usedRange.rowIndex + offset,
          usedRange.columnIndex,
          rowCount,
          usedRange.columnCount
        );
        range.load("values,text,formulas,numberFormat");
        await context.sync();
        for (let rowIndex = 0; rowIndex < range.values.length; rowIndex += 1) {
          include(
            range.values[rowIndex].map((value, columnIndex) => [
              normalizeCellValue(
                value,
                range.text[rowIndex]?.[columnIndex] ?? "",
                range.numberFormat[rowIndex]?.[columnIndex] ?? "General"
              ),
              range.formulas[rowIndex]?.[columnIndex] ?? null,
              range.numberFormat[rowIndex]?.[columnIndex] ?? "General"
            ])
          );
        }
      }
    }
    return `fnv1a32x2:${primary.toString(16).padStart(8, "0")}${secondary
      .toString(16)
      .padStart(8, "0")}`;
  });
}

function excelErrorDetail(reason: unknown): string {
  if (!(reason instanceof Error)) return String(reason);
  const error = reason as Error & {
    code?: string;
    debugInfo?: {
      errorLocation?: string;
      statement?: string;
      surroundingStatements?: string[];
    };
  };
  const details = [
    error.message,
    error.code ? `代码 ${error.code}` : "",
    error.debugInfo?.errorLocation
      ? `位置 ${error.debugInfo.errorLocation}`
      : "",
    error.debugInfo?.statement
      ? `语句 ${error.debugInfo.statement}`
      : ""
  ].filter(Boolean);
  return details.join("；");
}

export function isRunningInExcel(): boolean {
  return typeof Office !== "undefined" && Office.context?.host === Office.HostType.Excel;
}

export async function captureSelectionContext(): Promise<{
  activeWorksheet: string;
  selectedRange: string;
}> {
  return Excel.run(async (context) => {
    const activeWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const selectedRange = context.workbook.getSelectedRange();
    activeWorksheet.load("name");
    selectedRange.load("address");
    await context.sync();
    return {
      activeWorksheet: activeWorksheet.name,
      selectedRange: selectedRange.address
    };
  });
}

/**
 * 试算公式：把 formula 临时写进首格，读回 Excel 真算值，随即还原原公式。
 * 用于 /function 预览——让用户看到 Excel 亲算的真实结果，而非模型自报。
 */
export async function previewFormulaFirstCell(
  sheetName: string,
  firstCell: string,
  formula: string
): Promise<{ sampleResult: string; sampleInput: string; formulaR1C1: string }> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const cell = sheet.getRange(firstCell);
    cell.load("formulas");
    await context.sync();

    const original = cell.formulas as unknown[][];

    cell.formulas = [[formula]];
    await context.sync();

    cell.load("text,formulasR1C1");
    await context.sync();
    const sampleResult = String((cell.text as string[][])[0]?.[0] ?? "");
    // 锚点无关的 R1C1 形式，写入其他位置时保持相对引用正确平移
    const formulaR1C1 = String(
      (cell.formulasR1C1 as unknown[][])[0]?.[0] ?? ""
    );

    // 还原
    cell.formulas = original as (string | number | boolean)[][];
    await context.sync();

    const originalCell = String(original[0]?.[0] ?? "");
    return { sampleResult, sampleInput: originalCell, formulaR1C1 };
  });
}

// 即刻读取当前工作表选区，返回 { sheet, address }（裸 A1，无 $）。
// 交互：用户先在表里点/框选目标，再点「拾取」，本函数读回那一刻的选区。
// 供 /function 写入目标拾取用；调用方按当前预览表校验是否跨表。
export async function readSelectedRange(): Promise<{
  sheet: string;
  address: string;
} | null> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const sheet = range.worksheet;
    range.load("address");
    sheet.load("name");
    await context.sync();
    const rawAddress = String(range.address ?? "");
    if (!rawAddress) return null;
    const bare = rawAddress.includes("!")
      ? rawAddress.split("!").pop()!
      : rawAddress;
    return { sheet: String(sheet.name ?? ""), address: bare.replace(/\$/g, "") };
  });
}

export async function captureWorkbook(
  dataSheetNames?: string[]
): Promise<WorkbookSnapshot> {
  return Excel.run(async (context) => {
    const workbook = context.workbook;
    const worksheets = workbook.worksheets;
    const activeWorksheet = worksheets.getActiveWorksheet();
    const selectedRange = workbook.getSelectedRange();

    worksheets.load("items/name");
    activeWorksheet.load("name");
    selectedRange.load("address");
    await context.sync();

    const pending = worksheets.items
      .filter((worksheet) => !isEBSystemSheet(worksheet.name))
      .map((worksheet) => {
        const usedRange = worksheet.getUsedRangeOrNullObject(true);
        usedRange.load("address,rowCount,columnCount,rowIndex,columnIndex,isNullObject");
        return { worksheet, usedRange };
      });
    await context.sync();

    const sheetsToRead = new Set(
      dataSheetNames?.length ? dataSheetNames : [activeWorksheet.name]
    );
    const dataRanges = pending.map(({ worksheet, usedRange }) => {
      if (usedRange.isNullObject || !sheetsToRead.has(worksheet.name)) {
        return null;
      }
      const rowCount = Math.min(usedRange.rowCount, DATA_ROW_LIMIT + 1);
      const columnCount = Math.min(usedRange.columnCount, DATA_COLUMN_LIMIT);
      const range = worksheet.getRangeByIndexes(
        usedRange.rowIndex,
        usedRange.columnIndex,
        rowCount,
        columnCount
      );
      range.load("values,text,numberFormat");
      return range;
    });
    await context.sync();

    const snapshots: WorksheetSnapshot[] = pending.map(
      ({ worksheet, usedRange }, index) => {
        if (usedRange.isNullObject) {
          return {
            name: worksheet.name,
            usedRange: null,
            rowCount: 0,
            columnCount: 0,
            headers: [],
            dataRows: [],
            truncated: false
          };
        }

        const dataRange = dataRanges[index];
        const values = dataRange?.values ?? [];
        const texts = dataRange?.text ?? [];
        const formats = dataRange?.numberFormat ?? [];
        const normalizedValues = values.map((row, rowIndex) =>
          row.map((value, columnIndex) =>
            normalizeCellValue(
              value,
              texts[rowIndex]?.[columnIndex] ?? "",
              formats[rowIndex]?.[columnIndex] ?? "General"
            )
          )
        );
        return {
          name: worksheet.name,
          usedRange: usedRange.address,
          rowCount: usedRange.rowCount,
          columnCount: usedRange.columnCount,
          headers: normalizedValues[0] ?? [],
          dataRows: normalizedValues.slice(1),
          displayRows: values.slice(1).map((row, rowIndex) =>
            row.map((value, columnIndex) =>
              displayCellValue(
                normalizeCellValue(value),
                texts[rowIndex + 1]?.[columnIndex] ?? ""
              )
            )
          ),
          truncated:
            usedRange.rowCount > DATA_ROW_LIMIT + 1 ||
            usedRange.columnCount > DATA_COLUMN_LIMIT
        };
      }
    );

    const snapshot: WorkbookSnapshot = {
      name: workbookNameFromDocumentUrl(Office.context.document.url),
      capturedAt: new Date().toISOString(),
      activeWorksheet: activeWorksheet.name,
      selectedRange: selectedRange.address,
      worksheets: snapshots
    };
    const sourceFingerprintSheets = [...sheetsToRead];
    snapshot.sourceFingerprintSheets = sourceFingerprintSheets;
    snapshot.sourceFingerprint = sourceFingerprintForSnapshot(
      snapshot,
      sourceFingerprintSheets
    );
    return snapshot;
  });
}

export async function captureWorkbookStructure(
  dataSheetNames?: string[]
): Promise<WorkbookSnapshot> {
  const cacheKey = JSON.stringify([...(dataSheetNames ?? [])].sort());
  if (structureCacheEnabled && structureCache?.key === cacheKey) {
    const cached = structuredClone(structureCache.snapshot);
    // 结构缓存只在工作表数据/集合变化时失效，切换活动工作表不会失效。
    // 命中缓存时用一次轻量选区读取刷新 activeWorksheet，避免拿过期值构建
    // 指纹，导致确认期与运行期不一致的误判。
    const selection = await captureSelectionContext();
    cached.activeWorksheet = selection.activeWorksheet;
    return cached;
  }
  return Excel.run(async (context) => {
    const workbook = context.workbook;
    const worksheets = workbook.worksheets;
    const activeWorksheet = worksheets.getActiveWorksheet();
    const selectedRange = workbook.getSelectedRange();

    worksheets.load("items/name");
    activeWorksheet.load("name");
    selectedRange.load("address");
    await context.sync();

    const pending = worksheets.items
      .filter((worksheet) => !isEBSystemSheet(worksheet.name))
      .map((worksheet) => {
        const usedRange = worksheet.getUsedRangeOrNullObject(true);
        usedRange.load(
          "address,rowCount,columnCount,rowIndex,columnIndex,isNullObject"
        );
        return { worksheet, usedRange };
      });
    await context.sync();

    const sheetsToRead = new Set(
      dataSheetNames?.length ? dataSheetNames : [activeWorksheet.name]
    );
    const previews = pending.map(({ worksheet, usedRange }) => {
      if (usedRange.isNullObject || !sheetsToRead.has(worksheet.name)) {
        return null;
      }
      const preview = worksheet.getRangeByIndexes(
        usedRange.rowIndex,
        usedRange.columnIndex,
        Math.min(usedRange.rowCount, STRUCTURE_ROW_LIMIT),
        Math.min(usedRange.columnCount, STRUCTURE_COLUMN_LIMIT)
      );
      preview.load("values");
      return preview;
    });
    await context.sync();

    const snapshot: WorkbookSnapshot = {
      name: workbookNameFromDocumentUrl(Office.context.document.url),
      capturedAt: new Date().toISOString(),
      activeWorksheet: activeWorksheet.name,
      selectedRange: selectedRange.address,
      worksheets: pending.map(({ worksheet, usedRange }, index) => {
        if (usedRange.isNullObject) {
          return {
            name: worksheet.name,
            usedRange: null,
            rowCount: 0,
            columnCount: 0,
            headers: [],
            dataRows: [],
            truncated: false
          };
        }
        const values = previews[index]?.values ?? [];
        return {
          name: worksheet.name,
          usedRange: usedRange.address,
          rowCount: usedRange.rowCount,
          columnCount: usedRange.columnCount,
          headers: detectSheetFields(
            values.map((row) => row.map(normalizeValue))
          ),
          dataRows: [],
          truncated:
            usedRange.rowCount > STRUCTURE_ROW_LIMIT ||
            usedRange.columnCount > STRUCTURE_COLUMN_LIMIT
        };
      })
    };
    const sourceFingerprintSheets = [...sheetsToRead];
    snapshot.sourceFingerprintSheets = sourceFingerprintSheets;
    snapshot.sourceFingerprint = sourceFingerprintForSnapshot(
      snapshot,
      sourceFingerprintSheets
    );
    if (structureCacheEnabled) {
      structureCache = { key: cacheKey, snapshot: structuredClone(snapshot) };
    }
    return snapshot;
  });
}

async function getOrCreateWorksheet(
  context: Excel.RequestContext,
  sheetName: string
): Promise<Excel.Worksheet> {
  const collection = context.workbook.worksheets;
  const existing = collection.getItemOrNullObject(sheetName);
  existing.load("isNullObject");
  await context.sync();
  return existing.isNullObject ? collection.add(sheetName) : existing;
}

function assertRectangular(values: CellValue[][]): void {
  if (values.length === 0) return;
  const width = values[0].length;
  if (width === 0 || values.some((row) => row.length !== width)) {
    throw new Error("写入数据必须是非空的规则二维数组");
  }
}

function columnNumber(name: string): number {
  return [...name.toUpperCase()].reduce(
    (result, character) => result * 26 + character.charCodeAt(0) - 64,
    0
  );
}

function columnName(number: number): string {
  let result = "";
  while (number > 0) {
    const remainder = (number - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    number = Math.floor((number - 1) / 26);
  }
  return result;
}

function matrixRange(
  startCell: string,
  rowCount: number,
  columnCount: number
): string | null {
  const match = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(startCell.trim());
  if (!match) return null;
  const startColumn = columnNumber(match[1]);
  const startRow = Number(match[2]);
  return `${match[1].toUpperCase()}${startRow}:${columnName(
    startColumn + columnCount - 1
  )}${startRow + rowCount - 1}`;
}

function inferredCriteria(
  plan: AnalysisPlan,
  dynamicCriteria: VerificationCriterion[] = []
): VerificationCriterion[] {
  const criteria: VerificationCriterion[] = [
    ...(plan.acceptanceCriteria ?? []),
    ...dynamicCriteria
  ];
  const criterionKeys = new Set(criteria.map((criterion) => JSON.stringify(criterion)));
  const addCriterion = (criterion: VerificationCriterion) => {
    const key = JSON.stringify(criterion);
    if (!criterionKeys.has(key)) {
      criteria.push(criterion);
      criterionKeys.add(key);
    }
  };
  const seenSheets = new Set<string>();
  for (const [actionIndex, action] of plan.actions.entries()) {
    const laterActions = plan.actions.slice(actionIndex + 1);
    const rangeIsSortedLater = (range: string) =>
      laterActions.some(
        (later) =>
          later.type === "sortRange" &&
          later.sheet.toLocaleLowerCase() === action.sheet.toLocaleLowerCase() &&
          normalizedRangeAddress(later.range) === normalizedRangeAddress(range)
      );
    const filterChangesLater = laterActions.some(
      (later) =>
        (later.type === "filterRange" || later.type === "clearFilter") &&
        later.sheet.toLocaleLowerCase() === action.sheet.toLocaleLowerCase()
    );
    const laterSameRange = <T extends ExcelAction["type"]>(type: T) =>
      laterActions.filter(
        (later): later is Extract<ExcelAction, { type: T }> =>
          later.type === type &&
          later.sheet.toLocaleLowerCase() === action.sheet.toLocaleLowerCase() &&
          "range" in later &&
          "range" in action &&
          normalizedRangeAddress(later.range) ===
            normalizedRangeAddress(action.range)
      );
    if (action.type === "deleteWorksheet") {
      addCriterion({ type: "worksheetMissing", sheet: action.sheet });
      continue;
    }
    if (!seenSheets.has(action.sheet)) {
      addCriterion({ type: "worksheetExists", sheet: action.sheet });
      seenSheets.add(action.sheet);
    }
    if (action.type === "writeValues" && !rangeIsSortedLater(action.range)) {
      addCriterion({
        type: "rangeEquals",
        sheet: action.sheet,
        range: action.range,
        expected: action.values
      });
    } else if (action.type === "writeTable") {
      const expected = [action.headers, ...action.rows];
      const range = matrixRange(
        action.startCell,
        expected.length,
        action.headers.length
      );
      if (range && !rangeIsSortedLater(range)) {
        addCriterion({
          type: "rangeEquals",
          sheet: action.sheet,
          range,
          expected
        });
      }
    } else if (action.type === "writeFormulas") {
      if (action.formulaR1C1) {
        addCriterion({
          type: "formulasR1C1Equal",
          sheet: action.sheet,
          range: action.range,
          expected: action.formulaR1C1
        });
      } else {
        addCriterion({
          type: "formulasEqual",
          sheet: action.sheet,
          range: action.range,
          expected: action.formulas
        });
      }
    } else if (
      action.type === "clearRange" &&
      (action.applyTo === "all" || action.applyTo === "contents")
    ) {
      addCriterion({
        type: "rangeEmpty",
        sheet: action.sheet,
        range: action.range
      });
    } else if (action.type === "sortRange") {
      addCriterion({
        type: "rangeSorted",
        sheet: action.sheet,
        range: action.range,
        keys: action.keys,
        hasHeaders: action.hasHeaders
      });
    } else if (action.type === "filterRange" && !filterChangesLater) {
      addCriterion({
        type: "filterApplied",
        sheet: action.sheet,
        range: action.range,
        column: action.column,
        values: action.values
      });
    } else if (action.type === "clearFilter" && !filterChangesLater) {
      addCriterion({ type: "filterCleared", sheet: action.sheet });
    } else if (action.type === "createTable") {
      addCriterion({
        type: "tableExists",
        sheet: action.sheet,
        range: action.range,
        name: action.name,
        hasHeaders: action.hasHeaders
      });
    } else if (
      action.type === "setFill" &&
      laterSameRange("setFill").length === 0
    ) {
      addCriterion({
        type: "rangeFormatMatches",
        sheet: action.sheet,
        range: action.range,
        fillColor: action.color
      });
    } else if (action.type === "setFont") {
      const laterFonts = laterSameRange("setFont");
      const bold = laterFonts.some((later) => later.bold !== undefined)
        ? undefined
        : action.bold;
      const fontColor = laterFonts.some((later) => Boolean(later.color))
        ? undefined
        : action.color;
      if (bold === undefined && !fontColor) continue;
      addCriterion({
        type: "rangeFormatMatches",
        sheet: action.sheet,
        range: action.range,
        bold,
        fontColor
      });
    } else if (
      action.type === "setNumberFormat" &&
      laterSameRange("setNumberFormat").length === 0
    ) {
      addCriterion({
        type: "rangeFormatMatches",
        sheet: action.sheet,
        range: action.range,
        numberFormat: action.formatCode
      });
    } else if (action.type === "setAlignment") {
      const laterAlignments = laterSameRange("setAlignment");
      const horizontal = laterAlignments.some((later) => Boolean(later.horizontal))
        ? undefined
        : action.horizontal;
      const vertical = laterAlignments.some((later) => Boolean(later.vertical))
        ? undefined
        : action.vertical;
      const wrapText = laterAlignments.some(
        (later) => later.wrapText !== undefined
      )
        ? undefined
        : action.wrapText;
      if (!horizontal && !vertical && wrapText === undefined) continue;
      addCriterion({
        type: "rangeFormatMatches",
        sheet: action.sheet,
        range: action.range,
        horizontal,
        vertical,
        wrapText
      });
    } else if (action.type === "resizeRange") {
      const laterResizes = laterSameRange("resizeRange");
      const rowHeight = laterResizes.some(
        (later) => later.rowHeight !== undefined
      )
        ? undefined
        : action.rowHeight;
      const columnWidth = laterResizes.some(
        (later) => later.columnWidth !== undefined
      )
        ? undefined
        : action.columnWidth;
      if (rowHeight === undefined && columnWidth === undefined) continue;
      addCriterion({
        type: "rangeFormatMatches",
        sheet: action.sheet,
        range: action.range,
        rowHeight,
        columnWidth
      });
    } else if (action.type === "setBorders") {
      const laterBorders = laterSameRange("setBorders");
      const sides = action.sides.filter(
        (side) => !laterBorders.some((later) => later.sides.includes(side))
      );
      if (sides.length === 0) continue;
      addCriterion({
        type: "bordersMatch",
        sheet: action.sheet,
        range: action.range,
        sides,
        style: action.style,
        color: action.color,
        weight: action.weight
      });
    } else if (
      action.type === "setDataValidation" &&
      laterSameRange("setDataValidation").length === 0
    ) {
      addCriterion({
        type: "dataValidationMatches",
        sheet: action.sheet,
        range: action.range,
        validationType: action.validationType,
        values: action.values,
        formula1: action.formula1,
        formula2: action.formula2,
        operator: action.operator,
        allowBlank: action.allowBlank,
        prompt: action.prompt,
        errorMessage: action.errorMessage
      });
    } else if (
      action.type === "freezePanes" &&
      !laterActions.some(
        (later) =>
          later.type === "freezePanes" &&
          later.sheet.toLocaleLowerCase() === action.sheet.toLocaleLowerCase()
      )
    ) {
      addCriterion({
        type: "freezePanesMatches",
        sheet: action.sheet,
        rows: action.rows,
        columns: action.columns
      });
    } else if (action.type === "createPivotTable") {
      addCriterion({
        type: "pivotTableExists",
        sheet: action.sheet,
        sourceSheet: action.sourceSheet,
        sourceRange: action.sourceRange,
        name: action.name,
        destinationCell: action.destinationCell,
        rowFields: action.rowFields,
        columnFields: action.columnFields,
        valueFields: action.valueFields
      });
    }
  }
  return criteria;
}

function normalizedRangeAddress(address: string): string {
  const localAddress = address.includes("!")
    ? address.slice(address.lastIndexOf("!") + 1)
    : address;
  return localAddress.replace(/\$/g, "").trim().toLocaleUpperCase();
}

export function verificationGaps(
  plan: AnalysisPlan,
  dynamicCriteria: VerificationCriterion[] = []
): UnverifiedAction[] {
  const criteria = [
    ...(plan.acceptanceCriteria ?? []),
    ...dynamicCriteria
  ];
  const hasCriterion = (
    predicate: (criterion: VerificationCriterion) => boolean
  ) => criteria.some(predicate);
  const sameRange = (
    criterion: VerificationCriterion,
    sheet: string,
    range: string
  ) =>
    "range" in criterion &&
    criterion.sheet.toLocaleLowerCase() === sheet.toLocaleLowerCase() &&
    normalizedRangeAddress(criterion.range) ===
      normalizedRangeAddress(range);

  return plan.actions.flatMap((action, index) => {
    let verified = false;
    const laterActions = plan.actions.slice(index + 1);
    const laterSameRange = (type: ExcelAction["type"]) =>
      laterActions.filter(
        (later) =>
          later.type === type &&
          later.sheet.toLocaleLowerCase() === action.sheet.toLocaleLowerCase() &&
          "range" in later &&
          "range" in action &&
          normalizedRangeAddress(later.range) ===
            normalizedRangeAddress(action.range)
      );
    switch (action.type) {
      case "createWorksheet":
        verified = hasCriterion(
          (criterion) =>
            criterion.type === "worksheetExists" &&
            criterion.sheet.toLocaleLowerCase() ===
              action.sheet.toLocaleLowerCase()
        );
        break;
      case "deleteWorksheet":
        verified = hasCriterion(
          (criterion) =>
            criterion.type === "worksheetMissing" &&
            criterion.sheet.toLocaleLowerCase() ===
              action.sheet.toLocaleLowerCase()
        );
        break;
      case "writeValues":
        verified = hasCriterion(
          (criterion) =>
            criterion.type === "rangeEquals" &&
            sameRange(criterion, action.sheet, action.range)
        );
        break;
      case "writeFormulas":
        verified = hasCriterion(
          (criterion) =>
            (criterion.type === "formulasEqual" ||
              criterion.type === "formulasR1C1Equal") &&
            sameRange(criterion, action.sheet, action.range)
        );
        break;
      case "writeTable": {
        const range = matrixRange(
          action.startCell,
          action.rows.length + 1,
          action.headers.length
        );
        verified =
          range !== null &&
          hasCriterion(
            (criterion) =>
              criterion.type === "rangeEquals" &&
              sameRange(criterion, action.sheet, range)
          );
        break;
      }
      case "clearRange":
        verified =
          (action.applyTo === "all" || action.applyTo === "contents") &&
          hasCriterion(
            (criterion) =>
              criterion.type === "rangeEmpty" &&
              sameRange(criterion, action.sheet, action.range)
          );
        break;
      case "copyRange":
        verified = hasCriterion(
          (criterion) =>
            (criterion.type === "rangeEquals" ||
              criterion.type === "formulasEqual") &&
            sameRange(criterion, action.sheet, action.targetRange)
        );
        break;
      case "splitGroupAggregate":
        verified = dynamicCriteria.length > 0;
        break;
      case "sortRange":
        verified = hasCriterion(
          (criterion) =>
            criterion.type === "rangeSorted" &&
            sameRange(criterion, action.sheet, action.range)
        );
        break;
      case "filterRange":
        verified =
          laterActions.some(
            (later) =>
              (later.type === "filterRange" || later.type === "clearFilter") &&
              later.sheet.toLocaleLowerCase() === action.sheet.toLocaleLowerCase()
          ) ||
          hasCriterion(
          (criterion) =>
            criterion.type === "filterApplied" &&
            sameRange(criterion, action.sheet, action.range) &&
            criterion.column === action.column
          );
        break;
      case "clearFilter":
        verified =
          laterActions.some(
            (later) =>
              (later.type === "filterRange" || later.type === "clearFilter") &&
              later.sheet.toLocaleLowerCase() === action.sheet.toLocaleLowerCase()
          ) ||
          hasCriterion(
          (criterion) =>
            criterion.type === "filterCleared" &&
            criterion.sheet.toLocaleLowerCase() ===
              action.sheet.toLocaleLowerCase()
          );
        break;
      case "createTable":
        verified = hasCriterion(
          (criterion) =>
            criterion.type === "tableExists" &&
            sameRange(criterion, action.sheet, action.range)
        );
        break;
      case "setFill":
        verified =
          laterSameRange("setFill").length > 0 ||
          hasCriterion(
          (criterion) =>
            criterion.type === "rangeFormatMatches" &&
            sameRange(criterion, action.sheet, action.range) &&
            criterion.fillColor?.toLocaleLowerCase() ===
              action.color.toLocaleLowerCase()
          );
        break;
      case "setFont":
        verified =
          (action.bold === undefined ||
            laterSameRange("setFont").some(
              (later) =>
                later.type === "setFont" && later.bold !== undefined
            ) ||
            hasCriterion(
              (criterion) =>
                criterion.type === "rangeFormatMatches" &&
                sameRange(criterion, action.sheet, action.range) &&
                criterion.bold === action.bold
            )) &&
          (!action.color ||
            laterSameRange("setFont").some(
              (later) => later.type === "setFont" && Boolean(later.color)
            ) ||
            hasCriterion(
              (criterion) =>
                criterion.type === "rangeFormatMatches" &&
                sameRange(criterion, action.sheet, action.range) &&
                criterion.fontColor?.toLocaleLowerCase() ===
                  action.color?.toLocaleLowerCase()
            ));
        break;
      case "setNumberFormat":
        verified =
          laterSameRange("setNumberFormat").length > 0 ||
          hasCriterion(
          (criterion) =>
            criterion.type === "rangeFormatMatches" &&
            sameRange(criterion, action.sheet, action.range) &&
            criterion.numberFormat === action.formatCode
          );
        break;
      case "setAlignment":
        verified =
          (!action.horizontal ||
            laterSameRange("setAlignment").some(
              (later) =>
                later.type === "setAlignment" && Boolean(later.horizontal)
            ) ||
            hasCriterion(
              (criterion) =>
                criterion.type === "rangeFormatMatches" &&
                sameRange(criterion, action.sheet, action.range) &&
                criterion.horizontal === action.horizontal
            )) &&
          (!action.vertical ||
            laterSameRange("setAlignment").some(
              (later) =>
                later.type === "setAlignment" && Boolean(later.vertical)
            ) ||
            hasCriterion(
              (criterion) =>
                criterion.type === "rangeFormatMatches" &&
                sameRange(criterion, action.sheet, action.range) &&
                criterion.vertical === action.vertical
            )) &&
          (action.wrapText === undefined ||
            laterSameRange("setAlignment").some(
              (later) =>
                later.type === "setAlignment" &&
                later.wrapText !== undefined
            ) ||
            hasCriterion(
              (criterion) =>
                criterion.type === "rangeFormatMatches" &&
                sameRange(criterion, action.sheet, action.range) &&
                criterion.wrapText === action.wrapText
            ));
        break;
      case "resizeRange":
        verified =
          (action.rowHeight === undefined ||
            laterSameRange("resizeRange").some(
              (later) =>
                later.type === "resizeRange" && later.rowHeight !== undefined
            ) ||
            hasCriterion(
              (criterion) =>
                criterion.type === "rangeFormatMatches" &&
                sameRange(criterion, action.sheet, action.range) &&
                criterion.rowHeight === action.rowHeight
            )) &&
          (action.columnWidth === undefined ||
            laterSameRange("resizeRange").some(
              (later) =>
                later.type === "resizeRange" &&
                later.columnWidth !== undefined
            ) ||
            hasCriterion(
              (criterion) =>
                criterion.type === "rangeFormatMatches" &&
                sameRange(criterion, action.sheet, action.range) &&
                criterion.columnWidth === action.columnWidth
            ));
        break;
      case "setBorders":
        verified = hasCriterion(
          (criterion) =>
            criterion.type === "bordersMatch" &&
            sameRange(criterion, action.sheet, action.range)
        );
        break;
      case "setDataValidation":
        verified =
          laterSameRange("setDataValidation").length > 0 ||
          hasCriterion(
          (criterion) =>
            criterion.type === "dataValidationMatches" &&
            sameRange(criterion, action.sheet, action.range)
          );
        break;
      case "freezePanes":
        verified =
          laterActions.some(
            (later) =>
              later.type === "freezePanes" &&
              later.sheet.toLocaleLowerCase() === action.sheet.toLocaleLowerCase()
          ) ||
          hasCriterion(
          (criterion) =>
            criterion.type === "freezePanesMatches" &&
            criterion.sheet.toLocaleLowerCase() ===
              action.sheet.toLocaleLowerCase() &&
            criterion.rows === action.rows &&
            criterion.columns === action.columns
          );
        break;
      case "createChart":
        verified = hasCriterion(
          (criterion) =>
            criterion.type === "chartExists" &&
            criterion.sheet.toLocaleLowerCase() ===
              action.sheet.toLocaleLowerCase() &&
            normalizedRangeAddress(criterion.sourceRange) ===
              normalizedRangeAddress(action.sourceRange)
        );
        break;
      case "createPivotTable":
        verified = hasCriterion(
          (criterion) =>
            criterion.type === "pivotTableExists" &&
            criterion.sheet.toLocaleLowerCase() ===
              action.sheet.toLocaleLowerCase() &&
            criterion.name.toLocaleLowerCase() === action.name.toLocaleLowerCase()
        );
        break;
    }
    return verified
      ? []
      : [
          {
            index,
            type: action.type,
            sheet: action.sheet,
            message: `第 ${index + 1} 步 ${action.type} 已执行，但当前验收协议不能独立验证该操作的具体效果`
          }
        ];
  });
}

export function valuesEqual(actual: CellValue, expected: CellValue): boolean {
  if ((actual === null || actual === "") && (expected === null || expected === "")) {
    return true;
  }
  if (typeof actual === "number" && typeof expected === "number") {
    return (
      actual === expected ||
      Math.abs(actual - expected) <=
        Number.EPSILON * 16 * Math.max(1, Math.abs(actual), Math.abs(expected))
    );
  }
  if (
    (typeof actual === "number" && typeof expected === "string") ||
    (typeof actual === "string" && typeof expected === "number")
  ) {
    const numberValue = typeof actual === "number" ? actual : expected;
    const textValue = String(
      typeof actual === "string" ? actual : expected
    ).trim();
    if (!textValue || !Number.isFinite(Number(textValue))) return false;
    const losesLeadingZero =
      /^[-+]?0\d/.test(textValue) && !/^[-+]?0(?:\.\d+)?$/.test(textValue);
    return !losesLeadingZero && Number(textValue) === numberValue;
  }
  return actual === expected;
}

function matricesEqual(actual: CellValue[][], expected: CellValue[][]): boolean {
  return (
    actual.length === expected.length &&
    actual.every(
      (row, rowIndex) =>
        row.length === expected[rowIndex].length &&
        row.every((value, columnIndex) =>
          valuesEqual(value, expected[rowIndex][columnIndex])
        )
    )
  );
}

function firstMatrixDifference(
  actual: CellValue[][],
  expected: CellValue[][]
): string | null {
  if (actual.length !== expected.length) {
    return `实际 ${actual.length} 行，预期 ${expected.length} 行`;
  }
  for (let rowIndex = 0; rowIndex < expected.length; rowIndex += 1) {
    if (actual[rowIndex].length !== expected[rowIndex].length) {
      return `第 ${rowIndex + 1} 行实际 ${actual[rowIndex].length} 列，预期 ${
        expected[rowIndex].length
      } 列`;
    }
    for (
      let columnIndex = 0;
      columnIndex < expected[rowIndex].length;
      columnIndex += 1
    ) {
      if (
        !valuesEqual(
          actual[rowIndex][columnIndex],
          expected[rowIndex][columnIndex]
        )
      ) {
        return `${columnName(columnIndex + 1)}${rowIndex + 1} 实际 ${JSON.stringify(
          actual[rowIndex][columnIndex]
        )}，预期 ${JSON.stringify(expected[rowIndex][columnIndex])}`;
      }
    }
  }
  return null;
}

function isBlankMatrix(values: CellValue[][]): boolean {
  return values.every((row) =>
    row.every((value) => value === null || value === "")
  );
}

function compareSortValues(left: CellValue, right: CellValue): number {
  const leftBlank = left === null || left === "";
  const rightBlank = right === null || right === "";
  if (leftBlank || rightBlank) {
    if (leftBlank && rightBlank) return 0;
    return leftBlank ? 1 : -1;
  }
  if (typeof left === "number" && typeof right === "number") {
    return left - right;
  }
  if (typeof left === "boolean" && typeof right === "boolean") {
    return Number(left) - Number(right);
  }
  return String(left).localeCompare(String(right), undefined, {
    sensitivity: "base"
  });
}

export function sortedRangePasses(
  values: CellValue[][],
  keys: Array<{ column: number; ascending: boolean }>,
  hasHeaders: boolean
): boolean {
  const rows = hasHeaders ? values.slice(1) : values;
  if (
    keys.length === 0 ||
    keys.some((key) => key.column < 0 || rows.some((row) => key.column >= row.length))
  ) {
    return false;
  }
  for (let index = 1; index < rows.length; index += 1) {
    for (const key of keys) {
      const left = rows[index - 1][key.column];
      const right = rows[index][key.column];
      const comparison = compareSortValues(left, right);
      if (comparison === 0) continue;
      const hasBlank =
        left === null || left === "" || right === null || right === "";
      if (hasBlank ? comparison > 0 : key.ascending ? comparison > 0 : comparison < 0) {
        return false;
      }
      break;
    }
  }
  return true;
}

function filterValueText(value: unknown): string {
  if (
    value &&
    typeof value === "object" &&
    "date" in value &&
    typeof (value as { date?: unknown }).date === "string"
  ) {
    return (value as { date: string }).date;
  }
  return String(value);
}

function sameFilterValues(actual: unknown[] | undefined, expected: CellValue[]): boolean {
  if (!actual) return false;
  const actualValues = actual.map(filterValueText).sort();
  const expectedValues = expected.map(String).sort();
  return (
    actualValues.length === expectedValues.length &&
    actualValues.every((value, index) => value === expectedValues[index])
  );
}

function normalizedColor(value: string | null | undefined): string {
  const text = (value ?? "").trim().replace(/^#/, "");
  return text.length >= 6 ? `#${text.slice(-6).toLocaleUpperCase()}` : text;
}

function normalizedEnum(value: unknown): string {
  return String(value ?? "")
    .replace(/[-_\s]/g, "")
    .toLocaleLowerCase();
}

function closeEnough(actual: number, expected: number): boolean {
  return Math.abs(actual - expected) <= 0.1;
}

function validationFormulaEqual(
  actual: unknown,
  expected: string | number | null | undefined
): boolean {
  if (expected === undefined || expected === null) {
    return actual === undefined || actual === null || actual === "";
  }
  const actualText = String(actual ?? "");
  if (typeof expected === "number") {
    return Number(actualText.replace(/^=/, "")) === expected;
  }
  return actualText === expected;
}

function dataValidationRuleMatches(
  actual: Excel.DataValidationRule,
  criterion: Extract<
    VerificationCriterion,
    { type: "dataValidationMatches" }
  >
): boolean {
  if (criterion.validationType === "list") {
    return (
      actual.list?.source === criterion.values.map(String).join(",") &&
      actual.list.inCellDropDown
    );
  }
  if (criterion.validationType === "custom") {
    return validationFormulaEqual(actual.custom?.formula, criterion.formula1);
  }
  const rule = actual[
    criterion.validationType
  ] as Excel.BasicDataValidation | Excel.DateTimeDataValidation | undefined;
  return Boolean(
    rule &&
      normalizedEnum(rule.operator) === normalizedEnum(criterion.operator) &&
      validationFormulaEqual(rule.formula1, criterion.formula1) &&
      validationFormulaEqual(rule.formula2, criterion.formula2)
  );
}

function textMatrix(values: CellValue[][]): string[][] {
  return values.map((row) =>
    row.map((value) => (value === null ? "" : String(value)))
  );
}

async function chartSourceMatches(
  context: Excel.RequestContext,
  chart: Excel.Chart,
  sourceRange: Excel.Range
): Promise<boolean> {
  chart.load("plotBy,chartType");
  chart.series.load("items/name");
  sourceRange.load("values");
  await context.sync();
  const source = textMatrix(
    sourceRange.values.map((row) => row.map(normalizeValue))
  );
  if (source.length < 2 || source[0].length < 2) return false;
  const byRows = normalizedEnum(chart.plotBy) === "rows";
  const expectedSeries = byRows
    ? source.slice(1).map((row) => ({
        name: row[0],
        categories: source[0].slice(1),
        values: row.slice(1)
      }))
    : source[0].slice(1).map((name, columnIndex) => ({
        name,
        categories: source.slice(1).map((row) => row[0]),
        values: source.slice(1).map((row) => row[columnIndex + 1])
      }));
  if (chart.series.items.length !== expectedSeries.length) return false;
  const scatter = ["xyscatter", "bubble"].some((type) =>
    normalizedEnum(chart.chartType).startsWith(type)
  );
  const results = chart.series.items.map((series) => ({
    categories: series.getDimensionValues(
      scatter ? "XValues" : "Categories"
    ),
    values: series.getDimensionValues(scatter ? "YValues" : "Values")
  }));
  await context.sync();
  return results.every((result, index) => {
    const expected = expectedSeries[index];
    return (
      result.categories.value.length === expected.categories.length &&
      result.categories.value.every(
        (value, valueIndex) => value === expected.categories[valueIndex]
      ) &&
      result.values.value.length === expected.values.length &&
      result.values.value.every(
        (value, valueIndex) => value === expected.values[valueIndex]
      )
    );
  });
}

function normalizedSourceReference(value: string): string {
  return value
    .replace(/\$/g, "")
    .replace(/'/g, "")
    .replace(/\[[^\]]+\]/g, "")
    .replace(/\s/g, "")
    .toLocaleUpperCase();
}

function rangeStartAddress(address: string): string {
  const localAddress = address.includes("!")
    ? address.slice(address.lastIndexOf("!") + 1)
    : address;
  return normalizedRangeAddress(localAddress.split(":")[0]);
}

async function verifyPlan(
  context: Excel.RequestContext,
  plan: AnalysisPlan,
  dynamicCriteria: VerificationCriterion[] = []
): Promise<VerificationReport> {
  const checks: VerificationCheck[] = [];
  const criteria = inferredCriteria(plan, dynamicCriteria);
  const sheetsByName = new Map<string, Excel.Worksheet>();
  for (const sheetName of new Set(criteria.map((criterion) => criterion.sheet))) {
    const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
    sheet.load("isNullObject");
    sheetsByName.set(sheetName, sheet);
  }
  await context.sync();
  const valueRanges = new Map<VerificationCriterion, Excel.Range>();
  for (const criterion of criteria) {
    if (
      criterion.type !== "rangeEquals" &&
      criterion.type !== "rangeEmpty" &&
      criterion.type !== "formulasEqual" &&
      criterion.type !== "formulasR1C1Equal" &&
      criterion.type !== "rangeSorted"
    ) {
      continue;
    }
    const sheet = sheetsByName.get(criterion.sheet)!;
    if (sheet.isNullObject) continue;
    const range = sheet.getRange(criterion.range);
    range.load(
      criterion.type === "formulasEqual"
        ? "formulas"
        : criterion.type === "formulasR1C1Equal"
          ? "formulasR1C1"
          : "values"
    );
    valueRanges.set(criterion, range);
  }
  if (valueRanges.size > 0) await context.sync();

  for (const criterion of criteria) {
    const sheet = sheetsByName.get(criterion.sheet)!;
    if (criterion.type === "worksheetMissing") {
      checks.push({
        criterion,
        passed: sheet.isNullObject,
        message: sheet.isNullObject
          ? `工作表「${criterion.sheet}」已删除`
          : `工作表「${criterion.sheet}」仍然存在`
      });
      continue;
    }
    if (sheet.isNullObject) {
      checks.push({
        criterion,
        passed: false,
        message: `未找到工作表「${criterion.sheet}」`
      });
      continue;
    }
    if (criterion.type === "worksheetExists") {
      checks.push({
        criterion,
        passed: true,
        message: `工作表「${criterion.sheet}」存在`
      });
      continue;
    }

    if (criterion.type === "filterApplied") {
      const autoFilter = sheet.autoFilter;
      const filterRange = autoFilter.getRangeOrNullObject();
      autoFilter.load("enabled,isDataFiltered,criteria");
      filterRange.load("isNullObject,address");
      await context.sync();
      const appliedCriterion = autoFilter.criteria[criterion.column];
      const passed =
        autoFilter.enabled &&
        autoFilter.isDataFiltered &&
        !filterRange.isNullObject &&
        normalizedRangeAddress(filterRange.address) ===
          normalizedRangeAddress(criterion.range) &&
        appliedCriterion?.filterOn === "Values" &&
        sameFilterValues(appliedCriterion.values, criterion.values);
      checks.push({
        criterion,
        passed,
        message: passed
          ? `「${criterion.sheet}」${criterion.range} 第 ${criterion.column + 1} 列筛选条件一致`
          : `「${criterion.sheet}」${criterion.range} 的筛选范围或条件与预期不一致`
      });
      continue;
    }
    if (criterion.type === "filterCleared") {
      const autoFilter = sheet.autoFilter;
      autoFilter.load("isDataFiltered");
      await context.sync();
      const passed = !autoFilter.isDataFiltered;
      checks.push({
        criterion,
        passed,
        message: passed
          ? `工作表「${criterion.sheet}」的筛选条件已清除`
          : `工作表「${criterion.sheet}」仍有筛选条件`
      });
      continue;
    }
    if (criterion.type === "tableExists") {
      const tables = sheet.tables;
      tables.load("items/name,items/showHeaders");
      await context.sync();
      const tableRanges = tables.items.map((table) => {
        const range = table.getRange();
        range.load("address");
        return { table, range };
      });
      await context.sync();
      const match = tableRanges.find(
        ({ table, range }) =>
          (!criterion.name ||
            table.name.toLocaleLowerCase() === criterion.name.toLocaleLowerCase()) &&
          normalizedRangeAddress(range.address) ===
            normalizedRangeAddress(criterion.range)
      );
      const passed = Boolean(
        match && match.table.showHeaders === criterion.hasHeaders
      );
      checks.push({
        criterion,
        passed,
        message: passed
          ? `「${criterion.sheet}」${criterion.range} 的表格范围与表头状态一致`
          : `「${criterion.sheet}」${criterion.range} 未找到符合预期的表格`
      });
      continue;
    }
    if (criterion.type === "rangeFormatMatches") {
      const range = sheet.getRange(criterion.range);
      const format = range.format;
      const fill = format.fill;
      const font = format.font;
      range.load("numberFormat");
      format.load(
        "horizontalAlignment,verticalAlignment,wrapText,rowHeight,columnWidth"
      );
      fill.load("color");
      font.load("bold,color");
      await context.sync();
      const numberFormatMatches =
        criterion.numberFormat == null ||
        range.numberFormat.every((row) =>
          row.every((value) => value === criterion.numberFormat)
        );
      const passed =
        (criterion.fillColor == null ||
          normalizedColor(fill.color) === normalizedColor(criterion.fillColor)) &&
        (criterion.bold == null || font.bold === criterion.bold) &&
        (criterion.fontColor == null ||
          normalizedColor(font.color) === normalizedColor(criterion.fontColor)) &&
        numberFormatMatches &&
        (criterion.horizontal == null ||
          normalizedEnum(format.horizontalAlignment) ===
            normalizedEnum(criterion.horizontal)) &&
        (criterion.vertical == null ||
          normalizedEnum(format.verticalAlignment) ===
            normalizedEnum(criterion.vertical)) &&
        (criterion.wrapText == null || format.wrapText === criterion.wrapText) &&
        (criterion.rowHeight == null ||
          closeEnough(format.rowHeight, criterion.rowHeight)) &&
        (criterion.columnWidth == null ||
          closeEnough(format.columnWidth, criterion.columnWidth));
      checks.push({
        criterion,
        passed,
        message: passed
          ? `「${criterion.sheet}」${criterion.range} 的格式属性一致`
          : `「${criterion.sheet}」${criterion.range} 的格式属性与预期不一致`
      });
      continue;
    }
    if (criterion.type === "bordersMatch") {
      const borderIndexes = {
        top: "EdgeTop",
        bottom: "EdgeBottom",
        left: "EdgeLeft",
        right: "EdgeRight",
        insideHorizontal: "InsideHorizontal",
        insideVertical: "InsideVertical"
      } as const;
      const borders = criterion.sides.map((side) => {
        const border = sheet
          .getRange(criterion.range)
          .format.borders.getItem(borderIndexes[side]);
        border.load("style,color,weight");
        return border;
      });
      await context.sync();
      const passed = borders.every(
        (border) =>
          normalizedEnum(border.style) === normalizedEnum(criterion.style) &&
          normalizedColor(border.color) === normalizedColor(criterion.color) &&
          normalizedEnum(border.weight) === normalizedEnum(criterion.weight)
      );
      checks.push({
        criterion,
        passed,
        message: passed
          ? `「${criterion.sheet}」${criterion.range} 的边框属性一致`
          : `「${criterion.sheet}」${criterion.range} 的边框属性与预期不一致`
      });
      continue;
    }
    if (criterion.type === "dataValidationMatches") {
      const validation = sheet.getRange(criterion.range).dataValidation;
      validation.load("type,rule,ignoreBlanks,prompt,errorAlert");
      await context.sync();
      const passed =
        normalizedEnum(validation.type) ===
          normalizedEnum(criterion.validationType) &&
        validation.ignoreBlanks === criterion.allowBlank &&
        dataValidationRuleMatches(validation.rule, criterion) &&
        (criterion.prompt == null ||
          (validation.prompt.showPrompt &&
            validation.prompt.message === criterion.prompt)) &&
        (criterion.errorMessage == null ||
          (validation.errorAlert.showAlert &&
            validation.errorAlert.message === criterion.errorMessage));
      checks.push({
        criterion,
        passed,
        message: passed
          ? `「${criterion.sheet}」${criterion.range} 的数据验证规则一致`
          : `「${criterion.sheet}」${criterion.range} 的数据验证规则与预期不一致`
      });
      continue;
    }
    if (criterion.type === "freezePanesMatches") {
      const location = sheet.freezePanes.getLocationOrNullObject();
      location.load("isNullObject,rowCount,columnCount");
      await context.sync();
      const actualRows = location.isNullObject
        ? 0
        : location.rowCount === 1048576
          ? 0
          : location.rowCount;
      const actualColumns = location.isNullObject
        ? 0
        : location.columnCount === 16384
          ? 0
          : location.columnCount;
      const passed =
        actualRows === criterion.rows && actualColumns === criterion.columns;
      checks.push({
        criterion,
        passed,
        message: passed
          ? `工作表「${criterion.sheet}」的冻结窗格位置一致`
          : `工作表「${criterion.sheet}」的冻结窗格位置与预期不一致`
      });
      continue;
    }
    if (criterion.type === "chartExists") {
      const chart = criterion.name
        ? sheet.charts.getItemOrNullObject(criterion.name)
        : null;
      if (!chart) {
        checks.push({
          criterion,
          passed: false,
          message: `图表验收缺少执行后生成的对象名称`
        });
        continue;
      }
      chart.load("isNullObject,name,chartType,left,top");
      await context.sync();
      if (chart.isNullObject) {
        checks.push({
          criterion,
          passed: false,
          message: `未找到图表「${criterion.name}」`
        });
        continue;
      }
      chart.title.load("text,visible");
      await context.sync();
      const sourceMatches = await chartSourceMatches(
        context,
        chart,
        sheet.getRange(criterion.sourceRange)
      );
      let positionMatches = true;
      if (criterion.targetRange) {
        const target = sheet.getRange(criterion.targetRange);
        target.load("left,top");
        await context.sync();
        positionMatches =
          closeEnough(chart.left, target.left) &&
          closeEnough(chart.top, target.top);
      }
      const passed =
        normalizedEnum(chart.chartType) === normalizedEnum(criterion.chartType) &&
        (criterion.title == null ||
          (chart.title.visible && chart.title.text === criterion.title)) &&
        sourceMatches &&
        positionMatches;
      checks.push({
        criterion,
        passed,
        message: passed
          ? `图表「${chart.name}」的类型、数据源和位置一致`
          : `图表「${chart.name}」的类型、数据源、标题或位置与预期不一致`
      });
      continue;
    }
    if (criterion.type === "pivotTableExists") {
      const pivot = sheet.pivotTables.getItemOrNullObject(criterion.name);
      pivot.load("isNullObject,name");
      await context.sync();
      if (pivot.isNullObject) {
        checks.push({
          criterion,
          passed: false,
          message: `未找到数据透视表「${criterion.name}」`
        });
        continue;
      }
      const layoutRange = pivot.layout.getRange();
      layoutRange.load("address");
      pivot.rowHierarchies.load("items/name,items/position");
      pivot.columnHierarchies.load("items/name,items/position");
      pivot.dataHierarchies.load("items/name,items/position,items/summarizeBy");
      const source = pivot.getDataSourceString();
      await context.sync();
      for (const hierarchy of pivot.dataHierarchies.items) {
        hierarchy.field.load("name");
      }
      await context.sync();
      const orderedNames = (
        items: Array<{ name: string; position: number }>
      ) =>
        [...items]
          .sort((left, right) => left.position - right.position)
          .map((item) => item.name);
      const rowFields = orderedNames(pivot.rowHierarchies.items);
      const columnFields = orderedNames(pivot.columnHierarchies.items);
      const valueFields = [...pivot.dataHierarchies.items]
        .sort((left, right) => left.position - right.position)
        .map((item) => ({
          field: item.field.name,
          aggregation: normalizedEnum(item.summarizeBy)
        }));
      const expectedSource = normalizedSourceReference(
        `${criterion.sourceSheet}!${criterion.sourceRange}`
      );
      const actualSource = normalizedSourceReference(source.value);
      const passed =
        actualSource.endsWith(expectedSource) &&
        rangeStartAddress(layoutRange.address) ===
          rangeStartAddress(criterion.destinationCell) &&
        JSON.stringify(rowFields) === JSON.stringify(criterion.rowFields) &&
        JSON.stringify(columnFields) ===
          JSON.stringify(criterion.columnFields) &&
        valueFields.length === criterion.valueFields.length &&
        valueFields.every(
          (value, index) =>
            value.field === criterion.valueFields[index].field &&
            value.aggregation ===
              normalizedEnum(criterion.valueFields[index].aggregation)
        );
      checks.push({
        criterion,
        passed,
        message: passed
          ? `数据透视表「${criterion.name}」的数据源、位置和字段配置一致`
          : `数据透视表「${criterion.name}」的数据源、位置或字段配置与预期不一致`
      });
      continue;
    }

    const range = valueRanges.get(criterion)!;
    const actual =
      criterion.type === "formulasEqual"
        ? range.formulas.map((row) => row.map(normalizeValue))
        : criterion.type === "formulasR1C1Equal"
          ? range.formulasR1C1.map((row) => row.map(normalizeValue))
          : range.values.map((row) => row.map(normalizeValue));
    // R1C1 验收：期望值是与区域同形、逐格相同的矩阵，复用矩阵比较与差异定位
    const expected =
      criterion.type === "formulasR1C1Equal"
        ? actual.map((row) => row.map(() => normalizeValue(criterion.expected)))
        : criterion.type === "rangeEquals" || criterion.type === "formulasEqual"
          ? criterion.expected
          : [];
    const passed =
      criterion.type === "rangeEmpty"
        ? isBlankMatrix(actual)
        : criterion.type === "rangeSorted"
          ? sortedRangePasses(actual, criterion.keys, criterion.hasHeaders)
          : matricesEqual(actual, expected);
    const difference =
      !passed &&
      criterion.type !== "rangeEmpty" &&
      criterion.type !== "rangeSorted"
        ? firstMatrixDifference(actual, expected)
        : null;
    checks.push({
      criterion,
      passed,
      message: passed
        ? criterion.type === "rangeEmpty"
          ? `「${criterion.sheet}」${criterion.range} 已清空`
          : criterion.type === "rangeSorted"
            ? `「${criterion.sheet}」${criterion.range} 排序顺序一致`
          : criterion.type === "formulasEqual" ||
              criterion.type === "formulasR1C1Equal"
            ? `「${criterion.sheet}」${criterion.range} 公式一致`
            : `「${criterion.sheet}」${criterion.range} 写入值一致`
        : `「${criterion.sheet}」${criterion.range} ${
            criterion.type === "rangeSorted" ? "排序顺序" : "内容"
          }与预期不一致${
            difference ? `（${difference}）` : ""
          }`,
      actual
    });
  }
  const unverifiedActions = verificationGaps(plan, dynamicCriteria);
  const status: VerificationReport["status"] = checks.some(
    (check) => !check.passed
  )
    ? "failed"
    : unverifiedActions.length > 0 || checks.length === 0
      ? "executed_unverified"
      : "verified";
  return {
    status,
    passed: status === "verified",
    checks,
    unverifiedActions
  };
}

function enumText(value: string): string {
  return value
    .trim()
    .replace(/[-_\s]+(.)/g, (_, character: string) => character.toUpperCase())
    .replace(/^./, (character) => character.toUpperCase());
}

function dataValidationRule(
  action: Extract<ExcelAction, { type: "setDataValidation" }>
): Excel.DataValidationRule {
  if (action.validationType === "list") {
    return {
      list: {
        inCellDropDown: true,
        source: action.values.map(String).join(",")
      }
    };
  }
  if (action.validationType === "custom") {
    return { custom: { formula: String(action.formula1 ?? "") } };
  }
  const basic = {
    formula1: action.formula1 ?? 0,
    ...(action.formula2 !== undefined && action.formula2 !== null
      ? { formula2: action.formula2 }
      : {}),
    operator: enumText(action.operator) as Excel.DataValidationOperator
  };
  if (action.validationType === "date") {
    return {
      date: {
        formula1: String(basic.formula1),
        operator: basic.operator,
        ...(basic.formula2 !== undefined
          ? { formula2: String(basic.formula2) }
          : {})
      }
    };
  }
  return { [action.validationType]: basic } as Excel.DataValidationRule;
}

function positionShape(
  shape: Excel.Shape,
  target: Excel.Range
): void {
  shape.left = target.left;
  shape.top = target.top;
  shape.width = target.width;
  shape.height = target.height;
  shape.placement = "TwoCell";
}

function minimumExcelApiVersion(action: ExcelAction): string {
  const versions: Partial<Record<ExcelAction["type"], string>> = {
    sortRange: "1.2",
    setConditionalFormat: "1.6",
    freezePanes: "1.7",
    setHyperlink: "1.7",
    setDataValidation: "1.8",
    createChart: "1.12",
    createPivotTable: "1.15",
    copyRange: "1.9",
    filterRange: "1.9",
    clearFilter: "1.9",
    addImage: "1.9",
    addShape: "1.9",
    addComment: "1.10",
    addNote: "1.18"
  };
  return versions[action.type] ?? "1.1";
}

function normalizedObjectName(value: string): string {
  return value.trim().toLocaleLowerCase();
}

function explicitA1RangeShape(
  address: string
): { rows: number; columns: number } | null {
  const match = address
    .trim()
    .match(
      /^\$?([A-Z]{1,3})\$?([1-9]\d*)(?::\$?([A-Z]{1,3})\$?([1-9]\d*))?$/i
    );
  if (!match) return null;
  const startColumn = columnNumber(match[1]);
  const startRow = Number(match[2]);
  const endColumn = columnNumber(match[3] ?? match[1]);
  const endRow = Number(match[4] ?? match[2]);
  if (
    startColumn > 16_384 ||
    endColumn > 16_384 ||
    startRow > 1_048_576 ||
    endRow > 1_048_576 ||
    endColumn < startColumn ||
    endRow < startRow
  ) {
    return null;
  }
  return {
    rows: endRow - startRow + 1,
    columns: endColumn - startColumn + 1
  };
}

export function isExplicitA1Address(address: string): boolean {
  if (explicitA1RangeShape(address)) return true;
  const normalized = address.trim().replace(/\$/g, "");
  const columns = normalized.match(/^([A-Z]{1,3}):([A-Z]{1,3})$/i);
  if (columns) {
    const start = columnNumber(columns[1]);
    const end = columnNumber(columns[2]);
    return start <= end && end <= 16_384;
  }
  const rows = normalized.match(/^([1-9]\d*):([1-9]\d*)$/);
  if (rows) {
    const start = Number(rows[1]);
    const end = Number(rows[2]);
    return start <= end && end <= 1_048_576;
  }
  return false;
}

function actionRangeAddresses(
  action: ExcelAction
): Array<{ label: string; address: string }> {
  const record = action as unknown as Record<string, unknown>;
  return [
    ["range", "区域"],
    ["startCell", "起始位置"],
    ["cell", "单元格"],
    ["sourceRange", "源区域"],
    ["targetRange", "目标区域"],
    ["destinationCell", "目标位置"]
  ].flatMap(([field, label]) =>
    typeof record[field] === "string" && record[field]
      ? [{ label, address: record[field] }]
      : []
  );
}

export function preflightPlanActions(
  plan: AnalysisPlan,
  existingSheetNames: string[],
  isApiSupported: (version: string) => boolean = () => true,
  catalog: PlanPreflightCatalog = {}
): PlanPreflightIssue[] {
  const issues: PlanPreflightIssue[] = [];
  const sheets = new Map(
    existingSheetNames.map((name) => [name.toLocaleLowerCase(), name])
  );
  const hasSheet = (name: string) =>
    sheets.has(name.toLocaleLowerCase());
  const addSheet = (name: string) =>
    sheets.set(name.toLocaleLowerCase(), name);
  const removeSheet = (name: string) =>
    sheets.delete(name.toLocaleLowerCase());
  const tableAndRangeNames = new Set(
    [...(catalog.tableNames ?? []), ...(catalog.namedRangeNames ?? [])].map(
      normalizedObjectName
    )
  );
  const pivotNames = new Set(
    (catalog.pivotTableNames ?? []).map(normalizedObjectName)
  );
  const knownRangeNames = new Set(
    (catalog.namedRangeNames ?? []).map(normalizedObjectName)
  );
  const shapeNamesBySheet = new Map(
    Object.entries(catalog.shapeNamesBySheet ?? {}).map(([sheet, names]) => [
      sheet.toLocaleLowerCase(),
      new Set(names.map(normalizedObjectName))
    ])
  );

  const claimName = (
    index: number,
    names: Set<string>,
    name: string,
    label: string
  ) => {
    const normalized = normalizedObjectName(name);
    if (names.has(normalized)) {
      issues.push({
        index,
        message: `${label}名称「${name}」已存在`
      });
      return;
    }
    names.add(normalized);
  };

  plan.actions.forEach((action, index) => {
    const requiredVersion = minimumExcelApiVersion(action);
    if (!isApiSupported(requiredVersion)) {
      issues.push({
        index,
        message: `当前 Excel 不支持此操作（需要 ExcelApi ${requiredVersion} 或更高版本）`
      });
      return;
    }

    if (action.type === "deleteWorksheet") {
      if (!hasSheet(action.sheet)) {
        issues.push({
          index,
          message: `未找到工作表「${action.sheet}」`
        });
        return;
      }
      if (sheets.size === 1) {
        issues.push({
          index,
          message: "不能删除工作簿中唯一的工作表"
        });
        return;
      }
      removeSheet(action.sheet);
      return;
    }

    if (
      action.type === "splitGroupAggregate" &&
      !hasSheet(action.sheet)
    ) {
      issues.push({
        index,
        message: `未找到源工作表「${action.sheet}」`
      });
      return;
    }
    if (
      action.type === "copyRange" &&
      !hasSheet(action.sourceSheet)
    ) {
      issues.push({
        index,
        message: `未找到复制源工作表「${action.sourceSheet}」`
      });
      return;
    }
    if (
      action.type === "createPivotTable" &&
      !hasSheet(action.sourceSheet)
    ) {
      issues.push({
        index,
        message: `未找到数据透视表源工作表「${action.sourceSheet}」`
      });
      return;
    }

    for (const { label, address } of actionRangeAddresses(action)) {
      if (
        !isExplicitA1Address(address) &&
        !knownRangeNames.has(normalizedObjectName(address))
      ) {
        issues.push({
          index,
          message: `${label}「${address}」不是有效的 A1 地址或现有命名区域`
        });
      }
    }

    if (action.type === "writeValues" || action.type === "writeFormulas") {
      // R1C1 模式按区域形状自动铺满，不要求 formulas 矩阵与区域同形
      if (action.type === "writeFormulas" && action.formulaR1C1) {
        return;
      }
      const matrix =
        action.type === "writeValues" ? action.values : action.formulas;
      const shape = explicitA1RangeShape(action.range);
      if (
        shape &&
        (shape.rows !== matrix.length ||
          shape.columns !== matrix[0]?.length)
      ) {
        issues.push({
          index,
          message: `目标区域 ${action.range} 是 ${shape.rows}×${shape.columns}，但写入矩阵是 ${matrix.length}×${matrix[0]?.length ?? 0}`
        });
      }
    }
    if (action.type === "writeTable") {
      const shape = explicitA1RangeShape(action.startCell);
      if (shape && (shape.rows !== 1 || shape.columns !== 1)) {
        issues.push({
          index,
          message: `表格起始位置必须是单个单元格，当前为 ${action.startCell}`
        });
      }
    }
    if (action.type === "sortRange") {
      const shape = explicitA1RangeShape(action.range);
      for (const key of action.keys) {
        if (shape && key.column >= shape.columns) {
          issues.push({
            index,
            message: `排序列索引 ${key.column} 超出 ${action.range} 的 ${shape.columns} 列范围`
          });
        }
      }
    }
    if (action.type === "filterRange") {
      const shape = explicitA1RangeShape(action.range);
      if (shape && action.column >= shape.columns) {
        issues.push({
          index,
          message: `筛选列索引 ${action.column} 超出 ${action.range} 的 ${shape.columns} 列范围`
        });
      }
    }
    if (action.type === "createPivotTable") {
      const shape = explicitA1RangeShape(action.destinationCell);
      if (shape && (shape.rows !== 1 || shape.columns !== 1)) {
        issues.push({
          index,
          message: `数据透视表目标位置必须是单个单元格，当前为 ${action.destinationCell}`
        });
      }
      claimName(index, pivotNames, action.name, "数据透视表");
    }
    if (action.type === "createTable" && action.name) {
      claimName(index, tableAndRangeNames, action.name, "表格或命名区域");
    }
    if (action.type === "addNamedRange") {
      claimName(
        index,
        tableAndRangeNames,
        action.name,
        "表格或命名区域"
      );
    }
    if (action.type === "addImage" && action.name) {
      const sheetKey = action.sheet.toLocaleLowerCase();
      const names =
        shapeNamesBySheet.get(sheetKey) ?? new Set<string>();
      shapeNamesBySheet.set(sheetKey, names);
      claimName(index, names, action.name, "图片或形状");
    }

    if (action.type !== "splitGroupAggregate") {
      addSheet(action.sheet);
    }
  });
  return issues;
}

export function incompleteActionResults(
  plan: AnalysisPlan,
  failedIndexes: ReadonlyMap<number, string>,
  succeeded: ActionExecutionResult[] = []
): ActionExecutionResult[] {
  const succeededIndexes = new Set(
    succeeded.map((result) => result.index)
  );
  return plan.actions.map((action, index) => ({
    index,
    type: action.type,
    sheet: action.sheet,
    status: succeededIndexes.has(index)
      ? "succeeded"
      : failedIndexes.has(index)
        ? "failed"
        : "not_run",
    ...(failedIndexes.has(index)
      ? { message: failedIndexes.get(index) }
      : {})
  }));
}

async function executeAction(
  context: Excel.RequestContext,
  action: ExcelAction
): Promise<VerificationCriterion[] | void> {
  if (action.type === "deleteWorksheet") {
    const existing = context.workbook.worksheets.getItemOrNullObject(action.sheet);
    existing.load("isNullObject");
    await context.sync();
    if (existing.isNullObject) {
      throw new Error(`未找到工作表「${action.sheet}」`);
    }
    existing.delete();
    return;
  }

  if (action.type === "splitGroupAggregate") {
    const worksheets = context.workbook.worksheets;
    const sourceSheet = worksheets.getItemOrNullObject(action.sheet);
    sourceSheet.load("isNullObject,name");
    try {
      await context.sync();
    } catch (reason) {
      throw new Error(`定位源工作表失败：${excelErrorDetail(reason)}`);
    }
    if (sourceSheet.isNullObject) {
      throw new Error(`未找到源工作表「${action.sheet}」`);
    }

    const sourceRange = action.sourceRange
      ? sourceSheet.getRange(action.sourceRange)
      : sourceSheet.getUsedRangeOrNullObject(true);
    sourceRange.load("isNullObject,values");
    try {
      await context.sync();
    } catch (reason) {
      throw new Error(`读取源数据失败：${excelErrorDetail(reason)}`);
    }
    if (sourceRange.isNullObject) {
      throw new Error(`源工作表「${action.sheet}」没有可读取的数据`);
    }
    const outputs = buildSplitAggregateOutputs(
      sourceRange.values.map((row) => row.map(normalizeValue)),
      {
        ...action,
        maxOutputSheets: action.maxOutputSheets ?? 200
      }
    );

    worksheets.load("items/name");
    await context.sync();
    const existingNames = new Map(
      worksheets.items.map((worksheet) => [
        worksheet.name.toLocaleLowerCase(),
        worksheet.name
      ])
    );
    const sourceName = sourceSheet.name.toLocaleLowerCase();
    const generatedNames = new Set<string>();
    const renamed = (base: string): string => {
      if (!existingNames.has(base.toLocaleLowerCase())) return base;
      let suffix = 2;
      while (true) {
        const marker = ` (${suffix})`;
        const candidate = `${base.slice(0, 31 - marker.length)}${marker}`;
        if (!existingNames.has(candidate.toLocaleLowerCase())) return candidate;
        suffix += 1;
      }
    };
    const criteria: VerificationCriterion[] = [];
    const createdInThisRun: string[] = [];
    const rollbackCreatedSheets = async (): Promise<void> => {
      if (createdInThisRun.length === 0) return;
      const candidates = createdInThisRun.map((name) => {
        const worksheet = worksheets.getItemOrNullObject(name);
        worksheet.load("isNullObject");
        return worksheet;
      });
      await context.sync();
      for (const worksheet of candidates) {
        if (!worksheet.isNullObject) worksheet.delete();
      }
      await context.sync();
    };

    for (const outputBatch of batchSplitAggregateOutputs(
      outputs,
      SPLIT_AGGREGATE_BATCH_SHEETS
    )) {
      const jobs: Array<{
        output: (typeof outputBatch)[number];
        targetName: string;
        targetSheet: Excel.Worksheet;
        matrix: CellValue[][];
        usedRange: Excel.Range | null;
      }> = [];

      for (const output of outputBatch) {
        const baseName = safeWorksheetBaseName(output.splitValue);
        const normalizedBase = baseName.toLocaleLowerCase();
        let targetName = baseName;
        let targetSheet: Excel.Worksheet;
        let usedRange: Excel.Range | null = null;
        const existingName = existingNames.get(normalizedBase);
        const collidesWithCurrentRun = generatedNames.has(normalizedBase);

        if (
          existingName &&
          action.existingSheetPolicy === "skip" &&
          !collidesWithCurrentRun
        ) {
          criteria.push({
            type: "worksheetExists",
            sheet: existingName
          });
          continue;
        }
        if (
          existingName &&
          action.existingSheetPolicy === "replace" &&
          !collidesWithCurrentRun
        ) {
          if (normalizedBase === sourceName) {
            throw new Error(`不能用拆分结果覆盖源工作表「${sourceSheet.name}」`);
          }
          targetSheet = worksheets.getItem(existingName);
          usedRange = targetSheet.getUsedRangeOrNullObject();
          usedRange.load("isNullObject");
        } else {
          if (existingName) targetName = renamed(baseName);
          targetSheet = worksheets.add(targetName);
          createdInThisRun.push(targetName);
          existingNames.set(targetName.toLocaleLowerCase(), targetName);
        }
        generatedNames.add(targetName.toLocaleLowerCase());
        jobs.push({
          output,
          targetName,
          targetSheet,
          matrix: [output.headers, ...output.rows],
          usedRange
        });
      }

      if (jobs.length === 0) continue;
      const batchLabel = jobs
        .slice(0, 3)
        .map((job) => job.targetName)
        .join("、");
      try {
        if (jobs.some((job) => job.usedRange)) {
          await context.sync();
          for (const job of jobs) {
            if (job.usedRange && !job.usedRange.isNullObject) {
              job.usedRange.clear("All");
            }
          }
        }

        for (const { output, targetSheet, matrix } of jobs) {
          const target = targetSheet
            .getRange("A1")
            .getResizedRange(matrix.length - 1, matrix[0].length - 1);
          for (
            let columnIndex = 0;
            columnIndex < action.groupBy.length + 1;
            columnIndex += 1
          ) {
            if (
              matrix
                .slice(1)
                .some((row) => typeof row[columnIndex] === "string")
            ) {
              const dimensionRange = targetSheet.getRange(
                `${columnName(columnIndex + 1)}1:${columnName(columnIndex + 1)}${
                  matrix.length
                }`
              );
              dimensionRange.numberFormat = Array.from(
                { length: matrix.length },
                () => ["@"]
              );
            }
          }
          target.values = matrix;
          const header = target.getRow(0);
          header.format.font.bold = true;
          header.format.fill.color = "#DFF3E4";
          for (const columnIndex of output.ratioColumnIndexes) {
            const ratioRange = targetSheet.getRange(
              `${columnName(columnIndex + 1)}2:${columnName(columnIndex + 1)}${
                matrix.length
              }`
            );
            ratioRange.numberFormat = Array.from(
              { length: matrix.length - 1 },
              () => ["0.00%"]
            );
          }
          target.format.autofitColumns();
          target.format.autofitRows();
        }
        await context.sync();
      } catch (reason) {
        try {
          await rollbackCreatedSheets();
        } catch {
          // Keep the original Excel error; cleanup is best-effort.
        }
        throw new Error(
          `生成工作表批次「${batchLabel}${
            jobs.length > 3 ? "等" : ""
          }」失败：${excelErrorDetail(reason)}`
        );
      }

      for (const { targetName, matrix } of jobs) {
        const range = matrixRange("A1", matrix.length, matrix[0].length);
        criteria.push(
          range && matrix.length <= 500
            ? {
                type: "rangeEquals",
                sheet: targetName,
                range,
                expected: matrix
              }
            : { type: "worksheetExists", sheet: targetName }
        );
      }
    }
    return criteria;
  }

  const sheet = await getOrCreateWorksheet(context, action.sheet);

  switch (action.type) {
    case "createWorksheet":
      return;
    case "writeTable": {
      const values = [action.headers, ...action.rows];
      assertRectangular(values);
      const start = sheet.getRange(action.startCell);
      const target = start.getResizedRange(values.length - 1, values[0].length - 1);
      target.values = values;
      if (action.headers.length > 0) {
        const header = start.getResizedRange(0, action.headers.length - 1);
        header.format.font.bold = true;
        header.format.fill.color = "#DFF3E4";
      }
      return;
    }
    case "writeValues": {
      assertRectangular(action.values);
      const target = sheet.getRange(action.range);
      target.values = action.values;
      return;
    }
    case "writeFormulas": {
      const target = sheet.getRange(action.range);
      if (action.formulaR1C1) {
        // R1C1 与锚点无关：同一字符串铺满整个区域，相对引用逐格自动平移
        target.load("rowCount,columnCount");
        await context.sync();
        const r1c1 = action.formulaR1C1;
        target.formulasR1C1 = Array.from({ length: target.rowCount }, () =>
          Array.from({ length: target.columnCount }, () => r1c1)
        );
        return;
      }
      target.formulas = action.formulas;
      return;
    }
    case "clearRange":
      sheet
        .getRange(action.range)
        .clear(
          ({ all: "All", contents: "Contents", formats: "Formats", hyperlinks: "Hyperlinks" } as const)[
            action.applyTo
          ]
        );
      return;
    case "insertRange":
      sheet.getRange(action.range).insert(action.shift === "down" ? "Down" : "Right");
      return;
    case "deleteRange":
      sheet.getRange(action.range).delete(action.shift === "up" ? "Up" : "Left");
      return;
    case "copyRange": {
      const sourceSheet = context.workbook.worksheets.getItem(action.sourceSheet);
      sheet.getRange(action.targetRange).copyFrom(
        sourceSheet.getRange(action.sourceRange),
        enumText(action.copyType) as Excel.RangeCopyType,
        action.skipBlanks,
        action.transpose
      );
      return;
    }
    case "sortRange":
      sheet.getRange(action.range).sort.apply(
        action.keys.map((key) => ({ key: key.column, ascending: key.ascending })),
        false,
        action.hasHeaders
      );
      return;
    case "filterRange":
      sheet.autoFilter.apply(sheet.getRange(action.range), action.column, {
        filterOn: "Values",
        values: action.values.map((value) => String(value))
      });
      return;
    case "clearFilter":
      sheet.autoFilter.clearCriteria();
      return;
    case "setDataValidation": {
      const validation = sheet.getRange(action.range).dataValidation;
      validation.rule = dataValidationRule(action);
      validation.ignoreBlanks = action.allowBlank;
      if (action.prompt) {
        validation.prompt = {
          showPrompt: true,
          title: "输入提示",
          message: action.prompt
        };
      }
      if (action.errorMessage) {
        validation.errorAlert = {
          showAlert: true,
          style: "Stop",
          title: "输入无效",
          message: action.errorMessage
        };
      }
      return;
    }
    case "setConditionalFormat": {
      const range = sheet.getRange(action.range);
      if (action.ruleType === "colorScale") {
        const format = range.conditionalFormats.add("ColorScale").colorScale;
        format.criteria = {
          minimum: {
            type: "LowestValue",
            color: action.minColor ?? "#F8696B"
          },
          ...(action.midColor
            ? {
                midpoint: {
                  type: "Percent",
                  formula: "50",
                  color: action.midColor
                }
              }
            : {}),
          maximum: {
            type: "HighestValue",
            color: action.maxColor ?? "#63BE7B"
          }
        };
      } else if (action.ruleType === "custom") {
        const format = range.conditionalFormats.add("Custom").custom;
        format.rule.formula = String(action.formula1 ?? "");
        format.format.fill.color = action.color ?? "#FFF2CC";
      } else {
        const format = range.conditionalFormats.add("CellValue").cellValue;
        const conditionalOperator = {
          greaterThanOrEqualTo: "GreaterThanOrEqual",
          lessThanOrEqualTo: "LessThanOrEqual"
        }[action.operator ?? ""] ?? enumText(action.operator ?? "equalTo");
        format.rule = {
          operator: conditionalOperator as Excel.ConditionalCellValueOperator,
          formula1: String(action.formula1 ?? ""),
          ...(action.formula2 !== undefined && action.formula2 !== null
            ? { formula2: String(action.formula2) }
            : {})
        };
        format.format.fill.color = action.color ?? "#FFF2CC";
      }
      return;
    }
    case "setNumberFormat": {
      const range = sheet.getRange(action.range);
      range.load("rowCount,columnCount");
      await context.sync();
      range.numberFormat = Array.from({ length: range.rowCount }, () =>
        Array.from({ length: range.columnCount }, () => action.formatCode)
      );
      return;
    }
    case "setBorders": {
      const borderIndexes = {
        top: "EdgeTop",
        bottom: "EdgeBottom",
        left: "EdgeLeft",
        right: "EdgeRight",
        insideHorizontal: "InsideHorizontal",
        insideVertical: "InsideVertical"
      } as const;
      const borders = sheet.getRange(action.range).format.borders;
      for (const side of action.sides) {
        const border = borders.getItem(borderIndexes[side]);
        border.style = enumText(action.style) as Excel.BorderLineStyle;
        border.color = action.color;
        border.weight = enumText(action.weight) as Excel.BorderWeight;
      }
      return;
    }
    case "setAlignment": {
      const format = sheet.getRange(action.range).format;
      if (action.horizontal) {
        format.horizontalAlignment = enumText(
          action.horizontal
        ) as Excel.HorizontalAlignment;
      }
      if (action.vertical) {
        format.verticalAlignment = enumText(
          action.vertical
        ) as Excel.VerticalAlignment;
      }
      if (action.wrapText !== undefined && action.wrapText !== null) {
        format.wrapText = action.wrapText;
      }
      return;
    }
    case "mergeCells":
      sheet.getRange(action.range).merge(action.across);
      return;
    case "unmergeCells":
      sheet.getRange(action.range).unmerge();
      return;
    case "resizeRange": {
      const format = sheet.getRange(action.range).format;
      if (action.rowHeight !== undefined && action.rowHeight !== null) {
        format.rowHeight = action.rowHeight;
      }
      if (action.columnWidth !== undefined && action.columnWidth !== null) {
        format.columnWidth = action.columnWidth;
      }
      return;
    }
    case "freezePanes":
      sheet.freezePanes.unfreeze();
      if (action.rows > 0) sheet.freezePanes.freezeRows(action.rows);
      if (action.columns > 0) sheet.freezePanes.freezeColumns(action.columns);
      return;
    case "setHyperlink":
      sheet.getRange(action.range).hyperlink = {
        address: action.address,
        ...(action.text ? { textToDisplay: action.text } : {}),
        ...(action.screenTip ? { screenTip: action.screenTip } : {})
      };
      return;
    case "addComment":
      context.workbook.comments.add(sheet.getRange(action.cell), action.text);
      return;
    case "addNote":
      sheet.notes.add(sheet.getRange(action.cell), action.text);
      return;
    case "createTable": {
      const table = sheet.tables.add(action.range, action.hasHeaders);
      if (action.name) table.name = action.name;
      if (action.style) table.style = action.style;
      return;
    }
    case "createChart": {
      const chart = sheet.charts.add(
        enumText(action.chartType) as Excel.ChartType,
        sheet.getRange(action.sourceRange)
      );
      if (action.title) {
        chart.title.text = action.title;
        chart.title.visible = true;
      }
      if (action.targetRange) {
        chart.setPosition(sheet.getRange(action.targetRange));
      }
      chart.load("name");
      await context.sync();
      return [
        {
          type: "chartExists",
          sheet: action.sheet,
          name: chart.name,
          chartType: action.chartType,
          sourceRange: action.sourceRange,
          title: action.title,
          targetRange: action.targetRange
        }
      ];
    }
    case "createPivotTable": {
      const sourceSheet = context.workbook.worksheets.getItem(action.sourceSheet);
      const pivot = sheet.pivotTables.add(
        action.name,
        sourceSheet.getRange(action.sourceRange),
        sheet.getRange(action.destinationCell)
      );
      for (const field of action.rowFields) {
        pivot.rowHierarchies.add(pivot.hierarchies.getItem(field));
      }
      for (const field of action.columnFields) {
        pivot.columnHierarchies.add(pivot.hierarchies.getItem(field));
      }
      for (const valueField of action.valueFields) {
        const hierarchy = pivot.dataHierarchies.add(
          pivot.hierarchies.getItem(valueField.field)
        );
        hierarchy.summarizeBy = enumText(
          valueField.aggregation
        ) as Excel.AggregationFunction;
      }
      return;
    }
    case "addNamedRange":
      context.workbook.names.add(
        action.name,
        sheet.getRange(action.range),
        action.comment ?? undefined
      );
      return;
    case "addImage": {
      const target = sheet.getRange(action.targetRange);
      target.load("left,top,width,height");
      await context.sync();
      const shape = sheet.shapes.addImage(action.base64);
      if (action.name) shape.name = action.name;
      positionShape(shape, target);
      return;
    }
    case "addShape": {
      const target = sheet.getRange(action.targetRange);
      target.load("left,top,width,height");
      await context.sync();
      const shape =
        action.shapeType === "line"
          ? sheet.shapes.addLine(
              target.left,
              target.top,
              target.left + target.width,
              target.top + target.height,
              "Straight"
            )
          : sheet.shapes.addGeometricShape(
              ({
                rectangle: "Rectangle",
                roundedRectangle: "RoundRectangle",
                ellipse: "Ellipse",
                triangle: "Triangle",
                diamond: "Diamond"
              } as Record<string, Excel.GeometricShapeType>)[action.shapeType] ??
                "Rectangle"
            );
      if (action.text) shape.textFrame.textRange.text = action.text;
      if (action.fillColor) shape.fill.setSolidColor(action.fillColor);
      positionShape(shape, target);
      return;
    }
    case "setFill":
      sheet.getRange(action.range).format.fill.color = action.color;
      return;
    case "setFont": {
      const font = sheet.getRange(action.range).format.font;
      if (action.bold !== undefined) font.bold = action.bold;
      if (action.color) font.color = action.color;
      return;
    }
    case "autofit": {
      const format = sheet.getRange(action.range).format;
      format.autofitColumns();
      format.autofitRows();
      return;
    }
    case "activateWorksheet":
      sheet.activate();
      return;
  }
}

function recoverableActionRange(action: ExcelAction): string | null {
  switch (action.type) {
    case "writeValues":
    case "writeFormulas":
    case "setFill":
    case "setFont":
      return action.range;
    case "clearRange":
      return action.applyTo === "contents" ? action.range : null;
    case "writeTable":
      return matrixRange(
        action.startCell,
        action.rows.length + 1,
        action.headers.length
      );
    default:
      return null;
  }
}

async function captureUndoRange(
  context: Excel.RequestContext,
  action: ExcelAction,
  actionIndex: number
): Promise<ExecutionUndoSnapshot["ranges"][number] | null> {
  const address = recoverableActionRange(action);
  if (!address) return null;
  const sheet = context.workbook.worksheets.getItemOrNullObject(action.sheet);
  sheet.load("isNullObject");
  await context.sync();
  if (sheet.isNullObject) return null;
  const range = sheet.getRange(address);
  range.load(
    "formulas,numberFormat,format/fill/color,format/font/bold,format/font/color"
  );
  await context.sync();
  return {
    actionIndex,
    sheet: action.sheet,
    range: address,
    formulas: range.formulas.map((row) => row.map(normalizeValue)),
    numberFormat: range.numberFormat,
    fillColor: range.format.fill.color,
    fontBold: range.format.font.bold,
    fontColor: range.format.font.color
  };
}

export async function undoExecution(
  snapshot: ExecutionUndoSnapshot
): Promise<void> {
  await Excel.run(async (context) => {
    for (const saved of [...snapshot.ranges].reverse()) {
      const range = context.workbook.worksheets
        .getItem(saved.sheet)
        .getRange(saved.range);
      range.formulas = saved.formulas;
      range.numberFormat = saved.numberFormat;
      range.format.fill.color = saved.fillColor;
      range.format.font.bold = saved.fontBold;
      range.format.font.color = saved.fontColor;
    }
    await context.sync();
  });
}

export async function executePlan(plan: AnalysisPlan): Promise<PlanExecutionResult> {
  if (plan.sourceFingerprint && plan.sourceFingerprintSheets?.length) {
    const current = isToolSchemaFingerprint(plan.sourceFingerprint)
      ? toolSchemaFingerprintForSnapshot(
          plan,
          await captureWorkbookStructure(plan.sourceFingerprintSheets)
        )
      : await captureWorkbookSourceFingerprint(plan.sourceFingerprintSheets);
    if (current !== plan.sourceFingerprint) {
      throw new PlanExecutionError(
        isToolSchemaFingerprint(plan.sourceFingerprint)
          ? "工具所需的来源工作表或字段已发生变化，请重新选择后再执行"
          : "预览后数据来源已发生变化，请重新生成预览后再执行",
        incompleteActionResults(plan, new Map())
      );
    }
  }
  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    const tables = context.workbook.tables;
    tables.load("items/name");
    const names = context.workbook.names;
    names.load("items/name");
    const supportsPivotCatalog =
      Office.context.requirements.isSetSupported("ExcelApi", "1.3");
    const pivotTables = supportsPivotCatalog
      ? context.workbook.pivotTables
      : null;
    pivotTables?.load("items/name");
    await context.sync();
    const needsShapeCatalog =
      Office.context.requirements.isSetSupported("ExcelApi", "1.9") &&
      plan.actions.some(
        (action) => action.type === "addImage" && Boolean(action.name)
      );
    const shapeCollections = needsShapeCatalog
      ? worksheets.items.map((worksheet) => {
          const shapes = worksheet.shapes;
          shapes.load("items/name");
          return { sheet: worksheet.name, shapes };
        })
      : [];
    if (shapeCollections.length > 0) await context.sync();
    const preflightIssues = preflightPlanActions(
      plan,
      worksheets.items.map((worksheet) => worksheet.name),
      (version) =>
        Office.context.requirements.isSetSupported("ExcelApi", version),
      {
        tableNames: tables.items.map((table) => table.name),
        namedRangeNames: names.items.map((name) => name.name),
        pivotTableNames:
          pivotTables?.items.map((pivotTable) => pivotTable.name) ?? [],
        shapeNamesBySheet: Object.fromEntries(
          shapeCollections.map(({ sheet, shapes }) => [
            sheet,
            shapes.items.map((shape) => shape.name)
          ])
        )
      }
    );
    if (preflightIssues.length > 0) {
      const failed = new Map(
        preflightIssues.map((issue) => [issue.index, issue.message])
      );
      throw new PlanExecutionError(
        `执行前检查未通过：${preflightIssues
          .map(
            (issue) => `第 ${issue.index + 1} 步 ${issue.message}`
          )
          .join("；")}`,
        incompleteActionResults(plan, failed)
      );
    }

    const actionResults: ActionExecutionResult[] = [];
    const dynamicCriteria: VerificationCriterion[] = [];
    const undoRanges: ExecutionUndoSnapshot["ranges"] = [];
    const executionStartedAt = performance.now();
    for (const [index, action] of plan.actions.entries()) {
      try {
        const undoRange = await captureUndoRange(context, action, index);
        const actionCriteria = await executeAction(context, action);
        if (actionCriteria) dynamicCriteria.push(...actionCriteria);
        await context.sync();
        if (undoRange) undoRanges.push(undoRange);
      } catch (reason) {
        const detail = excelErrorDetail(reason);
        throw new PlanExecutionError(
          `第 ${index + 1} 步执行失败：${detail}`,
          incompleteActionResults(
            plan,
            new Map([[index, detail]]),
            actionResults
          )
        );
      }
      actionResults.push({
        index,
        type: action.type,
        sheet: action.sheet,
        status: "succeeded"
      });
    }
    const verificationStartedAt = performance.now();
    const executionMs = verificationStartedAt - executionStartedAt;
    const verification = await verifyPlan(context, plan, dynamicCriteria);
    const verificationMs = performance.now() - verificationStartedAt;
    return {
      actionResults,
      verification,
      executionMs,
      verificationMs,
      undoSnapshot:
        undoRanges.length > 0
          ? {
              planId: plan.id,
              capturedAt: new Date().toISOString(),
              ranges: undoRanges
            }
          : null
    };
  });
}
