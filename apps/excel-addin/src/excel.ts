import { isTrustedLiteralPlan } from "@excel-bro/core/literalPlanOrigin";
import {
  formulaSheetReferences,
  formulaSourceEvidenceRequirement,
  verifyFormulaResumeSources
} from "@excel-bro/core/formulaWriteback";
import { copyRangePlan } from "@excel-bro/core/copyRangePlan";
import {
  BAKED_WRITE_ACTION_TYPES,
  judgePlanSourceGate,
  PLAN_SOURCE_MISSING_MESSAGE,
  planHasBakedValues,
  planSourceAnchor,
  planSourceComparisonNeeded,
  planSourceGateMessage
} from "@excel-bro/core/planSourceAnchor";
import { filteredSourceRows } from "./filteredSourceRows";
import { writeTableValues, WRITE_TABLE_HEADER_FILL } from "@excel-bro/core/writeTableLayout";
import type {
  AnalysisPlan,
  ActionExecutionResult,
  BorderSide,
  CellBorderState,
  CellValue,
  DataToolRequest,
  DirectCellFormatValues,
  ExecutionUndoSnapshot,
  ExcelAction,
  PlanExecutionResult,
  UndoCellDirectFormat,
  UndoCellFormat,
  UndoCreatedWorksheetState,
  UndoRangeFormat,
  UndoRangeSnapshot,
  UndoRestoreScope,
  UndoWorksheetSnapshot,
  VerificationCheck,
  VerificationCriterion,
  VerificationReport,
  TaskProgressEvent,
  UnverifiedAction,
  WorkbookSnapshot,
  WorksheetSnapshot
} from "./contracts";
// 值导入（不是 type）：边框的缺省补齐与格式属性判据必须共用同一份实现，
// 否则"写入用的 style"和"回读核对用的 style"在模型省略字段时会各算各的。
import { normalizeBorderAction } from "./contracts";
import type { ExcelFormatCapability, ExcelFormatField } from "./excelFormatCapability";
import {
  EXCEL_DIRECT_FORMAT_REQUIREMENT,
  EXCEL_DISPLAYED_FORMAT_REQUIREMENT,
  blockedFormatActionIndexes,
  excelFormatActionFields,
  resolveExcelFormatCapability
} from "./excelFormatCapability";
import type { WorkbookLocator } from "@excel-bro/core/contracts";
import type {
  ChangeReviewStatus,
  RangeChangeReview
} from "@excel-bro/core/changeReview";
import { buildRangeChangeReview } from "@excel-bro/core/changeReview";
// 「实际写入了几次」的唯一口径（与 WPS、服务端读同一份定义）。
import type { HostWriteAudit } from "@excel-bro/core/hostWriteAudit";
// EB 系统遗留表的名单与判定在共享层：WPS 端的工作目录勾选也要用同一份名单。
import { isEBSystemSheet } from "@excel-bro/core/folderScope";
import {
  locatorFromOfficeIdentity,
  officeIdentityFromUrl
} from "@excel-bro/core/locator";
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
import { assertSafeFormula, assertSafeHyperlink } from "./formulaSafety";
import { assertRectangularValues, parseExplicitA1Rect, writeMatrixShapeIssue } from "./planCore";
import {
  resolveFilteredRowIndices,
  createQueryTableAccumulator,
  readSheetIntoAccumulator,
  DataToolExecutionError
} from "./dataTools";

const DATA_ROW_LIMIT = capabilities.snapshot.dataRows;
const DATA_COLUMN_LIMIT = capabilities.snapshot.dataColumns;
const STRUCTURE_ROW_LIMIT = capabilities.snapshot.structureRows;
const STRUCTURE_COLUMN_LIMIT = capabilities.snapshot.structureColumns;
// 删除工作表撤销的拍照格数上限：超过则该表不提供撤销（内存保护）。
const UNDO_WORKSHEET_MAX_CELLS = 100_000;
const FINGERPRINT_CHUNK_ROWS = capabilities.queryTable.chunkRows;
const FINGERPRINT_CHUNK_CELLS =
  capabilities.queryTable.chunkRows * capabilities.queryTable.maxColumns;
const SPLIT_AGGREGATE_BATCH_SHEETS =
  capabilities.excelExecution.splitAggregateBatchSheets;
// 超过此规模的区域跳过完整格式捕获：逐格读边框会让整表操作的执行耗时显著上升。
const UNDO_FORMAT_CAPTURE_MAX_CELLS =
  capabilities.undo.formatCaptureMaxCells;
// 双栏对照的规模上限：只影响"展示与采集"，不影响写入与撤销的正确性。
// 超过 captureMaxCells 的区域不再采集显示值（valuesCaptured=false，界面显示"显示值未采集"），
// 但仍保留 formulas 用于判定改动。
const CHANGE_REVIEW_CAPTURE_MAX_CELLS = capabilities.changeReview.captureMaxCells;
const CHANGE_REVIEW_LIMIT = {
  maxRows: capabilities.changeReview.maxRows,
  maxTotalCells: capabilities.changeReview.maxTotalCells
};
// EB 系统遗留表的名单与判定已搬到共享层（`@excel-bro/core/folderScope`）。
// 本文件内部照旧用 `isEBSystemSheet`，同时把名字转出去，
// 保持旧的导入路径（`./excel`）不变 —— officeAuthorizedTarget.ts 等还在用。
export { isEBSystemSheet };

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

// 执行完计划后，若含增删工作表的结构操作，先失效结构缓存再重扫。
// 不能只依赖 worksheets.onAdded/onDeleted 事件：执行后紧跟 scan() 时事件可能
// 尚未派发，scan 会拿旧缓存回填，导致整轮对话都读到陈旧的工作表集合。
export function invalidateStructureCacheForActions(
  actions: ExcelAction[]
): void {
  if (
    actions.some(
      (action) =>
        action.type === "createWorksheet" ||
        action.type === "deleteWorksheet" ||
        action.type === "splitGroupAggregate"
    )
  ) {
    invalidateWorkbookStructureCache();
  }
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
  const nameChangedHandler = async (
    event: Excel.WorksheetNameChangedEventArgs
  ) => {
    idToName.set(event.worksheetId, event.nameAfter);
    invalidateWorkbookStructureCache();
    bumpSheetDataEpoch(event.nameAfter);
    onInvalidated?.();
  };
  const supportsCollectionChanges =
    Office.context.requirements.isSetSupported("ExcelApi", "1.9");
  structureCacheEnabled = supportsCollectionChanges;
  const supportsCollectionLifecycle =
    Office.context.requirements.isSetSupported("ExcelApi", "1.7");
  const supportsNameChanged =
    Office.context.requirements.isSetSupported("ExcelApi", "1.17");
  const watched = await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    if (supportsCollectionChanges) {
      worksheets.onChanged.add(handler);
    }
    if (supportsCollectionLifecycle) {
      worksheets.onAdded.add(handler);
      worksheets.onDeleted.add(handler);
    }
    if (supportsNameChanged) {
      worksheets.onNameChanged.add(nameChangedHandler);
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
      if (supportsNameChanged) {
        worksheets.onNameChanged.remove(nameChangedHandler);
      }
      await context.sync();
    });
  };
}

export class PlanExecutionError extends Error {
  constructor(
    message: string,
    public readonly actionResults: ActionExecutionResult[],
    public readonly undoSnapshot: ExecutionUndoSnapshot | null = null
  ) {
    super(message);
    this.name = "PlanExecutionError";
  }
}

/**
 * 计划里含烘焙矩阵、且预览之后源数据变了。
 *
 * 这不是硬错误 —— 计划本身依然可执行，只是其中的固定值来自旧数据。所以单独
 * 分一个错误类型，让 UI 能出「仍要执行」的二次确认，而不是像字段缺失那样
 * 直接堵死。
 */
export class StaleSourceError extends PlanExecutionError {
  constructor(
    message: string,
    actionResults: ActionExecutionResult[],
    /** 计划里携带烘焙矩阵的动作数量，用于文案里说清影响范围。 */
    public readonly bakedActionCount: number
  ) {
    super(message, actionResults);
    this.name = "StaleSourceError";
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
    if (action.type === "computeValues") {
      add(action.sourceSheet, [
        ...(action.query.fields ?? []),
        ...(action.query.groupBy ?? []),
        ...(action.query.filters ?? []).map((filter) => filter.field),
        ...(action.query.metrics ?? []).map((metric) => metric.field),
        action.query.profileField,
        action.query.sortBy
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

export function isToolSchemaFingerprint(value: string): boolean {
  return value.startsWith(TOOL_SCHEMA_FINGERPRINT_PREFIX);
}

/**
 * 直接体检：计划点名的来源工作表、以及执行时求值动作引用的字段，现在还在不在。
 *
 * 为什么不复用 schema 指纹比对：指纹只能拿「计划生成时存下的那串哈希」当基线，
 * 而 live 计划存的往往是值指纹（旧规则），没有 schema 基线可比。拿现算的哈希跟
 * 自己比是恒等式，什么都挡不住 —— 所以这里改成逐项查在不在，并把缺了什么直接
 * 报出来。
 */
export function missingSchemaRequirements(
  plan: AnalysisPlan,
  snapshot: Pick<WorkbookSnapshot, "worksheets">
): string[] {
  const requiredBySheet = requiredToolFieldsBySheet(plan);
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
  const missing: string[] = [];
  for (const name of new Set(plan.sourceFingerprintSheets ?? [])) {
    const key = normalizeSchemaName(name);
    const headers = available.get(key);
    if (!headers) {
      missing.push(`工作表「${name}」`);
      continue;
    }
    for (const field of requiredBySheet.get(key) ?? []) {
      if (!headers.has(field)) missing.push(`「${name}」的字段「${field}」`);
    }
  }
  return missing;
}

/**
 * 烘焙动作：动作体里携带「从源数据算出来的」矩阵。源数据一变，这些矩阵就是
 * 过期结果，照原样执行会把旧数字写进表里。
 *
 * 清单本体只有一份，在 core（`BAKED_WRITE_ACTION_TYPES`，与 WPS 共用同一份判据）；
 * 这里只做 `Set` 包装，不再复述类型名 —— 本文件曾另存一份清单，两份漂移时
 * `planLiveness` 会说"live"而 `planHasBakedValues` 说"有烘焙"，直接踩空。
 */
const BAKED_ACTION_TYPES = new Set<string>(BAKED_WRITE_ACTION_TYPES);

export type PlanLiveness = "live" | "mixed";

/**
 * live = 所有写入都在执行那一刻求值，源数据改了照样成立；
 * mixed = 至少有一处烘焙矩阵，源数据改了这部分结果就过期。
 */
export function planLiveness(plan: AnalysisPlan): PlanLiveness {
  return plan.actions.some((action) => BAKED_ACTION_TYPES.has(action.type))
    ? "mixed"
    : "live";
}

function updateFingerprint(seed: number, text: string): number {
  let hash = seed;
  for (let index = 0; index < text.length; index += 1) {
    hash ^= text.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return hash >>> 0;
}

/**
 * 读取**规则指定**的源区域，供本地列映射计算用。
 *
 * 与 `captureWorkbookSourceFingerprint` 的区别：那只算指纹（不保留值），
 * 这个保留值矩阵 —— 因为调用方要拿它逐格算目标列。**值只留在本端**：
 * 算完的矩阵不回传服务端（单射映射可逆，上传标签列等于上传源列）。
 *
 * 边界取"声明的已用区域"（`usedRange`）而不是区域本身：映射规则里的区域
 * 可能超界，越界必须在 `sliceMappingSource` 按**真实**已用边界拒绝，而不是
 * 让 `getRange` 直接抛宿主错误（那样用户看不到"超出已读取范围"的可读原因）。
 * 因此这里读**整张表的已用区域**，交给切片函数按声明宽度判界。
 *
 * 读取上限沿用 `snapshot` 的 `dataRows`/`dataColumns`：源列映射的输入本来就是
 * "一列"，行上限与既有快照一致即可。
 */
export async function readSheetUsedRange(
  sheetName: string
): Promise<{
  name: string;
  startRow: number;
  startColumn: number;
  columnCount: number;
  rowCount: number;
  values: CellValue[][];
} | null> {
  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();
    const worksheet = worksheets.items.find((item) => item.name === sheetName);
    if (!worksheet) {
      throw new Error(`来源工作表「${sheetName}」不存在，请重新生成预览`);
    }
    const usedRange = worksheet.getUsedRangeOrNullObject(true);
    usedRange.load(
      "isNullObject,rowIndex,columnIndex,rowCount,columnCount"
    );
    await context.sync();
    if (usedRange.isNullObject) return null;

    // 只读一列宽也没用：列越界要按声明宽度判定，得拿到整块已用区域。
    const rowCount = Math.min(usedRange.rowCount, DATA_ROW_LIMIT + 1);
    const columnCount = Math.min(usedRange.columnCount, DATA_COLUMN_LIMIT);
    const range = worksheet.getRangeByIndexes(
      usedRange.rowIndex,
      usedRange.columnIndex,
      rowCount,
      columnCount
    );
    range.load("values,text,numberFormat");
    await context.sync();
    const values = range.values.map((row, rowIndex) =>
      row.map((value, columnIndex) =>
        normalizeCellValue(
          value,
          range.text[rowIndex]?.[columnIndex] ?? "",
          range.numberFormat[rowIndex]?.[columnIndex] ?? "General"
        )
      )
    );
    return {
      name: worksheet.name,
      // Office.js 的 rowIndex/columnIndex 是 0 基；映射切片用 1 基。
      startRow: usedRange.rowIndex + 1,
      startColumn: usedRange.columnIndex + 1,
      columnCount: usedRange.columnCount,
      rowCount: usedRange.rowCount,
      values
    };
  });
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
  formula: string,
  allowedExternal?: ReadonlySet<string>
): Promise<{ sampleResult: string; sampleInput: string; formulaR1C1: string }> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const cell = sheet.getRange(firstCell);
    cell.load("formulas");
    await context.sync();

    const original = cell.formulas as unknown[][];

    assertSafeFormula(formula, `${firstCell} 的公式`, allowedExternal);
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
      /*
       * 身份必须**当场**记进快照，不能只当成一个局部变量：
       * 快照是"这一轮工作簿是谁"的唯一凭据，工具桥绑定（`sdkBridge.begin`）与
       * 断点恢复账本都读它。留空会让两边退化成 `name:<文件名>`，而执行前复核
       * （`bridgeLiveContext`）读的是 `url:` 身份 —— 同一个工作簿自比不自等，
       * 复核必然判 `source_changed`，工具桥在已保存的工作簿上永远读不到数据。
       * 未保存的工作簿没有 URL，这里如实留空，两边**同时**退化成同名兜底。
       */
      workbookIdentity: currentOfficeWorkbookIdentity() ?? null,
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

/** Scope catalog only. Startup and scope refresh must not read cells or used ranges. */
export async function captureWorkbookCatalog(): Promise<WorkbookSnapshot> {
  const identity = currentOfficeWorkbookIdentity();
  const documentRef = Office.context.document;
  const snapshot = await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    const active = worksheets.getActiveWorksheet();
    worksheets.load("items/name");
    active.load("name");
    await context.sync();
    return {
      name: workbookNameFromDocumentUrl(Office.context.document.url),
      catalogOnly: true,
      capturedAt: new Date().toISOString(),
      activeWorksheet: active.name,
      // 目录快照同样要带上身份：它在启动与范围刷新时就成为 `latestWorkbook`，
      // 后面规划时 `sdkBridge.begin` 用的正是这份快照的身份。
      workbookIdentity: identity ?? null,
      worksheets: worksheets.items.filter((sheet) => !isEBSystemSheet(sheet.name)).map((sheet) => ({
        name: sheet.name, usedRange: null, rowCount: 0, columnCount: 0,
        headers: [], dataRows: [], truncated: false
      }))
    };
  });
  if (identity !== currentOfficeWorkbookIdentity() || documentRef !== Office.context.document) {
    throw new Error("文件身份在目录检查期间发生变化。");
  }
  return snapshot;
}

export async function captureWorkbookStructure(
  dataSheetNames?: string[]
): Promise<WorkbookSnapshot> {
  const cacheKey = JSON.stringify([...(dataSheetNames ?? [])].sort());
  if (structureCacheEnabled && structureCache?.key === cacheKey) {
    const cached = structuredClone(structureCache.snapshot);
    cached.name = workbookNameFromDocumentUrl(Office.context.document.url);
    // 身份与 name 一样是**取快照那一刻的实时属性**，不是被缓存的结构本身；
    // 连 name 都刷新了却不刷新身份，会留下"文件名是新的、身份是上一个工作簿的"
    // 这种自相矛盾的快照，桥绑定随后必被复核判成换簿。
    cached.workbookIdentity = currentOfficeWorkbookIdentity() ?? null;
    // 结构缓存只在工作表数据/集合变化时失效，切换活动工作表不会失效。
    // 命中缓存时用一次轻量选区读取刷新活动表与选区，避免拿过期值构建
    // 指纹，导致确认期与运行期不一致的误判。
    const selection = await captureSelectionContext();
    cached.activeWorksheet = selection.activeWorksheet;
    cached.selectedRange = selection.selectedRange;
    // 活动表不在缓存的工作表集合里，说明集合已变（如新建/删除表）但缓存
    // 尚未失效；此时放弃缓存走全量重扫，避免 activeWorksheet 悬空导致
    // 前端构造出空 sheets 被后端拒绝。
    if (
      !cached.worksheets.some(
        (sheet) => sheet.name === selection.activeWorksheet
      )
    ) {
      structureCache = null;
      return captureWorkbookStructure(dataSheetNames);
    }
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
      // 与 `captureWorkbook` 同口径：身份当场落进快照，未保存时如实为空。
      workbookIdentity: currentOfficeWorkbookIdentity() ?? null,
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

export function matrixVerificationCriteria(
  sheet: string,
  matrix: CellValue[][]
): VerificationCriterion[] {
  if (matrix.length === 0 || matrix[0].length === 0) {
    return [{ type: "worksheetExists", sheet }];
  }
  const criteria: VerificationCriterion[] = [];
  for (let offset = 0; offset < matrix.length; offset += 500) {
    const expected = matrix.slice(offset, offset + 500);
    const range = matrixRange(
      `A${offset + 1}`,
      expected.length,
      matrix[0].length
    );
    if (!range) return [{ type: "worksheetExists", sheet }];
    criteria.push({ type: "rangeEquals", sheet, range, expected });
  }
  return criteria;
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
      // 与 models.py 的 `add_deterministic_acceptance_criteria` 同口径：后面还有同范围的
      // setFont 也给了同一字段时，最终值由后者决定，本条不认领它。
      const laterFonts = laterSameRange("setFont");
      const superseded = (field: "bold" | "italic" | "name" | "size" | "color") =>
        laterFonts.some((later) => later[field] != null);
      const bold = superseded("bold") ? undefined : (action.bold ?? undefined);
      const fontItalic = superseded("italic")
        ? undefined
        : (action.italic ?? undefined);
      const fontName = superseded("name")
        ? undefined
        : (action.name ?? undefined);
      const fontSize = superseded("size")
        ? undefined
        : (action.size ?? undefined);
      const fontColor = superseded("color")
        ? undefined
        : (action.color ?? undefined);
      if (
        bold === undefined &&
        fontItalic === undefined &&
        fontName === undefined &&
        fontSize === undefined &&
        fontColor === undefined
      ) {
        continue;
      }
      addCriterion({
        type: "rangeFormatMatches",
        sheet: action.sheet,
        range: action.range,
        bold,
        fontItalic,
        fontName,
        fontSize,
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
      // 期望值走与执行器同一条缺省补齐：模型省略 style/color/weight 时，
      // 写入用的是默认值，核对也必须用同一组默认值，否则"写对了却核不过"。
      const expected = normalizeBorderAction(action);
      addCriterion({
        type: "bordersMatch",
        sheet: action.sheet,
        range: action.range,
        sides,
        style: expected.style,
        color: expected.color,
        weight: expected.weight
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
          !action.filters?.length &&
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
        ) || Boolean(action.filters?.length && dynamicCriteria.some(criterion =>
          (criterion.type === "rangeEquals" || criterion.type === "formulasEqual") &&
          criterion.sheet.toLocaleLowerCase() === action.sheet.toLocaleLowerCase() &&
          normalizedRangeAddress(criterion.range).split(":")[0] === normalizedRangeAddress(action.targetRange).split(":")[0]
        ));
        break;
      case "splitGroupAggregate":
        verified = dynamicCriteria.length > 0;
        break;
      case "computeValues":
        verified = dynamicCriteria.length > 0;
        break;
      case "sortRange":
        verified = hasCriterion(
          (criterion) =>
            criterion.type === "rangeSorted" &&
            sameRange(criterion, action.sheet, action.range)
        );
        break;
      case "removeDuplicates":
        verified = hasCriterion(
          (criterion) =>
            criterion.type === "rangeEquals" &&
            "range" in criterion &&
            criterion.sheet.toLocaleLowerCase() ===
              action.sheet.toLocaleLowerCase() &&
            normalizedRangeAddress(criterion.range).split(":")[0] ===
              normalizedRangeAddress(action.range).split(":")[0]
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
      // 字体族/字号/斜体与加粗、字色一样，都是**读宿主当前真实格式**，
      // 不是"值没变就算过"。五个一起 load，避免多次往返。
      font.load("bold,italic,name,size,color");
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
        (criterion.fontItalic == null || font.italic === criterion.fontItalic) &&
        (criterion.fontName == null ||
          String(font.name) === criterion.fontName) &&
        (criterion.fontSize == null ||
          closeEnough(font.size, criterion.fontSize)) &&
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
    removeDuplicates: "1.4",
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
  return parseExplicitA1Rect(address);
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
  catalog: PlanPreflightCatalog = {},
  authorizedSheetNames: readonly string[] = existingSheetNames
): PlanPreflightIssue[] {
  const issues: PlanPreflightIssue[] = [];
  const sheets = new Map(
    existingSheetNames.map((name) => [name.toLocaleLowerCase(), name])
  );
  const existingSheets = new Set(sheets.keys());
  const authorizedSheets = new Set(
    authorizedSheetNames.map((name) => name.toLocaleLowerCase())
  );
  const isUnauthorizedExistingSheet = (name: string) => {
    const normalized = name.toLocaleLowerCase();
    return existingSheets.has(normalized) && !authorizedSheets.has(normalized);
  };
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

    const sourceSheet =
      action.type === "copyRange" ||
      action.type === "createPivotTable" ||
      action.type === "computeValues"
        ? action.sourceSheet
        : action.type === "splitGroupAggregate"
          ? action.sheet
          : null;
    if (sourceSheet && isUnauthorizedExistingSheet(sourceSheet)) {
      issues.push({
        index,
        message: `动作 ${action.type} 尝试读取未授权的工作表「${sourceSheet}」`
      });
      return;
    }
    if (isUnauthorizedExistingSheet(action.sheet)) {
      issues.push({
        index,
        message: `动作 ${action.type} 尝试写入未授权的工作表「${action.sheet}」`
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
    if (action.type === "copyRange" && action.filters?.length && !["values", "link"].includes(action.copyType)) {
      issues.push({index, message:"条件复制目前支持值或链接，不能把完整格式或相对公式复制降级成值复制"});
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
    if (action.type === "computeValues") {
      if (!hasSheet(action.sourceSheet)) {
        issues.push({
          index,
          message: `未找到计算源工作表「${action.sourceSheet}」`
        });
        return;
      }
      const destination = explicitA1RangeShape(action.destinationCell);
      if (!destination || destination.rows !== 1 || destination.columns !== 1) {
        issues.push({
          index,
          message: `计算结果的起始单元格「${action.destinationCell}」必须是单个单元格`
        });
        return;
      }
      const referenced = [
        ...action.query.fields ?? [],
        ...action.query.groupBy ?? [],
        ...(action.query.filters ?? []).map((filter) => filter.field),
        ...(action.query.metrics ?? []).map((metric) => metric.field)
      ];
      if (referenced.some((field) => !String(field ?? "").trim())) {
        issues.push({
          index,
          message: "计算动作引用了空字段名"
        });
        return;
      }
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

    if (action.type === "writeFormulas") {
      /*
       * **公式的跨表依赖藏在文本里**：`=Other!A1` 会读一张没被勾选的工作表，而动作
       * 自身只声明了写入目标 —— 只看 `action.sheet` 的授权闸完全看不到它。
       *
       * 引用解析（含 `'My Sheet'!A1`、双引号字面量里的 `!`、外部引用、`INDIRECT` / `OFFSET`）
       * 在共享层 `@excel-bro/core/formulaWriteback`，与 WPS 端读同一份；这里只补宿主侧
       * "哪张表存在、哪张表被授权"。
       *
       * 只拒**已存在且未授权**的工作表，与上面读/写闸同一口径：
       *  - 计划自己新建的表此刻还不存在，不能当成未授权；
       *  - 外部工作簿引用（`[Book.xlsx]Sheet1`）不是本工作簿的表，由 `assertSafeFormula`
       *    的 `allowedExternalFiles` 单独把关，不在这里重复一套。
       */
      const references = formulaSheetReferences(action.formulas);
      if (references.dynamic) {
        issues.push({
          index,
          message: "公式含动态引用（INDIRECT / OFFSET），无法核对依赖的工作表"
        });
        return;
      }
      const unauthorized = references.sheets.filter(isUnauthorizedExistingSheet);
      if (unauthorized.length > 0) {
        issues.push({
          index,
          message: `公式引用了未授权的工作表「${unauthorized.join("、")}」`
        });
        return;
      }
    }
    if (action.type === "writeValues" || action.type === "writeFormulas") {
      // R1C1 模式按区域形状自动铺满，不要求 formulas 矩阵与区域同形
      if (action.type === "writeFormulas" && action.formulaR1C1) {
        return;
      }
      const matrix =
        action.type === "writeValues" ? action.values : action.formulas;
      const shapeIssue = writeMatrixShapeIssue(action.range, matrix);
      if (shapeIssue) {
        issues.push({
          index,
          message: shapeIssue
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
    if (action.type === "removeDuplicates") {
      const shape = explicitA1RangeShape(action.range);
      for (const column of action.columns) {
        if (shape && column >= shape.columns) {
          issues.push({
            index,
            message: `去重依据列索引 ${column} 超出 ${action.range} 的 ${shape.columns} 列范围`
          });
        }
      }
    }
    if (
      (action.type === "removeDuplicates" ||
        action.type === "clearRange" ||
        action.type === "deleteRange" ||
        action.type === "copyRange") &&
      action.filters?.length &&
      !action.hasHeaders
    ) {
      issues.push({
        index,
        message: `带过滤条件的 ${action.type} 操作需要表头行（hasHeaders），否则无法按字段名定位列`
      });
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

  // —— 格式能力检查**前移到预检**：宿主连逐格格式读写所需的最低 API（ExcelApi 1.9）都没有时，
  // 格式动作必然会在执行期被拒。与其等用户确认后才报一句笼统错误，不如在这里就点名
  // "缺哪个 API、影响哪几步"。
  //
  // 注意这只拦 `unavailable`：1.9–1.17 的宿主要走**兼容路径**（`getCellProperties`），
  // 不该被拦下来。兼容路径的逐区域拒绝（条件格式 / 合并区域）需要宿主实读，只能在
  // 写入前判定，因此由预览阶段的**非阻断提示**（`excelFormatCapabilityNotice`）先告知。
  const formatCapability = resolveExcelFormatCapability((set, version) =>
    set === "ExcelApi" ? isApiSupported(version) : false
  );
  if (formatCapability.path === "unavailable") {
    for (const index of blockedFormatActionIndexes(plan.actions, formatCapability)) {
      const action = plan.actions[index];
      issues.push({
        index,
        message:
          `当前 Excel 缺少 ${formatCapability.missingApis.join("、")}，` +
          `无法为 ${action.type} 读取或写回逐格格式，该步无法执行`
      });
    }
  }
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

// 动作执行过程中收集的撤销数据：新建工作表、被覆盖(替换)工作表原有内容。
interface ActionUndoCollector {
  createdWorksheets: string[];
  overwrittenRanges: ExecutionUndoSnapshot["ranges"][number][];
}

interface ActionProgressUpdate {
  stepId: string;
  title: string;
  summary?: string;
  current?: number;
  total?: number;
  unit?: TaskProgressEvent["unit"];
  target?: TaskProgressEvent["target"];
}

export async function executeAction(
  context: Excel.RequestContext,
  action: ExcelAction,
  actionIndex: number,
  collector: ActionUndoCollector,
  allowedExternalFiles?: ReadonlySet<string>,
  signal?: AbortSignal,
  onProgress?: (update: ActionProgressUpdate) => void
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
    onProgress?.({
      stepId: `execution.action.${actionIndex + 1}.sheets`,
      title: "正在生成拆分工作表",
      summary: `已计算出 ${outputs.length} 张结果工作表，正在分批写入。`,
      current: 0,
      total: outputs.length,
      unit: "sheets",
      target: { sheet: action.sheet }
    });

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

    let processedOutputCount = 0;
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
          usedRange.load(
            "isNullObject,address,formulas,numberFormat,format/fill/color,format/font/bold,format/font/color"
          );
        } else {
          if (existingName) targetName = renamed(baseName);
          targetSheet = worksheets.add(targetName);
          createdInThisRun.push(targetName);
          collector.createdWorksheets.push(targetName);
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
            if (!job.usedRange) continue;
            // 撤销时按快照数组的逆序回放：先清空本次写入区，再把原占用区
            // 内容写回，因此这里先推「原内容恢复」，后推「写入区清空」。
            if (!job.usedRange.isNullObject) {
              // 覆盖前拍下原有内容，撤销时原样写回。
              collector.overwrittenRanges.push({
                actionIndex,
                sheet: job.targetName,
                range:
                  job.usedRange.address.split("!").pop() ??
                  job.usedRange.address,
                formulas: job.usedRange.formulas.map((row) =>
                  row.map(normalizeValue)
                ),
                numberFormat: job.usedRange.numberFormat,
                fillColor: job.usedRange.format.fill.color,
                fontBold: job.usedRange.format.font.bold,
                fontColor: job.usedRange.format.font.color
              });
              job.usedRange.clear("All");
            }
            // 清空本次写入区（含写入区超出原占用区的部分），被替换的
            // 空表因此也能还原为空白。
            const writeArea = matrixRange(
              "A1",
              job.matrix.length,
              job.matrix[0].length
            );
            if (writeArea) {
              const emptyMatrix = Array.from(
                { length: job.matrix.length },
                () => Array<CellValue>(job.matrix[0].length).fill(null)
              );
              collector.overwrittenRanges.push({
                actionIndex,
                sheet: job.targetName,
                range: writeArea,
                formulas: emptyMatrix,
                numberFormat: Array.from(
                  { length: job.matrix.length },
                  () => Array(job.matrix[0].length).fill("")
                ),
                fillColor: "",
                fontBold: false,
                fontColor: ""
              });
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
          header.format.fill.color = WRITE_TABLE_HEADER_FILL;
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
        criteria.push(...matrixVerificationCriteria(targetName, matrix));
      }
      processedOutputCount += outputBatch.length;
      onProgress?.({
        stepId: `execution.action.${actionIndex + 1}.sheets`,
        title: "正在生成拆分工作表",
        summary: `已处理 ${processedOutputCount} / ${outputs.length} 张结果工作表。`,
        current: processedOutputCount,
        total: outputs.length,
        unit: "sheets",
        target: {
          sheet: jobs.length > 0 ? jobs[jobs.length - 1].targetName : action.sheet
        }
      });
    }
    return criteria;
  }

  if (action.type === "computeValues") {
    const worksheets = context.workbook.worksheets;
    const sourceSheet = worksheets.getItemOrNullObject(action.sourceSheet);
    sourceSheet.load("isNullObject,name");
    try {
      await context.sync();
    } catch (reason) {
      throw new Error(`定位源工作表失败：${excelErrorDetail(reason)}`);
    }
    if (sourceSheet.isNullObject) {
      throw new Error(`未找到源工作表「${action.sourceSheet}」`);
    }

    const sourceRange = action.sourceRange
      ? sourceSheet.getRange(action.sourceRange)
      : sourceSheet.getUsedRangeOrNullObject(true);
    sourceRange.load("isNullObject,rowCount,columnCount,rowIndex,columnIndex");
    try {
      await context.sync();
    } catch (reason) {
      throw new Error(`读取源数据范围失败：${excelErrorDetail(reason)}`);
    }
    if (sourceRange.isNullObject) {
      throw new Error(`源工作表「${action.sourceSheet}」没有可读取的数据`);
    }

    const request: DataToolRequest = {
      id: "computeValues",
      tool: "query_table",
      arguments: action.query
    };
    const accumulator = createQueryTableAccumulator(request);
    try {
      await readSheetIntoAccumulator(
        context,
        {
          name: action.sourceSheet,
          worksheet: sourceSheet,
          rowIndex: sourceRange.rowIndex,
          columnIndex: sourceRange.columnIndex,
          rowCount: sourceRange.rowCount,
          columnCount: sourceRange.columnCount
        },
        accumulator,
        {
          queryArguments: action.query,
          headerMode: action.hasHeaders ? "detect" : "synthesize"
        }
      );
    } catch (reason) {
      if (reason instanceof DataToolExecutionError) throw reason;
      throw new Error(`计算数据失败：${excelErrorDetail(reason)}`);
    }

    const result = accumulator.finish();
    const outputRows: typeof result.rows = action.includeHeaders
      ? [result.headers as typeof result.rows[number], ...result.rows]
      : result.rows;

    if (outputRows.length === 0) {
      return [{ type: "worksheetExists", sheet: action.sheet }];
    }

    // 检查目标工作表是否已存在（决定撤销行为）。
    const existingDestSheet = worksheets.getItemOrNullObject(action.sheet);
    existingDestSheet.load("isNullObject");
    await context.sync();
    const destSheet = existingDestSheet.isNullObject
      ? worksheets.add(action.sheet)
      : existingDestSheet;
    if (existingDestSheet.isNullObject) {
      collector.createdWorksheets.push(action.sheet);
    }

    const destStart = destSheet.getRange(action.destinationCell);
    const destRange = destStart.getResizedRange(
      outputRows.length - 1,
      (outputRows[0]?.length ?? 1) - 1
    );
    // 输出形状要算完才知道，所以撤销快照不能走 recoverableActionRange（那里
    // 只认动作体里写死的 range），必须在这里、写入之前手动拍一张。
    if (!existingDestSheet.isNullObject) {
      destRange.load(
        "address,formulas,numberFormat,format/fill/color,format/font/bold,format/font/color"
      );
      try {
        await context.sync();
        collector.overwrittenRanges.push({
          actionIndex,
          sheet: action.sheet,
          range: destRange.address.split("!").pop() ?? destRange.address,
          formulas: destRange.formulas.map((row) => row.map(normalizeValue)),
          numberFormat: destRange.numberFormat,
          fillColor: destRange.format.fill.color,
          fontBold: destRange.format.font.bold,
          fontColor: destRange.format.font.color
        });
      } catch {
        // 拍照失败不阻塞写入，只是这一步失去撤销能力。
      }
    }
    destRange.values = outputRows;
    if (action.includeHeaders && result.headers.length > 0) {
      const headerRange = destStart.getResizedRange(0, result.headers.length - 1);
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = "#DFF3E4";
    }
    destRange.format.autofitColumns();
    try {
      await context.sync();
    } catch (reason) {
      throw new Error(`写入计算结果失败：${excelErrorDetail(reason)}`);
    }
    return [{ type: "worksheetExists", sheet: action.sheet }];
  }

  const sheet = await getOrCreateWorksheet(context, action.sheet);

  switch (action.type) {
    case "createWorksheet":
      return;
    case "writeTable": {
      const values = writeTableValues(action);
      const start = sheet.getRange(action.startCell);
      const target = start.getResizedRange(values.length - 1, values[0].length - 1);
      target.values = values;
      if (action.headers.length > 0) {
        const header = start.getResizedRange(0, action.headers.length - 1);
        header.format.font.bold = true;
        header.format.fill.color = WRITE_TABLE_HEADER_FILL;
      }
      return;
    }
    case "writeValues": {
      assertRectangularValues(action.values);
      const target = sheet.getRange(action.range);
      target.values = action.values;
      return;
    }
    case "writeFormulas": {
      if (action.formulaR1C1) {
        assertSafeFormula(
          action.formulaR1C1,
          `${action.range} 的公式`,
          allowedExternalFiles
        );
      } else {
        action.formulas.forEach((row, rowIndex) =>
          row.forEach((formula, columnIndex) =>
            assertSafeFormula(
              formula,
              `${action.range} 第 ${rowIndex + 1} 行第 ${columnIndex + 1} 列`,
              allowedExternalFiles
            )
          )
        );
      }
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
    case "clearRange": {
      const clearType = (
        {
          all: "All",
          contents: "Contents",
          formats: "Formats",
          hyperlinks: "Hyperlinks"
        } as const
      )[action.applyTo];
      if (!action.filters?.length) {
        sheet.getRange(action.range).clear(clearType);
        return;
      }
      const target = sheet.getRange(action.range);
      target.load("values");
      await context.sync();
      const values = target.values as CellValue[][];
      const matchIndices = resolveFilteredRowIndices(
        values,
        (action.hasHeaders ?? true) ? 0 : -1,
        action.filters
      );
      for (const rowIndex of matchIndices) {
        target.getRow(rowIndex).clear(clearType);
      }
      return;
    }
    case "insertRange":
      sheet.getRange(action.range).insert(action.shift === "down" ? "Down" : "Right");
      return;
    case "deleteRange": {
      if (!action.filters?.length) {
        sheet
          .getRange(action.range)
          .delete(action.shift === "up" ? "Up" : "Left");
        return;
      }
      const target = sheet.getRange(action.range);
      target.load("values");
      await context.sync();
      const values = target.values as CellValue[][];
      const matchIndices = resolveFilteredRowIndices(
        values,
        (action.hasHeaders ?? true) ? 0 : -1,
        action.filters
      );
      // 从下往上删除命中行，避免行号错位
      for (let i = matchIndices.length - 1; i >= 0; i--) {
        target.getRow(matchIndices[i]).delete("Up");
      }
      return;
    }
    case "copyRange": {
      if (!action.filters?.length) {
        const sourceSheet = context.workbook.worksheets.getItem(
          action.sourceSheet
        );
        sheet.getRange(action.targetRange).copyFrom(
          sourceSheet.getRange(action.sourceRange),
          enumText(action.copyType) as Excel.RangeCopyType,
          action.skipBlanks,
          action.transpose
        );
        return;
      }
      const sourceSheet = context.workbook.worksheets.getItem(
        action.sourceSheet
      );
      const source = sourceSheet.getRange(action.sourceRange);
      source.load("values");
      await context.sync();
      const values = source.values as CellValue[][];
      const selected = filteredSourceRows(values, action.filters, action.hasHeaders ?? true);
      if (!selected.length) throw new Error("没有命中复制条件的行，原表未修改。");
      const writes = copyRangePlan(action, values, [0, ...selected]);
      const targetRect = parseExplicitA1Rect(action.targetRange)!;
      const outputRows = action.transpose ? values[0].length : selected.length + 1;
      const outputColumns = action.transpose ? selected.length + 1 : values[0].length;
      const outputAddress = matrixRange(action.targetRange.split(":")[0], outputRows, outputColumns)!;
      const outputRange = sheet.getRange(outputAddress);
      outputRange.load("values,formulas");
      await context.sync();
      const expected: CellValue[][] = (action.copyType === "link" ? outputRange.formulas : outputRange.values)
        .map(row => row.map(normalizeValue));
      // Save each actual write span: skipped blanks are neither written nor restored.
      const capturedWrites: Array<{write:typeof writes[number]; snapshot:ExecutionUndoSnapshot["ranges"][number]}> = [];
      for (const write of writes) {
        const captured = await captureUndoRange(context, write, actionIndex);
        if (!captured.snapshot) throw new Error("复制目标无法保存恢复依据，未写入。");
        capturedWrites.push({write,snapshot:captured.snapshot});
        const rect = parseExplicitA1Rect(write.range)!;
        const matrix = write.type === "writeValues" ? write.values : write.formulas;
        matrix.forEach((row, r) => row.forEach((value, c) => {
          expected[rect.startRow - targetRect.startRow + r][rect.startColumn - targetRect.startColumn + c] = value;
        }));
      }
      // Capturing spans required async host reads; verify again before queuing any writes.
      source.load("values");
      const targets = capturedWrites.map(({write})=>sheet.getRange(write.range));
      targets.forEach(target=>target.load("formulas"));
      await context.sync();
      if (JSON.stringify(source.values) !== JSON.stringify(values)) throw new Error("复制来源在检查期间发生变化，未写入。");
      for (let index=0;index<targets.length;index++) {
        if (JSON.stringify(targets[index].formulas.map(row=>row.map(normalizeValue))) !== JSON.stringify(capturedWrites[index].snapshot.formulas)) {
          throw new Error("复制目标在检查期间发生变化，未写入。");
        }
      }
      for (const [index,{write,snapshot}] of capturedWrites.entries()) {
        const target = targets[index];
        collector.overwrittenRanges.push(snapshot);
        if (write.type === "writeValues") target.values = write.values;
        else target.formulas = write.formulas;
      }
      return action.copyType === "link"
        ? [{type:"formulasEqual",sheet:action.sheet,range:outputAddress,expected:expected.map(row=>row.map(value=>String(value ?? "")))}]
        : [{type:"rangeEquals",sheet:action.sheet,range:outputAddress,expected}];
    }
    case "sortRange":
      sheet.getRange(action.range).sort.apply(
        action.keys.map((key) => ({ key: key.column, ascending: key.ascending })),
        false,
        action.hasHeaders
      );
      return;
    case "removeDuplicates": {
      if (!action.filters?.length) {
        sheet
          .getRange(action.range)
          .removeDuplicates(action.columns, action.hasHeaders);
        return;
      }
      const target = sheet.getRange(action.range);
      target.load("values");
      await context.sync();
      const values = target.values as CellValue[][];
      const headerRowIndex = action.hasHeaders ? 0 : -1;
      const matchIndices = resolveFilteredRowIndices(
        values,
        headerRowIndex,
        action.filters
      );
      const dataStart = headerRowIndex >= 0 ? headerRowIndex + 1 : 0;
      const seen = new Set<string>();
      const rowsToDelete: number[] = [];
      for (const rowIndex of matchIndices) {
        if (rowIndex < dataStart) continue;
        const key = action.columns
          .map((column) => {
            const value = values[rowIndex][column];
            return value === null || value === undefined
              ? "∅"
              : `${typeof value}:${String(value)}`;
          })
          .join("\u001f");
        if (seen.has(key)) {
          rowsToDelete.push(rowIndex);
        } else {
          seen.add(key);
        }
      }
      // 从下往上删除命中行，避免行号错位
      for (let i = rowsToDelete.length - 1; i >= 0; i--) {
        target.getRow(rowsToDelete[i]).delete("Up");
      }
      return;
    }
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
    case "mergeCells":
      sheet.getRange(action.range).merge(action.across);
      return;
    case "unmergeCells":
      sheet.getRange(action.range).unmerge();
      return;
    case "freezePanes":
      sheet.freezePanes.unfreeze();
      if (action.rows > 0) sheet.freezePanes.freezeRows(action.rows);
      if (action.columns > 0) sheet.freezePanes.freezeColumns(action.columns);
      return;
    case "setHyperlink":
      assertSafeHyperlink(action.address, `${action.range} 的超链接`);
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
    // 六个格式动作统一走 `applyFormatAction`：执行器与原生副本审阅必须同源，
    // 否则"副本里试算出来的样子"与"真正写回原表的样子"会分叉。
    case "setFill":
    case "setFont":
    case "setBorders":
    case "setAlignment":
    case "setNumberFormat":
    case "resizeRange":
      await applyFormatAction(context, sheet.getRange(action.range), action);
      return;
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

/**
 * 格式动作各自的恢复依据（用户要求：按动作**实际修改的属性**检查恢复能力，
 * 不笼统"纯格式动作"一刀切）。**恢复范围的唯一来源是 `undoRestoreScope`**，
 * 这里只说依据，不再另立一份动作名单 —— 两份名单迟早会对不上。
 *  - `setFill`/`setFont`/`setBorders`/`setAlignment`（`scope.cellFormat`）：改逐格格式，
 *    恢复依据是 `getDisplayedCellProperties`（ExcelApi 1.19）的**显示值**。
 *    宿主 1.9–1.17（批量许可永久版最高到 1.17）走**直接格式兼容路径**：逐格读
 *    `getCellProperties`（只返回直接设置在单元格上的属性 + 命名样式），撤销时
 *    先清空区域直接格式、再写回捕获到的直接格式集合 ⇒ 直接格式逐格精确还原、
 *    未直接设置的属性自动回到继承/默认。宿主连 1.9 都没有 ⇒ 执行前拒绝。
 *  - `setNumberFormat`：改 `numberFormat`，恢复依据是 `range.numberFormat`（独立捕获，
 *    不依赖 1.19）。
 *  - `resizeRange`：改行高/列宽，恢复依据是 `getRowProperties`/`getColumnProperties`
 *    （1.9，含 `useStandardHeight`/`useStandardWidth`）。
 */

/**
 * 当前宿主的格式能力。**唯一入口**，`captureUndoRange`、预检与错误文案共用，
 * 避免"一处按 1.19 判、另一处按 1.9 判"这种两套口径。
 */
export function currentExcelFormatCapability(): ExcelFormatCapability {
  // 注意签名要**收下两个参数**：`resolveExcelFormatCapability` 会传 `("ExcelApi", "1.19")`。
  // 只声明一个参数会把 `version` 接成 `"ExcelApi"`，从而永远判成"不支持"。
  return resolveExcelFormatCapability((_set, version) =>
    Office.context.requirements.isSetSupported("ExcelApi", version)
  );
}

/**
 * 动作**实际影响** → 撤销恢复范围。纯函数，不依赖宿主，便于单测与变异。
 *
 * 逐条依据（都是"没有它就不用恢复"）：
 *  - `writeValues`/`writeFormulas`：只动值/公式，**不动任何格式**。
 *  - `writeTable`：写表头与数据行，同样是值。
 *  - `clearRange`：能走到撤销的只有 `applyTo === "contents"`（见 `recoverableActionRange`），
 *    清的是内容，不是格式。
 *  - `setNumberFormat`：只动数字格式。
 *  - `setFill`/`setFont`/`setBorders`/`setAlignment`：只动逐格格式 ⇒ `cellFormat`，
 *    并**连带 `numberFormat`**（兼容路径恢复用 `clear(Formats)`，它连数字格式一起清）。
 *  - `resizeRange`：行高/列宽是行级/列级属性，**绝不碰单元格格式**；只在该动作确实给了
 *    对应字段时才恢复那一个。
 *  - 其余（删表/建表等）没有可撤销的范围内格式。
 */
export function undoRestoreScope(action: ExcelAction): UndoRestoreScope {
  const none: UndoRestoreScope = {
    formulas: false,
    numberFormat: false,
    cellFormat: false,
    rowHeight: false,
    columnWidth: false
  };
  switch (action.type) {
    case "writeValues":
    case "writeFormulas":
    case "writeTable":
    case "clearRange":
      return { ...none, formulas: true };
    case "setNumberFormat":
      return { ...none, numberFormat: true };
    case "setFill":
    case "setFont":
    case "setBorders":
    case "setAlignment":
      // cellFormat 连带 numberFormat：兼容路径的 `clear(Formats)` 会清掉直接数字格式，
      // 不一起恢复就等于把用户的数字格式弄丢。机制约束，写在这里以免调用方漏掉。
      return { ...none, cellFormat: true, numberFormat: true };
    case "resizeRange":
      return {
        ...none,
        rowHeight: action.rowHeight !== undefined && action.rowHeight !== null,
        columnWidth: action.columnWidth !== undefined && action.columnWidth !== null
      };
    default:
      return none;
  }
}

/**
 * 旧恢复点（无 `scope`）的隐含范围：**全范围恢复**。升级不能改变旧记录的行为 ——
 * 把一个没有范围信息的旧快照按"只恢复值"处理，会让它该恢复的格式恢复不回来。
 */
const LEGACY_FULL_SCOPE: UndoRestoreScope = {
  formulas: true,
  numberFormat: true,
  cellFormat: true,
  rowHeight: true,
  columnWidth: true
};

/** 快照的恢复范围；缺失即旧恢复点，按全范围处理。 */
function scopeOf(saved: { scope?: UndoRestoreScope }): UndoRestoreScope {
  return saved.scope ?? LEGACY_FULL_SCOPE;
}

/**
 * 一个动作的**可恢复目标**（单个 A1 矩形）；结构类动作与多矩形清除返回 `null`。
 *
 * 导出给「重开后重建撤销」用（`App` 把计划动作翻成目标清单喂给共享判据）：
 * 它必须与本文件里恢复/撤销真正用的地址**同一个函数**，否则"撤销的目标"和
 * "执行的目标"就是两份互相独立的口径。
 */
export function recoverableActionRange(action: ExcelAction): string | null {
  switch (action.type) {
    case "writeValues":
    case "writeFormulas":
    case "setFill":
    case "setFont":
    // 以下动作只改格式，但格式现在也会被捕获与恢复。
    case "setBorders":
    case "setAlignment":
    case "setNumberFormat":
    case "resizeRange":
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

// Excel.BorderIndex 是运行时枚举，模块加载期取值会在非 Excel 环境（测试、
// 文件夹模式）里抛 ReferenceError，因此推迟到实际捕获格式时才求值。
function undoBorderSides(): Excel.BorderIndex[] {
  return [
    Excel.BorderIndex.edgeTop,
    Excel.BorderIndex.edgeBottom,
    Excel.BorderIndex.edgeLeft,
    Excel.BorderIndex.edgeRight,
    Excel.BorderIndex.insideHorizontal,
    Excel.BorderIndex.insideVertical
  ];
}

/**
 * `Excel.BorderIndex` 枚举值 → 协议里的边名。
 *
 * 只覆盖 `undoBorderSides()` 那六条，所以是一张固定的小表。**不要去拼枚举成员名**
 * （`edgeTop` → `top`）：成员名与协议边名本来就不同形，靠 `toLowerCase`/`slice`
 * 去凑，遇到 `insideHorizontal` 这种就错了，而且错得静默。
 */
function borderSideName(sideIndex: Excel.BorderIndex): BorderSide {
  switch (sideIndex) {
    case Excel.BorderIndex.edgeTop:
      return "top";
    case Excel.BorderIndex.edgeBottom:
      return "bottom";
    case Excel.BorderIndex.edgeLeft:
      return "left";
    case Excel.BorderIndex.edgeRight:
      return "right";
    case Excel.BorderIndex.insideHorizontal:
      return "insideHorizontal";
    case Excel.BorderIndex.insideVertical:
      return "insideVertical";
    default:
      throw new Error(`未知的边框位置：${String(sideIndex)}`);
  }
}

/**
 * 逐格采集的请求参数。**两条路径各一份**：
 *  - 显示值路径（1.19）用宿主的显示值语义；
 *  - 直接格式兼容路径（1.9–1.17）走 `clear(Formats)`，**凡是能被 `clear` 重置的属性
 *    都要请求**（字体删除线/上下标/底纹、填充图案色/底纹、八条边与边底纹、缩进/
 *    收缩适应/文字方向/阅读顺序/自动缩进/单元格保护），否则"清空再写回"会静默丢属性。
 *    **不请求就没有、不写回就等于丢**，所以这里是本机可表达的全量清单。
 *
 * **仍然不是"全部格式"**：单元格内部富文本（`textRuns`，ExcelApi 1.18）在 1.17 宿主上
 * 不可请求 —— **既读不到、也写不回**，因此本机无法判断某个单元格是否含富文本，也无法在
 * 富文本被改变时还原它。
 *
 * 关于 `clear(Formats)`：官方 Excel 语义说它**不清除**单元格内字符级格式
 * （"Clear Formats 不改单元格内单个字符的格式"）。这是**对 clear 这一步的语义依据**，
 * **不等于"含富文本的区域可以安全执行并撤销"** —— 完整写入与撤销链路**未经真实宿主验证**，
 * 且一旦宿主行为与官方语义不符，本机既发现不了也恢复不了。故此处只声明依据，不给出整体安全结论。
 * 如实声明边界，不补默认值、不声称已覆盖（见 `DirectCellFormatValues` 的同类说明）。
 */
const DISPLAYED_FORMAT_LOAD_OPTIONS = {
  fill: { color: true, pattern: true },
  font: { name: true, size: true, bold: true, italic: true, color: true, underline: true },
  borders: { style: true, color: true, weight: true },
  horizontalAlignment: true,
  verticalAlignment: true,
  wrapText: true,
  indentLevel: true
} as const;

export const DIRECT_FORMAT_LOAD_OPTIONS = {
  horizontalAlignment: true,
  verticalAlignment: true,
  wrapText: true,
  indentLevel: true,
  shrinkToFit: true,
  textOrientation: true,
  readingOrder: true,
  autoIndent: true,
  // `CellPropertiesFormatLoadOptions.protection` 是布尔开关（不是对象）。
  protection: true,
  fill: {
    color: true,
    pattern: true,
    patternColor: true,
    patternTintAndShade: true,
    tintAndShade: true
  },
  font: {
    name: true,
    size: true,
    bold: true,
    italic: true,
    color: true,
    underline: true,
    strikethrough: true,
    subscript: true,
    superscript: true,
    tintAndShade: true
  },
  borders: { style: true, color: true, weight: true, tintAndShade: true }
} as const;

/**
 * `getCellProperties` 的入参。**`style` 与 `format` 同级**（`CellPropertiesLoadOptions.style`，
 * ExcelApi 1.9）—— 单元格命名样式不是 `format` 的子字段，请求时必须并排给。
 */
export const DIRECT_CELL_PROPERTIES_LOAD = {
  style: true,
  format: DIRECT_FORMAT_LOAD_OPTIONS
} as const;

/**
 * 目标区域的格式**无法可靠恢复**（兼容路径下的条件格式 / 合并区域，或宿主缺 API）。
 *
 * 刻意做成独立类型：调用方要把它与"区域过大/矩阵不完整"（返回 null）区分开，
 * 因为两者的用户指引不同 —— 前者是"这个区域本身不可恢复"，后者是"这次没采全"。
 */
class FormatRangeUnrecoverableError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "FormatRangeUnrecoverableError";
  }
}

/**
 * 直接格式路径要采的边名：`Excel.CellBorderCollection` 的**八条**。
 *
 * 不能只采四条 —— 内部边（`horizontal`/`vertical`）与斜线（`diagonalDown`/`diagonalUp`）
 * 同样是"直接设置过"的格式，兼容路径恢复要先 `clear(Formats)`，漏采即静默丢失。
 */
const CELL_BORDER_SIDES: readonly (keyof NonNullable<
  DirectCellFormatValues["borders"]
>)[] = [
  "top",
  "bottom",
  "left",
  "right",
  "horizontal",
  "vertical",
  "diagonalDown",
  "diagonalUp"
];

/**
 * 把 `getCellProperties` 的返回值读成**直接格式子集**。
 *
 * 关键：**只装宿主确实返回了的属性**。没有的字段一概不出现 —— 缺席表示"该属性未直接设置"，
 * 是"继承样式/条件格式/默认"这一确定状态。绝不在这里补 `defaultCellStyle()` 之类的默认值，
 * 否则"未直接设置"会被写成"显式设置成默认"，继承语义当场丢失。
 *
 * `style` 是 `CellPropertiesLoadOptions.style` 读回来的**命名样式名**，与 `format` 同级，
 * 因此由调用方单独传进来（而不是从 `cellFormat` 里取）。
 */
function readDirectCellFormat(
  cellFormat: Excel.CellPropertiesFormat | undefined,
  style?: string
): DirectCellFormatValues {
  const direct: DirectCellFormatValues = {};
  if (style != null) direct.style = style;
  if (!cellFormat) return direct;
  const fill = cellFormat.fill;
  if (fill) {
    if (fill.pattern != null) direct.fillPattern = String(fill.pattern);
    if (fill.color != null) direct.fillColor = fill.color;
    if (fill.patternColor != null) direct.fillPatternColor = fill.patternColor;
    if (fill.patternTintAndShade != null) {
      direct.fillPatternTintAndShade = fill.patternTintAndShade;
    }
    if (fill.tintAndShade != null) direct.fillTintAndShade = fill.tintAndShade;
  }
  const font = cellFormat.font;
  if (font) {
    if (font.name != null) direct.fontName = String(font.name);
    if (font.size != null) direct.fontSize = font.size;
    if (font.bold != null) direct.fontBold = font.bold;
    if (font.italic != null) direct.fontItalic = font.italic;
    if (font.color != null) direct.fontColor = font.color;
    if (font.underline != null) direct.fontUnderline = String(font.underline);
    if (font.strikethrough != null) direct.fontStrikethrough = font.strikethrough;
    if (font.subscript != null) direct.fontSubscript = font.subscript;
    if (font.superscript != null) direct.fontSuperscript = font.superscript;
    if (font.tintAndShade != null) direct.fontTintAndShade = font.tintAndShade;
  }
  if (cellFormat.horizontalAlignment != null) {
    direct.horizontalAlignment = String(cellFormat.horizontalAlignment);
  }
  if (cellFormat.verticalAlignment != null) {
    direct.verticalAlignment = String(cellFormat.verticalAlignment);
  }
  if (cellFormat.wrapText != null) direct.wrapText = cellFormat.wrapText;
  if (cellFormat.indentLevel != null) direct.indentLevel = cellFormat.indentLevel;
  if (cellFormat.shrinkToFit != null) direct.shrinkToFit = cellFormat.shrinkToFit;
  if (cellFormat.textOrientation != null) direct.textOrientation = cellFormat.textOrientation;
  if (cellFormat.readingOrder != null) direct.readingOrder = String(cellFormat.readingOrder);
  if (cellFormat.autoIndent != null) direct.autoIndent = cellFormat.autoIndent;
  const protection = cellFormat.protection;
  if (protection) {
    if (protection.locked != null) direct.locked = protection.locked;
    if (protection.formulaHidden != null) direct.formulaHidden = protection.formulaHidden;
  }
  const cellBorders = cellFormat.borders;
  if (cellBorders) {
    const sides: NonNullable<DirectCellFormatValues["borders"]> = {};
    for (const side of CELL_BORDER_SIDES) {
      const state = cellBorders[side];
      // 只有"确实带了线型"的边才算直接设置过；`style` 缺席 = 该边未直接设置。
      if (state && state.style != null) {
        const captured: CellBorderState = {
          style: String(state.style),
          color: state.color ?? "#000000",
          weight: state.weight == null ? "Thin" : String(state.weight)
        };
        if (state.tintAndShade != null) captured.tintAndShade = state.tintAndShade;
        sides[side] = captured;
      }
    }
    if (Object.keys(sides).length > 0) direct.borders = sides;
  }
  return direct;
}

/**
 * 某格**直接格式**在指定属性上的取值（副本审阅用它做逐格"是否变化"的比对）。
 *
 * 未直接设置 ⇒ `null`（"继承/默认"这一确定状态），不是"未知"。返回值只装格式化后的
 * 取值，可安全 `JSON.stringify` 比较。
 */
export function directFormatFieldValue(
  cellFormat: Excel.CellPropertiesFormat | undefined,
  field: ExcelFormatField
): unknown {
  const direct = readDirectCellFormat(cellFormat);
  switch (field) {
    case "fill":
      return direct.fillPattern == null
        ? null
        : { pattern: direct.fillPattern, color: direct.fillColor ?? null };
    case "fontName":
      return direct.fontName ?? null;
    case "fontSize":
      return direct.fontSize ?? null;
    case "fontBold":
      return direct.fontBold ?? null;
    case "fontItalic":
      return direct.fontItalic ?? null;
    case "fontColor":
      return direct.fontColor ?? null;
    case "fontUnderline":
      return direct.fontUnderline ?? null;
    case "horizontalAlignment":
      return direct.horizontalAlignment ?? null;
    case "verticalAlignment":
      return direct.verticalAlignment ?? null;
    case "wrapText":
      return direct.wrapText ?? null;
    case "indentLevel":
      return direct.indentLevel ?? null;
    case "borders":
      return direct.borders ?? null;
    default:
      return null;
  }
}

/** 该格的**直接格式**里是否包含某个属性（用于判定它这次是"直接"还是"继承"）。 */
function directFormatHasField(
  direct: DirectCellFormatValues,
  field: ExcelFormatField
): boolean {
  switch (field) {
    case "fill":
      return direct.fillPattern != null;
    case "fontName":
      return direct.fontName != null;
    case "fontSize":
      return direct.fontSize != null;
    case "fontBold":
      return direct.fontBold != null;
    case "fontItalic":
      return direct.fontItalic != null;
    case "fontColor":
      return direct.fontColor != null;
    case "fontUnderline":
      return direct.fontUnderline != null;
    case "horizontalAlignment":
      return direct.horizontalAlignment != null;
    case "verticalAlignment":
      return direct.verticalAlignment != null;
    case "wrapText":
      return direct.wrapText != null;
    case "indentLevel":
      return direct.indentLevel != null;
    case "borders":
      return direct.borders != null;
    case "rowHeight":
    case "columnWidth":
    case "numberFormat":
      // 这三个不是逐格属性：行高/列宽走 rowHeights/columnWidths，numberFormat 走
      // `range.numberFormat`，都不在逐格矩阵里，因此不参与"直接/继承"的逐格判定。
      return true;
    default:
      return true;
  }
}

/**
 * `ExcelFormatField` → 它管辖的**直接格式**键。
 *
 * 这是"按动作实际影响拆分恢复"落到**属性级**的那一层：`setFont{bold}` 只该恢复
 * `fontBold`，不该顺手把 `textOrientation`/`shrinkToFit`/`locked` 一起写回 —— 后者是
 * 用户可能在本动作之后自己改的属性，写回"执行前的值"就等于覆盖用户的编辑。
 */
const FIELD_DIRECT_KEYS: ReadonlyArray<[ExcelFormatField, readonly (keyof DirectCellFormatValues)[]]> = [
  [
    "fill",
    [
      "fillPattern",
      "fillColor",
      "fillPatternColor",
      "fillPatternTintAndShade",
      "fillTintAndShade"
    ]
  ],
  ["fontName", ["fontName"]],
  ["fontSize", ["fontSize"]],
  ["fontBold", ["fontBold"]],
  ["fontItalic", ["fontItalic"]],
  ["fontColor", ["fontColor"]],
  ["fontUnderline", ["fontUnderline"]],
  ["horizontalAlignment", ["horizontalAlignment"]],
  ["verticalAlignment", ["verticalAlignment"]],
  ["wrapText", ["wrapText"]],
  ["indentLevel", ["indentLevel"]],
  ["borders", ["borders"]]
];

/** 本动作实际修改的属性所管辖的直接格式键集合。 */
function directKeysOfFields(fields: readonly string[]): Set<keyof DirectCellFormatValues> {  const keys = new Set<keyof DirectCellFormatValues>();
  for (const [field, owned] of FIELD_DIRECT_KEYS) {
    if (fields.includes(field)) owned.forEach((key) => keys.add(key));
  }
  return keys;
}

/** `ExcelFormatField` → 它管辖的**显示值**字段（`UndoCellFormat` 的键）。 */
const FIELD_DISPLAYED_KEYS: ReadonlyArray<[ExcelFormatField, readonly (keyof UndoCellFormat)[]]> = [
  ["fill", ["fillPattern", "fillColor"]],
  ["fontName", ["fontName"]],
  ["fontSize", ["fontSize"]],
  ["fontBold", ["fontBold"]],
  ["fontItalic", ["fontItalic"]],
  ["fontColor", ["fontColor"]],
  ["fontUnderline", ["fontUnderline"]],
  ["horizontalAlignment", ["horizontalAlignment"]],
  ["verticalAlignment", ["verticalAlignment"]],
  ["wrapText", ["wrapText"]],
  ["indentLevel", ["indentLevel"]],
  ["borders", ["borders"]]
];

/** 本动作实际修改的属性所管辖的显示值字段集合。 */
function displayedKeysOfFields(fields: readonly string[]): Set<keyof UndoCellFormat> {
  const keys = new Set<keyof UndoCellFormat>();
  for (const [field, owned] of FIELD_DISPLAYED_KEYS) {
    if (fields.includes(field)) owned.forEach((key) => keys.add(key));
  }
  return keys;
}

/**
 * 旧恢复点没有 `formatFields` 时的兜底：按**全部**逐格格式属性算摘要/恢复。
 * 它比按动作收窄的版本更严，但对旧记录只能如此 —— 缺信息时宁可保守，也不能假装知道
 * "这次只改了哪几个属性"。它不会恢复出错误结果，只会更早报冲突。
 */
const ALL_DIRECT_FORMAT_FIELDS: readonly string[] = FIELD_DIRECT_KEYS.map(([field]) => field);

/** 从一条记录里只挑出指定键（缺席的键不出现在结果里 ⇒ 不会被误当成"写成 null"）。 */
function pickKeys<T extends object>(
  source: T,
  keys: ReadonlySet<keyof T>
): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(source)) {
    if (value !== undefined && keys.has(key as keyof T)) out[key] = value;
  }
  return out;
}

/**
 * 冲突摘要用的投影：只保留本动作管辖的键，缺席记为 `null`（"未直接设置"是确定状态）。
 *
 * `keys` 收 `ReadonlySet<string>`（而不是 `keyof T`）：聚合摘要（`Set<keyof DirectCellFormatValues>`）
 * 与逐属性摘要（同一集合的子集）都要用它，统一到同一份实现才能保证两者口径一致。
 */
function projectKeys(source: object, keys: ReadonlySet<string>): Record<string, unknown> {
  const record = source as Record<string, unknown>;
  const out: Record<string, unknown> = {};
  for (const key of keys) out[key] = record[key] ?? null;
  return out;
}

/** `Set<具体字面量联合>` → `ReadonlySet<string>`（只为通过类型，不改运行期集合）。 */
function asKeySet(keys: ReadonlySet<unknown>): ReadonlySet<string> {
  return keys as unknown as ReadonlySet<string>;
}

/**
 * 把一个**格式动作**应用到给定区域。
 *
 * 抽成导出函数是为了让"执行器"与"Excel 原生副本审阅"用**同一份**写入逻辑：
 * 副本里试算的格式必须与真正写回原表时逐字节同源，否则"审阅通过、写回却不一样"，
 * 副本审阅就失去了意义。属性判空一律用 `!= null`（`bold: false` 是"取消加粗"，
 * 与"没给"不同）。
 */
export async function applyFormatAction(
  context: Excel.RequestContext,
  range: Excel.Range,
  action: ExcelAction
): Promise<void> {
  switch (action.type) {
    case "setFill":
      range.format.fill.color = action.color;
      return;
    case "setFont": {
      const font = range.format.font;
      if (action.bold != null) font.bold = action.bold;
      if (action.italic != null) font.italic = action.italic;
      if (action.name != null) font.name = action.name;
      if (action.size != null) font.size = action.size;
      if (action.color != null) font.color = action.color;
      return;
    }
    case "setNumberFormat": {
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
      // 缺省补齐只此一处（`normalizeBorderAction`），回读判据走同一条 ——
      // 模型省略 style/color/weight 时，写入值与期望值必须来自同一个默认。
      const borders = normalizeBorderAction(action);
      const target = range.format.borders;
      for (const side of borders.sides) {
        const border = target.getItem(borderIndexes[side]);
        border.style = enumText(borders.style) as Excel.BorderLineStyle;
        border.color = borders.color;
        border.weight = enumText(borders.weight) as Excel.BorderWeight;
      }
      return;
    }
    case "setAlignment": {
      const format = range.format;
      if (action.horizontal) {
        format.horizontalAlignment = enumText(action.horizontal) as Excel.HorizontalAlignment;
      }
      if (action.vertical) {
        format.verticalAlignment = enumText(action.vertical) as Excel.VerticalAlignment;
      }
      if (action.wrapText !== undefined && action.wrapText !== null) {
        format.wrapText = action.wrapText;
      }
      return;
    }
    case "resizeRange": {
      const format = range.format;
      if (action.rowHeight !== undefined && action.rowHeight !== null) {
        format.rowHeight = action.rowHeight;
      }
      if (action.columnWidth !== undefined && action.columnWidth !== null) {
        format.columnWidth = action.columnWidth;
      }
      return;
    }
    default:
      throw new Error(`不是格式动作：${action.type}`);
  }
}

/**
 * 显示值路径下把一格的 `format` 投影成 `UndoCellFormat`（全字段都是**确定的显示值**）。
 *
 * 抽出来是为了让"采集快照"与"执行后状态摘要"用**同一份投影** —— 两处各写一份，
 * 迟早会漂移成"摘要说没变、快照却说变了"。`format` 缺失（宿主异常）⇒ 返回 null，
 * 由调用方决定是拒绝（采集）还是计入摘要（比对）。
 */
function displayedCellFormatOf(
  cellFormat: Excel.CellPropertiesFormat | undefined
): UndoCellFormat | null {
  if (!cellFormat) return null;
  const cellBorders = cellFormat.borders ?? {};
  const toCellBorder = (state: Excel.CellBorder | undefined): CellBorderState => ({
    style: state?.style == null ? "None" : String(state.style),
    color: state?.color ?? "#000000",
    weight: state?.weight == null ? "Thin" : String(state.weight)
  });
  const font = cellFormat.font ?? {};
  const fill = cellFormat.fill;
  // 显示填充：`pattern` 缺失或为 `"None"` 都表示"显示上没有填充"（明确无填充）。
  // 注意：这是**显示值**的语义 —— 若填充来自样式/条件格式，`pattern` 会是 `"Solid"` 等，
  // 而不是 `"None"`；`"None"` 只表示"显示上确实没有填充"。
  const fillPattern = fill?.pattern ?? "None";
  return {
    fillPattern,
    fillColor: fillPattern === "None" ? null : (fill?.color ?? null),
    fontName: font.name ?? "",
    fontSize: font.size ?? 11,
    fontBold: font.bold ?? false,
    fontItalic: font.italic ?? false,
    fontColor: font.color ?? "",
    fontUnderline: font.underline == null ? "None" : String(font.underline),
    horizontalAlignment:
      cellFormat.horizontalAlignment == null
        ? "General"
        : String(cellFormat.horizontalAlignment),
    verticalAlignment:
      cellFormat.verticalAlignment == null ? "Bottom" : String(cellFormat.verticalAlignment),
    wrapText: cellFormat.wrapText ?? false,
    indentLevel: cellFormat.indentLevel ?? 0,
    borders: {
      top: toCellBorder(cellBorders.top),
      bottom: toCellBorder(cellBorders.bottom),
      left: toCellBorder(cellBorders.left),
      right: toCellBorder(cellBorders.right)
    }
  };
}

/**
 * 捕获超出常用格式之外的属性，**按属性实际粒度**：
 *  - 逐格字体/对齐/填充/边框：一次读回二维矩阵，混合值不会被折叠成"区域级单值"
 *    （这是旧实现丢失混合格式的根因）。
 *  - 逐行行高：`range.getRowProperties()`（`format.rowHeight`，行级属性）。
 *  - 逐列列宽：`range.getColumnProperties()`（`format.columnWidth`，列级属性）。
 *
 * 两条逐格路径，由宿主能力决定（见 `currentExcelFormatCapability`）：
 *  - **显示值路径**（ExcelApi ≥ 1.19）：`getDisplayedCellProperties` 返回单元格**实际显示**
 *    的格式，无论来源是直接设置还是表格样式/命名样式/条件格式 ⇒ 写进 `cells`。
 *  - **直接格式兼容路径**（ExcelApi 1.9–1.17；批量许可永久版的实际上限）：`getCellProperties`
 *    只返回**直接设置在单元格上**的属性，缺席即"未直接设置（继承/默认）"这一**确定状态**
 *    ⇒ 写进 `directCells`，并记录本次动作**实际修改的属性**里哪些是继承态。
 *    兼容路径下目标区域若含**条件格式**或**合并单元格**，显示状态无法在执行前确定，
 *    且条件格式/表格样式可能覆盖本次写入的显示效果，因此**明确拒绝**
 *    （`FormatRangeUnrecoverableError`），不写默认值冒充"采集成功"。
 *
 * 关键语义（"未采集 ≠ 默认值"）：两条路径都**不填默认值** —— `cells` 的每个字段都来自
 * 宿主返回值，`directCells` 只装宿主确认直接设置过的字段；矩阵形状不符/元素非对象一律
 * 视为"不可完整捕获"（返回 null），由调用方在**原表修改前**拒绝执行。
 *
 * `modifiedFields` 是本次动作**实际要改**的属性（`excelFormatActionFields`）。它只用于
 * 记录来源（`inherited`），不影响恢复 —— 恢复写回的是**整份直接格式集合**。
 *
 * `scope` 决定**采什么**（用户要求：按动作实际影响拆分恢复）：
 *  - `scope.cellFormat` 假 ⇒ 不请求逐格矩阵，也不查条件格式/合并区域（纯 `resizeRange`
 *    与"区域里有条件格式"无关，不该被误拒）；
 *  - `scope.rowHeight`/`scope.columnWidth` 皆假 ⇒ 不请求逐行行高/逐列列宽。
 * 调用方对"没有任何格式范围"的动作（纯值/公式、`setNumberFormat`）干脆不调用本函数，
 * 快照的 `format` 直接为 null —— **撤销因此一个格式写入都不做**。
 */
async function captureRangeFormat(
  context: Excel.RequestContext,
  range: Excel.Range,
  cellCount: number,
  rowCount: number,
  columnCount: number,
  modifiedFields: readonly ExcelFormatField[],
  scope: UndoRestoreScope
): Promise<UndoRangeFormat | null> {
  if (cellCount > UNDO_FORMAT_CAPTURE_MAX_CELLS) return null;
  const capability = currentExcelFormatCapability();
  // 连 1.9（getCellProperties/setCellProperties）都没有：逐格格式既读不出也写不回，
  // 直接拒绝而不是交一份"半份快照"出去 —— 半份快照会让撤销静默失效。
  if (capability.path === "unavailable" && scope.cellFormat) {
    throw new FormatRangeUnrecoverableError(
      `本机 Excel 缺少 ${EXCEL_DIRECT_FORMAT_REQUIREMENT}，无法读取或写回逐格格式`
    );
  }
  const format = range.format;
  format.load(
    "horizontalAlignment,verticalAlignment,wrapText,rowHeight,columnWidth,indentLevel"
  );
  format.font.load("name,size,italic,underline");
  const borders = undoBorderSides().map((sideIndex) => {
    const border = format.borders.getItem(sideIndex);
    border.load("style,color,weight");
    return { sideIndex, border };
  });

  // 逐格采集请求：显示值路径与兼容路径**二选一**，同一批 sync 取回，不额外增加往返。
  //
  // 只在动作**确实会改逐格格式**时请求。值/公式动作与 `setNumberFormat` 的范围里没有
  // `cellFormat`，撤销时也就一个格式写入都不做 ⇒ 它们的混合格式**不采也不会坏**
  // （不被写回 = 原样保留）。这比"为了撤销而先把格式采一遍"更省，也避免了在
  // 只改值的动作上因为"区域含条件格式"而被无端拒绝。
  const displayedAvailable = capability.path === "displayed";
  const needCellFormat = scope.cellFormat;
  const cellProperties =
    needCellFormat && displayedAvailable
      ? range.getDisplayedCellProperties({ format: DISPLAYED_FORMAT_LOAD_OPTIONS })
      : null;
  const directProperties =
    needCellFormat && !displayedAvailable
      ? range.getCellProperties(DIRECT_CELL_PROPERTIES_LOAD)
      : null;
  // 兼容路径的两个"能不能可靠恢复"判据。**只有真的会写回逐格格式时才查**：
  // 纯 `resizeRange`（行高/列宽）与"区域里有条件格式"无关，不该被误拒。
  const conditionalFormats = directProperties ? range.conditionalFormats.getCount() : null;
  const mergedAreas = directProperties ? range.getMergedAreasOrNullObject() : null;
  mergedAreas?.load("isNullObject");
  const dataValidation = directProperties ? range.dataValidation : null;
  dataValidation?.load("type");

  // 逐行行高 / 逐列列宽：行高是行级、列宽是列级，区域级 `format.rowHeight/columnWidth`
  // 对混合尺寸不可靠，必须逐行/逐列读。`useStandardHeight`/`useStandardWidth` 一并请求，
  // 用于区分"直接设置的行高/列宽"与"未直接设置（标准行高/列宽）"—— 后者是确定状态，
  // 不是"无法确定"。**只在动作真的会改行高/列宽时请求**（即 `resizeRange`）。
  const needRowHeight = scope.rowHeight;
  const needColumnWidth = scope.columnWidth;
  const rowProperties = needRowHeight
    ? range.getRowProperties({ format: { rowHeight: true, useStandardHeight: true } })
    : null;
  const columnProperties = needColumnWidth
    ? range.getColumnProperties({ format: { columnWidth: true, useStandardWidth: true } })
    : null;
  await context.sync();

  // 兼容路径的硬性拒绝：条件格式支配显示、合并区域不逐格成立。都在**写入前**抛出，
  // 且文案指名缺哪个 API / 哪类区域，避免"笼统报错"。
  if (conditionalFormats && conditionalFormats.value > 0) {
    throw new FormatRangeUnrecoverableError(
      `目标区域含 ${conditionalFormats.value} 条条件格式；本机缺少 ${EXCEL_DISPLAYED_FORMAT_REQUIREMENT}，` +
        "条件格式支配的显示状态无法在执行前确定，也不能保证本次写入的显示生效"
    );
  }
  if (mergedAreas && !mergedAreas.isNullObject) {
    throw new FormatRangeUnrecoverableError(
      "目标区域含合并单元格；合并区域的格式不逐格成立，无法按格恢复执行前的格式来源"
    );
  }
  // `clear(Formats)` 会清除区域上的数据验证，而兼容路径既不采集也不逐格还原数据验证
  // （数据验证不是格式动作的管辖范围，本机也读不回逐格验证来源）。因此含数据验证的区域
  // 必须在写入前 fail-closed 拒绝，与条件格式 / 合并单元格同口径，而不是静默丢验证。
  if (dataValidation && dataValidation.type !== "None") {
    throw new FormatRangeUnrecoverableError(
      "目标区域含数据验证；恢复用的 clear(Formats) 会清除数据验证且无法逐格还原，本机缺少可验证的保留路径"
    );
  }

  // —— 矩阵尺寸 / 行列数量 / 缺项校验：任一不满足即视为"格式不可完整捕获"（返回 null），
  // 由调用方对纯格式动作执行前拒绝，而不是执行完才发现撤销不了。
  // 未请求的行高/列宽（动作不改它）不参与校验 —— 缺席就是"不在范围内"。
  const rowValues = rowProperties?.value;
  const columnValues = columnProperties?.value;
  if (needRowHeight && (!Array.isArray(rowValues) || rowValues.length !== rowCount)) return null;
  if (needColumnWidth && (!Array.isArray(columnValues) || columnValues.length !== columnCount)) {
    return null;
  }

  let cells: UndoCellFormat[][] | null = null;
  if (cellProperties) {
    const cellMatrix = cellProperties.value;
    if (!Array.isArray(cellMatrix) || cellMatrix.length !== rowCount) return null;
    for (const row of cellMatrix) {
      if (!Array.isArray(row) || row.length !== columnCount) return null;
    }
    // 采集失败判据：矩阵里出现**非对象**的元素（宿主异常、返回被截断成 undefined）。
    // `getDisplayedCellProperties` 对每个单元格都应返回一个对象（含 `format`），元素为
    // null/非对象即宿主异常。
    if (
      cellMatrix.some((row) =>
        row.some((cell) => cell == null || typeof cell !== "object")
      )
    ) {
      return null;
    }

    cells = cellMatrix.map((row) =>
      row.map((cell) => {
        // `getDisplayedCellProperties` 返回的是**显示值**，`format` 是确定对象。
        // 若某格的 `format` 缺失（宿主异常），该格显示值不可确定 ⇒ 整个快照不可靠。
        const projected = displayedCellFormatOf((cell as Excel.CellProperties).format);
        if (!projected) throw new DisplayedFormatUnavailableError();
        return projected;
      })
    );
  }

  // —— 兼容路径（宿主 < 1.19）：逐格读**直接格式**，并按本次动作实际修改的属性
  // 记录"哪些是继承态"。缺席 ≠ 默认值：缺席就是"未直接设置"，不补任何默认。
  let directCells: UndoCellDirectFormat[][] | null = null;
  const inheritedFields = new Set<string>();
  if (directProperties) {
    const cellMatrix = directProperties.value;
    if (!Array.isArray(cellMatrix) || cellMatrix.length !== rowCount) return null;
    for (const row of cellMatrix) {
      if (!Array.isArray(row) || row.length !== columnCount) return null;
    }
    if (
      cellMatrix.some((row) =>
        row.some((cell) => cell == null || typeof cell !== "object")
      )
    ) {
      return null;
    }
    directCells = cellMatrix.map((row) =>
      row.map((cell) => {
        // `style` 与 `format` 同级（`CellPropertiesLoadOptions.style`），单独取出来。
        const cellProperties = cell as Excel.CellProperties;
        const direct = readDirectCellFormat(cellProperties.format, cellProperties.style);
        const inherited = modifiedFields.filter(
          (field) => !directFormatHasField(direct, field)
        );
        inherited.forEach((field) => inheritedFields.add(field));
        return { direct, inherited: [...inherited] };
      })
    );
  }

  return {
    horizontalAlignment: String(format.horizontalAlignment),
    verticalAlignment: String(format.verticalAlignment),
    wrapText: format.wrapText,
    rowHeight: format.rowHeight,
    columnWidth: format.columnWidth,
    indentLevel: format.indentLevel,
    fontName: format.font.name,
    fontSize: format.font.size,
    fontItalic: format.font.italic,
    fontUnderline: String(format.font.underline),
    borders: borders.map(({ sideIndex, border }) => ({
      sideIndex,
      style: String(border.style),
      color: border.color,
      weight: String(border.weight)
    })),
    cells,
    directCells,
    inheritedFields: directCells ? [...inheritedFields].sort() : undefined,
    compatPath: directCells != null,
    // 行高三态：`number` = 直接设置；`"standard"` = `useStandardHeight` 为真（未直接设置）；
    // `null` = 两者都拿不到（宿主异常，无法确定）。**不在范围内就不采**（数组缺席），
    // 恢复侧据此判定"这次撤销不碰行高/列宽"。
    rowHeights: rowValues?.map((row) => {
      const f = row.format;
      if (f?.rowHeight != null) return f.rowHeight;
      if (f?.useStandardHeight === true) return "standard";
      return null;
    }),
    columnWidths: columnValues?.map((column) => {
      const f = column.format;
      if (f?.columnWidth != null) return f.columnWidth;
      if (f?.useStandardWidth === true) return "standard";
      return null;
    })
  };
}

/** 逐格显示格式不可用：`getDisplayedCellProperties` 返回的某格 `format` 缺失（宿主异常）。 */
class DisplayedFormatUnavailableError extends Error {
  constructor() {
    super("无法确定单元格的显示格式，不能可靠保存恢复依据");
    this.name = "DisplayedFormatUnavailableError";
  }
}

/**
 * 双栏对照用的"单元格此刻的样子"。
 *
 * values 与 formulas 分开存：展示看 values，判定写入成败看 formulas
 * （计算值会因重算时机漂移，不能拿它当验收依据）。
 * valuesCaptured=false 表示这次没采到显示值（区域过大），此时 values 是与
 * formulas 同形的 null 占位 —— 界面必须显示"显示值未采集"，不许拿公式顶替。
 */
type RangeChangeSnapshot = {
  values: CellValue[][];
  formulas: (string | null)[][];
  valuesCaptured: boolean;
};

/**
 * 独立验收的三种结论映射到双栏状态。
 *
 * 抽成纯函数是为了能被单测锁住：只要验收没走到 verified，
 * 双栏就不许落到 applied（core 的 canClaimWriteSuccess 只对 applied 放行）。
 */
export function changeReviewStatusFromVerification(
  verification: VerificationReport["status"]
): ChangeReviewStatus {
  if (verification === "failed") return "readback_mismatch";
  if (verification === "executed_unverified") return "executed_unverified";
  return "applied";
}

/**
 * 撤销快照里的 formulas 对非公式单元格存的其实是值（"abc"/5）。
 * 双栏的"原公式"只认以 = 开头的那些，其余一律 null —— 与 WPS 侧
 * `hasAmbiguousLogicalCell` 的判定口径一致。
 */
export function toChangeFormulas(matrix: CellValue[][]): (string | null)[][] {
  return matrix.map((row) =>
    row.map((value) =>
      typeof value === "string" && value.startsWith("=") ? value : null
    )
  );
}

function formulaMatrix(raw: unknown[][]): (string | null)[][] {
  return raw.map((row) =>
    row.map((value) => {
      if (value === null || value === undefined) return null;
      return typeof value === "string" ? value : String(value);
    })
  );
}

/**
 * 读一个区域当前的值与公式矩阵，用于"修改前｜修改后"对照。
 *
 * 只读本次已经（或即将）写入的授权区域，不做全簿扫描、不做轮询；
 * 区域超过 captureMaxCells 时放弃采集值，只保留公式。
 */
async function readRangeChangeSnapshot(
  context: Excel.RequestContext,
  sheetName: string,
  address: string
): Promise<RangeChangeSnapshot | null> {
  const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
  sheet.load("isNullObject");
  await context.sync();
  if (sheet.isNullObject) return null;
  const range = sheet.getRange(address);
  range.load("values,formulas");
  await context.sync();
  const formulas = formulaMatrix(range.formulas as unknown[][]);
  const cellCount = formulas.length * (formulas[0]?.length ?? 0);
  if (cellCount > CHANGE_REVIEW_CAPTURE_MAX_CELLS) {
    return {
      values: formulas.map((row) => row.map(() => null)),
      formulas,
      valuesCaptured: false
    };
  }
  return {
    values: (range.values as unknown[][]).map((row) => row.map(normalizeValue)),
    formulas,
    valuesCaptured: true
  };
}

/**
 * 汇总本次执行的"实际变更"：逐格拿（写前快照 + 写后回读）交给共用层比对。
 *
 * 关键点：after 必须是**这一次**回读出来的，不能用计划值顶替。
 * 读不到整块区域时该区域的 status 落到 readback_failed，绝不假装成功。
 */
async function collectChangeReviews(
  context: Excel.RequestContext,
  ranges: ExecutionUndoSnapshot["ranges"],
  beforeByAction: Map<number, RangeChangeSnapshot>
): Promise<RangeChangeReview[]> {
  const reviews: RangeChangeReview[] = [];
  if (!ranges.length) return reviews;
  for (const saved of ranges) {
    const before = beforeByAction.get(saved.actionIndex);
    if (!before) continue;
    let after: RangeChangeSnapshot | null = null;
    try {
      after = await readRangeChangeSnapshot(context, saved.sheet, saved.range);
    } catch {
      after = null;
    }
    if (!after) {
      reviews.push({
        sheet: saved.sheet,
        range: saved.range,
        phase: "actual",
        status: "readback_failed",
        rows: [],
        changedCount: 0,
        totalCount: 0,
        truncated: false
      });
      continue;
    }
    reviews.push(
      buildRangeChangeReview({
        sheet: saved.sheet,
        range: saved.range,
        phase: "actual",
        status: "applied",
        before,
        after,
        limit: CHANGE_REVIEW_LIMIT
      })
    );
  }
  return reviews;
}

async function captureUndoRange(
  context: Excel.RequestContext,
  action: ExcelAction,
  actionIndex: number
): Promise<{
  snapshot: ExecutionUndoSnapshot["ranges"][number] | null;
  /** 写前显示值；null 表示区域过大未采集。与 snapshot 分开，避免把矩阵塞进持久化快照。 */
  beforeValues: CellValue[][] | null;
}> {
  const address = recoverableActionRange(action);
  if (!address) return { snapshot: null, beforeValues: null };
  const sheet = context.workbook.worksheets.getItemOrNullObject(action.sheet);
  sheet.load("isNullObject");
  await context.sync();
  if (sheet.isNullObject) return { snapshot: null, beforeValues: null };
  const range = sheet.getRange(address);
  // values 只在双栏对照里用（展示"修改前"到底是什么值），
  // 与 formulas 同一批 load，不额外增加一次 sync。
  range.load(
    "values,formulas,numberFormat,format/fill/color,format/font/bold,format/font/color"
  );
  await context.sync();
  const cellCount = range.formulas.length * (range.formulas[0]?.length ?? 0);
  const rowCount = range.formulas.length;
  const columnCount = range.formulas[0]?.length ?? 0;
  // 本次动作**实际要改**的属性：捕获、来源记录与拒绝判据都按它取，不做动作名一刀切。
  const modifiedFields = excelFormatActionFields(action);
  // 本次动作**实际要恢复**的范围（`undoRestoreScope`）。它是"采什么""恢复写什么"
  // "冲突比什么"三处的唯一依据 —— 三处口径不同就会出现"没采却要写"或"写了却没比"。
  const scope = undoRestoreScope(action);
  const hasFormatScope = scope.cellFormat || scope.rowHeight || scope.columnWidth;
  let format: UndoRangeFormat | null = null;
  if (hasFormatScope) {
    try {
      format = await captureRangeFormat(
        context,
        range,
        cellCount,
        rowCount,
        columnCount,
        modifiedFields,
        scope
      );
    } catch (error) {
      if (error instanceof DisplayedFormatUnavailableError) {
        // 逐格显示格式不可确定 ⇒ 视为"格式不可完整捕获"，走与 null 相同的拒绝路径。
        format = null;
      } else if (error instanceof FormatRangeUnrecoverableError) {
        // 区域本身不可恢复（条件格式/合并区域/缺 API）：**指名道姓**地拒绝，
        // 让用户在写入前就知道是哪一类区域、缺哪个 API，而不是笼统报错。
        throw new Error(
          `第 ${actionIndex + 1} 步 ${action.type} 的目标区域无法在执行前可靠保存恢复依据，已拒绝执行以避免产生不可撤销的修改：${error.message}`
        );
      } else {
        throw error;
      }
    }
  }
  // 按动作**实际修改的属性**分别校验恢复能力（用户要求：不笼统"纯格式动作"一刀切）：
  //  - setFill/setFont/setBorders/setAlignment 改逐格格式（`scope.cellFormat`）⇒ 需要逐格矩阵：
  //    * 1.19 宿主看 `cells`（显示值）；
  //    * 1.9–1.17 宿主看 `directCells`（直接格式兼容路径）；
  //    两者都缺（区域过大 / 矩阵不匹配 / 显示值缺失 / 连 1.9 都没有）⇒ 执行前拒绝。
  //  - resizeRange 改行高/列宽 ⇒ 依赖 rowHeights/columnWidths 三态；含 null（无法确定）时
  //    执行前拒绝。
  //  - setNumberFormat 改 numberFormat ⇒ 依赖 range.numberFormat（独立捕获，总是可得）。
  // 值/公式类动作（writeValues 等）**根本不在格式范围内**，不采也不写，不在拒绝之列。
  if (scope.cellFormat) {
    const hasCellMatrix = format?.cells != null || format?.directCells != null;
    if (!hasCellMatrix) {
      throw new Error(
        `第 ${actionIndex + 1} 步 ${action.type} 的目标区域逐格格式无法在执行前可靠保存恢复依据，已拒绝执行以避免产生不可撤销的修改`
      );
    }
  }
  if (scope.rowHeight || scope.columnWidth) {
    // 行高/列宽的可靠恢复约束（**只校验本动作真的要改的那一项**）：
    //  - 含 null（无法确定该行/列状态）⇒ 拒绝；
    //  - 混合了"直接设置值"与"标准行高/列宽"⇒ `setRowProperties`/`setColumnProperties`
    //    只能逐行写具体值、`useStandardHeight`/`useStandardWidth` 只能整区域统一重置，
    //    无法逐行恢复"标准"行 ⇒ 拒绝；
    //  - 全部直接设置 或 全部标准 ⇒ 可可靠恢复（分别走逐行写回 / 整区域重置）。
    const heights = scope.rowHeight ? format?.rowHeights : undefined;
    const widths = scope.columnWidth ? format?.columnWidths : undefined;
    const heightsMixed =
      heights && heights.some((h) => typeof h === "number") && heights.some((h) => h === "standard");
    const widthsMixed =
      widths && widths.some((w) => typeof w === "number") && widths.some((w) => w === "standard");
    const heightBad =
      scope.rowHeight &&
      (!heights || heights.some((h) => h == null) || Boolean(heightsMixed));
    const widthBad =
      scope.columnWidth &&
      (!widths || widths.some((w) => w == null) || Boolean(widthsMixed));
    if (heightBad || widthBad) {
      throw new Error(
        `第 ${actionIndex + 1} 步 resizeRange 的目标区域行高/列宽无法在执行前可靠确定，已拒绝执行以避免产生不可撤销的修改`
      );
    }
  }
  const beforeValues =
    cellCount > CHANGE_REVIEW_CAPTURE_MAX_CELLS
      ? null
      : (range.values as unknown[][]).map((row) => row.map(normalizeValue));
  return {
    snapshot: {
      actionIndex,
      sheet: action.sheet,
      range: address,
      formulas: range.formulas.map((row) => row.map(normalizeValue)),
      numberFormat: range.numberFormat,
      fillColor: range.format.fill.color,
      fontBold: range.format.font.bold,
      fontColor: range.format.font.color,
      scope,
      formatFields: [...modifiedFields],
      format
    },
    beforeValues
  };
}

interface DeletedWorksheetCapture {
  snapshot: UndoWorksheetSnapshot | null;
  skipped: boolean;
}

async function captureDeletedWorksheet(
  context: Excel.RequestContext,
  sheetName: string,
  originalPosition?: number
): Promise<DeletedWorksheetCapture> {
  const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
  sheet.load("isNullObject,position");
  await context.sync();
  if (sheet.isNullObject) return { snapshot: null, skipped: false };
  const used = sheet.getUsedRangeOrNullObject(true);
  used.load(
    "address,rowCount,columnCount,rowIndex,columnIndex,isNullObject"
  );
  await context.sync();
  if (used.isNullObject) return { snapshot: null, skipped: false };
  if (used.rowCount * used.columnCount > UNDO_WORKSHEET_MAX_CELLS) {
    return { snapshot: null, skipped: true };
  }
  const range = sheet.getRangeByIndexes(
    used.rowIndex,
    used.columnIndex,
    used.rowCount,
    used.columnCount
  );
  range.load("formulas,numberFormat");
  await context.sync();
  return {
    snapshot: {
      sheet: sheetName,
      position: originalPosition ?? sheet.position,
      usedRange: used.address.split("!").pop() ?? used.address,
      formulas: range.formulas.map((row) => row.map(normalizeValue)),
      numberFormat: range.numberFormat
    },
    skipped: false
  };
}

function localStateFingerprint(value: unknown): string {
  const serialized = JSON.stringify(value);
  return `fnv1a32:${updateFingerprint(0x811c9dc5, serialized)
    .toString(16)
    .padStart(8, "0")}`;
}

function loadedRangePostFingerprint(range: Excel.Range): string {
  return localStateFingerprint({
    formulas: range.formulas.map((row) => row.map(normalizeValue)),
    numberFormat: range.numberFormat,
    fillColor: range.format.fill.color,
    fontBold: range.format.font.bold,
    fontColor: range.format.font.color
  });
}

/**
 * 撤销守卫要读的全部字段。**一次 load 取回两代判据需要的所有属性** ——
 * 分两次读会让每个范围多一次 `context.sync()`，而守卫在撤销与执行后各跑一遍。
 */
const UNDO_GUARD_RANGE_LOAD =
  "formulas,numberFormat,format/fill/color,format/font/bold,format/font/color," +
  "format/horizontalAlignment,format/verticalAlignment,format/wrapText," +
  "format/rowHeight,format/columnWidth,format/indentLevel," +
  "format/font/name,format/font/size,format/font/italic,format/font/underline";

/**
 * **第二代**执行后判据：在 `loadedRangePostFingerprint` 的六项之外，把
 * 对齐 / 自动换行 / 行高列宽 / 缩进 / 字体名・字号・斜体・下划线 / 六条边框
 * 一起摘要进来。
 *
 * 载荷是**严格超集**：第一代那六项也在里面。所以撤销前只要这一项存在，
 * 就只用比它 —— 比得过就意味着第一代也对得上，不用再比一次。
 *
 * 存在的原因很具体：`setBorders` / `setAlignment` / `resizeRange` 执行完之后，
 * 用户手动把某个边框改细、把对齐改回左对齐，第一代判据会认为"现场没变"，
 * 撤销就照常覆盖了用户后来的编辑。
 */
function extendedFormatDigest(
  range: Excel.Range,
  borders: ReadonlyArray<{ sideIndex: Excel.BorderIndex; border: Excel.RangeBorder }>
): string {
  return localStateFingerprint({
    formulas: range.formulas.map((row) => row.map(normalizeValue)),
    numberFormat: range.numberFormat,
    fillColor: range.format.fill.color,
    fontBold: range.format.font.bold,
    fontColor: range.format.font.color,
    horizontalAlignment: String(range.format.horizontalAlignment),
    verticalAlignment: String(range.format.verticalAlignment),
    wrapText: range.format.wrapText,
    rowHeight: range.format.rowHeight,
    columnWidth: range.format.columnWidth,
    indentLevel: range.format.indentLevel,
    fontName: range.format.font.name,
    fontSize: range.format.font.size,
    fontItalic: range.format.font.italic,
    fontUnderline: String(range.format.font.underline),
    borders: borders.map(({ sideIndex, border }) => [
      String(sideIndex),
      String(border.style),
      normalizedColor(border.color),
      String(border.weight)
    ])
  });
}

/**
 * 第三代判据的逐字段摘要：**只含本次撤销真正会写回的字段**（`scope` 为真的那几项）。
 *
 * 与 `extendedFormatDigest` 的区别是**覆盖范围跟着动作走**而非固定的全属性集合。
 * 只改行高的动作不会因为用户改了字色而拒绝撤销（字色根本不会被写回，保留即可）；
 * 但只要用户在**会被写回**的字段上动过手，摘要立刻就变，撤销在写入前报冲突。
 */
type UndoScopeFieldDigests = Partial<Record<keyof UndoRestoreScope, string>>;

/**
 * 读一次撤销守卫现场，同时给出三代判据。
 *
 * 边框必须逐条 `getItem().load()` —— `range.load()` 的路径串表达不了集合元素，
 * 所以这里先 load 范围属性、再把六条边框挂上，最后**一次** `context.sync()`。
 *
 * 第三代判据（`scopeFieldDigests`/`scopeDigest`）**按 `scope` 决定读什么**：
 *  - `cellFormat` ⇒ 逐格读一遍（显示值或直接格式，与采集同源），摘要按**逐格矩阵**算，
 *    这样"区域里某一格被改过"也能被发现（区域级单值在混合状态下会钝化）；
 *  - `rowHeight`/`columnWidth` ⇒ 逐行/逐列读，摘要按**逐行/逐列数组**算；
 *  - `formulas`/`numberFormat` ⇒ 直接摘要矩阵。
 * 前两代判据照旧算，供**旧恢复点**走原判据。
 */
async function readUndoGuardState(
  context: Excel.RequestContext,
  range: Excel.Range,
  scope: UndoRestoreScope,
  formatFields?: readonly string[]
): Promise<{
  postFingerprint: string;
  extendedFormatDigest: string;
  scopeFieldDigests: UndoScopeFieldDigests;
  /** 逐格式属性摘要（键为 `ExcelFormatField`）：冲突文案指名具体属性用。 */
  cellFormatFieldDigests: Record<string, string>;
  scopeDigest: string;
  /** 兼容路径下顺带读回的**当前逐格直接格式**（撤销时用来保住非本动作属性的后续编辑）。 */
  observedDirectCells: DirectCellFormatValues[][] | null;
}> {
  range.load(UNDO_GUARD_RANGE_LOAD);
  const borders = undoBorderSides().map((sideIndex) => {
    const border = range.format.borders.getItem(sideIndex);
    border.load("style,color,weight");
    return { sideIndex, border };
  });
  // 逐格格式摘要：与采集同源（显示值路径 / 直接格式兼容路径二选一）。宿主连 1.9 都没有时
  // 读不了 ⇒ 该字段不进摘要（它的恢复本来也不可能发生）。
  const capability = currentExcelFormatCapability();
  const displayed = capability.path === "displayed";
  const cellProperties =
    scope.cellFormat && capability.path !== "unavailable"
      ? displayed
        ? range.getDisplayedCellProperties({ format: DISPLAYED_FORMAT_LOAD_OPTIONS })
        : range.getCellProperties(DIRECT_CELL_PROPERTIES_LOAD)
      : null;
  const rowProperties = scope.rowHeight
    ? range.getRowProperties({ format: { rowHeight: true, useStandardHeight: true } })
    : null;
  const columnProperties = scope.columnWidth
    ? range.getColumnProperties({ format: { columnWidth: true, useStandardWidth: true } })
    : null;
  await context.sync();

  const fieldDigests: UndoScopeFieldDigests = {};
  // 逐**格式属性**摘要：冲突文案要指名具体属性（"加粗"），而不是笼统的"单元格格式"。
  // 它与聚合摘要读**同一份**中间表示，因此两者不可能互相矛盾。
  const cellFormatFieldDigests: Record<string, string> = {};
  if (scope.formulas) {
    fieldDigests.formulas = localStateFingerprint(
      range.formulas.map((row) => row.map(normalizeValue))
    );
  }
  if (scope.numberFormat) {
    fieldDigests.numberFormat = localStateFingerprint(range.numberFormat);
  }
  let observedDirectCells: DirectCellFormatValues[][] | null = null;
  if (cellProperties) {
    // **摘要只覆盖本动作真正会写回的属性**（`formatFields`；旧记录没有这份信息就按全集）。
    // 于是"用户改了某个本动作没碰过的属性"不会让摘要变化 ⇒ 撤销照常执行且**保留**该编辑；
    // 只有用户改了**会被写回**的属性，摘要才会变 ⇒ 在写入前报冲突。
    const scopedFields = formatFields ?? ALL_DIRECT_FORMAT_FIELDS;
    const rows = cellProperties.value;
    // 先把每格归一成一份中间表示（显示值路径 / 直接格式路径二选一），再按键集合投影。
    // 聚合摘要与逐属性摘要都从这份中间表示派生 —— 同一个源，不会互相矛盾。
    let sources: (Record<string, unknown> | null)[][];
    if (displayed) {
      sources = rows.map((row) =>
        row.map((cell) => {
          const projected = displayedCellFormatOf((cell as Excel.CellProperties).format);
          return projected ? (projected as unknown as Record<string, unknown>) : null;
        })
      );
    } else {
      observedDirectCells = rows.map((row) =>
        row.map((cell) => {
          const properties = cell as Excel.CellProperties;
          return readDirectCellFormat(properties.format, properties.style);
        })
      );
      sources = observedDirectCells.map((row) =>
        row.map((direct) => direct as unknown as Record<string, unknown>)
      );
    }
    const projectAll = (keys: ReadonlySet<string>) =>
      sources.map((row) =>
        row.map((source) => (source ? projectKeys(source, keys) : null))
      );
    const keysOfFields = (fields: readonly string[]) =>
      asKeySet(displayed ? displayedKeysOfFields(fields) : directKeysOfFields(fields));
    fieldDigests.cellFormat = localStateFingerprint(projectAll(keysOfFields(scopedFields)));
    for (const field of scopedFields) {
      cellFormatFieldDigests[field] = localStateFingerprint(projectAll(keysOfFields([field])));
    }
  }
  if (rowProperties) {
    fieldDigests.rowHeight = localStateFingerprint(
      rowProperties.value.map((row) => row.format?.rowHeight ?? null)
    );
  }
  if (columnProperties) {
    fieldDigests.columnWidth = localStateFingerprint(
      columnProperties.value.map((column) => column.format?.columnWidth ?? null)
    );
  }
  return {
    postFingerprint: loadedRangePostFingerprint(range),
    extendedFormatDigest: extendedFormatDigest(range, borders),
    scopeFieldDigests: fieldDigests,
    cellFormatFieldDigests,
    // 字段名集合本身也进摘要：范围收窄/放宽都不是"现场没变"。
    scopeDigest: localStateFingerprint(fieldDigests),
    observedDirectCells
  };
}

/**
 * 出网身份字符串。归一规则已下沉到 core 的 `officeIdentityFromUrl`,
 * 这里不再自行拼接 `url:` 前缀 —— 两份实现一旦漂移,已落盘账本里的旧记录
 * 就会因身份对不上而失去续跑与撤销入口。
 */
export function currentOfficeWorkbookIdentity(): string | undefined {
  const url = currentOfficeDocumentUrl();
  if (url === undefined) return undefined;
  return officeIdentityFromUrl(url) ?? undefined;
}

/**
 * 桥/读路径的工作簿身份（**唯一折叠点**：规划请求、桥绑定、执行前复核、续跑证据
 * 都读这一份）。
 *
 * 已保存 → `url:<...>`，与写入账本、恢复检查点同一口径。
 *
 * **未保存 → 会话内对象身份**（`unsaved:<序号>-<随机>`），按宿主的工作簿对象分配：
 * 同一个对象每次拿到同一个令牌，换一个未保存簿就是另一个令牌。这样做是因为
 * Office.js 在文件未保存时 `document.url` 为空，退回 `name:<名字>` 会让两个都叫
 * 「工作簿1」的未保存簿在桥的匹配（`hostId + workbookName + workbookIdentity`）
 * 上**三项全等** —— 授权是给 A 的，读表却可能落到 B。
 *
 * 作用域**只在本次加载内**：窗格重开 = 新的 JS 运行环境 = 新的文档对象 = 新令牌。
 * 因此它只能证明"同一次加载里是同一个对象"，不能证明重开后还是同一本簿；服务端
 * 拿它跟检查点里记录的身份比对时会判 `PLANNING_CONTEXT_TARGET_CHANGED`，界面上表现
 * 为"不能继续旧任务，请按当前工作簿重新发起需求"。这是宿主边界，不是缺陷。
 *
 * 刻意**不**用这些东西当身份：
 * - 名字：两个未保存簿可以同名（这正是本函数要解决的撞车）；
 * - 表结构或数据指纹：编辑一下就变，"同一本簿"会被判成"换了一本"；
 * - `document.settings` 里的随机 GUID：那会把"读一次表"变成"改用户的文件"，
 *   也等于要求用户先保存。宁可只保证会话内身份。
 */
const unsavedWorkbookIdentities = new WeakMap<object, string>();
let unsavedWorkbookCounter = 0;

export function currentOfficeBridgeWorkbookIdentity(): string | undefined {
  const stable = currentOfficeWorkbookIdentity();
  if (stable !== undefined) return stable;
  if (typeof Office === "undefined") return undefined;
  const documentRef = (Office as { context?: { document?: unknown } }).context
    ?.document;
  if (!documentRef || typeof documentRef !== "object") return undefined;
  const key = documentRef as object;
  const existing = unsavedWorkbookIdentities.get(key);
  if (existing !== undefined) return existing;
  unsavedWorkbookCounter += 1;
  const token = `unsaved:${unsavedWorkbookCounter}-${Math.random()
    .toString(16)
    .slice(2, 10)}`;
  unsavedWorkbookIdentities.set(key, token);
  return token;
}

/** 共享层内部表示。折叠回字符串只在出网时发生（方案 §2.3 第 1 条）。 */
export function currentOfficeWorkbookLocator(): WorkbookLocator | undefined {
  const identity = currentOfficeWorkbookIdentity();
  if (identity === undefined) return undefined;
  return locatorFromOfficeIdentity(identity) ?? undefined;
}

function currentOfficeDocumentUrl(): string | undefined {
  if (typeof Office === "undefined") return undefined;
  const url = (Office as any).context?.document?.url;
  if (typeof url !== "string" || !url.trim()) return undefined;
  return url;
}

const OFFICE_RECOVERY_ACTION_TYPES = new Set<ExcelAction["type"]>([
  "writeValues",
  "writeFormulas",
  "setFill",
  "setFont",
  "setNumberFormat",
  // 这三个只改格式，`recoverableActionRange` 与 `captureRangeFormat` 都已覆盖它们；
  // 之前漏在这里，表现是"能执行、能撤销，但重启后无法核对在途动作"。
  // `autofit` 仍不在名单里：它的结果由宿主按内容算，没有确定性期望值可比对。
  "setBorders",
  "setAlignment",
  "resizeRange"
]);

export interface OfficeRecoverySupport {
  supported: boolean;
  reason?: string;
  /**
   * 公式计划登记前必须钉住的来源表（= 公式文本真正读的那几张表）。
   * 登记方按它现场取一次值指纹写进 `plan.sourceDependencyFingerprint`；
   * 续跑时按同一个集合重新读、重新比。空/缺省 = 这份计划没有需要证据的公式。
   */
  sourceEvidenceSheets?: string[];
}

export function officeRecoverySupport(plan: AnalysisPlan): OfficeRecoverySupport {
  /*
   * 读写重叠只看**实际来源依赖**，不看指纹覆盖范围。
   *
   * 指纹覆盖范围端上取的是本次授权范围（含写入目标），拿它判重叠会把每一个
   * 正常计划都判成"会修改来源表"，登记断点恢复因此永远不可达。来源依赖由
   * `sourceDependencySheets` 承载，只来自本端可信记账（确定性本地读取／工具
   * 执行记录／已保存工具模板声明）；缺失时**退回指纹范围** —— 也就是既有的
   * 保守拒绝条件原样保留：依赖不确定就说不确定，不猜，也绝不用"授权范围减去
   * 写入目标"伪造来源集合。
   */
  const declaredDependency = plan.sourceDependencySheets?.length
    ? plan.sourceDependencySheets
    : null;
  const sourceSheets = new Set(
    (declaredDependency ?? plan.sourceFingerprintSheets ?? []).map((name) =>
      name.toLocaleLowerCase()
    )
  );
  for (const [index, action] of plan.actions.entries()) {
    if (!OFFICE_RECOVERY_ACTION_TYPES.has(action.type)) {
      return {
        supported: false,
        reason: `第 ${index + 1} 步 ${action.type} 暂不支持跨重启自动续跑`
      };
    }
    if (action.type === "writeFormulas" && action.formulaR1C1) {
      return {
        supported: false,
        reason: `第 ${index + 1} 步 R1C1 铺满公式暂不支持跨重启效果核对`
      };
    }
    if (sourceSheets.has(action.sheet.toLocaleLowerCase())) {
      return {
        supported: false,
        reason: declaredDependency
          ? `第 ${index + 1} 步会修改来源工作表，暂不能区分任务自身修改与恢复期间的外部修改`
          : `第 ${index + 1} 步会修改本次覆盖的来源范围，而这次没有"本机实际读过哪些表"的可信记账，无法确定来源与目标是否重叠`
      };
    }
    if (
      action.type === "setFont" &&
      action.bold == null &&
      action.italic == null &&
      action.name == null &&
      action.size == null &&
      action.color == null
    ) {
      return {
        supported: false,
        reason: `第 ${index + 1} 步没有可核对的字体变更`
      };
    }
  }
  /*
   * **公式依赖面**：`writeFormulas` 写的是公式文本，宿主执行时自己求值，所以它不会
   * 落进上面的"修改来源表"判断（公式依赖并不等于 `sourceDependencySheets` 里那些
   * 读过的表）。于是恢复要回答的是另一个问题：公式读的那几张表，我们有没有证据？
   *
   * 三种情况一律不登记精确恢复：
   * - 依赖面本身就不可分离（同表引用／动态／外部工作簿／名称／引用表=写入目标）——
   *   `formulaSourceEvidenceRequirement` 给理由，这里只转发，不另写一套判断；
   * - 依赖表不在本机可信记账里（`sourceDependencySheets`）：没有"这台机器真读过"的
   *   凭据，也就没了"当时是什么值"的基线；
   * - 记账为缺失（`undefined`）而不是空数组：空 = 没读过任何表，缺失 = 不知道读过谁。
   *
   * 通过时把**这个集合**交给登记方去钉指纹（口径写在 `sourceDependencyFingerprint`）。
   */
  const formulaRequirement = formulaSourceEvidenceRequirement(plan);
  if (!formulaRequirement.ok) {
    return {
      supported: false,
      reason: `公式精确恢复：${formulaRequirement.reason}`
    };
  }
  if (formulaRequirement.sheets.length > 0) {
    if (!declaredDependency) {
      return {
        supported: false,
        reason:
          "公式引用的来源表没有「本机实际读过哪些表」的可信记账，无法登记来源证据，不能精确恢复"
      };
    }
    const recorded = new Set(
      declaredDependency.map((name) => name.toLocaleLowerCase())
    );
    const uncovered = formulaRequirement.sheets.filter(
      (name) => !recorded.has(name.toLocaleLowerCase())
    );
    if (uncovered.length > 0) {
      return {
        supported: false,
        reason: `公式引用了「${uncovered.join("、")}」，但这几张表没有本机读取记录，无法为它们登记来源证据`
      };
    }
    return { supported: true, sourceEvidenceSheets: [...formulaRequirement.sheets] };
  }
  return plan.actions.length > 0
    ? { supported: true }
    : { supported: false, reason: "空计划不需要断点续跑" };
}

export interface OfficeRangeObservedState {
  formulas: CellValue[][];
  numberFormat: string[][];
  fillColor: string;
  fontBold: boolean;
  fontColor: string;
  /**
   * 扩展格式现场。**全部可选**：`observedRangeFingerprint` 的载荷刻意保持不变
   * （改它会让已经落盘的在途检查点指纹对不上），所以这些字段只服务于
   * `classifyOfficeActionState` 对格式动作的判定，不参与指纹。
   */
  fontItalic?: boolean;
  fontName?: string;
  fontSize?: number;
  horizontal?: string;
  vertical?: string;
  wrapText?: boolean;
  rowHeight?: number;
  columnWidth?: number;
  borders?: Partial<
    Record<BorderSide, { style: string; color: string; weight: string }>
  >;
}

export type OfficeActionReconciliationOutcome =
  | "applied"
  | "not_applied"
  | "unknown";

function sameMatrix(left: unknown[][], right: unknown[][]): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

function sameColor(left: string, right: string): boolean {
  return left.trim().toLocaleUpperCase() === right.trim().toLocaleUpperCase();
}

/**
 * 数值比较的"现场缺值 ⇒ 不通过"版本。
 *
 * `closeEnough` 收的是两个 `number`，而格式动作的现场值是按需读取的可选字段：
 * 没读到就是 `undefined`。把 `undefined` 直接喂给 `closeEnough` 在类型上过不去，
 * 用 `?? 0` 之类去凑更糟 —— 那会把"没读到"当成"读到 0"，判成通过。
 * 所以这里显式表达"没读到就不算匹配"。
 */
function closeEnoughObserved(left: number | undefined, right: number): boolean {
  return typeof left === "number" && closeEnough(left, right);
}

function observedRangeFingerprint(state: OfficeRangeObservedState): string {
  return localStateFingerprint({
    formulas: state.formulas.map((row) => row.map(normalizeValue)),
    numberFormat: state.numberFormat,
    fillColor: state.fillColor,
    fontBold: state.fontBold,
    fontColor: state.fontColor
  });
}

function undoRangeBeforeFingerprint(saved: ExecutionUndoSnapshot["ranges"][number]): string {
  return localStateFingerprint({
    formulas: saved.formulas.map((row) => row.map(normalizeValue)),
    numberFormat: saved.numberFormat,
    fillColor: saved.fillColor,
    fontBold: saved.fontBold,
    fontColor: saved.fontColor
  });
}

export function classifyOfficeActionState(
  action: ExcelAction,
  before: ExecutionUndoSnapshot["ranges"][number],
  observed: OfficeRangeObservedState
): OfficeActionReconciliationOutcome {
  let matchesExpected = false;
  switch (action.type) {
    case "writeValues":
      matchesExpected = sameMatrix(
        observed.formulas,
        action.values.map((row) => row.map(normalizeValue))
      );
      break;
    case "writeFormulas":
      matchesExpected =
        !action.formulaR1C1 && sameMatrix(observed.formulas, action.formulas);
      break;
    case "setFill":
      matchesExpected = sameColor(observed.fillColor, action.color);
      break;
    case "setFont":
      // 五个属性各自独立判：`!= null` 表示"这次要改它"。
      matchesExpected =
        (action.bold == null || observed.fontBold === action.bold) &&
        (action.italic == null || observed.fontItalic === action.italic) &&
        (action.name == null || observed.fontName === action.name) &&
        (action.size == null || closeEnoughObserved(observed.fontSize, action.size)) &&
        (action.color == null || sameColor(observed.fontColor, action.color));
      break;
    case "setBorders": {
      const expected = normalizeBorderAction(action);
      matchesExpected = expected.sides.every((side) => {
        const seen = observed.borders?.[side];
        if (!seen) return false;
        return (
          normalizedEnum(seen.style) === normalizedEnum(expected.style) &&
          sameColor(seen.color, expected.color) &&
          normalizedEnum(seen.weight) === normalizedEnum(expected.weight)
        );
      });
      break;
    }
    case "setAlignment":
      matchesExpected =
        (action.horizontal == null ||
          normalizedEnum(observed.horizontal) ===
            normalizedEnum(action.horizontal)) &&
        (action.vertical == null ||
          normalizedEnum(observed.vertical) === normalizedEnum(action.vertical)) &&
        (action.wrapText == null || observed.wrapText === action.wrapText);
      break;
    case "resizeRange":
      matchesExpected =
        (action.rowHeight == null ||
          closeEnoughObserved(observed.rowHeight, action.rowHeight)) &&
        (action.columnWidth == null ||
          closeEnoughObserved(observed.columnWidth, action.columnWidth));
      break;
    case "setNumberFormat":
      matchesExpected = observed.numberFormat.every((row) =>
        row.every((value) => value === action.formatCode)
      );
      break;
    default:
      return "unknown";
  }
  if (matchesExpected) return "applied";
  if (observedRangeFingerprint(observed) === undoRangeBeforeFingerprint(before)) {
    return "not_applied";
  }
  return "unknown";
}

export function appendOfficeRangeCheckpoint(input: {
  plan: AnalysisPlan;
  operationId: string;
  workbookIdentity: string;
  previous?: ExecutionUndoSnapshot | null;
  before: ExecutionUndoSnapshot["ranges"][number];
  actionIndex: number;
  postFingerprint: string;
}): ExecutionUndoSnapshot {
  const savedRange = { ...input.before, postFingerprint: input.postFingerprint };
  const previousRanges = (input.previous?.ranges ?? []).filter(
    (item) => item.actionIndex !== input.actionIndex
  );
  const completedActionCount = input.actionIndex + 1;
  return {
    planId: input.plan.id,
    executionId: input.operationId,
    workbookIdentity: input.workbookIdentity,
    capturedAt: new Date().toISOString(),
    coverage:
      completedActionCount === input.plan.actions.length ? "full" : "partial",
    completedActionCount,
    uncertainActionIndexes: [],
    postStateCoverage: "complete",
    ranges: [...previousRanges, savedRange],
    createdWorksheets: [],
    createdWorksheetStates: [],
    deletedWorksheets: [],
    skippedWorksheets: []
  };
}

export async function reconcileOfficeRangeAction(input: {
  plan: AnalysisPlan;
  operationId: string;
  workbookIdentity: string;
  actionIndex: number;
  before: ExecutionUndoSnapshot["ranges"][number];
  previous?: ExecutionUndoSnapshot | null;
}): Promise<{
  outcome: OfficeActionReconciliationOutcome;
  checkpoint?: ExecutionUndoSnapshot;
}> {
  const currentIdentity = currentOfficeWorkbookIdentity();
  if (!currentIdentity || currentIdentity !== input.workbookIdentity) {
    throw new Error("当前工作簿与断点记录不一致，已停止自动续跑。");
  }
  const action = input.plan.actions[input.actionIndex];
  if (!action || !officeRecoverySupport({ ...input.plan, actions: [action] }).supported) {
    throw new Error("在途动作不在 Office 自动恢复白名单内。");
  }
  const address = recoverableActionRange(action);
  if (!address || action.sheet !== input.before.sheet || address !== input.before.range) {
    throw new Error("在途动作的目标范围与写前快照不一致。");
  }
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(action.sheet);
    sheet.load("isNullObject");
    await context.sync();
    if (sheet.isNullObject) return { outcome: "unknown" as const };
    const range = sheet.getRange(address);
    range.load(
      "formulas,numberFormat,format/fill/color,format/font/bold,format/font/color," +
        "format/font/name,format/font/size,format/font/italic," +
        "format/horizontalAlignment,format/verticalAlignment,format/wrapText," +
        "format/rowHeight,format/columnWidth"
    );
    const borderObjects = undoBorderSides().map((sideIndex) => {
      const border = range.format.borders.getItem(sideIndex);
      border.load("style,color,weight");
      return { sideIndex, border };
    });
    await context.sync();
    const observed: OfficeRangeObservedState = {
      formulas: range.formulas.map((row) => row.map(normalizeValue)),
      numberFormat: range.numberFormat,
      fillColor: range.format.fill.color,
      fontBold: range.format.font.bold,
      fontColor: range.format.font.color,
      fontItalic: range.format.font.italic,
      fontName: String(range.format.font.name),
      fontSize: range.format.font.size,
      horizontal: String(range.format.horizontalAlignment),
      vertical: String(range.format.verticalAlignment),
      wrapText: range.format.wrapText,
      rowHeight: range.format.rowHeight,
      columnWidth: range.format.columnWidth,
      borders: Object.fromEntries(
        borderObjects.map(({ sideIndex, border }) => [
          borderSideName(sideIndex),
          {
            style: String(border.style),
            color: border.color,
            weight: String(border.weight)
          }
        ])
      )
    };
    const outcome = classifyOfficeActionState(action, input.before, observed);
    if (outcome !== "applied") return { outcome };
    return {
      outcome,
      checkpoint: appendOfficeRangeCheckpoint({
        ...input,
        postFingerprint: observedRangeFingerprint(observed)
      })
    };
  });
}

async function readCreatedWorksheetState(
  context: Excel.RequestContext,
  sheetName: string
): Promise<UndoCreatedWorksheetState | null> {
  const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
  sheet.load("isNullObject");
  await context.sync();
  if (sheet.isNullObject) return null;
  const used = sheet.getUsedRangeOrNullObject(true);
  used.load("address,rowCount,columnCount,rowIndex,columnIndex,isNullObject");
  await context.sync();
  if (used.isNullObject) {
    return { sheet: sheetName, postFingerprint: localStateFingerprint({ empty: true }) };
  }
  if (used.rowCount * used.columnCount > UNDO_WORKSHEET_MAX_CELLS) {
    return { sheet: sheetName };
  }
  const range = sheet.getRangeByIndexes(
    used.rowIndex,
    used.columnIndex,
    used.rowCount,
    used.columnCount
  );
  range.load(
    "formulas,numberFormat,format/fill/color,format/font/bold,format/font/color"
  );
  await context.sync();
  return {
    sheet: sheetName,
    postFingerprint: localStateFingerprint({
      address: used.address.split("!").pop() ?? used.address,
      formulas: range.formulas.map((row) => row.map(normalizeValue)),
      numberFormat: range.numberFormat,
      fillColor: range.format.fill.color,
      fontBold: range.format.font.bold,
      fontColor: range.format.font.color
    })
  };
}

async function captureUndoPostState(
  context: Excel.RequestContext,
  ranges: ExecutionUndoSnapshot["ranges"],
  createdWorksheetNames: string[],
  deletedWorksheets: UndoWorksheetSnapshot[]
): Promise<{
  complete: boolean;
  createdWorksheetStates: UndoCreatedWorksheetState[];
  appliedDeletedWorksheets: UndoWorksheetSnapshot[];
}> {
  let complete = true;
  for (const saved of ranges) {
    const sheet = context.workbook.worksheets.getItemOrNullObject(saved.sheet);
    sheet.load("isNullObject");
    await context.sync();
    if (sheet.isNullObject) {
      complete = false;
      continue;
    }
    const range = sheet.getRange(saved.range);
    const observed = await readUndoGuardState(
      context,
      range,
      scopeOf(saved),
      saved.formatFields
    );
    saved.postFingerprint = observed.postFingerprint;
    saved.extendedFormatDigest = observed.extendedFormatDigest;
    // 第三代判据（按恢复范围、逐格/逐行）—— 新记录的主判据。
    saved.scopeFieldDigests = observed.scopeFieldDigests;
    // 主判据的按属性再拆分：只用于冲突文案指名具体属性，不参与判定。
    saved.cellFormatFieldDigests = observed.cellFormatFieldDigests;
    saved.scopeDigest = observed.scopeDigest;
  }

  const createdWorksheetStates: UndoCreatedWorksheetState[] = [];
  for (const sheetName of [...new Set(createdWorksheetNames)]) {
    const state = await readCreatedWorksheetState(context, sheetName);
    if (!state) continue;
    if (!state.postFingerprint) complete = false;
    createdWorksheetStates.push(state);
  }

  const appliedDeletedWorksheets: UndoWorksheetSnapshot[] = [];
  for (const saved of deletedWorksheets) {
    const sheet = context.workbook.worksheets.getItemOrNullObject(saved.sheet);
    sheet.load("isNullObject");
    await context.sync();
    if (sheet.isNullObject) appliedDeletedWorksheets.push(saved);
  }
  return { complete, createdWorksheetStates, appliedDeletedWorksheets };
}

async function buildExecutionUndoSnapshot(
  context: Excel.RequestContext,
  input: {
    planId: string;
    executionId: string;
    coverage: "full" | "partial";
    completedActionCount: number;
    uncertainActionIndexes: number[];
    ranges: ExecutionUndoSnapshot["ranges"];
    createdWorksheets: string[];
    deletedWorksheets: UndoWorksheetSnapshot[];
    skippedWorksheets: string[];
  }
): Promise<ExecutionUndoSnapshot | null> {
  if (
    input.ranges.length === 0 &&
    input.createdWorksheets.length === 0 &&
    input.deletedWorksheets.length === 0
  ) {
    return null;
  }
  let postState: Awaited<ReturnType<typeof captureUndoPostState>>;
  try {
    postState = await captureUndoPostState(
      context,
      input.ranges,
      input.createdWorksheets,
      input.deletedWorksheets
    );
  } catch {
    postState = {
      complete: false,
      createdWorksheetStates: [],
      appliedDeletedWorksheets: input.deletedWorksheets
    };
  }
  const createdWorksheets = postState.createdWorksheetStates.map(
    (state) => state.sheet
  );
  if (
    input.ranges.length === 0 &&
    createdWorksheets.length === 0 &&
    postState.appliedDeletedWorksheets.length === 0
  ) {
    return null;
  }
  return {
    planId: input.planId,
    executionId: input.executionId,
    workbookIdentity: currentOfficeWorkbookIdentity(),
    capturedAt: new Date().toISOString(),
    coverage: input.coverage,
    completedActionCount: input.completedActionCount,
    uncertainActionIndexes: input.uncertainActionIndexes,
    postStateCoverage: postState.complete ? "complete" : "partial",
    ranges: input.ranges,
    createdWorksheets,
    createdWorksheetStates: postState.createdWorksheetStates,
    deletedWorksheets: postState.appliedDeletedWorksheets,
    skippedWorksheets: [...new Set(input.skippedWorksheets)]
  };
}

/** 恢复范围字段的中文名，仅用于冲突报错**指名到底哪一项对不上**。 */
const UNDO_SCOPE_FIELD_LABELS: Record<keyof UndoRestoreScope, string> = {
  formulas: "值/公式",
  numberFormat: "数字格式",
  cellFormat: "单元格格式",
  rowHeight: "行高",
  columnWidth: "列宽"
};

/** `ExcelFormatField` → 冲突文案里的中文属性名。 */
const FORMAT_FIELD_LABELS: Record<string, string> = {
  fill: "填充",
  fontName: "字体",
  fontSize: "字号",
  fontBold: "加粗",
  fontItalic: "斜体",
  fontColor: "字色",
  fontUnderline: "下划线",
  horizontalAlignment: "水平对齐",
  verticalAlignment: "垂直对齐",
  wrapText: "自动换行",
  indentLevel: "缩进",
  borders: "边框"
};

/**
 * 列出两代按字段摘要里**对不上**的字段名（供冲突文案指名）。
 *
 * `cellFormat` 一项还会再往下拆到具体属性 —— "单元格格式对不上"对用户没有可操作性，
 * 而"单元格格式（加粗）"直接指出该去哪一格看什么。
 */
function changedScopeFields(
  saved: UndoScopeFieldDigests | undefined,
  observed: UndoScopeFieldDigests,
  savedFormatFields?: Record<string, string>,
  observedFormatFields?: Record<string, string>
): string {
  const keys = Object.keys(observed) as (keyof UndoRestoreScope)[];
  const changed = keys.filter((key) => saved?.[key] !== observed[key]);
  if (changed.length === 0) return "未知字段";
  return changed
    .map((key) => {
      const label = UNDO_SCOPE_FIELD_LABELS[key] ?? String(key);
      if (key !== "cellFormat") return label;
      const current = observedFormatFields ?? {};
      const fields = Object.keys(current).filter(
        (field) => (savedFormatFields?.[field] ?? null) !== current[field]
      );
      if (fields.length === 0) return label;
      return `${label}（${fields.map((field) => FORMAT_FIELD_LABELS[field] ?? field).join("、")}）`;
    })
    .join("、");
}

/** 单个区域「执行后现场」的核对结论（只读）。 */
type UndoPostStateIssue =
  | { kind: "no_evidence"; detail: string }
  | { kind: "sheet_missing"; detail: string }
  | { kind: "changed"; detail: string };

/**
 * 只读核对**一个**区域现在还是不是"这次执行写入后的样子"。
 *
 * 抽出来的理由：撤销前的核对（`assertUndoPostState`）与续跑前/重开后的现场核对
 * （`observeOfficeExecutionPostState`）问的是**同一个问题**，只是处置不同 ——
 * 前者必须抛错停下，后者要把"逐动作的观测"交给共享判据（`verifyResumeScene`）去
 * 汇总结论。两份各写一遍"哪一代摘要优先"迟早会漂移，而漂移的后果是"同一个现场，
 * 撤销说变了、续跑说没变"。
 *
 * 判据优先级（由强到弱，逐级退回是为兼容旧恢复点，**不是**为放行）：
 *  1. 第三代 `scopeDigest` —— 覆盖范围恰好等于实际恢复范围，逐格/逐行；
 *  2. 第二代 `extendedFormatDigest` —— 固定的全属性集合；
 *  3. 第一代 `postFingerprint` —— 最旧的记录，只看值/公式/数字格式/填充色/加粗/字色。
 */
async function undoRangePostStateIssue(
  context: Excel.RequestContext,
  saved: UndoRangeSnapshot
): Promise<UndoPostStateIssue | null> {
  if (!saved.postFingerprint) {
    return { kind: "no_evidence", detail: "撤销数据缺少目标范围的执行后校验，已停止自动覆盖。" };
  }
  const sheet = context.workbook.worksheets.getItemOrNullObject(saved.sheet);
  sheet.load("isNullObject");
  await context.sync();
  if (sheet.isNullObject) {
    return { kind: "sheet_missing", detail: `工作表「${saved.sheet}」已不存在，已停止撤销。` };
  }
  const range = sheet.getRange(saved.range);
  const scope = scopeOf(saved);
  const observed = await readUndoGuardState(context, range, scope, saved.formatFields);
  if (saved.scopeDigest) {
    if (observed.scopeDigest !== saved.scopeDigest) {
      return {
        kind: "changed",
        detail:
          `「${saved.sheet}」${saved.range} 在本次撤销将要写回的字段（${changedScopeFields(
            saved.scopeFieldDigests,
            observed.scopeFieldDigests,
            saved.cellFormatFieldDigests,
            observed.cellFormatFieldDigests
          )}）上又发生了变化，已停止撤销以免覆盖后续编辑。`
      };
    }
    return null;
  }
  const stale = saved.extendedFormatDigest
    ? observed.extendedFormatDigest !== saved.extendedFormatDigest
    : observed.postFingerprint !== saved.postFingerprint;
  if (!stale) return null;
  return {
    kind: "changed",
    detail: `「${saved.sheet}」${saved.range} 在执行后又发生了变化，已停止撤销以免覆盖后续编辑。`
  };
}

/**
 * 只读：这次执行**写入过的每个目标**现在还是不是执行后的样子。
 *
 * 这是「续跑前现场核对」的取数口（判据在 core 的 `verifyResumeScene`）。为什么不能省：
 * 撤销成功之后如果处置留痕那次上报失败，账本里就只剩「已完成 2 / 共 3 步」，
 * 而表已经被撤销清空了 —— 照账本续跑会在一次已经回退掉的执行后面补写第 3 步。
 *
 * 读不到就报 `unreadable`，**不**当作"没变"：读不到不等于没变。
 */
export async function observeOfficeExecutionPostState(
  snapshot: Pick<ExecutionUndoSnapshot, "ranges" | "postStateCoverage">
): Promise<Array<{ actionIndex: number; state: "matched" | "changed" | "unreadable" }>> {
  return Excel.run(async (context) => {
    const observations: Array<{
      actionIndex: number;
      state: "matched" | "changed" | "unreadable";
    }> = [];
    for (const saved of snapshot.ranges) {
      let state: "matched" | "changed" | "unreadable";
      try {
        const issue = await undoRangePostStateIssue(context, saved);
        state = !issue ? "matched" : issue.kind === "changed" || issue.kind === "sheet_missing"
          ? "changed"
          : "unreadable";
      } catch {
        state = "unreadable";
      }
      observations.push({ actionIndex: saved.actionIndex, state });
    }
    return observations;
  });
}

export async function assertUndoPostState(
  context: Excel.RequestContext,
  snapshot: ExecutionUndoSnapshot
): Promise<void> {
  if (snapshot.postStateCoverage !== "complete") {
    throw new Error(
      "这份撤销数据缺少完整的执行后状态校验，已停止自动覆盖；请先人工核对当前工作簿。"
    );
  }
  const currentIdentity = currentOfficeWorkbookIdentity();
  if (
    snapshot.workbookIdentity &&
    currentIdentity &&
    snapshot.workbookIdentity !== currentIdentity
  ) {
    throw new Error("当前已不是生成这份撤销数据的工作簿，已停止撤销。");
  }
  for (const saved of snapshot.ranges) {
    const issue = await undoRangePostStateIssue(context, saved);
    if (issue) throw new Error(issue.detail);
  }

  const createdStates = new Map(
    (snapshot.createdWorksheetStates ?? []).map((state) => [state.sheet, state])
  );
  for (const sheetName of snapshot.createdWorksheets) {
    const expected = createdStates.get(sheetName);
    if (!expected?.postFingerprint) {
      throw new Error(`新建工作表「${sheetName}」缺少执行后校验，已停止自动删除。`);
    }
    const current = await readCreatedWorksheetState(context, sheetName);
    if (!current || current.postFingerprint !== expected.postFingerprint) {
      throw new Error(
        `新建工作表「${sheetName}」在执行后又发生了变化，已停止撤销以免删除后续编辑。`
      );
    }
  }
  for (const saved of snapshot.deletedWorksheets) {
    const sheet = context.workbook.worksheets.getItemOrNullObject(saved.sheet);
    sheet.load("isNullObject");
    await context.sync();
    if (!sheet.isNullObject) {
      throw new Error(
        `工作表「${saved.sheet}」已被重新创建，已停止撤销以免覆盖现有内容。`
      );
    }
  }
}

/**
 * 读取区域内**此刻**的逐格直接格式（兼容路径专用）。
 *
 * 为什么必须"此刻重新读"：撤销是**反序**执行的，而 `clear(Formats)` 之后要把
 * "非本动作属性"的当前值补回去（否则会连带清掉其它属性）。这个"当前"必须是
 * **每个动作恢复的那一刻**的现场 —— 反序恢复时，较早的动作看到的是较晚动作
 * **刚刚恢复过**的状态。若沿用"复核阶段一次性读回"的那份快照，较早的动作会把
 * 较晚动作刚撤销掉的属性又写回来（同区域"先加粗、后居中"就是这样：加粗恢复时
 * 用旧快照把居中写回，撤销后居中仍在）。
 *
 * 与 `readUndoGuardState` 用**同一套原语**（`DIRECT_CELL_PROPERTIES_LOAD` +
 * `readDirectCellFormat`），两份表示不可能漂移。
 */
async function readCurrentDirectCells(
  context: Excel.RequestContext,
  range: Excel.Range
): Promise<DirectCellFormatValues[][] | null> {
  // 只有兼容（直接格式）路径需要它；显示值路径与不可用宿主都返回 null。
  if (currentExcelFormatCapability().path !== "direct") return null;
  const properties = range.getCellProperties(DIRECT_CELL_PROPERTIES_LOAD);
  await context.sync();
  return properties.value.map((row) =>
    row.map((cell) => {
      const typed = cell as Excel.CellProperties;
      return readDirectCellFormat(typed.format, typed.style);
    })
  );
}

/**
 * 把捕获到的**直接格式子集**翻译成 `setCellProperties` 的载荷。
 *
 * 只写**确实出现过**的字段：没有的字段一律不出现在载荷里 ⇒ 该属性保持"未直接设置"。
 * 这一点是"未采集 ≠ 默认值"的执行侧体现 —— 绝不在这里补默认值。
 *
 * `style`（单元格命名样式）是 `SettableCellProperties` 的同级字段，**不进 `format`**；
 * 拼错层级不会报错，只会静默不生效。
 */
function buildDirectCellProperties(
  direct: DirectCellFormatValues
): Excel.SettableCellProperties {
  const format: Record<string, unknown> = {};
  const fill: Record<string, unknown> = {};
  if (direct.fillPattern != null) fill.pattern = direct.fillPattern;
  if (direct.fillColor != null) fill.color = direct.fillColor;
  if (direct.fillPatternColor != null) fill.patternColor = direct.fillPatternColor;
  if (direct.fillPatternTintAndShade != null) {
    fill.patternTintAndShade = direct.fillPatternTintAndShade;
  }
  if (direct.fillTintAndShade != null) fill.tintAndShade = direct.fillTintAndShade;
  if (Object.keys(fill).length > 0) format.fill = fill;
  const font: Record<string, unknown> = {};
  if (direct.fontName != null) font.name = direct.fontName;
  if (direct.fontSize != null) font.size = direct.fontSize;
  if (direct.fontBold != null) font.bold = direct.fontBold;
  if (direct.fontItalic != null) font.italic = direct.fontItalic;
  if (direct.fontColor != null) font.color = direct.fontColor;
  if (direct.fontUnderline != null) font.underline = direct.fontUnderline;
  if (direct.fontStrikethrough != null) font.strikethrough = direct.fontStrikethrough;
  if (direct.fontSubscript != null) font.subscript = direct.fontSubscript;
  if (direct.fontSuperscript != null) font.superscript = direct.fontSuperscript;
  if (direct.fontTintAndShade != null) font.tintAndShade = direct.fontTintAndShade;
  if (Object.keys(font).length > 0) format.font = font;
  if (direct.horizontalAlignment != null) format.horizontalAlignment = direct.horizontalAlignment;
  if (direct.verticalAlignment != null) format.verticalAlignment = direct.verticalAlignment;
  if (direct.wrapText != null) format.wrapText = direct.wrapText;
  if (direct.indentLevel != null) format.indentLevel = direct.indentLevel;
  if (direct.shrinkToFit != null) format.shrinkToFit = direct.shrinkToFit;
  if (direct.textOrientation != null) format.textOrientation = direct.textOrientation;
  if (direct.readingOrder != null) format.readingOrder = direct.readingOrder;
  if (direct.autoIndent != null) format.autoIndent = direct.autoIndent;
  if (direct.locked != null || direct.formulaHidden != null) {
    const protection: Record<string, unknown> = {};
    if (direct.locked != null) protection.locked = direct.locked;
    if (direct.formulaHidden != null) protection.formulaHidden = direct.formulaHidden;
    format.protection = protection;
  }
  if (direct.borders) format.borders = direct.borders;
  const properties: Excel.SettableCellProperties = {};
  if (direct.style != null) properties.style = direct.style;
  properties.format = format as Excel.CellPropertiesFormat;
  return properties;
}

/**
 * 恢复完整格式；捕获时因区域过大而降级的（null）直接跳过。
 *
 * 优先走**按粒度**字段：兼容路径的 `directCells`（直接格式集合）优先于显示值 `cells`，
 * 两者都没有才退回旧版区域级单值。
 */
function restoreRangeFormat(
  range: Excel.Range,
  saved: UndoRangeFormat | null | undefined,
  scope: UndoRestoreScope,
  formatFields?: readonly string[],
  observedDirectCells?: DirectCellFormatValues[][] | null
): void {
  if (!saved) return;
  // **只写本次动作真正会影响的部分**（`scope`）。写回"执行前的值"看起来无害，实则会把
  // 用户在执行之后对这些属性做的编辑一起抹掉 —— 本动作没改过的属性就不该被写。
  if (scope.cellFormat) restoreCellFormats(range, saved, formatFields, observedDirectCells);
  if (scope.rowHeight) restoreRowHeights(range, saved);
  if (scope.columnWidth) restoreColumnWidths(range, saved);
}

/**
 * 显示值路径下按**属性子集**拼 `setCellProperties` 载荷。
 *
 * 只含被请求的键：显示值路径不需要 `clear`（每格写的是确定的显示值），所以没被写回的
 * 属性**原样保留** —— 包括用户在本动作之后自己改的那些。
 */
function buildDisplayedCellFormat(
  cell: UndoCellFormat,
  keys: ReadonlySet<keyof UndoCellFormat>
): Record<string, unknown> {
  const format: Record<string, unknown> = {};
  if (keys.has("fillPattern") || keys.has("fillColor")) {
    const fill: Record<string, unknown> = { pattern: cell.fillPattern };
    if (cell.fillColor != null) fill.color = cell.fillColor;
    format.fill = fill;
  }
  const font: Record<string, unknown> = {};
  if (keys.has("fontName")) font.name = cell.fontName;
  if (keys.has("fontSize")) font.size = cell.fontSize;
  if (keys.has("fontBold")) font.bold = cell.fontBold;
  if (keys.has("fontItalic")) font.italic = cell.fontItalic;
  if (keys.has("fontColor")) font.color = cell.fontColor;
  if (keys.has("fontUnderline")) font.underline = cell.fontUnderline;
  if (Object.keys(font).length > 0) format.font = font;
  if (keys.has("horizontalAlignment")) format.horizontalAlignment = cell.horizontalAlignment;
  if (keys.has("verticalAlignment")) format.verticalAlignment = cell.verticalAlignment;
  if (keys.has("wrapText")) format.wrapText = cell.wrapText;
  if (keys.has("indentLevel")) format.indentLevel = cell.indentLevel;
  if (keys.has("borders")) {
    format.borders = {
      top: cell.borders.top,
      bottom: cell.borders.bottom,
      left: cell.borders.left,
      right: cell.borders.right
    };
  }
  return format;
}

/**
 * 逐格格式恢复（`scope.cellFormat` 为真时才会被调用）。
 *
 * **只写回本动作实际修改的属性**（`formatFields`）。写回"执行前的整份格式"看起来更保险，
 * 实则会把用户在本动作之后对**别的属性**（文字方向、收缩适应、保护……）做的编辑一起抹掉；
 * 而"本动作没碰过的属性"根本不该出现在恢复写入里。
 */
function restoreCellFormats(
  range: Excel.Range,
  saved: UndoRangeFormat,
  formatFields?: readonly string[],
  observedDirectCells?: DirectCellFormatValues[][] | null
): void {
  const format = range.format;
  // 旧恢复点没有 `formatFields` ⇒ 按全集处理（保守，且与旧行为一致）。
  const fields = formatFields ?? ALL_DIRECT_FORMAT_FIELDS;
  // 按粒度恢复优先：逐格格式矩阵优先，都没有才退回旧版区域级单值。
  // 新捕获（有 `cells`/`directCells`）时绝不走区域级单值路径 —— 否则"先拍平成单值、
  // 再逐格覆盖"不仅冗余，还会把"未逐格采集却因旧单值而短暂写坏"的隐患带进来。
  if (saved.directCells) {
    const directKeys = directKeysOfFields(fields);
    if (directKeys.size === 0) return;
    // —— 兼容路径（宿主 < ExcelApi 1.19）：恢复的是**格式的来源语义**，不是显示值。
    //
    // 只有当本动作确实把某个"执行前**未**直接设置"的属性设上过，才必须先 `clear(Formats)`：
    // Excel 没有"只取消某一个属性的直接设置"的 API，要让那个属性回到继承态只能整格清掉。
    // 其余情况**一律不清**（用户要求：不能动不动就 clear），按属性写回即可。
    const needsClear = saved.directCells.some((row) =>
      row.some((cell) => cell.inherited.length > 0)
    );
    const matrix = saved.directCells.map((row, rowIndex) =>
      row.map((cell, columnIndex) => {
        // 本动作管辖的属性 ⇒ 回到执行前的取值；当初没直接设置的键就**不出现**（保持继承）。
        const restored = pickKeys(cell.direct, directKeys) as DirectCellFormatValues;
        if (!needsClear) return buildDirectCellProperties(restored);
        // `clear` 会带走**整格**直接格式（含数字格式与命名样式），所以必须把
        // "非本动作属性"的**当前**取值补回去 —— 否则执行前就有的其它属性、以及用户在本动作
        // 之后的编辑都会一起丢。这就是"撤销只覆盖实际恢复范围"在 clear 路径上的落实。
        const current = observedDirectCells?.[rowIndex]?.[columnIndex] ?? {};
        const merged: Record<string, unknown> = { ...current };
        for (const key of Object.keys(restored)) {
          merged[key] = (restored as Record<string, unknown>)[key];
        }
        // 本动作管辖但当初未直接设置的键 ⇒ 显式剔除（清空后保持"未直接设置"）。
        for (const key of directKeys) {
          if ((restored as Record<string, unknown>)[key] === undefined) delete merged[key];
        }
        return buildDirectCellProperties(merged as DirectCellFormatValues);
      })
    );
    if (needsClear) range.clear(Excel.ClearApplyTo.formats);
    range.setCellProperties(matrix as Excel.SettableCellProperties[][]);
    return;
  }
  if (saved.cells) {
    const displayedKeys = displayedKeysOfFields(fields);
    if (displayedKeys.size === 0) return;
    // 逐格写回**执行前的显示值**里、属于本动作属性的那一部分。其余属性不写 ⇒ 原样保留。
    range.setCellProperties(
      saved.cells.map((row) =>
        row.map((cell) => ({
          format: buildDisplayedCellFormat(cell, displayedKeys)
        }))
      ) as Excel.SettableCellProperties[][]
    );
    return;
  }
  // 旧恢复点（无逐格矩阵）退回区域级单值。
  format.horizontalAlignment =
    saved.horizontalAlignment as Excel.HorizontalAlignment;
  format.verticalAlignment =
    saved.verticalAlignment as Excel.VerticalAlignment;
  if (saved.wrapText !== null) format.wrapText = saved.wrapText;
  format.indentLevel = saved.indentLevel;
  format.font.name = saved.fontName;
  format.font.size = saved.fontSize;
  format.font.italic = saved.fontItalic;
  format.font.underline = saved.fontUnderline as Excel.RangeUnderlineStyle;
  for (const border of saved.borders) {
    const target = format.borders.getItem(border.sideIndex);
    target.style = border.style as Excel.BorderLineStyle;
    target.color = border.color;
    target.weight = border.weight as Excel.BorderWeight;
  }
}

/**
 * 行高恢复（`scope.rowHeight` 为真时才会被调用）。三态：
 *  - 全部 `number`（直接设置值）：逐行写回。
 *  - 全部 `"standard"`（未直接设置，标准行高）：`SettableRowProperties` 无
 *    `useStandardHeight` 字段，无法逐行写"标准"，只能整区域 `format.useStandardHeight=true`
 *    一次性重置。
 *  - 混合 / 含 null：`captureUndoRange` 已对 resizeRange 执行前拒绝，不会走到撤销；
 *    这里仅作旧恢复点兼容（退回区域级单值兜底）。
 */
function restoreRowHeights(range: Excel.Range, saved: UndoRangeFormat): void {
  const format = range.format;
  if (saved.rowHeights) {
    const allStandard = saved.rowHeights.every((h) => h === "standard");
    const allNumber = saved.rowHeights.every((h) => typeof h === "number");
    if (allStandard) {
      format.useStandardHeight = true;
    } else if (allNumber) {
      range.setRowProperties(
        saved.rowHeights.map((height) => ({ format: { rowHeight: height as number } }))
      );
    } else {
      format.rowHeight = saved.rowHeight;
    }
  } else {
    format.rowHeight = saved.rowHeight;
  }
}

/** 列宽恢复（`scope.columnWidth` 为真时才会被调用）。三态语义同 `restoreRowHeights`。 */
function restoreColumnWidths(range: Excel.Range, saved: UndoRangeFormat): void {
  const format = range.format;
  if (saved.columnWidths) {
    const allStandard = saved.columnWidths.every((w) => w === "standard");
    const allNumber = saved.columnWidths.every((w) => typeof w === "number");
    if (allStandard) {
      format.useStandardWidth = true;
    } else if (allNumber) {
      range.setColumnProperties(
        saved.columnWidths.map((width) => ({ format: { columnWidth: width as number } }))
      );
    } else {
      format.columnWidth = saved.columnWidth;
    }
  } else {
    format.columnWidth = saved.columnWidth;
  }
}

export async function undoExecution(
  snapshot: ExecutionUndoSnapshot
): Promise<void> {
  await Excel.run(async (context) => {
    // 所有检查均在首个写操作之前完成；不满足条件时整次撤销保持只读失败。
    await assertUndoPostState(context, snapshot);
    // 先恢复被删除的工作表，让引用它们的单元格撤销能落回原表。
    const worksheets = context.workbook.worksheets;
    // 按原始位置升序逐个恢复：每个被删表原本位于其前面的工作表（无论
    // 残留还是已恢复）数量不小于其原始位置，因此直接赋回位置不会越界。
    const deletedSorted = [...snapshot.deletedWorksheets].sort(
      (a, b) => a.position - b.position
    );
    for (const saved of deletedSorted) {
      const ws = worksheets.add(saved.sheet);
      const range = ws.getRange(saved.usedRange);
      range.formulas = saved.formulas;
      range.numberFormat = saved.numberFormat;
      ws.position = saved.position;
    }
    await context.sync();
    // 再恢复单元格值/公式/常用格式。**每一项都按本次动作实际影响的范围来写**：
    // 只改行高的动作不写值、不写格式；只改值的动作不写任何格式。
    for (const saved of [...snapshot.ranges].reverse()) {
      const range = context.workbook.worksheets
        .getItem(saved.sheet)
        .getRange(saved.range);
      const scope = scopeOf(saved);
      if (scope.formulas) range.formulas = saved.formulas;
      // `clear(Formats)` 之后要补回的"非本动作属性当前值"必须是**此刻**的现场：
      // 反序恢复时较早的动作会看到较晚动作刚恢复过的状态。用复核阶段那份一次性
      // 旧快照，就会把较晚动作刚撤销掉的属性又写回来。
      const observedDirect = scope.cellFormat
        ? await readCurrentDirectCells(context, range)
        : null;
      restoreRangeFormat(
        range,
        saved.format,
        scope,
        saved.formatFields,
        observedDirect
      );
      // 旧恢复点（区域过大而未逐格采集）退回区域级常用格式；新记录一律逐格恢复，
      // **不再做区域级重写**（那正是把"没被本动作改过的格式"拍平的来源）。
      const hasCellMatrix = saved.format?.cells != null || saved.format?.directCells != null;
      if (scope.cellFormat && !hasCellMatrix) {
        range.format.fill.color = saved.fillColor;
        range.format.font.bold = saved.fontBold;
        range.format.font.color = saved.fontColor;
      }
      // 数字格式放在**格式恢复之后**：兼容路径的 `clear(Formats)` 会把直接设置的数字格式
      // 一并清掉（Excel 的"清除格式"就包含数字格式），写在它前面会被清没。
      if (scope.numberFormat) range.numberFormat = saved.numberFormat;
      await context.sync();
      range.load('formulas,numberFormat');
      await context.sync();
      if ((scope.formulas && !sameMatrix(range.formulas.map(row => row.map(value => value ?? "")), saved.formulas.map(row => row.map(value => value ?? "")))) ||
          (scope.numberFormat && !sameMatrix(range.numberFormat, saved.numberFormat))) {
        throw new Error(`恢复后回读不一致：「${saved.sheet}」${saved.range}。部分恢复可能已完成，请保留现场核对。`);
      }
      if (saved.format && (scope.cellFormat || scope.rowHeight || scope.columnWidth)) {
        const observed = await captureRangeFormat(context, range,
          saved.formulas.length * (saved.formulas[0]?.length ?? 0), saved.formulas.length,
          saved.formulas[0]?.length ?? 0, (saved.formatFields ?? ALL_DIRECT_FORMAT_FIELDS) as ExcelFormatField[], scope);
        if (!observed) throw new Error('恢复后无法核实格式，部分恢复可能已完成，请保留现场核对。');
        const fields = saved.formatFields ?? ALL_DIRECT_FORMAT_FIELDS;
        const direct = Boolean(saved.format.directCells);
        const keys = asKeySet(direct ? directKeysOfFields(fields) : displayedKeysOfFields(fields));
        const project = (matrix: unknown[][] | null | undefined) => matrix?.map(row => row.map(cell => cell ? projectKeys(cell as Record<string, unknown>, keys) : null));
        const expectedCells = direct ? saved.format.directCells : saved.format.cells;
        const actualCells = direct ? observed.directCells : observed.cells;
        if ((scope.cellFormat && expectedCells && localStateFingerprint(project(expectedCells)) !== localStateFingerprint(project(actualCells))) ||
            (scope.rowHeight && localStateFingerprint(observed.rowHeights ?? observed.rowHeight) !== localStateFingerprint(saved.format.rowHeights ?? saved.format.rowHeight)) ||
            (scope.columnWidth && localStateFingerprint(observed.columnWidths ?? observed.columnWidth) !== localStateFingerprint(saved.format.columnWidths ?? saved.format.columnWidth))) {
          throw new Error(`恢复后格式回读不一致：「${saved.sheet}」${saved.range}。部分恢复可能已完成，请保留现场核对。`);
        }
      }
    }
    // 最后移除本次执行新建的工作表。
    for (const name of snapshot.createdWorksheets) {
      const ws = context.workbook.worksheets.getItemOrNullObject(name);
      ws.load("isNullObject");
      await context.sync();
      if (!ws.isNullObject) ws.delete();
    }
    await context.sync();
  });
}

export interface ExecutePlanOptions {
  /** Read-only native review guard, after durable callbacks and immediately before first mutation. */
  beforeFirstWrite?: (context: Excel.RequestContext) => Promise<void>;
  /** 用户本次明确授权的现有工作表名称；新输出表可由计划创建。 */
  authorizedSheetNames: readonly string[];
  /** 本次会话已勾选的外部工作簿文件名集合，用于放行跨文件公式引用。 */
  allowedExternalFiles?: ReadonlySet<string>;
  /**
   * 当计划含烘焙矩阵且源数据已变时，UI 二次确认后带上此标志重试。
   * 不带此标志时，数据已变的 mixed 计划会抛 StaleSourceError 而非直接执行。
   */
  acknowledgeStaleSource?: boolean;
  /** 取消信号，透传给沙箱脚本执行。 */
  signal?: AbortSignal;
  /** 前端生成的单次操作标识，用于归并、去重与恢复进度。 */
  operationId?: string;
  /** 由确定性执行器产生的结构化进度，不包含模型隐藏推理。 */
  onProgress?: (event: TaskProgressEvent) => void;
  /**
   * 仅供已登记的 Office 区域动作恢复使用。回调必须在写入前/后持久提交；
   * 任一回调失败都会停止后续动作，避免产生未经登记的写入。
   */
  recovery?: {
    startActionIndex: number;
    undoSnapshot?: ExecutionUndoSnapshot | null;
    onBeforeAction: (input: {
      actionIndex: number;
      beforeSnapshot: ExecutionUndoSnapshot["ranges"][number];
    }) => Promise<void>;
    onActionCompleted: (input: {
      actionIndex: number;
      checkpoint: ExecutionUndoSnapshot;
    }) => Promise<void>;
  };
  /**
   * 拿到本次执行的"实际变更"对照。
   *
   * 走回调而不是塞进 PlanExecutionResult：结果会被服务端持久化，
   * 而完整前后矩阵按隐私边界只允许停留在会话内内存。
   */
  onChangeReview?: (reviews: RangeChangeReview[]) => void;
  /**
   * 每个动作的**宿主写入事实**（A1-E3 的"实际写入次数"口径）。
   *
   * 只报计数与状态，不含单元格值或公式正文。口径与判据见
   * `@excel-bro/core/hostWriteAudit`：一个动作可能分块写（多次刷新），所以
   * "动作数"不等于"宿主写调用数"，只有这里数字才是真的向宿主发起的次数。
   * 回调是**观测**：它抛错不应该挡住写入，调用方自己兜住。
   */
  onHostWrite?: (entry: HostWriteAudit) => void;
}

export async function executePlan(
  plan: AnalysisPlan,
  options: ExecutePlanOptions
): Promise<PlanExecutionResult> {
  const operationId =
    options.operationId ?? `execution-${Date.now().toString(36)}`;
  const startActionIndex = options.recovery?.startActionIndex ?? 0;
  if (startActionIndex < 0 || startActionIndex > plan.actions.length) {
    throw new PlanExecutionError(
      "Office 断点的下一动作位置无效，已停止续跑。",
      incompleteActionResults(plan, new Map())
    );
  }
  if (startActionIndex > 0 && !options.recovery?.undoSnapshot) {
    throw new PlanExecutionError(
      "Office 断点缺少已完成动作的恢复快照，已停止续跑；请保留现场并重新核对，不能仅为剩余动作创建撤销点。",
      incompleteActionResults(plan, new Map())
    );
  }
  if (
    options.recovery?.undoSnapshot &&
    (options.recovery.undoSnapshot.planId !== plan.id ||
      options.recovery.undoSnapshot.executionId !== operationId)
  ) {
    throw new PlanExecutionError(
      "Office 断点的计划或执行标识不一致，已停止续跑。",
      incompleteActionResults(plan, new Map())
    );
  }
  let progressSequence = 0;
  const emitProgress = (
    event: Omit<
      TaskProgressEvent,
      "schemaVersion" | "operationId" | "sequence" | "timestamp"
    >
  ): void => {
    if (!options.onProgress) return;
    progressSequence += 1;
    options.onProgress({
      schemaVersion: 1,
      operationId,
      sequence: progressSequence,
      timestamp: new Date().toISOString(),
      ...event
    });
  };
  emitProgress({
    phase: "preflight",
    stepId: "execution.preflight",
    status: "started",
    title: "正在执行前检查",
    summary: `正在核对 ${plan.actions.length} 步操作的授权范围、目标和来源结构。`,
    current: 0,
    total: plan.actions.length,
    unit: "actions"
  });
  const authorizedSheetKeys = new Set(
    options.authorizedSheetNames.map((name) => name.toLocaleLowerCase())
  );
  const unauthorizedFingerprintSheet = plan.sourceFingerprintSheets?.find(
    (name) => !authorizedSheetKeys.has(name.toLocaleLowerCase())
  );
  if (unauthorizedFingerprintSheet) {
    throw new PlanExecutionError(
      `数据来源指纹尝试读取未授权的工作表「${unauthorizedFingerprintSheet}」`,
      incompleteActionResults(plan, new Map())
    );
  }

  /*
   * **写前来源闸门（计划级）**：这份计划里的固定值，是对哪一版源算出来的。
   *
   * 以前这一段的入口条件是 `plan.sourceFingerprint && plan.sourceFingerprintSheets?.length`
   * —— 没有指纹就整块跳过，于是一份"写了固定值、却不知道自己按哪一版算的"计划
   * 会被直接写进工作表。判据现在拆到共享层（`judgePlanSourceGate`，与 WPS 同一份）：
   * 缺锚点、锚点不完整、覆盖范围与当前授权不一致，一律**拒绝**；
   * **不拿"现在读一次"补一个基线**（那正是 2026-10-25 那个缺陷的另一种写法）。
   */
  const anchor = planSourceAnchor(plan);
  if (!isTrustedLiteralPlan(plan) && (anchor || planHasBakedValues(plan))) {
    if (anchor && isToolSchemaFingerprint(anchor.fingerprint)) {
      // 固化工具：基线本身就是 schema 指纹，照原样比。
      const current = toolSchemaFingerprintForSnapshot(
        plan,
        await captureWorkbookStructure([...anchor.sheets])
      );
      if (current !== anchor.fingerprint) {
        throw new PlanExecutionError(
          "工具所需的来源工作表或字段已发生变化，请重新选择后再执行",
          incompleteActionResults(plan, new Map())
        );
      }
    } else if (planLiveness(plan) === "live") {
      // 写入全在执行时求值：源数据变了正是它该重算的场景，只要表和字段还在
      // 就放行。
      //
      // 这里**不写 `anchor!`**：入口条件是 `anchor || planHasBakedValues(plan)`，
      // "live 且没有锚点"仍然可能出现（例如计划自己声明了覆盖表之外的写法）。
      // 非空断言会把一次可读的拒绝变成 TypeError（`Cannot read properties of null`）。
      // 烘焙清单已经统一到 core 一份，"两份漂移导致的 live 空锚点"不再是通路；
      // 但判据本身与清单无关，留作 fail-closed。
      const liveSourceSheets = anchor ? [...anchor.sheets] : [];
      if (liveSourceSheets.length === 0) {
        throw new PlanExecutionError(
          "这份计划写入的内容要在执行时求值，但没有记录它依赖哪些工作表，无法核对来源；未执行。",
          incompleteActionResults(plan, new Map())
        );
      }
      const missing = missingSchemaRequirements(
        plan,
        await captureWorkbookStructure(liveSourceSheets)
      );
      if (missing.length > 0) {
        throw new PlanExecutionError(
          `计划所需的来源已发生变化，缺少：${missing.join("、")}，请重新选择后再执行`,
          incompleteActionResults(plan, new Map())
        );
      }
    } else {
      // 含烘焙矩阵（或根本没有锚点）：这些值是预览时算出来的，源数据一变就过期。
      const current = planSourceComparisonNeeded({ plan, anchor })
        ? await captureWorkbookSourceFingerprint([...anchor!.sheets])
        : null;
      const verdict = judgePlanSourceGate({
        plan,
        anchor,
        authorizedSheets: options.authorizedSheetNames,
        liveFingerprint: current
      });
      const bakedActionCount = plan.actions.filter((action) =>
        BAKED_ACTION_TYPES.has(action.type)
      ).length;
      /*
       * 「仍要执行旧结果」只覆盖**已确认变化**这一档：用户知道源变了还坚持写，
       * 那是一个明确的决定。`missing` / `unverifiable` / 覆盖范围不一致不给他
       * 这个选项 —— 那不是"我知道旧但仍然要"，而是"根本不知道它是按哪一版算的"。
       */
      if (verdict.verdict === "changed") {
        if (!options.acknowledgeStaleSource) {
          throw new StaleSourceError(
            `数据来源在预览后发生了变化，这份计划里有 ${bakedActionCount} 处写入的是预览时算出的固定值，执行会写入旧结果。建议重新生成预览；也可以坚持执行。`,
            incompleteActionResults(plan, new Map()),
            bakedActionCount
          );
        }
      } else if (verdict.verdict !== "ok") {
        throw new PlanExecutionError(
          planSourceGateMessage(verdict) ?? PLAN_SOURCE_MISSING_MESSAGE,
          incompleteActionResults(plan, new Map())
        );
      }
    }
  }
  /*
   * **续跑前的公式来源校验**（产品规则 2026-10-30）：公式精确恢复要求登记的来源证据
   * 仍然一致。放在计划级来源闸**之后**，两级判据各说各的事实：
   *
   * - 前面那级回答"表还在不在、结构对不对"（`missingSchemaRequirements`）；
   * - 这一级回答"那几张表里的**值**是不是还是登记时那一版"。
   *
   * 顺序不能颠倒：表都删了却说"来源值变了"是误导。也不能合并：一句"来源有问题"
   * 会让用户不知道是去恢复表，还是去重新读数据。
   *
   * 只在续跑时判（`startActionIndex > 0`）：首次执行写下去的公式由宿主按那一刻的源值
   * 求值，中间没有"前后两段用不同源值算"的问题 —— 那正是"不放宽正常重算"的意思。
   * 判据本体在 core（`verifyFormulaResumeSources`），WPS 读同一份。
   */
  if (startActionIndex > 0) {
    const resumeSourceVerdict = await verifyFormulaResumeSources(plan, (sheets) =>
      captureWorkbookSourceFingerprint([...sheets])
    );
    if (resumeSourceVerdict.verdict === "refused") {
      throw new PlanExecutionError(
        resumeSourceVerdict.reason,
        incompleteActionResults(plan, new Map())
      );
    }
  }
  return Excel.run(async (context) => {
    /*
     * 「实际写入了几次」的观测点。
     *
     * Office.js 把排队的写操作在 `sync()` 时按批提交 —— 真正到达宿主的是**刷新**，
     * 不是某个 `range.values = ...` 赋值。所以要数的就是动作窗口内的刷新次数：
     * 它同时覆盖"一个动作分块写多次"（每个块各刷一次）与"还没写就失败"（0 次）。
     * 只包这一个实例，不动宿主对象原型，也不改任何写入语义。
     */
    let hostFlushes = 0;
    const flushContext = context.sync.bind(context) as (
      ...args: unknown[]
    ) => Promise<unknown>;
    (context as unknown as { sync: (...args: unknown[]) => Promise<unknown> }).sync = (
      ...args: unknown[]
    ) => {
      hostFlushes += 1;
      return flushContext(...args);
    };
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name,items/position");
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
    // 记录各工作表在计划执行前的原始位置，供删除撤销时精确还原顺序。
    const positionByName = new Map<string, number>(
      worksheets.items.map((ws) => [ws.name, ws.position])
    );
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
      },
      options.authorizedSheetNames
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
    emitProgress({
      phase: "preflight",
      stepId: "execution.preflight",
      status: "completed",
      title: "执行前检查已通过",
      summary: `已确认 ${plan.actions.length} 步操作均在本次授权与白名单范围内。`,
      current: plan.actions.length,
      total: plan.actions.length,
      unit: "actions",
      trace: {
        kind: "evidence",
        summary: "执行前检查已通过。",
        basis: "工作表授权、动作白名单、目标命名和运行环境能力检查"
      }
    });

    const actionResults: ActionExecutionResult[] = plan.actions
      .slice(0, startActionIndex)
      .map((action, index) => ({
        index,
        type: action.type,
        sheet: action.sheet,
        status: "succeeded"
      }));
    const dynamicCriteria: VerificationCriterion[] = [];
    const undoRanges: ExecutionUndoSnapshot["ranges"] = [
      ...(options.recovery?.undoSnapshot?.ranges ?? [])
    ];
    // 双栏对照的"修改前"矩阵。只活在本次执行的内存里，不进撤销快照、
    // 不进任何持久化或出站结构（隐私边界见指导文件 §7）。
    const changeBefore = new Map<number, RangeChangeSnapshot>();
    const createdWorksheets: string[] = [];
    const deletedWorksheets: UndoWorksheetSnapshot[] = [];
    const skippedWorksheets: string[] = [];
    const executionStartedAt = performance.now();
    let recoveryCheckpoint = options.recovery?.undoSnapshot ?? null;
    for (let index = startActionIndex; index < plan.actions.length; index += 1) {
      const action = plan.actions[index];
      let undoRange: ExecutionUndoSnapshot["ranges"][number] | null = null;
      // 本动作的宿主观测：只在成功路径报一次，失败路径由 catch 兜底报一次。
      let hostWriteReported = false;
      let flushesBeforeAction = hostFlushes;
      let undoCollected = false;
      let deletedCapture: Awaited<ReturnType<typeof captureDeletedWorksheet>> | null = null;
      const actionUndo: ActionUndoCollector = {
        createdWorksheets: [],
        overwrittenRanges: []
      };
      try {
        emitProgress({
          phase: "backup",
          stepId: `execution.action.${index + 1}.undo`,
          status: "started",
          title: `正在保存第 ${index + 1} 步撤销数据`,
          actionIndex: index + 1,
          actionCount: plan.actions.length,
          actionType: action.type,
          target: { sheet: action.sheet }
        });
        const captured = await captureUndoRange(context, action, index);
        undoRange = captured.snapshot;
        if (captured.snapshot) {
          // 双栏的"修改前"：值是单独采的，公式沿用撤销快照，两者都留，
          // 但不进持久化快照（beforeValues 与 snapshot 分开）。
          changeBefore.set(index, {
            values:
              captured.beforeValues ??
              captured.snapshot.formulas.map((row) => row.map(() => null)),
            formulas: toChangeFormulas(captured.snapshot.formulas),
            valuesCaptured: captured.beforeValues !== null
          });
        }
        deletedCapture =
          action.type === "deleteWorksheet"
            ? await captureDeletedWorksheet(
                context,
                action.sheet,
                positionByName.get(action.sheet)
              )
            : null;
        if (options.recovery) {
          if (!undoRange) {
            throw new Error(
              `第 ${index + 1} 步缺少完整的区域写前快照，不能按恢复模式执行`
            );
          }
          await options.recovery.onBeforeAction({
            actionIndex: index,
            beforeSnapshot: undoRange
          });
        }
        emitProgress({
          phase: "backup",
          stepId: `execution.action.${index + 1}.undo`,
          status: "completed",
          title: `第 ${index + 1} 步撤销数据已准备`,
          summary:
            undoRange || deletedCapture
              ? "已保存本步执行前状态；结构化输出的覆盖快照会在写入时继续记录。"
              : "本步没有可预先捕获的单元格范围；如创建新表，执行器会记录新表名称。",
          actionIndex: index + 1,
          actionCount: plan.actions.length,
          actionType: action.type,
          target: { sheet: action.sheet }
        });
        emitProgress({
          phase: "executing",
          stepId: `execution.action.${index + 1}`,
          status: "started",
          title: `正在执行第 ${index + 1} / ${plan.actions.length} 步`,
          summary: `${action.type} · ${action.sheet}`,
          current: index,
          total: plan.actions.length,
          unit: "actions",
          actionIndex: index + 1,
          actionCount: plan.actions.length,
          actionType: action.type,
          target: { sheet: action.sheet }
        });
        if (index === startActionIndex) await options.beforeFirstWrite?.(context);
        // 窗口起点：撤销快照的读取刷新不算本动作的写入。
        flushesBeforeAction = hostFlushes;
        const actionCriteria = await executeAction(
          context,
          action,
          index,
          actionUndo,
          options.allowedExternalFiles,
          options.signal,
          (update) =>
            emitProgress({
              phase: "executing",
              stepId: update.stepId,
              status: "progress",
              title: update.title,
              summary: update.summary,
              current: update.current,
              total: update.total,
              unit: update.unit,
              actionIndex: index + 1,
              actionCount: plan.actions.length,
              actionType: action.type,
              target: update.target
            })
        );
        if (actionCriteria) dynamicCriteria.push(...actionCriteria);
        await context.sync();
        if (undoRange) undoRanges.push(undoRange);
        if (actionUndo.overwrittenRanges.length > 0) {
          undoRanges.push(...actionUndo.overwrittenRanges);
        }
        if (action.type === "createWorksheet") {
          createdWorksheets.push(action.sheet);
        }
        if (actionUndo.createdWorksheets.length > 0) {
          createdWorksheets.push(...actionUndo.createdWorksheets);
        }
        if (deletedCapture?.snapshot) {
          deletedWorksheets.push(deletedCapture.snapshot);
        } else if (deletedCapture?.skipped) {
          skippedWorksheets.push(action.sheet);
        }
        undoCollected = true;
        /*
         * 上报本动作的宿主观测。`hostWriteCalls` 是窗口内的刷新次数：
         * 0 次说明**没有真的向宿主提交过写入**（此时下面若失败只能记 failed_before_write）。
         * 回读口径：恢复模式下动作级写后指纹已核对 ⇒ verified；其余模式本次不做
         * 动作级回读（整份计划的回读核对在 executePlan 之后），如实记 not_applicable。
         */
        hostWriteReported = true;
        try {
          options.onHostWrite?.({
            actionIndex: index,
            actionType: action.type,
            hostWriteCalls: Math.max(hostFlushes - flushesBeforeAction, 0),
            mutatingCalls: 1,
            outcome: "written",
            readback: options.recovery ? "verified" : "not_applicable"
          });
        } catch {
          // 观测失败不许影响写入：这是统计，不是判据。
        }
        if (options.recovery) {
          if (!undoRange) {
            throw new Error(`第 ${index + 1} 步无法生成写后核对证据`);
          }
          const postState = await captureUndoPostState(context, [undoRange], [], []);
          if (!postState.complete || !undoRange.postFingerprint) {
            throw new Error(`第 ${index + 1} 步写后状态未能完整核对`);
          }
          const workbookIdentity = currentOfficeWorkbookIdentity();
          if (!workbookIdentity) {
            throw new Error("当前工作簿缺少稳定身份，不能提交 Office 动作回执");
          }
          recoveryCheckpoint = appendOfficeRangeCheckpoint({
            plan,
            operationId,
            workbookIdentity,
            previous: recoveryCheckpoint,
            before: undoRange,
            actionIndex: index,
            postFingerprint: undoRange.postFingerprint
          });
          await options.recovery.onActionCompleted({
            actionIndex: index,
            checkpoint: recoveryCheckpoint
          });
        }
      } catch (reason) {
        if (!hostWriteReported) {
          /*
           * 失败也要如实记一次，而且**区分有没有进过宿主**：
           * - 窗口内刷新 0 次 ⇒ 写入根本没提交（failed_before_write，不能算已写入）；
           * - 刷新过但随后失败 ⇒ 可能写了一部分（failed_after_write，要回读核对）。
           * 这一条是"认领后、写入前失败不能被统计成已写入"的依据。
           */
          hostWriteReported = true;
          try {
            options.onHostWrite?.({
              actionIndex: index,
              actionType: action.type,
              hostWriteCalls: Math.max(hostFlushes - flushesBeforeAction, 0),
              mutatingCalls: 1,
              outcome:
                hostFlushes > flushesBeforeAction
                  ? "failed_after_write"
                  : "failed_before_write",
              readback: "not_applicable",
              errorCode: "EXCEL_ACTION_FAILED"
            });
          } catch {
            // 同上：统计失败不改变执行结果。
          }
        }
        if (!undoCollected) {
          if (undoRange) undoRanges.push(undoRange);
          if (actionUndo.overwrittenRanges.length > 0) {
            undoRanges.push(...actionUndo.overwrittenRanges);
          }
          if (action.type === "createWorksheet") {
            createdWorksheets.push(action.sheet);
          }
          if (actionUndo.createdWorksheets.length > 0) {
            createdWorksheets.push(...actionUndo.createdWorksheets);
          }
          if (deletedCapture?.snapshot) {
            deletedWorksheets.push(deletedCapture.snapshot);
          } else if (deletedCapture?.skipped) {
            skippedWorksheets.push(action.sheet);
          }
        }
        const detail = excelErrorDetail(reason);
        emitProgress({
          phase: "executing",
          stepId: `execution.action.${index + 1}`,
          status: "failed",
          title: `第 ${index + 1} 步执行失败`,
          summary: detail,
          current: index,
          total: plan.actions.length,
          unit: "actions",
          actionIndex: index + 1,
          actionCount: plan.actions.length,
          actionType: action.type,
          target: { sheet: action.sheet }
        });
        const partialSnapshot = await buildExecutionUndoSnapshot(context, {
          planId: plan.id,
          executionId: operationId,
          coverage: "partial",
          completedActionCount: actionResults.length,
          uncertainActionIndexes: [index],
          ranges: undoRanges,
          createdWorksheets,
          deletedWorksheets,
          skippedWorksheets
        });
        throw new PlanExecutionError(
          `第 ${index + 1} 步执行失败：${detail}`,
          incompleteActionResults(
            plan,
            new Map([[index, detail]]),
            actionResults
          ),
          partialSnapshot
        );
      }
      actionResults.push({
        index,
        type: action.type,
        sheet: action.sheet,
        status: "succeeded"
      });
      emitProgress({
        phase: "executing",
        stepId: `execution.action.${index + 1}`,
        status: "completed",
        title: `第 ${index + 1} / ${plan.actions.length} 步已完成`,
        summary: `${action.type} 已写入 ${action.sheet}`,
        current: index + 1,
        total: plan.actions.length,
        unit: "actions",
        actionIndex: index + 1,
        actionCount: plan.actions.length,
        actionType: action.type,
        target: { sheet: action.sheet }
      });
    }
    const completedSnapshot = await buildExecutionUndoSnapshot(context, {
      planId: plan.id,
      executionId: operationId,
      coverage: "full",
      completedActionCount: actionResults.length,
      uncertainActionIndexes: [],
      ranges: undoRanges,
      createdWorksheets,
      deletedWorksheets,
      skippedWorksheets
    });
    // 事件点补读：写入刚完成、还在同一个 RequestContext 里，按本次真正写过的
    // 区域读一次"修改后"矩阵。只对这些区域读，不做全簿扫描，也不轮询。
    // 回读本身失败**不阻断**主流程，但会把状态标记成 readback_failed。
    const changeReviews = await collectChangeReviews(context, undoRanges, changeBefore);
    const verificationStartedAt = performance.now();
    const executionMs = verificationStartedAt - executionStartedAt;
    emitProgress({
      phase: "verifying",
      stepId: "execution.verify",
      status: "started",
      title: "正在独立验证写入结果"
    });
    let verification: VerificationReport;
    try {
      verification = await verifyPlan(context, plan, dynamicCriteria);
    } catch (reason) {
      const detail = excelErrorDetail(reason);
      throw new PlanExecutionError(
        `写入已经完成，但独立验证失败：${detail}`,
        actionResults,
        completedSnapshot
      );
    }
    const verificationMs = performance.now() - verificationStartedAt;
    const passedChecks = verification.checks.filter((check) => check.passed).length;
    emitProgress({
      phase: "verifying",
      stepId: "execution.verify",
      status: verification.status === "failed" ? "warning" : "completed",
      title:
        verification.status === "failed"
          ? "执行已完成，但验证发现不一致"
          : verification.status === "executed_unverified"
            ? "执行已完成，部分结果暂无法独立验证"
            : "写入结果验证通过",
      summary: `${passedChecks} / ${verification.checks.length} 项检查符合预期。`,
      current: passedChecks,
      total: verification.checks.length,
      unit: "checks",
      trace: {
        kind: verification.status === "failed" ? "conflict" : "evidence",
        summary: `${passedChecks} / ${verification.checks.length} 项检查符合预期。`,
        basis: "执行后重新读取工作簿并按计划验收条件比较"
      }
    });
    // 双栏的成败口径跟随独立验收：验收判失败就是不一致，验收没核完就是
    // executed_unverified —— 两种情况都不许让界面宣称成功（core 的
    // canClaimWriteSuccess 只对 applied 返回 true）。
    const reviewStatus = changeReviewStatusFromVerification(verification.status);
    const finalReviews = changeReviews.map((review) =>
      review.status === "readback_failed"
        ? review
        : { ...review, status: reviewStatus }
    );
    options.onChangeReview?.(finalReviews);
    return {
      actionResults,
      verification,
      executionMs,
      verificationMs,
      undoSnapshot: completedSnapshot
    };
  });
}
