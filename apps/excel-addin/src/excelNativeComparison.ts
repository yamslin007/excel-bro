import capabilities from "../../../config/capabilities.json";
import type { AnalysisPlan, CellValue, ExcelAction } from "./contracts";
import { assertRectangularValues, parseExplicitA1Rect, writeMatrixShapeIssue } from "./planCore";
import { assertSafeFormula } from "./formulaSafety";
import { applyFormatAction, captureWorkbookSourceFingerprint, captureWorkbookStructure, currentOfficeWorkbookIdentity,
  DIRECT_FORMAT_LOAD_OPTIONS, directFormatFieldValue,
  isToolSchemaFingerprint, missingSchemaRequirements, planLiveness, toolSchemaFingerprintForSnapshot, valuesEqual } from "./excel";
import { excelFormatActionFields, expectedDirectFormat, formatDirectValueMatches } from "./excelFormatCapability";
import { nativeColumnName, nativeReviewSeed } from "./nativeReviewFile";
import { buildManualReviewFile } from "./manualReviewFile";
import { prepareLocalExcelWindows, arrangeLocalExcelWindows, closeLocalExcelWindows } from "./excelWindowLayout";
import { ExcelReviewLayoutVersionError } from "./api";
import type { ReviewCopyCloseOutcome, ReviewCopyContentVerdict } from "@excel-bro/core/reviewCopyLifecycle";
import { describeCopyDiscardSafety } from "@excel-bro/core/reviewCopyLifecycle";

type RegionAction = Extract<ExcelAction, { type: "writeValues" | "writeFormulas" | "clearRange" }>;
/**
 * 原生副本审阅支持的**格式动作**（`setFont`/`setAlignment`/`resizeRange` 等六个）。
 *
 * 之前这里一律拒绝，导致"加粗 / 水平居中 / 行高"这类计划根本进不了副本审阅 —— 用户只能在
 * 确认后直接写原表，看不到试算效果。现在格式动作走与执行器**同源**的
 * `applyFormatAction`，并在副本上回读核对。
 */
export type NativeReviewFormatAction = Extract<
  ExcelAction,
  { type: "setFill" | "setFont" | "setBorders" | "setAlignment" | "setNumberFormat" | "resizeRange" }
>;
type PlanAction = RegionAction | NativeReviewFormatAction;
const NATIVE_FORMAT_ACTION_TYPES = new Set<string>([
  "setFill",
  "setFont",
  "setBorders",
  "setAlignment",
  "setNumberFormat",
  "resizeRange"
]);
export function isNativeReviewFormatAction(action: ExcelAction): action is NativeReviewFormatAction {
  return NATIVE_FORMAT_ACTION_TYPES.has(action.type);
}
export type SheetSnapshot = {
  id: string; name: string; address: string;
  startRow: number; startColumn: number;
  values: CellValue[][]; formulas: CellValue[][]; numberFormat: string[][];
  properties: Excel.CellProperties[][];
  formulaCells: string[];
  rowHeights: number[]; columnWidths: number[];
};
/** 副本某动作范围的**执行后状态**：值/公式 + 该动作实际修改属性的逐格直接格式。 */
type ActionSnapshot = {
  values: CellValue[][];
  formulas: CellValue[][];
  /** 仅格式动作有：逐格"直接格式"在受影响属性上的取值（未直接设置 = null）。 */
  format?: Array<Array<unknown[]>>;
  rowHeights?: number[];
  columnWidths?: number[];
  numberFormat?: string[][];
};
export type ExcelNativeComparison = {
  recoveryCopyPath?: string | null;
  mode?: "native" | "manual";
  windowLayout?: "arranged" | "manual" | "service_outdated";
  synchronousScrolling?: "enabled" | "different_instance" | "unavailable";
  plan: AnalysisPlan; authorization: readonly string[]; workbookIdentity: string;
  original: SheetSnapshot; originalWindowName: string;
  comparisonNotice?: string;
  copyWindowName: string; copySheetName: string; copySheetId: string;
  copyEvidence: string; actionEvidence: string; changedCells: number;
  /**
   * 副本里**格式动作范围**的逐格直接格式摘要（无格式动作时为空串）。
   *
   * 与 `copyEvidence`（值/公式/数字格式）互补：`copyEvidence` 看不见"副本里的格式被改过"，
   * 而用户审阅格式动作时看的正是格式。缺失（旧会话） ⇒ 跳过这项比对。
   */
  copyFormatEvidence?: string;
};
const error = (message: string) => new Error(message);
// Copy fidelity is not write/undo safety. Only this explicit capability failure
// may degrade optional comparison; authorization and host failures still propagate.
class ComparisonUnavailableError extends Error {}
export type ExcelBeforeResult = SheetSnapshot | { comparisonUnavailable: string };
const LIMITS = capabilities.excelNativeComparison;

export function nativeComparisonPlanIssue(plan: AnalysisPlan, authorization: readonly string[]): string | null {
  if (authorization.length !== 1 || !authorization[0] || !plan.actions.length) return "原生并排审阅目前只支持单张已授权工作表。";
  if (plan.sourceFingerprintSheets?.some(name => name !== authorization[0])) return "计划来源超出当前单表授权，不能生成副本。";
  let cells = 0;
  const rectangles: { rect: NonNullable<ReturnType<typeof parseExplicitA1Rect>>; isFormat: boolean }[] = [];
  for (const action of plan.actions) {
    if (action.sheet !== authorization[0]) return "计划包含其他工作表，不能扩大授权生成副本。";
    const isFormat = isNativeReviewFormatAction(action);
    if (!isFormat && !["writeValues", "writeFormulas", "clearRange"].includes(action.type)) return "此计划含尚不支持原生副本审阅的动作；未写入原表。";
    const a = action as PlanAction;
    const rect = parseExplicitA1Rect(a.range);
    if (!rect) return "原生审阅需要明确的 A1 矩形区域。";
    if (rectangles.some(({ rect: other, isFormat: otherFormat }) =>
      !(isFormat && otherFormat) && rect.startRow <= other.endRow && rect.endRow >= other.startRow &&
      rect.startColumn <= other.endColumn && rect.endColumn >= other.startColumn)) return "计划含重叠写入区域，请先整理为互不重叠的动作。";
    rectangles.push({ rect, isFormat });
    cells += rect.rows * rect.columns;
    if (cells > LIMITS.maxCells) return "计划写入范围超过原生审阅限额。";
    if (!isFormat) {
      const region = a as RegionAction;
      if (region.type === "clearRange" && region.applyTo !== "contents") return "原生审阅只支持清空内容，不支持清除格式。";
      if (region.type === "writeValues" || region.type === "writeFormulas") {
        if (region.type === "writeFormulas" && region.formulaR1C1) return "原生审阅尚不支持 R1C1 铺满公式。";
        const matrix = region.type === "writeValues" ? region.values : region.formulas;
        if (!matrix) return "计划缺少完整写入矩阵。";
        try { assertRectangularValues(matrix); } catch { return "计划写入矩阵不是规则二维数组。"; }
        const issue = writeMatrixShapeIssue(region.range, matrix, rect);
        if (issue) return issue;
        for (const value of matrix.flat()) {
          if (region.type === "writeValues" && typeof value === "string" && /^[=+\-@]/.test(value)) return "写值计划包含可能被解释为公式的文本，请重新生成安全计划。";
          if (region.type === "writeFormulas" && typeof value === "string") {
            if (!value.startsWith("=")) return "公式计划必须提供真实公式，不混入普通文本。";
            try { assertLocalFormula(value); } catch { return "公式含外部、跨表或不安全引用，不能在隔离副本试算。"; }
          }
          if (region.type === "writeFormulas" && typeof value !== "string") return "公式矩阵只能含公式文本。";
        }
      }
    } else if (excelFormatActionFields(action).length === 0) {
      // 空格式动作（一个属性都没给）= 什么也不做。计划校验已有这条规矩，这里复述一次：
      // 副本审阅不能为一个"看起来成功、实际没改"的动作开放写回。
      return "格式动作没有指定任何要修改的属性，请重新生成计划。";
    }
  }
  return null;
}
function assertLocalFormula(formula: string): void {
  assertSafeFormula(formula, "副本公式");
  // This initial adapter has no other sheets, workbook names or table names.
  // Reject qualified/structured/indirect references instead of importing unauthorized dependencies.
  const body = formula.replace(/"(?:[^"]|"")*"/g, '""');
  if (/[!\[\]]/.test(body) || /\b(?:INDIRECT|OFFSET)\s*\(/i.test(body)) throw new ComparisonUnavailableError("副本公式需要未复制的引用，不能隔离试算。");
}
export function excelNativeReviewMode(): "native" | "manual" | "unavailable" {
  if (typeof Office === "undefined" || typeof Excel === "undefined" ||
      typeof Excel.createWorkbook !== "function" ||
      !Office.context?.requirements?.isSetSupported("ExcelApi", "1.13")) return "unavailable";
  return Office.context.requirements.isSetSupported("ExcelApiDesktop", "1.1") ? "native" : "manual";
}
function exactAddress(range: Excel.Range): string {
  return `${nativeColumnName(range.columnIndex + 1)}${range.rowIndex + 1}:${nativeColumnName(range.columnIndex + range.columnCount)}${range.rowIndex + range.rowCount}`;
}
async function capture(context: Excel.RequestContext, name: string, allowEmpty = false): Promise<SheetSnapshot> {
  const sheet = context.workbook.worksheets.getItem(name);
  sheet.load("id,name"); sheet.protection.load("protected");
  let used = sheet.getUsedRangeOrNullObject(); used.load("isNullObject,rowIndex,columnIndex,rowCount,columnCount");
  await context.sync();
  if (used.isNullObject && allowEmpty) {
    // An empty output sheet has a valid blank before-image; still inspect its metadata.
    used = sheet.getRange("A1");
    used.load("isNullObject,rowIndex,columnIndex,rowCount,columnCount");
    await context.sync();
  }
  if (sheet.protection.protected) throw error("原表已保护，不能确认写回。");
  if (used.isNullObject || used.rowCount * used.columnCount > LIMITS.maxCells) throw error("工作表为空或已用范围超过原生审阅限额。");
  // Read metadata only for this authorized sheet. Complex objects are explicitly unsupported.
  const tables = sheet.tables.getCount(), charts = sheet.charts.getCount(), shapes = sheet.shapes.getCount();
  const pivots = sheet.pivotTables.getCount(), names = sheet.names.getCount(), bookNames = context.workbook.names.getCount();
  const conditions = used.conditionalFormats.getCount();
  const merges = used.getMergedAreasOrNullObject(); merges.load("isNullObject");
  const validations = used.getSpecialCellsOrNullObject(Excel.SpecialCellType.dataValidations); validations.load("isNullObject");
  const formulaAreas = used.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas); formulaAreas.load("isNullObject");
  await context.sync();
  if (tables.value || charts.value || shapes.value || pivots.value || names.value || bookNames.value || conditions.value || !merges.isNullObject || !validations.isNullObject) {
    throw new ComparisonUnavailableError("此表含表格、合并、图表、名称或其他复杂对象，当前不能完整复制；未写入原表。");
  }
  used.load("values,formulas,numberFormat");
  if (!formulaAreas.isNullObject) formulaAreas.areas.load("items/rowIndex,items/columnIndex,items/rowCount,items/columnCount");
  const properties = used.getCellProperties({ format: {
    // `pattern` 必须一起请求：只读 `color` 的话，"有填充"与"无填充"在直接格式里
    // 都退化成"color 有/无"，逐格比对 fill 时会误判（见 `directFormatFieldValue`）。
    fill: { color: true, pattern: true }, font: { name: true, size: true, bold: true, italic: true, color: true, underline: true },
    borders: { color: true, style: true, weight: true },
    horizontalAlignment: true, verticalAlignment: true, wrapText: true, indentLevel: true, shrinkToFit: true, textOrientation: true
  } });
  const rows = Array.from({ length: used.rowCount }, (_, i) => used.getRow(i).format);
  const columns = Array.from({ length: used.columnCount }, (_, i) => used.getColumn(i).format);
  rows.forEach(row => row.load("rowHeight")); columns.forEach(column => column.load("columnWidth"));
  await context.sync();
  const formulaCells: string[] = [];
  if (!formulaAreas.isNullObject) for (const area of formulaAreas.areas.items) {
    for (let r = 0; r < area.rowCount; r++) for (let c = 0; c < area.columnCount; c++) {
      const sr = area.rowIndex + r - used.rowIndex, sc = area.columnIndex + c - used.columnIndex;
      const formula = used.formulas[sr]?.[sc];
      if (typeof formula !== "string" || !formula.startsWith("=")) throw error("无法确认来源公式类型，未生成副本。");
      assertLocalFormula(formula);
      formulaCells.push(`${nativeColumnName(area.columnIndex + c + 1)}${area.rowIndex + r + 1}`);
    }
  }
  return JSON.parse(JSON.stringify({ id: sheet.id, name: sheet.name, address: exactAddress(used),
    startRow: used.rowIndex + 1, startColumn: used.columnIndex + 1,
    values: used.values, formulas: used.formulas, numberFormat: used.numberFormat,
    properties: properties.value, formulaCells, rowHeights: rows.map(row => row.rowHeight), columnWidths: columns.map(column => column.columnWidth)
  })) as SheetSnapshot;
}

/** A read-only pre-write image for optional comparison after execution. */
export function resultReviewPlanIssue(plan: AnalysisPlan, authorization: readonly string[]): string | null {
  const targets = [...new Set(plan.actions.map(action => action.sheet))];
  if (targets.length !== 1 || !targets[0] || !authorization.includes(targets[0])) {
    return '结果审阅需要单张已授权写入目标表。';
  }
  if (plan.sourceFingerprintSheets?.some(name => !authorization.includes(name)) ||
      plan.sourceDependencySheets?.some(name => !authorization.includes(name))) {
    return '计划来源超出本次授权，未执行。';
  }
  // Reuse action/shape limits only; source checks still use the full original plan.
  return nativeComparisonPlanIssue({ ...plan, sourceFingerprintSheets: [targets[0]] }, targets);
}

export async function captureExcelBeforeResult(plan: AnalysisPlan, authorization: readonly string[]): Promise<ExcelBeforeResult> {
  const issue = resultReviewPlanIssue(plan, authorization);
  if (issue) throw error(issue);
  const rects = plan.actions.map(action => parseExplicitA1Rect('range' in action ? action.range : '')!);
  const startRow = Math.min(...rects.map(r => r.startRow)), endRow = Math.max(...rects.map(r => r.endRow));
  const startColumn = Math.min(...rects.map(r => r.startColumn)), endColumn = Math.max(...rects.map(r => r.endColumn));
  const height = endRow - startRow + 1, width = endColumn - startColumn + 1;
  if (height * width > LIMITS.maxCells) return { comparisonUnavailable: '本次修改区域跨度超过双视图快照限额；请在主表审阅结果。' };
  return Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItem(plan.actions[0].sheet);
    sheet.load('id,name'); sheet.protection.load('protected');
    // Read only action rectangles. Gaps in the display canvas remain blank.
    const reads = rects.map(rect => {
      const range = sheet.getRange(`${nativeColumnName(rect.startColumn)}${rect.startRow}:${nativeColumnName(rect.endColumn)}${rect.endRow}`);
      range.load('values,numberFormat');
      const properties = range.getCellProperties({ format: DIRECT_FORMAT_LOAD_OPTIONS });
      const rows = Array.from({ length: rect.rows }, (_, i) => range.getRow(i).format);
      const columns = Array.from({ length: rect.columns }, (_, i) => range.getColumn(i).format);
      rows.forEach(row => row.load('rowHeight')); columns.forEach(column => column.load('columnWidth'));
      return { range, properties, rows, columns, rect };
    });
    await context.sync();
    if (sheet.protection.protected) throw error('原表已保护，不能执行。');
    const values: CellValue[][] = Array.from({ length: height }, () => Array(width).fill(''));
    const numberFormat: string[][] = Array.from({ length: height }, () => Array(width).fill('General'));
    const properties: Excel.CellProperties[][] = Array.from({ length: height }, () => Array.from({ length: width }, () => ({})));
    const rowHeights = Array(height).fill(15), columnWidths = Array(width).fill(8.43);
    for (const read of reads) {
      if (read.range.values.length !== read.rect.rows || read.properties.value.length !== read.rect.rows) throw error('执行前快照读取不完整。');
      for (let r = 0; r < read.rect.rows; r++) {
        if (read.range.values[r]?.length !== read.rect.columns || read.properties.value[r]?.length !== read.rect.columns) throw error('执行前快照读取不完整。');
        const row = read.rect.startRow - startRow + r;
        rowHeights[row] = read.rows[r].rowHeight;
        for (let c = 0; c < read.rect.columns; c++) {
          const column = read.rect.startColumn - startColumn + c;
          values[row][column] = read.range.values[r][c];
          numberFormat[row][column] = read.range.numberFormat[r][c];
          properties[row][column] = read.properties.value[r][c];
          columnWidths[column] = read.columns[c].columnWidth;
        }
      }
    }
    // No formulas are exported: cached values cannot acquire new dependencies or recalculate.
    return JSON.parse(JSON.stringify({ id: sheet.id, name: sheet.name,
      address: `${nativeColumnName(startColumn)}${startRow}:${nativeColumnName(endColumn)}${endRow}`,
      startRow, startColumn, values, formulas: values, formulaCells: [], numberFormat, properties, rowHeights, columnWidths
    })) as SheetSnapshot;
  });
}

/** Open the saved BEFORE image, never re-run the plan to manufacture a comparison. */
export async function showExcelBeforeResult(original: ExcelBeforeResult, identity: string, plan: AnalysisPlan,
  authorization: readonly string[]): Promise<ExcelNativeComparison> {
  if (currentOfficeWorkbookIdentity() !== identity) throw error('原工作簿已变化，未打开对比副本。');
  if ('comparisonUnavailable' in original) throw error(original.comparisonUnavailable);
  const copySheetName = `EB_${crypto.randomUUID().replace(/-/g, '').slice(0, 20)}`;
  const artifact = buildManualReviewFile(copySheetName, original, { ...plan, actions: [] }, LIMITS.maxCells, LIMITS.maxBytes);
  let ticket: Awaited<ReturnType<typeof prepareLocalExcelWindows>> | null = null;
  let comparisonNotice: string | undefined;
  const layoutFailure = (reason: unknown) => `执行前副本已打开，但未能自动并排：${reason instanceof Error ? reason.message : '本机窗口排列不可用'}。可手动排列窗口；确认变更和放弃变更仍可使用。`;
  try { ticket = await prepareLocalExcelWindows(identity, original.name, copySheetName); }
  catch (reason) { comparisonNotice = layoutFailure(reason); }
  if (currentOfficeWorkbookIdentity() !== identity) throw error('原工作簿已变化，未打开对比副本。');
  await Excel.createWorkbook(artifact.base64);
  const session: ExcelNativeComparison = { mode: 'manual', plan, authorization: [...authorization], workbookIdentity: identity,
    original, originalWindowName: '', copyWindowName: '', copySheetName, copySheetId: '',
    copyEvidence: '', actionEvidence: '', changedCells: 0, windowLayout: 'manual', comparisonNotice };
  if (ticket) {
    try {
      const layout = await arrangeLocalExcelWindows(ticket);
      session.windowLayout = 'arranged'; session.synchronousScrolling = layout.synchronousScrolling;
    } catch (reason) { session.comparisonNotice = layoutFailure(reason); }
  }
  return session;
}
async function windows(context: Excel.RequestContext): Promise<Excel.Window[]> {
  const collection = context.workbook.application.windows;
  collection.load("items/name"); await context.sync(); return collection.items;
}
async function copyWindow(context: Excel.RequestContext, session: ExcelNativeComparison): Promise<Excel.Window> {
  const found = (await windows(context)).filter(window => window.name === session.copyWindowName);
  if (found.length !== 1) throw error("无法确认审阅副本窗口仍存在，不能写回原表。");
  const sheet = found[0].activeWorksheet; sheet.load("id,name"); await context.sync();
  if (sheet.id !== session.copySheetId || sheet.name !== session.copySheetName) throw error("副本身份或活动表已变化，不能写回原表。");
  return found[0];
}
async function actionEvidence(context: Excel.RequestContext, sheet: Excel.Worksheet, plan: AnalysisPlan): Promise<ActionSnapshot[]> {
  const actions = plan.actions as PlanAction[];
  const ranges = actions.map(action => sheet.getRange(action.range));
  ranges.forEach(range => range.load("values,formulas")); await context.sync();
  // 格式动作额外读**逐格直接格式**（与执行器同口径，不用显示值）：1.9 起可用，
  // 因此这条证据在批量许可永久版（≤1.17）上同样成立。
  const formatReads = actions.map((action, index) =>
    isNativeReviewFormatAction(action)
      ? ranges[index].getCellProperties({ format: DIRECT_FORMAT_LOAD_OPTIONS })
      : null
  );
  const dimensionReads = actions.map((action, index) => {
    const rect = parseExplicitA1Rect(action.range)!;
    // Review compares effective sizes, including inherited standard dimensions.
    const rows = action.type === "resizeRange" && action.rowHeight != null
      ? Array.from({ length: rect.rows }, (_, r) => ranges[index].getRow(r).format) : null;
    const columns = action.type === "resizeRange" && action.columnWidth != null
      ? Array.from({ length: rect.columns }, (_, c) => ranges[index].getColumn(c).format) : null;
    rows?.forEach(format => format.load("rowHeight"));
    columns?.forEach(format => format.load("columnWidth"));
    return { rows, columns };
  });
  actions.forEach((action, index) => {
    if (action.type === "setNumberFormat") ranges[index].load("numberFormat");
  });
  await context.sync();
  return actions.map((action, index) => {
    const snapshot: ActionSnapshot = {
      values: ranges[index].values as CellValue[][],
      formulas: ranges[index].formulas as CellValue[][]
    };
    const read = formatReads[index];
    if (read) {
      const fields = excelFormatActionFields(action);
      snapshot.format = (read.value as Excel.CellProperties[][]).map(row =>
        row.map(cell =>
          fields.map(field => directFormatFieldValue(cell?.format, field))
        )
      );
    }
    const dimensions = dimensionReads[index];
    const rect = parseExplicitA1Rect(action.range)!;
    if (dimensions.rows) {
      snapshot.rowHeights = dimensions.rows.map(row => row.rowHeight);
      if (snapshot.rowHeights.length !== rect.rows || snapshot.rowHeights.some(value => !Number.isFinite(value))) {
        throw error("无法完整回读副本行高证据，未开放写回原表。");
      }
    }
    if (dimensions.columns) {
      snapshot.columnWidths = dimensions.columns.map(column => column.columnWidth);
      if (snapshot.columnWidths.length !== rect.columns || snapshot.columnWidths.some(value => !Number.isFinite(value))) {
        throw error("无法完整回读副本列宽证据，未开放写回原表。");
      }
    }
    if (action.type === "setNumberFormat") snapshot.numberFormat = ranges[index].numberFormat;
    return snapshot;
  });
}
async function evidence(context: Excel.RequestContext, sheet: Excel.Worksheet): Promise<string> {
  const range = sheet.getUsedRangeOrNullObject(); range.load("isNullObject,rowIndex,columnIndex,rowCount,columnCount"); await context.sync();
  if (range.isNullObject || range.rowCount * range.columnCount > LIMITS.maxCells) throw error("副本范围已变化，请重新预览。");
  range.load("values,formulas,numberFormat"); await context.sync();
  return JSON.stringify({ address: exactAddress(range), values: range.values, formulas: range.formulas, numberFormat: range.numberFormat });
}

/** 逐格直接格式的摘要（副本审阅期间用于发现"被改过"，不参与写回判据）。 */
async function formatEvidenceDigest(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  plan: AnalysisPlan
): Promise<string> {
  const actions = (plan.actions as PlanAction[]).filter(isNativeReviewFormatAction);
  if (actions.length === 0) return "";
  const snapshots = await actionEvidence(context, sheet, { ...plan, actions });
  return JSON.stringify(snapshots.map(({ format, rowHeights, columnWidths, numberFormat }) =>
    ({ format, rowHeights, columnWidths, numberFormat })));
}
/**
 * 副本里某个**格式动作**是否真的按计划生效（逐格直接格式回读）。
 *
 * 为什么用**直接格式**而不是显示值：这条判据要在批量许可永久版（ExcelApi ≤ 1.17，无
 * `getDisplayedCellProperties`）上同样成立，而格式动作改的就是"直接设置在单元格上的属性"。
 * 返回 `null` 表示一致，否则返回一句人话说明是哪个属性对不上。
 */
async function assertCopyFormatApplied(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  action: NativeReviewFormatAction
): Promise<string | null> {
  const expected = expectedDirectFormat(action);
  // 行高/列宽是**行级/列级**属性，不在逐格矩阵里；单独走逐行/逐列与数字格式回读。
  if (action.type === "resizeRange") {
    const range = sheet.getRange(action.range);
    const rows = action.rowHeight != null ? range.getRowProperties({ format: { rowHeight: true } }) : null;
    const columns = action.columnWidth != null ? range.getColumnProperties({ format: { columnWidth: true } }) : null;
    await context.sync();
    if (action.rowHeight != null) {
      const heights = (rows!.value as Array<{ format?: { rowHeight?: number } }>).map(item => item.format?.rowHeight);
      if (heights.length === 0 || heights.some(height => typeof height !== "number" || Math.abs(height - action.rowHeight!) > 0.1)) {
        return "行高与计划要求不一致";
      }
    }
    if (action.columnWidth != null) {
      const widths = (columns!.value as Array<{ format?: { columnWidth?: number } }>).map(item => item.format?.columnWidth);
      if (widths.length === 0 || widths.some(width => typeof width !== "number" || Math.abs(width - action.columnWidth!) > 0.1)) {
        return "列宽与计划要求不一致";
      }
    }
    return null;
  }
  if (action.type === "setNumberFormat") {
    const range = sheet.getRange(action.range);
    range.load("numberFormat");
    await context.sync();
    const seen = range.numberFormat as string[][];
    if (!seen.length || seen.some(row => row.some(code => code !== action.formatCode))) {
      return "数字格式与计划要求不一致";
    }
    return null;
  }
  if (expected.length === 0) return "格式动作没有指定任何属性";
  const range = sheet.getRange(action.range);
  const read = range.getCellProperties({ format: DIRECT_FORMAT_LOAD_OPTIONS });
  await context.sync();
  const rows = read.value as Excel.CellProperties[][];
  if (!Array.isArray(rows) || rows.length === 0) return "副本格式矩阵为空";
  for (const row of rows) {
    if (!Array.isArray(row) || row.length === 0) return "副本格式矩阵形状不符";
    for (const cell of row) {
      for (const { field, value } of expected) {
        const seen = directFormatFieldValue(cell?.format, field);
        if (!formatDirectValueMatches(field, seen, value)) {
          return `${field} 回读值与计划要求不一致`;
        }
      }
    }
  }
  return null;
}

export async function assertExcelNativeComparisonUnchanged(context: Excel.RequestContext, session: ExcelNativeComparison, authorization: readonly string[], plan: AnalysisPlan): Promise<void> {
  if (JSON.stringify(plan) !== JSON.stringify(session.plan) || JSON.stringify(authorization) !== JSON.stringify(session.authorization)) throw error("计划或授权范围已变化，请重新生成并排预览。");
  if (currentOfficeWorkbookIdentity() !== session.workbookIdentity) throw error("原工作簿身份已变化，不能写回。");
  if (JSON.stringify(await capture(context, session.original.name)) !== JSON.stringify(session.original)) throw error("原表在审阅期间已变化，请重新生成副本；未写入原表。");
  // Compatibility mode has no handle to the opened copy. User explicitly reviews
  // the generated artifact; only the immutable plan is committed, never copy edits.
  if (session.mode === "manual") return;
  const right = await copyWindow(context, session);
  if (await evidence(context, right.activeWorksheet) !== session.copyEvidence) throw error("副本数据或公式已变化，请重新生成审阅副本。");
  // 格式动作：副本里的**格式**也要没被改过。旧会话没有这项证据 ⇒ 跳过（不把旧会话判死）。
  if (session.copyFormatEvidence) {
    const current = await formatEvidenceDigest(context, right.activeWorksheet, session.plan);
    if (current !== session.copyFormatEvidence) throw error("副本中的格式已变化，请重新生成审阅副本。");
  }
}
export async function stageExcelNativeComparison(plan: AnalysisPlan, authorization: readonly string[]): Promise<ExcelNativeComparison> {
  const mode = excelNativeReviewMode();
  if (mode === "unavailable") throw error("当前 Excel 缺少副本创建或基础读取 API（ExcelApi 1.13）；未生成副本，未写入原表。");
  const issue = nativeComparisonPlanIssue(plan, authorization); if (issue) throw error(issue);
  const identity = currentOfficeWorkbookIdentity();
  if (!identity) throw error("请先保存原工作簿，再生成独立副本审阅。");
  // Snapshot plan and authorization before any await: edited plans cannot inherit this review.
  const confirmedPlan = JSON.parse(JSON.stringify(plan)) as AnalysisPlan;
  const scope = [...authorization];
  if (confirmedPlan.sourceFingerprint && confirmedPlan.sourceFingerprintSheets?.length) {
    if (isToolSchemaFingerprint(confirmedPlan.sourceFingerprint)) {
      if (toolSchemaFingerprintForSnapshot(confirmedPlan, await captureWorkbookStructure(confirmedPlan.sourceFingerprintSheets)) !== confirmedPlan.sourceFingerprint) throw error("工具来源结构已变化，请重新生成计划；未生成副本。");
    } else if (planLiveness(confirmedPlan) === "live") {
      if (missingSchemaRequirements(confirmedPlan, await captureWorkbookStructure(confirmedPlan.sourceFingerprintSheets)).length) throw error("公式所需来源结构已变化，请重新生成计划；未生成副本。");
    } else if (await captureWorkbookSourceFingerprint(confirmedPlan.sourceFingerprintSheets) !== confirmedPlan.sourceFingerprint) throw error("计划来源已变化，请重新生成计划；未生成副本。");
  }
  const before = await Excel.run(async context => {
    const original = await capture(context, scope[0]);
    if (mode === "manual") return { original, names: [] as string[], leftName: "" };
    const open = await windows(context);
    const left = context.workbook.application.activeWindow; left.load("name"); left.activeWorksheet.load("id"); await context.sync();
    if (left.activeWorksheet.id !== original.id) throw error("请先激活原工作表窗口，再生成并排副本。");
    return { original, names: open.map(window => window.name), leftName: left.name };
  });
  const sourceRect = parseExplicitA1Rect(before.original.address)!;
  const rects = [sourceRect, ...confirmedPlan.actions.map(action => parseExplicitA1Rect((action as RegionAction).range)!)];
  if ((Math.max(...rects.map(r => r.endRow)) - Math.min(...rects.map(r => r.startRow)) + 1) *
      (Math.max(...rects.map(r => r.endColumn)) - Math.min(...rects.map(r => r.startColumn)) + 1) > LIMITS.maxCells) throw error("副本合并显示范围超过限额，未打开副本。");
  const copySheetName = `EB_${crypto.randomUUID().replace(/-/g, "").slice(0, 20)}`;
  // The bridge owns both visible-frame layout and explicit native scrolling,
  // including hosts with Desktop 1.1. No bridge failure grants a write.
  let ticket: Awaited<ReturnType<typeof prepareLocalExcelWindows>> | undefined;
  let windowLayout: ExcelNativeComparison['windowLayout'] = 'manual';
  try { ticket = await prepareLocalExcelWindows(identity, before.original.name, copySheetName); }
  catch (cause) { if (cause instanceof ExcelReviewLayoutVersionError) windowLayout = 'service_outdated'; }
  if (mode === "manual") {
    const artifact = buildManualReviewFile(copySheetName, before.original, confirmedPlan, LIMITS.maxCells, LIMITS.maxBytes);
    const session: ExcelNativeComparison = { mode, plan: confirmedPlan, authorization: scope, workbookIdentity: identity,
      original: before.original, originalWindowName: "", copyWindowName: "", copySheetName,
      copySheetId: "", copyEvidence: "", actionEvidence: "", changedCells: artifact.markedCells };
    await Excel.run(context => assertExcelNativeComparisonUnchanged(context, session, scope, confirmedPlan));
    // Capture the original and pre-existing windows BEFORE opening the copy.
    // A layout failure affects positioning only, never permission to write.
    session.windowLayout = windowLayout;
    await Excel.run(context => assertExcelNativeComparisonUnchanged(context, session, scope, confirmedPlan));
    await Excel.createWorkbook(artifact.base64);
    await Excel.run(context => assertExcelNativeComparisonUnchanged(context, session, scope, confirmedPlan));
    if (ticket) {
      try {
        const layout = await arrangeLocalExcelWindows(ticket);
        session.windowLayout = 'arranged';
        session.synchronousScrolling = layout.synchronousScrolling;
      }
      catch (cause) { if (cause instanceof ExcelReviewLayoutVersionError) session.windowLayout = 'service_outdated'; }
    }
    await Excel.run(context => assertExcelNativeComparisonUnchanged(context, session, scope, confirmedPlan));
    return session;
  }
  const base64 = nativeReviewSeed(copySheetName, before.original.startRow, before.original.startColumn, before.original.values, LIMITS.maxBytes);
  await Excel.createWorkbook(base64);
  const session = await Excel.run(async context => {
    const candidates = (await windows(context)).filter(window => !before.names.includes(window.name));
    if (candidates.length !== 1) throw error("副本已请求打开，但无法唯一识别新窗口。请关闭新副本后重试；原表未写入。");
    const right = candidates[0], sheet = right.activeWorksheet;
    sheet.load("name,id"); await context.sync();
    if (sheet.name !== copySheetName) throw error("新窗口不是本次生成的副本，已停止；原表未写入。");
    const source = before.original;
    if (currentOfficeWorkbookIdentity() !== identity || JSON.stringify(await capture(context, source.name)) !== JSON.stringify(source)) throw error("原表在复制期间已变化，请关闭副本并重新预览。");
    // All following writes are to the uniquely identified NEW worksheet, never context.workbook.
    const copied = sheet.getRange(source.address);
    // Constants (including text beginning with '=') stay inline strings in the OOXML seed.
    // Never replay a whole formulas matrix: only host-confirmed formula cells are restored.
    for (const address of source.formulaCells) {
      const cell = parseExplicitA1Rect(address)!;
      sheet.getRange(address).formulas = [[source.formulas[cell.startRow - source.startRow][cell.startColumn - source.startColumn]]];
    }
    copied.numberFormat = source.numberFormat;
    copied.setCellProperties(source.properties);
    source.rowHeights.forEach((height, i) => { copied.getRow(i).format.rowHeight = height; });
    source.columnWidths.forEach((width, i) => { copied.getColumn(i).format.columnWidth = width; });
    for (const action of confirmedPlan.actions as PlanAction[]) {
      const range = sheet.getRange(action.range);
      if (action.type === "writeValues") range.values = action.values;
      else if (action.type === "writeFormulas") range.formulas = action.formulas!;
      else if (action.type === "clearRange") range.clear(Excel.ClearApplyTo.contents);
      // 格式动作与执行器**同源**（同一份 `applyFormatAction`），副本试算的样子
      // 必须与将来写回原表的样子逐字一致。
      else {
        // Validate each action immediately; later overlapping actions may intentionally
        // replace it. Review and commit evidence below describe the final plan state.
        await applyFormatAction(context, range, action);
        const mismatch = await assertCopyFormatApplied(context, sheet, action);
        if (mismatch) throw error(`${action.type} 在副本中的回读结果与计划不符（${mismatch}），未开放写回原表。`);
      }
    }
    sheet.calculate(true);
    await context.sync();
    const observed = await actionEvidence(context, sheet, confirmedPlan);
    const changedAddresses = new Set<string>();
    for (const [index, action] of (confirmedPlan.actions as PlanAction[]).entries()) {
      const actual = observed[index];
      if (action.type === "writeValues" && (actual.values.length !== action.values.length || actual.values.some((row, r) => row.length !== action.values[r].length || row.some((value, c) => !valuesEqual(value, action.values[r][c]))))) throw error("副本写值回读不符，未开放写回原表。");
      if (action.type === "writeFormulas" && JSON.stringify(actual.formulas) !== JSON.stringify(action.formulas)) throw error("副本公式回读不符，未开放写回原表。");
      if (action.type === "clearRange" && actual.values.flat().some(value => value !== null && value !== "")) throw error("副本清空回读不符，未开放写回原表。");
      const rect = parseExplicitA1Rect(action.range)!;
      for (let r = 0; r < rect.rows; r++) for (let c = 0; c < rect.columns; c++) {
        const sr = rect.startRow + r - source.startRow, sc = rect.startColumn + c - source.startColumn;
        const same = (a: CellValue | undefined, b: CellValue | undefined) => (a ?? "") === (b ?? "");
        const valuesDiffer =
          !same(actual.values[r][c], source.values[sr]?.[sc]) ||
          !same(actual.formulas[r][c], source.formulas[sr]?.[sc]);
        // 格式动作：逐格比**受影响属性上的直接格式**（未直接设置 = null），
        // 变化了的格才标记 —— 动作没造成实际变化的格不该被标成"已修改"。
        const formatMatrix = actual.format;
        const formatDiffers = formatMatrix
          ? excelFormatActionFields(action).some((field, fieldIndex) => {
              const before = directFormatFieldValue(source.properties[sr]?.[sc]?.format, field);
              const after = formatMatrix[r]?.[c]?.[fieldIndex];
              return JSON.stringify(before) !== JSON.stringify(after);
            })
          : false;
        const dimensionDiffers =
          (actual.rowHeights !== undefined && actual.rowHeights[r] !== source.rowHeights[sr]) ||
          (actual.columnWidths !== undefined && actual.columnWidths[c] !== source.columnWidths[sc]);
        const numberFormatDiffers = actual.numberFormat !== undefined && actual.numberFormat[r]?.[c] !== source.numberFormat[sr]?.[sc];
        if (valuesDiffer || formatDiffers || dimensionDiffers || numberFormatDiffers) {
          const address = `${nativeColumnName(rect.startColumn + c)}${rect.startRow + r}`;
          // A format preview must display the actual fill, including unchanged fill.
          // Painting a green marker would change the very result being reviewed.
          if (!isNativeReviewFormatAction(action)) sheet.getRange(address).format.fill.color = "#E3F1E8";
          changedAddresses.add(address);
        }
      }
    }
    const left = (await windows(context)).find(window => window.name === before.leftName);
    if (!left) throw error("原窗口已关闭，无法并排审阅；未写入原表。");
    left.activate();
    context.workbook.application.windows.compareCurrentSideBySideWith(right.name);
    context.workbook.application.windows.resetPositionsSideBySide();
    left.scrollRow = right.scrollRow = source.startRow;
    left.scrollColumn = right.scrollColumn = source.startColumn;
    await context.sync();
    const session: ExcelNativeComparison = { mode, plan: confirmedPlan, authorization: scope, workbookIdentity: identity,
      original: source, originalWindowName: before.leftName, copyWindowName: right.name, copySheetName,
      copySheetId: sheet.id,
      copyEvidence: await evidence(context, sheet),
      // 格式动作的副本逐格直接格式摘要一并进 `copyEvidence`：审阅期间副本格式被改动
      // （外部编辑、误操作）在写回前会被发现。
      copyFormatEvidence: await formatEvidenceDigest(context, sheet, confirmedPlan),
      actionEvidence: JSON.stringify(observed), changedCells: changedAddresses.size };
    await assertExcelNativeComparisonUnchanged(context, session, scope, confirmedPlan);
    return session;
  });
  session.windowLayout = windowLayout;
  if (ticket) {
    try {
      const layout = await arrangeLocalExcelWindows(ticket);
      session.windowLayout = 'arranged';
      session.synchronousScrolling = layout.synchronousScrolling;
    } catch (cause) {
      if (cause instanceof ExcelReviewLayoutVersionError) session.windowLayout = 'service_outdated';
    }
  }
  await Excel.run(context => assertExcelNativeComparisonUnchanged(context, session, scope, confirmedPlan));
  return session;
}
export async function compareExcelNativeCommittedResult(session: ExcelNativeComparison): Promise<"matched" | "different" | "unavailable"> {
  if (session.mode === "manual") return "unavailable";
  try {
    if (currentOfficeWorkbookIdentity() !== session.workbookIdentity) return "unavailable";
    return await Excel.run(async context => {
      const original = context.workbook.worksheets.getItem(session.original.name);
      original.load("id"); await context.sync();
      if (original.id !== session.original.id) return "unavailable";
      const actual = await actionEvidence(context, original, session.plan);
      const expected = JSON.parse(session.actionEvidence) as ActionSnapshot[];
      const same = (a: CellValue[][] | undefined, b: CellValue[][] | undefined) =>
        Array.isArray(a) && Array.isArray(b) && a.length === b.length &&
        a.every((row, r) => Array.isArray(row) && row.length === b[r].length && row.every((value, c) => valuesEqual(value, b[r][c])));
      // 格式动作：把写回后原表的**逐格直接格式**与副本里审阅过的样子逐字比。
      // 少了这一项，"格式写回成功"就只剩一句口头承诺。
      const sameFormat = (a: ActionSnapshot["format"], b: ActionSnapshot["format"]) =>
        a === undefined && b === undefined
          ? true
          : JSON.stringify(a) === JSON.stringify(b);
      return actual.length === expected.length && actual.every((range, i) =>
        same(range.values, expected[i].values) &&
        same(range.formulas, expected[i].formulas) &&
        sameFormat(range.format, expected[i].format) &&
        JSON.stringify(range.rowHeights) === JSON.stringify(expected[i].rowHeights) &&
        JSON.stringify(range.columnWidths) === JSON.stringify(expected[i].columnWidths) &&
        JSON.stringify(range.numberFormat) === JSON.stringify(expected[i].numberFormat)
      ) ? "matched" : "different";
    });
  } catch { return "unavailable"; }
}
/** v4 bridge archives the retained Workbook before closing it. No WM_CLOSE fallback. */
export async function closeCompatReviewCopy(session: ExcelNativeComparison): Promise<ReviewCopyCloseOutcome> {
  try {
    const result = await closeLocalExcelWindows(session.workbookIdentity, session.original.name, session.copySheetName);
    session.recoveryCopyPath = result.recoveryCopyPath;
    // Only claim a close when the receipt also says the copy is gone (the bridge polled
    // for the window to disappear; a still-visible window is `close_unconfirmed`).
    if (result.status === "closed" || result.status === "already_closed") return result.copyClosed ? result.status : "failed";
    return result.status;
  } catch {
    // 服务过旧或不可达：没有可核实的答复，一律降级为手动关闭，绝不宣称已关闭。
    return "unsupported";
  }
}

/** Close status is separate from writeback status. Native copies retain their content guard. */
export async function closeExcelNativeComparison(session: ExcelNativeComparison): Promise<ReviewCopyCloseOutcome> {
  if (session.mode === "manual") {
    return closeCompatReviewCopy(session);
  }
  /*
   * 第三项判断：**关掉之前**必须先证明"副本里没有额外的人工修改"。
   *
   * 副本身份可信（copyWindowName/copySheetId/copySheetName）与"没被另存"都成立时，
   * 用户仍然可以在并排审阅期间往副本里敲字 —— `Window.close()` 不会替我们判断这件事，
   * 所以不能拿前两项顶替。用独立一轮读取：**读不到 ⇒ `unknown`**，绝不当作"没改过"。
   * 只有逐格读回与本次写入后的证据完全一致，才允许继续关闭。
   */
  const content: ReviewCopyContentVerdict | null = await Excel.run(async (context): Promise<ReviewCopyContentVerdict | null> => {
    if (!(await windows(context)).some(window => window.name === session.copyWindowName)) return null;
    try {
      const copy = await copyWindow(context, session);
      const unchanged = (await evidence(context, copy.activeWorksheet) === session.copyEvidence) &&
        (!session.copyFormatEvidence ||
          (await formatEvidenceDigest(context, copy.activeWorksheet, session.plan)) === session.copyFormatEvidence);
      return unchanged ? "unchanged" : "changed";
    } catch { return "unknown"; }
  });
  if (content === null) return "already_closed";
  const blocked = describeCopyDiscardSafety(content);
  if (blocked) return blocked;
  return Excel.run(async context => {
    const copy = await copyWindow(context, session);
    const result = copy.close(); await context.sync();
    /*
     * `Window.close()` 返回 false 是**歧义**信号（被取消 / 被推迟 / 被拒），
     * 而且我们**没有**对话框证据 —— 所以只说"关闭尚未确认"，不推断宿主在问是否保存。
     */
    return result.value ? "closed" : "close_unconfirmed";
  });
}
