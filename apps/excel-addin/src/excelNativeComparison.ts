import capabilities from "../../../config/capabilities.json";
import type { AnalysisPlan, CellValue, ExcelAction } from "./contracts";
import { assertRectangularValues, parseExplicitA1Rect, writeMatrixShapeIssue } from "./planCore";
import { assertSafeFormula } from "./formulaSafety";
import { captureWorkbookSourceFingerprint, captureWorkbookStructure, currentOfficeWorkbookIdentity,
  isToolSchemaFingerprint, missingSchemaRequirements, planLiveness, toolSchemaFingerprintForSnapshot, valuesEqual } from "./excel";
import { nativeColumnName, nativeReviewSeed } from "./nativeReviewFile";
import { buildManualReviewFile } from "./manualReviewFile";
import { prepareLocalExcelWindows, arrangeLocalExcelWindows } from "./excelWindowLayout";
import { ExcelReviewLayoutVersionError } from "./api";

type RegionAction = Extract<ExcelAction, { type: "writeValues" | "writeFormulas" | "clearRange" }>;
export type SheetSnapshot = {
  id: string; name: string; address: string;
  startRow: number; startColumn: number;
  values: CellValue[][]; formulas: CellValue[][]; numberFormat: string[][];
  properties: Excel.CellProperties[][];
  formulaCells: string[];
  rowHeights: number[]; columnWidths: number[];
};
export type ExcelNativeComparison = {
  mode?: "native" | "manual";
  windowLayout?: "arranged" | "manual" | "service_outdated";
  synchronousScrolling?: "enabled" | "different_instance" | "unavailable";
  plan: AnalysisPlan; authorization: readonly string[]; workbookIdentity: string;
  original: SheetSnapshot; originalWindowName: string;
  copyWindowName: string; copySheetName: string; copySheetId: string;
  copyEvidence: string; actionEvidence: string; changedCells: number;
};
const error = (message: string) => new Error(message);
const LIMITS = capabilities.excelNativeComparison;

export function nativeComparisonPlanIssue(plan: AnalysisPlan, authorization: readonly string[]): string | null {
  if (authorization.length !== 1 || !authorization[0] || !plan.actions.length) return "原生并排审阅目前只支持单张已授权工作表。";
  if (plan.sourceFingerprintSheets?.some(name => name !== authorization[0])) return "计划来源超出当前单表授权，不能生成副本。";
  let cells = 0;
  const rectangles: NonNullable<ReturnType<typeof parseExplicitA1Rect>>[] = [];
  for (const action of plan.actions) {
    if (action.sheet !== authorization[0]) return "计划包含其他工作表，不能扩大授权生成副本。";
    if (!["writeValues", "writeFormulas", "clearRange"].includes(action.type)) return "此计划含尚不支持原生副本审阅的动作；未写入原表。";
    const a = action as RegionAction;
    const rect = parseExplicitA1Rect(a.range);
    if (!rect) return "原生审阅需要明确的 A1 矩形区域。";
    if (rectangles.some(other => rect.startRow <= other.endRow && rect.endRow >= other.startRow && rect.startColumn <= other.endColumn && rect.endColumn >= other.startColumn)) return "计划含重叠写入区域，请先整理为互不重叠的动作。";
    rectangles.push(rect);
    cells += rect.rows * rect.columns;
    if (cells > LIMITS.maxCells) return "计划写入范围超过原生审阅限额。";
    if (a.type === "clearRange" && a.applyTo !== "contents") return "原生审阅只支持清空内容，不支持清除格式。";
    if (a.type === "writeValues" || a.type === "writeFormulas") {
      if (a.type === "writeFormulas" && a.formulaR1C1) return "原生审阅尚不支持 R1C1 铺满公式。";
      const matrix = a.type === "writeValues" ? a.values : a.formulas;
      if (!matrix) return "计划缺少完整写入矩阵。";
      try { assertRectangularValues(matrix); } catch { return "计划写入矩阵不是规则二维数组。"; }
      const issue = writeMatrixShapeIssue(a.range, matrix, rect);
      if (issue) return issue;
      for (const value of matrix.flat()) {
        if (a.type === "writeValues" && typeof value === "string" && /^[=+\-@]/.test(value)) return "写值计划包含可能被解释为公式的文本，请重新生成安全计划。";
        if (a.type === "writeFormulas" && typeof value === "string") {
          if (!value.startsWith("=")) return "公式计划必须提供真实公式，不混入普通文本。";
          try { assertLocalFormula(value); } catch { return "公式含外部、跨表或不安全引用，不能在隔离副本试算。"; }
        }
        if (a.type === "writeFormulas" && typeof value !== "string") return "公式矩阵只能含公式文本。";
      }
    }
  }
  return null;
}
function assertLocalFormula(formula: string): void {
  assertSafeFormula(formula, "副本公式");
  // This initial adapter has no other sheets, workbook names or table names.
  // Reject qualified/structured/indirect references instead of importing unauthorized dependencies.
  const body = formula.replace(/"(?:[^"]|"")*"/g, '""');
  if (/[!\[\]]/.test(body) || /\b(?:INDIRECT|OFFSET)\s*\(/i.test(body)) throw error("副本公式需要未复制的引用，不能隔离试算。");
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
async function capture(context: Excel.RequestContext, name: string): Promise<SheetSnapshot> {
  const sheet = context.workbook.worksheets.getItem(name);
  sheet.load("id,name"); sheet.protection.load("protected");
  const used = sheet.getUsedRangeOrNullObject(); used.load("isNullObject,rowIndex,columnIndex,rowCount,columnCount");
  await context.sync();
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
    throw error("此表含表格、合并、图表、名称或其他复杂对象，当前不能完整复制；未写入原表。");
  }
  used.load("values,formulas,numberFormat");
  if (!formulaAreas.isNullObject) formulaAreas.areas.load("items/rowIndex,items/columnIndex,items/rowCount,items/columnCount");
  const properties = used.getCellProperties({ format: {
    fill: { color: true }, font: { name: true, size: true, bold: true, italic: true, color: true, underline: true },
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
async function actionEvidence(context: Excel.RequestContext, sheet: Excel.Worksheet, plan: AnalysisPlan): Promise<string> {
  const ranges = (plan.actions as RegionAction[]).map(action => sheet.getRange(action.range));
  ranges.forEach(range => range.load("values,formulas")); await context.sync();
  return JSON.stringify(ranges.map(range => ({ values: range.values, formulas: range.formulas })));
}
async function evidence(context: Excel.RequestContext, sheet: Excel.Worksheet): Promise<string> {
  const range = sheet.getUsedRangeOrNullObject(); range.load("isNullObject,rowIndex,columnIndex,rowCount,columnCount"); await context.sync();
  if (range.isNullObject || range.rowCount * range.columnCount > LIMITS.maxCells) throw error("副本范围已变化，请重新预览。");
  range.load("values,formulas,numberFormat"); await context.sync();
  return JSON.stringify({ address: exactAddress(range), values: range.values, formulas: range.formulas, numberFormat: range.numberFormat });
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
    for (const action of confirmedPlan.actions as RegionAction[]) {
      const range = sheet.getRange(action.range);
      if (action.type === "writeValues") range.values = action.values;
      else if (action.type === "writeFormulas") range.formulas = action.formulas!;
      else range.clear(Excel.ClearApplyTo.contents);
    }
    sheet.calculate(true);
    await context.sync();
    const observed = JSON.parse(await actionEvidence(context, sheet, confirmedPlan)) as Array<{ values: CellValue[][]; formulas: CellValue[][] }>;
    let changedCells = 0;
    for (const [index, action] of (confirmedPlan.actions as RegionAction[]).entries()) {
      const actual = observed[index];
      if (action.type === "writeValues" && (actual.values.length !== action.values.length || actual.values.some((row, r) => row.length !== action.values[r].length || row.some((value, c) => !valuesEqual(value, action.values[r][c]))))) throw error("副本写值回读不符，未开放写回原表。");
      if (action.type === "writeFormulas" && JSON.stringify(actual.formulas) !== JSON.stringify(action.formulas)) throw error("副本公式回读不符，未开放写回原表。");
      if (action.type === "clearRange" && actual.values.flat().some(value => value !== null && value !== "")) throw error("副本清空回读不符，未开放写回原表。");
      const rect = parseExplicitA1Rect(action.range)!;
      for (let r = 0; r < rect.rows; r++) for (let c = 0; c < rect.columns; c++) {
        const sr = rect.startRow + r - source.startRow, sc = rect.startColumn + c - source.startColumn;
        const same = (a: CellValue | undefined, b: CellValue | undefined) => (a ?? "") === (b ?? "");
        if (!same(actual.values[r][c], source.values[sr]?.[sc]) || !same(actual.formulas[r][c], source.formulas[sr]?.[sc])) {
          sheet.getRange(`${nativeColumnName(rect.startColumn + c)}${rect.startRow + r}`).format.fill.color = "#E3F1E8";
          changedCells++;
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
      copySheetId: sheet.id, copyEvidence: await evidence(context, sheet), actionEvidence: JSON.stringify(observed), changedCells };
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
      const actual = JSON.parse(await actionEvidence(context, original, session.plan)) as Array<{ values: CellValue[][]; formulas: CellValue[][] }>;
      const expected = JSON.parse(session.actionEvidence) as typeof actual;
      const same = (a: CellValue[][], b: CellValue[][]) => a.length === b.length && a.every((row, r) => row.length === b[r].length && row.every((value, c) => valuesEqual(value, b[r][c])));
      return actual.length === expected.length && actual.every((range, i) => same(range.values, expected[i].values) && same(range.formulas, expected[i].formulas)) ? "matched" : "different";
    });
  } catch { return "unavailable"; }
}
export async function closeExcelNativeComparison(session: ExcelNativeComparison): Promise<boolean> {
  if (session.mode === "manual") return false; // No API handle: never claim that the window was closed.
  return Excel.run(async context => {
    if (!(await windows(context)).some(window => window.name === session.copyWindowName)) return true;
    const right = await copyWindow(context, session);
    const result = right.close(); await context.sync();
    // Excel may prompt to save the temporary workbook. Cancellation is NOT a successful close.
    return result.value;
  });
}
