import capabilities from "../../../config/capabilities.json";
import type { CellValue } from "../../excel-addin/src/contracts";
import type { RangeChangeReview } from "@excel-bro/core/changeReview";
import type { WpsApplication, WpsRange, WpsWorkbook, WpsWorksheet } from "./wpsHost";
import { normalizeWpsMatrix, workbookFullName } from "./wpsWorkbookReader";
import {
  assertWpsPreparedPlanUnchanged,
  matrixFingerprintV2,
  prepareWpsPlan,
  wpsActionChangeReview,
  type StoredAction,
  type WpsActionAfter,
  type WpsPreparedAction,
  type WpsPreparedPlan
} from "./wpsPlanExecutor";

export class WpsFormulaTrialError extends Error {}

/** A disposable, unsaved workbook. Only the original plan can later be committed. */
export type WpsNativeComparison = {
  original: WpsWorkbook;
  copy: WpsWorkbook;
  copyIdentity: WpsComparisonWorkbookIdentity;
  plan: WpsPreparedPlan;
  reviews: RangeChangeReview[];
  changedCells: string[];
  originalBefore: WpsActionAfter[];
  copyFingerprints: string[];
  copyAfter: WpsActionAfter[];
};

type WpsComparisonWorkbookIdentity = {
  name: string;
  fullName: string | null;
  path: string | null;
};

function comparisonWorkbookIdentity(workbook: WpsWorkbook): WpsComparisonWorkbookIdentity {
  return {
    name: String(workbook.Name ?? ""),
    fullName: workbookFullName(workbook),
    path: String(workbook.Path ?? "").trim() || null
  };
}

function sameWorkbookIdentity(left: WpsComparisonWorkbookIdentity, right: WpsComparisonWorkbookIdentity): boolean {
  return left.name === right.name && left.fullName === right.fullName && left.path === right.path;
}

function columnName(index: number): string {
  let result = "";
  for (let value = index; value > 0; value = Math.floor((value - 1) / 26)) {
    result = String.fromCharCode(65 + (value - 1) % 26) + result;
  }
  return result;
}

function target(sheet: WpsWorksheet, address: string): WpsRange {
  const range = sheet.Range?.(address);
  if (!range || range.Formula === undefined) {
    throw new Error(`副本无法读取「${sheet.Name}!${address}」的公式与单元格值。`);
  }
  return range;
}

function snapshot(range: WpsRange, action: WpsPreparedAction): WpsActionAfter & { fingerprint: string } {
  const values = normalizeWpsMatrix(range.Value2, action.rowCount, action.columnCount);
  const formulas = normalizeWpsMatrix(range.Formula, action.rowCount, action.columnCount);
  return {
    values,
    formulas,
    valuesCaptured: range.Value2 !== undefined && range.Value2 !== null,
    fingerprint: matrixFingerprintV2(values, formulas)
  };
}

function logicalCell(values: CellValue[][], formulas: CellValue[][], row: number, column: number): string {
  const formula = formulas[row]?.[column];
  const value = typeof formula === "string" && formula.startsWith("=")
    ? ["formula", formula]
    : [typeof values[row]?.[column], values[row]?.[column]];
  return JSON.stringify(value);
}

export function changedWpsCellAddresses(action: StoredAction, after: WpsActionAfter): string[] {
  const addresses: string[] = [];
  for (let row = 0; row < action.rowCount; row += 1) {
    for (let column = 0; column < action.columnCount; column += 1) {
      if (logicalCell(action.beforeValues, action.beforeFormulas, row, column) !==
          logicalCell(after.values, after.formulas, row, column)) {
        addresses.push(`${columnName(action.rect.startColumn + column)}${action.rect.startRow + row}`);
      }
    }
  }
  return addresses;
}

function chunkAddress(action: WpsPreparedAction, offset: number, rows: number): string {
  const start = `${columnName(action.rect.startColumn)}${action.rect.startRow + offset}`;
  const end = `${columnName(action.rect.endColumn)}${action.rect.startRow + offset + rows - 1}`;
  return start === end ? start : `${start}:${end}`;
}

function assertOriginalUnchanged(application: WpsApplication, prepared: WpsPreparedPlan): WpsWorkbook {
  assertWpsPreparedPlanUnchanged(prepared);
  const expected: WpsComparisonWorkbookIdentity = {
    name: prepared.authorization.workbookName,
    fullName: prepared.authorization.workbookFullName,
    path: prepared.authorization.workbookPath
  };
  const active = application.ActiveWorkbook;
  // WPS may return a fresh JSAPI wrapper for the same workbook on every access.
  // A saved workbook is bound by its path/name and the plan's source/target snapshots.
  if (!active || !expected.path || !sameWorkbookIdentity(comparisonWorkbookIdentity(active), expected)) {
    throw new Error("当前活动工作簿已不是预览时的原表，未复制或写入。请重新预览。");
  }
  const indexes = prepared.authorization.sheets.map((sheet) => sheet.index);
  const checked = prepareWpsPlan(application, indexes, prepared.plan, prepared.sourceSnapshots);
  if (checked.workbookIdentity !== prepared.workbookIdentity ||
      JSON.stringify(checked.actions) !== JSON.stringify(prepared.actions) ||
      JSON.stringify(checked.sources) !== JSON.stringify(prepared.sources)) {
    throw new Error("原表或授权范围在预览后发生变化，请重新预览并确认。");
  }
  const confirmed = application.ActiveWorkbook;
  if (!confirmed || !sameWorkbookIdentity(comparisonWorkbookIdentity(confirmed), expected)) {
    throw new Error("当前活动工作簿已不是预览时的原表，未复制或写入。请重新预览。");
  }
  return confirmed;
}

function closeCopyWithoutSaving(copy: WpsWorkbook): void {
  if (typeof copy.Close !== "function" || copy.Close(false) === false) {
    throw new Error("无法自动关闭审阅副本；原表未因此写入。请手动关闭副本且不要保存。");
  }
}

/**
 * The native split view is deliberately limited to one authorized sheet. Copying the whole
 * workbook would silently include sheets the user did not authorize. No file is saved here.
 */
export function stageWpsNativeComparison(
  application: WpsApplication,
  prepared: WpsPreparedPlan
): WpsNativeComparison {
  if (prepared.authorization.sheets.length !== 1 ||
      prepared.actions.some((action) => action.sheetName !== prepared.authorization.sheets[0].name) ||
      prepared.sources.some((source) => source.sheetName !== prepared.authorization.sheets[0].name)) {
    throw new Error("原生并排审阅当前只支持一张已授权工作表；跨表计划不会复制整本工作簿。");
  }
  const original = assertOriginalUnchanged(application, prepared);
  const originalSheet = original.Worksheets.Item(prepared.authorization.sheets[0].index);
  const originalBefore = prepared.actions.map((action) => snapshot(target(originalSheet, action.range), action));
  if (originalBefore.some((item) => !item.valuesCaptured)) {
    throw new Error("宿主未提供原表目标区域的显示值，无法建立可靠的并排审阅基线；原表未写入。");
  }
  if (typeof originalSheet.Copy !== "function") throw new Error("当前 WPS 宿主未提供工作表副本接口，原表未写入。");
  if (typeof original.Activate !== "function" || typeof application.Windows?.CompareSideBySideWith !== "function") {
    throw new Error("当前 WPS 宿主未提供原生并排比较接口，原表未写入。");
  }
  originalSheet.Copy();
  const copy = application.ActiveWorkbook;
  const copyIdentity = copy ? comparisonWorkbookIdentity(copy) : null;
  if (!copy || !copyIdentity || !copyIdentity.name || copyIdentity.path ||
      copyIdentity.name === prepared.authorization.workbookName ||
      copyIdentity.fullName === prepared.authorization.workbookFullName || copy.Worksheets.Count !== 1 ||
      copy.Worksheets.Item(1).Name !== originalSheet.Name) {
    throw new Error("WPS 未能确认新建了独立的单表副本；为安全起见未执行写入。请检查是否留下未保存副本。");
  }

  try {
    const sheet = copy.Worksheets.Item(1);
    const reviews: RangeChangeReview[] = [];
    const changedCells: string[] = [];
    const copyFingerprints: string[] = [];
    const copyAfter: WpsActionAfter[] = [];
    for (const [index, action] of prepared.actions.entries()) {
      if (sheet.ProtectContents !== false) throw new Error("审阅副本的工作表受保护或状态未知。");
      const range = target(sheet, action.range);
      if (range.MergeCells !== false) throw new Error("审阅副本的目标区域包含合并单元格或状态未知。");
      const before = snapshot(range, action);
      if (!before.valuesCaptured || before.fingerprint !== action.previewFingerprint ||
          JSON.stringify(before.values) !== JSON.stringify(originalBefore[index].values)) {
        throw new Error("审阅副本与原表的预览快照不同，不能据此生成可靠对照。原表未写入。");
      }
      const stored: StoredAction = {
        ...action,
        beforeValues: before.values,
        beforeFormulas: before.formulas,
        beforeFingerprint: before.fingerprint,
        afterFingerprint: null,
        appliedRows: 0
      };
      // Mirror the production executor's bounded write shape; the copy is disposable but
      // the WPS bridge still has the same per-call limits.
      for (let offset = 0; offset < action.rowCount; offset += capabilities.wpsWrite.chunkRows) {
        const rows = Math.min(capabilities.wpsWrite.chunkRows, action.rowCount - offset);
        const chunkValues = action.values.slice(offset, offset + rows);
        const chunk = target(sheet, chunkAddress(action, offset, rows));
        try {
          chunk[action.type === "writeFormulas" ? "Formula" : "Value2"] =
            rows === 1 && action.columnCount === 1 ? chunkValues[0][0] : chunkValues;
        } catch (error) {
          if (action.type === "writeFormulas") {
            throw new WpsFormulaTrialError("审阅副本无法试算该公式；原表未写入。", { cause: error });
          }
          throw error;
        }
      }
      stored.appliedRows = action.rowCount;
      let after: WpsActionAfter & { fingerprint: string };
      try {
        after = snapshot(target(sheet, action.range), action);
      } catch (error) {
        if (action.type === "writeFormulas") {
          throw new WpsFormulaTrialError("审阅副本无法回读公式计算值；原表未写入。", { cause: error });
        }
        throw error;
      }
      if (!after.valuesCaptured) {
        if (action.type === "writeFormulas") throw new WpsFormulaTrialError("审阅副本的公式计算值未能回读；原表未写入。");
        throw new Error("宿主未提供审阅副本写后的显示值，不能完成可靠对照；原表未写入。");
      }
      const observed = action.type === "writeFormulas" ? after.formulas : after.values;
      const matched = JSON.stringify(observed) === JSON.stringify(action.values);
      reviews.push(wpsActionChangeReview(stored, after, matched));
      if (!matched) {
        if (action.type === "writeFormulas") throw new WpsFormulaTrialError("审阅副本的公式文本回读不符；原表未写入。");
        throw new Error(`审阅副本「${action.sheetName}!${action.range}」回读与计划不一致；原表未写入。`);
      }
      const expectedValues = prepared.plan.acceptanceCriteria?.find((criterion) =>
        criterion.type === "rangeEquals" && criterion.sheet === action.sheetName && criterion.range === action.range
      );
      if (action.type === "writeFormulas" && prepared.plan.id.startsWith("wps-map-column-") && !expectedValues) {
        throw new Error("本地映射公式缺少计算值验收条件，原表未写入。");
      }
      if (action.type === "writeFormulas" && expectedValues?.type === "rangeEquals" &&
          JSON.stringify(after.values) !== JSON.stringify(expectedValues.expected)) {
        throw new WpsFormulaTrialError("审阅副本的公式计算值与确定性映射结果不同；原表未写入。");
      }
      copyFingerprints.push(after.fingerprint);
      copyAfter.push(after);
      changedCells.push(...changedWpsCellAddresses(stored, after));
    }
    // Color is applied only to the disposable right-hand workbook, never to the original.
    for (const address of changedCells) {
      const cell = sheet.Range?.(address);
      if (!cell?.Interior) throw new Error("WPS 未提供副本单元格底色接口，无法完整标记差异；原表未写入。");
      cell.Interior.Color = 0xE8F5EA;
    }
    original.Activate();
    assertOriginalUnchanged(application, prepared);
    const copyWindowName = String(copy.Windows?.Item(1)?.Caption ?? copy.Name);
    if (application.Windows!.CompareSideBySideWith!(copyWindowName) !== true) {
      throw new Error("WPS 未能打开原生并排比较；原表未写入。");
    }
    if (application.Windows) application.Windows.SyncScrollingSideBySide = true;
    return { original, copy, copyIdentity, plan: prepared, reviews, changedCells, originalBefore, copyFingerprints, copyAfter };
  } catch (error) {
    let cleanupFailed = false;
    try { closeCopyWithoutSaving(copy); } catch { cleanupFailed = true; }
    try { original.Activate(); } catch { /* The original workbook is never written in this path. */ }
    if (cleanupFailed) {
      throw new Error(`${error instanceof Error ? error.message : "原生并排审阅失败。"} 审阅副本未能自动关闭；请手动关闭且不要保存。原表未写入。`);
    }
    throw error;
  }
}

export function verifyWpsNativeComparison(
  application: WpsApplication,
  staged: WpsNativeComparison
): void {
  if (typeof staged.copy.Activate !== "function") throw new Error("无法激活审阅副本，已停止写回。");
  staged.copy.Activate();
  const activeCopy = application.ActiveWorkbook;
  if (!activeCopy || !sameWorkbookIdentity(comparisonWorkbookIdentity(activeCopy), staged.copyIdentity) ||
      activeCopy.Worksheets.Count !== 1) {
    throw new Error("审阅副本身份已变化，已停止写回。");
  }
  const sheet = activeCopy.Worksheets.Item(1);
  if (sheet.Name !== staged.plan.authorization.sheets[0].name) throw new Error("审阅副本工作表已变化，已停止写回。");
  for (const [index, action] of staged.plan.actions.entries()) {
    const current = snapshot(target(sheet, action.range), action);
    if (!current.valuesCaptured || current.fingerprint !== staged.copyFingerprints[index] ||
        JSON.stringify(current.values) !== JSON.stringify(staged.copyAfter[index].values)) {
      throw new Error("审阅副本在生成对照后发生变化，请重新生成预览；不会把未经确认的内容写回原表。");
    }
  }
  if (typeof staged.original.Activate !== "function") throw new Error("无法激活原表，已停止写回。");
  staged.original.Activate();
  const activeOriginal = assertOriginalUnchanged(application, staged.plan);
  const originalSheet = activeOriginal.Worksheets.Item(staged.plan.authorization.sheets[0].index);
  for (const [index, action] of staged.plan.actions.entries()) {
    const current = snapshot(target(originalSheet, action.range), action);
    if (!current.valuesCaptured ||
        JSON.stringify(current.values) !== JSON.stringify(staged.originalBefore[index].values) ||
        JSON.stringify(current.formulas) !== JSON.stringify(staged.originalBefore[index].formulas)) {
      throw new Error("原表在并排审阅后发生变化，请重新预览；不会覆盖新状态。");
    }
  }
}

export function discardWpsNativeComparison(staged: WpsNativeComparison): void {
  closeCopyWithoutSaving(staged.copy);
  staged.original.Activate?.();
}

/** Formula text can match while calculated values differ between a one-sheet copy and the original. */
export function compareCommittedWpsResultToCopy(staged: WpsNativeComparison): "matched" | "different" | "unavailable" {
  const sheet = staged.original.Worksheets.Item(staged.plan.authorization.sheets[0].index);
  let unavailable = false;
  for (const [index, action] of staged.plan.actions.entries()) {
    const actual = snapshot(target(sheet, action.range), action);
    const copied = staged.copyAfter[index];
    if (!actual.valuesCaptured || !copied.valuesCaptured) unavailable = true;
    if (JSON.stringify(actual.formulas) !== JSON.stringify(copied.formulas)) return "different";
    if (actual.valuesCaptured && copied.valuesCaptured &&
        JSON.stringify(actual.values) !== JSON.stringify(copied.values)) return "different";
  }
  return unavailable ? "unavailable" : "matched";
}
