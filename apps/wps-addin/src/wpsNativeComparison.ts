import capabilities from "../../../config/capabilities.json";
import { readWpsRangeValues } from "./wpsRangeValues";
import { writeWpsMatrix } from "./wpsWriteMatrix";
import type { CellValue } from "../../excel-addin/src/contracts";
import type { RangeChangeReview } from "@excel-bro/core/changeReview";
// 审阅副本生命周期的**共享语义**：两端共用同一套关闭结果与文案，各自实现宿主适配。
import type { ReviewCopyCloseOutcome, ReviewCopyContentVerdict } from "@excel-bro/core/reviewCopyLifecycle";
import { describeCopyDiscardSafety } from "@excel-bro/core/reviewCopyLifecycle";
import type { WpsApplication, WpsRange, WpsWorkbook, WpsWorksheet } from "./wpsHost";
import {
  normalizeWpsMatrix,
  workbookFullName,
  type WpsAuthorizedSheet
} from "./wpsWorkbookReader";
import {
  applyFormatAction,
  assertWpsPreparedPlanUnchanged,
  asFormatFields,
  composedFormatSnapshot,
  formatActionApplied,
  formatEvidenceFingerprint,
  matrixFingerprintV2,
  prepareWpsPlan,
  wpsActionChangeReview,
  type StoredAction,
  type WpsActionAfter,
  type WpsPreparedAction,
  type WpsPreparedPlan
} from "./wpsPlanExecutor";
import { captureWpsFormatSnapshot, sameWpsFormatSnapshot, type WpsFormatSnapshot } from "./wpsFormatSnapshot";

export class WpsFormulaTrialError extends Error {}

/** A disposable, unsaved workbook. Only the original plan can later be committed. */
export type WpsNativeComparison = {
  original: WpsWorkbook;
  copy: WpsWorkbook;
  copyIdentity: WpsComparisonWorkbookIdentity;
  /**
   * 副本承载的**写入目标表**。
   *
   * 它不等于"授权表里的第一张"：计划可能依赖别的已授权表（例如两表关联后
   * 汇总到第三张表，来源表只读、目标表写入）。副本里只要有写入目标那一张
   * 就够了 —— 来源的核对发生在**原表**上（`assertOriginalUnchanged` 会用全部
   * 授权序号重新读一遍），不需要进副本。
   */
  writeSheet: WpsAuthorizedSheet;
  plan: WpsPreparedPlan;
  reviews: RangeChangeReview[];
  changedCells: string[];
  originalBefore: WpsActionAfter[];
  copyFingerprints: string[];
  copyAfter: WpsActionAfter[];
  /**
   * 格式动作的**副本侧最终**格式快照（整份计划全部执行完之后采集）。按 actionIndex
   * 与 `plan.actions` 对齐；值动作为 `undefined`。
   *
   * 确认（`verifyWpsNativeComparison`）与提交后对照（`compareCommittedWpsResultToCopy`）
   * 都拿它作"整份计划产出的样子"来比 —— **不能**用逐动作中间快照：同区域
   * "加粗→居中"里，动作 0 的中间快照只有加粗、动作 1 之后才是加粗+居中，
   * 拿中间快照比最终现场会把正常组合误报成"副本发生了变化"。
   */
  copyFormatFinal: Array<WpsFormatSnapshot | undefined>;
  /**
   * 格式动作的**原表侧**格式快照（执行前）。按 actionIndex 对齐；值动作为
   * `undefined`。用于核对"原表格式在并排审阅期间没有被外部改掉"。
   *
   * 原表在审阅期间不做任何修改，所以这里天然就是同一份"写前"状态；即便多个动作
   * 区域部分重叠，按各自区域采集得到的仍是同一次未变更的现场。
   */
  originalFormatBefore: Array<WpsFormatSnapshot | undefined>;
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
  const { values, valuesCaptured } = readWpsRangeValues(range, action.rowCount, action.columnCount);
  const formulas = normalizeWpsMatrix(range.Formula, action.rowCount, action.columnCount);
  return {
    values,
    formulas,
    valuesCaptured,
    fingerprint: matrixFingerprintV2(values, formulas)
  };
}

function isFormatActionType(type: string): type is "setFont" | "setAlignment" | "resizeRange" {
  return type === "setFont" || type === "setAlignment" || type === "resizeRange";
}

/** 格式动作覆盖的全部格地址（格式动作不改值，`changedCells` 需按整块区域标记）。 */
function rangeCellAddresses(action: WpsPreparedAction): string[] {
  const addresses: string[] = [];
  for (let row = 0; row < action.rowCount; row += 1) {
    for (let column = 0; column < action.columnCount; column += 1) {
      addresses.push(`${columnName(action.rect.startColumn + column)}${action.rect.startRow + row}`);
    }
  }
  return addresses;
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

function valueKind(value: CellValue | undefined): string {
  if (value === undefined) return "未取得值";
  if (value === null) return "空单元格";
  if (value === "") return "空字符串";
  if (typeof value === "number") return "数值";
  if (typeof value === "boolean") return "布尔值";
  return "文本";
}

/** Local diagnostic only: identify the failing cell without exposing its content. */
function firstValueDifference(action: WpsPreparedAction, observed: CellValue[][]): string {
  for (let row = 0; row < action.rowCount; row += 1) {
    for (let column = 0; column < action.columnCount; column += 1) {
      const expected = action.values[row]?.[column];
      const actual = observed[row]?.[column];
      if (JSON.stringify(expected) !== JSON.stringify(actual)) {
        const address = `${columnName(action.rect.startColumn + column)}${action.rect.startRow + row}`;
        const sameKind = valueKind(expected) === valueKind(actual) ? "（内容不同）" : "";
        return `首个差异 ${address}：计划为${valueKind(expected)}，回读为${valueKind(actual)}${sameKind}。`;
      }
    }
  }
  return "回读矩阵尺寸与计划不同。";
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

/**
 * **第一、二项判断**：这是本次流程创建的那一份副本，并且没有被另存/替换。
 *
 * 只管"是不是我们的副本"，**管不着**"副本里有没有额外人工修改"（那是第三项，
 * 见 `wpsCopyContentVerdict`）。返回 `null` 表示前两项通过。
 *
 *  - 读不出身份 ⇒ 副本已不在（用户自己关过）⇒ `already_closed`，幂等收尾；
 *  - 身份变了 ⇒ 用户另存/替换过 ⇒ `user_saved`：那是他自己的文件，**不关不删**。
 *
 * `expected` 省略时跳过 —— 只有"副本是我们刚刚创建、还攥在手里"的 staging 失败路径
 * 才允许这样调用。
 */
function copyOwnershipVerdict(copy: WpsWorkbook, expected?: WpsComparisonWorkbookIdentity): ReviewCopyCloseOutcome | null {
  if (!expected) return null;
  let current: WpsComparisonWorkbookIdentity;
  try { current = comparisonWorkbookIdentity(copy); }
  catch { return "already_closed"; }
  return sameWorkbookIdentity(current, expected) ? null : "user_saved";
}

/**
 * **第三项判断（WPS）**：副本相对"整份计划写入后的最终状态"有没有变化。
 *
 * 身份可信（①）、没被另存（②）都成立，**也推不出**③"副本里没有额外人工修改"：
 * 用户在并排审阅期间完全可以在副本上敲字。所以这里逐格回读副本：
 *  - 值/公式按生成对照时记下的 `copyFingerprints` + `copyAfter`；
 *  - 格式动作按 `copyFormatFinal`（整份计划执行完之后的最终格式，不是中间快照）；
 *  - 任何一步**读不到**（宿主不给读数、副本不可激活、工作表被换掉）⇒ `unknown`，
 *    **不**当成"没改过"。只有全部读回且逐格一致才 `unchanged`。
 *
 * 与 `verifyWpsNativeComparison`（写回侧）共用同一套判据，避免两条路各说一套。
 */
export function wpsCopyContentVerdict(staged: WpsNativeComparison): ReviewCopyContentVerdict {
  try {
    if (typeof staged.copy.Activate !== "function") return "unknown";
    staged.copy.Activate();
    const sheet = staged.copy.Worksheets?.Item?.(1);
    if (!sheet || sheet.Name !== staged.writeSheet.name) return "unknown";
    for (const [index, action] of staged.plan.actions.entries()) {
      const current = snapshot(target(sheet, action.range), action);
      if (!current.valuesCaptured) return "unknown";
      if (current.fingerprint !== staged.copyFingerprints[index] ||
          JSON.stringify(current.values) !== JSON.stringify(staged.copyAfter[index].values)) return "changed";
      if (isFormatActionType(action.type)) {
        const finalFormat = staged.copyFormatFinal[index];
        const currentFormat = captureWpsFormatSnapshot(target(sheet, action.range), action.rowCount, action.columnCount);
        if (!finalFormat?.captured || !currentFormat.captured) return "unknown";
        if (!sameWpsFormatSnapshot(currentFormat, finalFormat)) return "changed";
      }
    }
    return "unchanged";
  } catch { return "unknown"; }
}

/**
 * 只负责把宿主对关闭动作的**真实反应**折成结果（前两项/第三项判断已在调用方完成）。
 *
 * 抛异常会把"副本没能自动关闭"和"审阅本身失败"混成同一件事，调用方就只能
 * 笼统地报"失败"。这里把宿主反应分开：
 *  - 没有 `Close` ⇒ `unsupported`，如实降级为手动关闭；
 *  - `Close(false)` 返回 false ⇒ `close_unconfirmed`。这是**歧义**信号（被取消 /
 *    被推迟 / 被拒），而 `Close` 不返回任何对话框信息 ⇒ 只报"关闭尚未确认"，
 *    **不推断**宿主正在询问是否保存；
 *  - 抛异常 ⇒ `failed`；成功 ⇒ `closed`。
 */
function requestCopyClose(copy: WpsWorkbook): ReviewCopyCloseOutcome {
  if (typeof copy.Close !== "function") return "unsupported";
  try { return copy.Close(false) === false ? "close_unconfirmed" : "closed"; }
  catch { return "failed"; }
}

/**
 * 副本要承载的那一张表 = **写入目标表**。
 *
 * 判据从"授权表只有一张"改成"**写入只落在一张已授权表上**"，理由是
 * 授权表张数与写入表张数根本不是一回事：
 *
 *  - 本机桥会按用户勾选的**全部**已授权表读源（两表关联汇总就是 3 张：两张来源 +
 *    一张写入目标），来源只被读、不被写；
 *  - 服务端按引用编译出的白名单计划同样可能跨表依赖（`sourceDependencySheets`）。
 *
 * 把"授权张数 ≠ 1"当成"不能审阅"，会把这类计划整片挡在确认写回之外 —— 而同一份
 * 计划在 Excel 端走的是另一条**已存在的**显式执行入口，两端就此分叉。
 *
 * 复制范围仍然严格等于**一张表**：绝不把整本工作簿（含未写入、甚至未授权的表）
 * 复制出去。跨表**写入**依旧拒绝 —— 那种计划的副本预览无法自洽。
 */
function writeSheetOf(prepared: WpsPreparedPlan): WpsAuthorizedSheet {
  const written = new Set(prepared.actions.map((action) => action.sheetName));
  if (written.size !== 1) {
    throw new Error("原生并排审阅当前只支持单张工作表上的写入计划；跨表写入不会复制整本工作簿。");
  }
  const name = prepared.actions[0].sheetName;
  const authorized = prepared.authorization.sheets.find((sheet) => sheet.name === name);
  if (!authorized) {
    throw new Error(`写入目标表「${name}」不在本次已授权范围内，原表未写入。`);
  }
  return authorized;
}

/**
 * The native split view is deliberately limited to the single sheet being written. Copying the
 * whole workbook would silently include sheets the user did not authorize. No file is saved here.
 */
export function stageWpsNativeComparison(
  application: WpsApplication,
  prepared: WpsPreparedPlan,
  beforeOnly = false
): WpsNativeComparison {
  const writeSheet = writeSheetOf(prepared);
  const original = assertOriginalUnchanged(application, prepared);
  const originalSheet = original.Worksheets.Item(writeSheet.index);
  const originalBefore = prepared.actions.map((action) => snapshot(target(originalSheet, action.range), action));
  if (originalBefore.some((item) => !item.valuesCaptured)) {
    throw new Error("宿主未提供原表目标区域的显示值，无法建立可靠的并排审阅基线；原表未写入。");
  }
  // 格式动作的原表侧格式快照（执行前），用于核对"并排审阅期间原表格式未被外部改掉"。
  const originalFormatBefore = prepared.actions.map((action) =>
    isFormatActionType(action.type)
      ? captureWpsFormatSnapshot(target(originalSheet, action.range), action.rowCount, action.columnCount)
      : undefined
  );
  for (const [index, action] of prepared.actions.entries()) {
    const formatBefore = originalFormatBefore[index];
    if (formatBefore && !formatBefore.captured) {
      throw new Error(`原表「${action.sheetName}!${action.range}」的格式无法逐格采集，无法建立格式基线；原表未写入。`);
    }
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
    /**
     * 副本侧格式动作快照，按 actionIndex 对齐。循环里先填**动作级执行证据**
     * （本动作执行后的现场），循环结束后再**整体覆盖成最终状态** —— 见下方 finalize 段。
     */
    const copyFormatFinal: Array<WpsFormatSnapshot | undefined> =
      prepared.actions.map(() => undefined);
    for (const [index, action] of prepared.actions.entries()) {
      if (sheet.ProtectContents !== false) throw new Error("审阅副本的工作表受保护或状态未知。");
      const range = target(sheet, action.range);
      if (range.MergeCells !== false) throw new Error("审阅副本的目标区域包含合并单元格或状态未知。");
      const before = snapshot(range, action);
      if (!before.valuesCaptured || before.fingerprint !== action.previewFingerprint ||
          JSON.stringify(before.values) !== JSON.stringify(originalBefore[index].values)) {
        throw new Error("审阅副本与原表的预览快照不同，不能据此生成可靠对照。原表未写入。");
      }
      if (beforeOnly) {
        copyFingerprints.push(before.fingerprint);
        copyAfter.push(before);
        continue;
      }
      const stored: StoredAction = {
        ...action,
        beforeValues: before.values,
        beforeFormulas: before.formulas,
        beforeFingerprint: before.fingerprint,
        afterFingerprint: null,
        appliedRows: 0
      };

      // 格式动作：不改值/公式，只在副本上改格式，回读格式快照判定。
      // 值/公式仍与 before 一致（格式动作的 `values === beforeValues`）。
      if (isFormatActionType(action.type)) {
        const beforeFormat = captureWpsFormatSnapshot(range, action.rowCount, action.columnCount);
        if (!beforeFormat.captured) {
          throw new Error("审阅副本无法逐格采集格式快照，不能据此生成可靠对照；原表未写入。");
        }
        /*
         * 副本的写前格式必须与**预览时登记的格式证据**一致。上面那条 `before.fingerprint`
         * 比的是值矩阵，而格式动作不改值 —— 只用它复核，格式在预览后被外部改过也发现不了。
         *
         * 基准是**合成后的**写前快照：同一区域里更早的格式动作按序改过本动作声明的属性时，
         * 副本此刻的状态本来就是"前序动作执行后"，直接拿预览原始快照比会误报。
         */
        const formatAction = asFormatFields(action);
        const expectedBefore = composedFormatSnapshot(prepared.actions, index);
        if (expectedBefore &&
            formatEvidenceFingerprint(formatAction, beforeFormat) !==
              formatEvidenceFingerprint(formatAction, expectedBefore)) {
          throw new Error("审阅副本的格式与原表预览时的状态不同，不能据此生成可靠对照。原表未写入。");
        }
        applyFormatAction(range, formatAction);
        stored.appliedRows = action.rowCount;
        const afterFormat = captureWpsFormatSnapshot(range, action.rowCount, action.columnCount);
        // 这是**动作级执行证据**：只判"本动作声明的属性改到了、未声明属性没被连带改动"。
        const matched = afterFormat.captured &&
          formatActionApplied(asFormatFields(action), beforeFormat, afterFormat);
        // 格式动作不改值：回读值矩阵应与写前一致。
        const after = snapshot(target(sheet, action.range), action);
        if (!after.valuesCaptured ||
            JSON.stringify(after.values) !== JSON.stringify(before.values) ||
            JSON.stringify(after.formulas) !== JSON.stringify(before.formulas)) {
          throw new Error("审阅副本在应用格式后值或公式发生了变化；原表未写入。");
        }
        reviews.push(wpsActionChangeReview(stored, after, matched));
        if (!matched) {
          throw new Error(`审阅副本「${action.sheetName}!${action.range}」的格式回读与计划不一致；原表未写入。`);
        }
        // 先记**动作级执行证据**；循环结束后会被整体覆盖成最终状态（见下方 finalize）。
        copyFormatFinal[index] = afterFormat;
        copyFingerprints.push(after.fingerprint);
        copyAfter.push(after);
        changedCells.push(...rangeCellAddresses(action));
        continue;
      }

      // 值动作：格式维度无状态。

      // Mirror the production executor's bounded write shape; the copy is disposable but
      // the WPS bridge still has the same per-call limits.
      for (let offset = 0; offset < action.rowCount; offset += capabilities.wpsWrite.chunkRows) {
        const rows = Math.min(capabilities.wpsWrite.chunkRows, action.rowCount - offset);
        const chunkValues = action.values.slice(offset, offset + rows);
        const chunk = target(sheet, chunkAddress(action, offset, rows));
        try {
          writeWpsMatrix(chunk, chunkValues, action.type === "writeFormulas" ? "Formula" : "Value2");
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
      // 严格比较即可：宿主写回读的语义等值差异（分析结果里的空字符串 ↔ 空单元格）
      // 已在**服务端结果计划编译**（`result_plan.result_matrix`）规范成 null，
      // 计划值与回读值这时本来就是同形的。此处不做归一 —— 归一会把"写入没生效"
      // 也放过去（`wpsEmptyStringBaseline` 钉的就是这条）。
      const matched = JSON.stringify(observed) === JSON.stringify(action.values);
      reviews.push(wpsActionChangeReview(stored, after, matched));
      if (!matched) {
        if (action.type === "writeFormulas") throw new WpsFormulaTrialError("审阅副本的公式文本回读不符；原表未写入。");
        throw new Error(`审阅副本「${action.sheetName}!${action.range}」回读与计划不一致；${firstValueDifference(action, observed)}原表未写入。`);
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
    // 整份计划执行后，把副本侧格式快照**整体覆盖成最终状态**（仍按各自区域采集）。
    // 确认（`verifyWpsNativeComparison`）与提交后对照（`compareCommittedWpsResultToCopy`）
    // 都拿它作"计划产出的样子" —— 上面循环里的是**动作级中间态**，代表不了同区域
    // 多格式动作的组合状态（"加粗→居中"里动作 0 的中间态还没有居中）。
    for (const [index, action] of prepared.actions.entries()) {
      if (!isFormatActionType(action.type)) continue;
      const finalFormat = captureWpsFormatSnapshot(target(sheet, action.range), action.rowCount, action.columnCount);
      if (!finalFormat.captured) {
        throw new Error(`审阅副本「${action.sheetName}!${action.range}」的最终格式无法逐格采集；原表未写入。`);
      }
      copyFormatFinal[index] = finalFormat;
    }
    original.Activate();
    assertOriginalUnchanged(application, prepared);
    const copyWindowName = String(copy.Windows?.Item(1)?.Caption ?? copy.Name);
    if (!beforeOnly && application.Windows!.CompareSideBySideWith!(copyWindowName) !== true) {
      throw new Error("WPS 未能打开原生并排比较；原表未写入。");
    }
    if (!beforeOnly && application.Windows) application.Windows.SyncScrollingSideBySide = true;
    return { original, copy, copyIdentity, writeSheet, plan: prepared, reviews, changedCells, originalBefore, copyFingerprints, copyAfter, copyFormatFinal, originalFormatBefore };
  } catch (error) {
    // 副本是这一路径自己创建的、此刻仍在手里，没有任何用户交互窗口：
    // 先尝试关掉且不保存。关不掉也不掩盖原始失败原因（原表在这个路径里从未被写入）。
    const closeOutcome = requestCopyClose(copy);
    try { original.Activate(); } catch { /* The original workbook is never written in this path. */ }
    if (closeOutcome !== "closed" && closeOutcome !== "already_closed") {
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
  if (sheet.Name !== staged.writeSheet.name) throw new Error("审阅副本工作表已变化，已停止写回。");
  /*
   * 副本侧的**第三项判断**与"放弃"路径共用同一套判据（`wpsCopyContentVerdict`）：
   * 两条路必须得出同一个结论，否则会出现"放弃说没改、写回说改了"这种自相矛盾。
   * 写回只提交**已审阅的计划**，副本里的人工修改一律不导入原表。
   */
  const content = wpsCopyContentVerdict(staged);
  if (content === "changed") throw new Error("审阅副本在生成对照后发生变化，请重新生成预览；不会把未经确认的内容写回原表。");
  if (content === "unknown") throw new Error("审阅副本无法逐格回读，不能确认它未被额外修改；已停止写回，原表未写入。");
  if (typeof staged.original.Activate !== "function") throw new Error("无法激活原表，已停止写回。");
  staged.original.Activate();
  const activeOriginal = assertOriginalUnchanged(application, staged.plan);
  const originalSheet = activeOriginal.Worksheets.Item(staged.writeSheet.index);
  for (const [index, action] of staged.plan.actions.entries()) {
    const current = snapshot(target(originalSheet, action.range), action);
    if (!current.valuesCaptured ||
        JSON.stringify(current.values) !== JSON.stringify(staged.originalBefore[index].values) ||
        JSON.stringify(current.formulas) !== JSON.stringify(staged.originalBefore[index].formulas)) {
      throw new Error("原表在并排审阅后发生变化，请重新预览；不会覆盖新状态。");
    }
    // 格式动作：核对原表格式快照仍与"执行前"一致（外部没改格式）。
    if (isFormatActionType(action.type)) {
      const currentFormat = captureWpsFormatSnapshot(target(originalSheet, action.range), action.rowCount, action.columnCount);
      if (!currentFormat.captured || !sameWpsFormatSnapshot(currentFormat, staged.originalFormatBefore[index]!)) {
        throw new Error("原表的格式在并排审阅后发生变化，请重新预览；不会覆盖新状态。");
      }
    }
  }
}

/**
 * 放弃/结束本次审阅：**先过三项判断**，再决定要不要关闭副本；返回**关闭结果**
 * （不再抛异常）。
 *
 * 顺序不可颠倒、也不可互相顶替：
 *  ① 身份可信；② 没被另存/替换（两者见 `copyOwnershipVerdict`）；
 *  ③ **副本里没有额外人工修改**（`wpsCopyContentVerdict` → `describeCopyDiscardSafety`）。
 *  只有三项都通过才调用 `Close(false)`。③ 不通过（有修改 / 读不到读数）就
 *  **不关副本**：用户自己的改动不会被静默丢弃，且这些改动**不会**导入原表。
 *
 * 原表永远不写、不关；只在真的要关闭副本时才把焦点还给用户。关闭被拒或结果
 * 不明时不动业务状态。
 */
export function discardWpsNativeComparison(staged: WpsNativeComparison): ReviewCopyCloseOutcome {
  const ownership = copyOwnershipVerdict(staged.copy, staged.copyIdentity);
  if (ownership) return ownership;
  const blocked = describeCopyDiscardSafety(wpsCopyContentVerdict(staged));
  if (blocked) return blocked;
  const outcome = requestCopyClose(staged.copy);
  try { staged.original.Activate?.(); } catch { /* 焦点归还失败不影响清理结果 */ }
  return outcome;
}

/**
 * 用户主动点「重新排列」时**重新套用宿主默认并排布局**。
 *
 * 只做两件宿主级操作，绝不写任何单元格：重新激活原表并再次
 * `Windows.CompareSideBySideWith(副本题注)`，若宿主提供 `ResetPositionsSideBySide`
 * 则先复位默认左右位置。WPS 的 `Window` 没有可写的宽度成员，**两表的窗口宽度由宿主决定，
 * 不由本加载项设置**；因此这里只承诺"已请求宿主重新并排"，不承诺 35/65。
 *
 * 身份在调用前后各核一次；任何证据不足或宿主拒绝都返回 `unavailable`，绝不谎报已排列。
 */
export function rearrangeWpsComparison(
  application: WpsApplication,
  staged: WpsNativeComparison
): "arranged" | "unavailable" {
  try {
    const caption = String(staged.copy.Windows?.Item(1)?.Caption ?? staged.copy.Name ?? "");
    // 只核身份与几何，不读任何单元格：窗格/副本被换掉时绝不重排别人的窗口。
    const identityHolds = (): boolean =>
      sameWorkbookIdentity(comparisonWorkbookIdentity(staged.copy), staged.copyIdentity) &&
      Number(staged.original.Worksheets?.Count ?? 0) > 0;
    if (!caption || typeof staged.original.Activate !== "function" ||
        typeof application.Windows?.CompareSideBySideWith !== "function" || !identityHolds()) return "unavailable";
    application.Windows.ResetPositionsSideBySide?.();
    staged.original.Activate();
    if (application.Windows.CompareSideBySideWith(caption) !== true) return "unavailable";
    if (!identityHolds()) return "unavailable";
    // 同步滚动是独立能力：重新排列后尽力恢复，失败只影响联动，不影响排列结论。
    try { if (application.Windows) application.Windows.SyncScrollingSideBySide = true; } catch { /* optional */ }
    return "arranged";
  } catch {
    return "unavailable";
  }
}

/** Formula text can match while calculated values differ between a one-sheet copy and the original. */
export function compareCommittedWpsResultToCopy(staged: WpsNativeComparison): "matched" | "different" | "unavailable" {
  const sheet = staged.original.Worksheets.Item(staged.writeSheet.index);
  let unavailable = false;
  for (const [index, action] of staged.plan.actions.entries()) {
    const actual = snapshot(target(sheet, action.range), action);
    const copied = staged.copyAfter[index];
    if (!actual.valuesCaptured || !copied.valuesCaptured) unavailable = true;
    // 这里比的是**两个不同工作簿**（原表提交后 vs 副本试算后）的读数，公式文本
    // 必须逐字一致（公式差异是真差异）；值同样严格比较 —— 分析结果里的空字符串
    // 已在服务端编译计划时规范成 null，两份读数应当同形。
    if (JSON.stringify(actual.formulas) !== JSON.stringify(copied.formulas)) return "different";
    if (actual.valuesCaptured && copied.valuesCaptured &&
        JSON.stringify(actual.values) !== JSON.stringify(copied.values)) return "different";
    // 格式动作：原表提交后格式 vs 副本最终格式，逐格逐行一致。
    if (isFormatActionType(action.type)) {
      const committedFormat = captureWpsFormatSnapshot(target(sheet, action.range), action.rowCount, action.columnCount);
      const copyFormat = staged.copyFormatFinal[index];
      if (!committedFormat.captured || !copyFormat || !copyFormat.captured) {
        unavailable = true;
      } else if (!sameWpsFormatSnapshot(committedFormat, copyFormat)) {
        return "different";
      }
    }
  }
  return unavailable ? "unavailable" : "matched";
}
