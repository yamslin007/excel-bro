/**
 * 「修改前｜修改后」双栏变更对照的共用数据模型与纯比较逻辑。
 *
 * 规格见 docs/DUAL_VIEW_CHANGE_REVIEW_GUIDE_2026-09-29.md。
 * 本文件是宿主无关层：不碰 Office.js / WPS JSAPI / openpyxl，
 * Excel、WPS、文件夹三种模式都把自己的读取结果交给这里的纯函数出对照结果。
 *
 * 三条不能松的口径：
 * 1. 空单元格（null）与空字符串（""）是两个状态，比较前**不得**折叠。
 *    WPS 侧现有两个 normalize：`wpsWorkbookReader.normalizeCell` 保留区别（双栏走这条），
 *    `wpsPlanExecutor.canonicalCell` 会把 "" 折成 null（只给指纹用）——本文件不复用后者。
 * 2. 值（value）与公式（formula）同时容纳。展示看值，判定写入成败要看公式本身，
 *    因为重算时机不同会让计算值漂移。
 * 3. 宿主没采到计算值时如实标注"显示值未采集"，禁止拿公式原文冒充值。
 */

import { columnLetterFromIndex } from "./range";
import { parseExplicitA1Rect } from "./planCore";

/** 单元格显示给用户的值：null = 空单元格，"" = 空字符串，两者不是一回事。 */
export type ChangeValue = string | number | boolean | null;

export type ChangeCell = {
  value: ChangeValue;
  /** 原公式；非公式单元格为 null。 */
  formula: string | null;
  /** false = 这一侧没采到计算值，只能展示公式，不得用公式冒充值。 */
  valueCaptured: boolean;
};

export type ChangeSnapshot = {
  values: ChangeValue[][];
  formulas: (string | null)[][];
  /** 整份快照有没有采到显示值。逐格时也可以在 payload 里按格降级。 */
  valuesCaptured: boolean;
};

export type ChangePhase = "expected" | "actual";

export type ChangeReviewStatus =
  | "ready"
  | "applied"
  /** 写入确实发生了、前后矩阵也拿到了，但独立验收没能全部通过。
   *  这种情况下双栏照常展示真实前后值，但**不允许**宣称成功。 */
  | "executed_unverified"
  | "readback_failed"
  | "readback_mismatch"
  | "not_executed";

export type ChangeKind =
  | "value"
  | "formula"
  | "formulaOnly"
  | "empty"
  | "same"
  | "not_executed";

export type ChangeReviewRow = {
  /** 稳定标识，供虚拟化 key 与滚动同步使用。 */
  key: string;
  address: string;
  before: ChangeCell;
  after: ChangeCell;
  kind: ChangeKind;
  changed: boolean;
  /** 修改前是公式且本次被覆盖（改成值或改成另一条公式）→ 双栏必须突出标记。 */
  overwritesFormula: boolean;
};

export type ChangeReviewLimit = {
  /** 双栏最多渲染多少条（每条是一个单元格），超出的走摘要与提示。 */
  maxRows: number;
  maxTotalCells: number;
};

export type RangeChangeReview = {
  sheet: string;
  range: string;
  phase: ChangePhase;
  status: ChangeReviewStatus;
  rows: ChangeReviewRow[];
  /** 全量统计，不等于被截断后的 rows.length。 */
  changedCount: number;
  totalCount: number;
  truncated: boolean;
};

/** 出口摘要：不含任何单元格内容，用于活动日志与任何可能被持久化的地方。 */
export type ChangeReviewSummary = {
  sheet: string;
  range: string;
  status: ChangeReviewStatus;
  phase: ChangePhase;
  changedCount: number;
  totalCount: number;
  truncated: boolean;
};

export const DEFAULT_CHANGE_REVIEW_LIMIT: ChangeReviewLimit = {
  maxRows: 200,
  maxTotalCells: 20000
};

/** 值未采集时给用户看的占位。必须是这一串，方便测试锁住"不许拿公式顶替"。 */
export const VALUE_NOT_CAPTURED_TEXT = "显示值未采集";

export type CellDisplayText = {
  text: string;
  /** 原公式，供下方小字展示；没有公式时是 null。 */
  formula: string | null;
  /** true = 没采到值，text 是占位文案而不是真实内容。 */
  unresolved: boolean;
};

/** 把一个单元格翻译成"值栏写什么、公式原文是什么"。 */
export function cellDisplayText(cell: ChangeCell): CellDisplayText {
  const formula = cell.formula;
  if (!cell.valueCaptured) {
    return { text: VALUE_NOT_CAPTURED_TEXT, formula, unresolved: true };
  }
  const value = cell.value;
  const text =
    value === null ? "（空）" : value === "" ? "（空字符串）" : String(value);
  return { text, formula, unresolved: false };
}

/**
 * 只有 UI 可以宣称"写入成功"的状态。
 * 回读失败、回读不符、未执行、只有预期（还没写）都不算——这条是兜底，
 * 防止将来有人为了让结果好看把 readback_failed 也放行。
 */
export function canClaimWriteSuccess(status: ChangeReviewStatus): boolean {
  return status === "applied";
}

function cellAt(
  snapshot: ChangeSnapshot,
  rowIndex: number,
  columnIndex: number
): ChangeCell {
  return {
    value: snapshot.values[rowIndex]?.[columnIndex] ?? null,
    formula: snapshot.formulas[rowIndex]?.[columnIndex] ?? null,
    valueCaptured: snapshot.valuesCaptured
  };
}

function isEmptyCell(cell: ChangeCell): boolean {
  return cell.value === null || cell.value === "";
}

function sameCell(left: ChangeCell, right: ChangeCell): boolean {
  if (left.formula !== right.formula) return false;
  // 两侧都采到值才比值；否则公式已经比过了，不允许用糊涂的值再判一次。
  if (left.valueCaptured && right.valueCaptured) {
    return left.value === right.value;
  }
  return true;
}

function classifyChange(
  before: ChangeCell,
  after: ChangeCell
): { kind: ChangeKind; changed: boolean; overwritesFormula: boolean } {
  const formulaChanged = before.formula !== after.formula;
  const overwritesFormula = before.formula !== null && formulaChanged;
  const beforeEmpty = isEmptyCell(before);
  const afterEmpty = isEmptyCell(after);
  let kind: ChangeKind;
  if (formulaChanged) kind = "formula";
  else if (!beforeEmpty && afterEmpty) kind = "empty";
  else if (!sameCell(before, after)) kind = "value";
  else kind = "same";
  return {
    kind,
    changed: kind !== "same",
    overwritesFormula
  };
}

function sheetLessRange(range: string): string {
  return range.includes("!") ? (range.split("!").pop() ?? range) : range;
}

function matrixShape(matrices: ChangeValue[][][]): { rows: number; columns: number } {
  let rows = 0;
  let columns = 0;
  for (const matrix of matrices) {
    rows = Math.max(rows, matrix.length);
    for (const row of matrix) columns = Math.max(columns, row.length);
  }
  return { rows, columns };
}

export type BuildRangeChangeReviewInput = {
  sheet: string;
  range: string;
  phase: ChangePhase;
  status: ChangeReviewStatus;
  before: ChangeSnapshot;
  after: ChangeSnapshot;
  /**
   * 真正写入的行数（从区域首行算起）。部分写入时，超出这一行数的行全部标
   * not_executed —— 它们既不能算改动，更不能让用户误以为已经写进去了。
   * 不传代表整块都已尝试写入。
   */
  appliedRows?: number;
  limit?: Partial<ChangeReviewLimit>;
};

/**
 * 由「写前快照 + 写后快照（或未核对整个区域的占位）」生成一份双栏对照。
 *
 * 注意入参语义：`after` 必须来自**同一次执行**在事件点的回读，
 * 不允许传计划值来顶替，也不允许用上一次执行的结果。
 */
export function buildRangeChangeReview(
  input: BuildRangeChangeReviewInput
): RangeChangeReview {
  const limit: ChangeReviewLimit = {
    maxRows: input.limit?.maxRows ?? DEFAULT_CHANGE_REVIEW_LIMIT.maxRows,
    maxTotalCells:
      input.limit?.maxTotalCells ?? DEFAULT_CHANGE_REVIEW_LIMIT.maxTotalCells
  };
  const shape = matrixShape([input.before.values, input.after.values]);
  const rect = parseExplicitA1Rect(sheetLessRange(input.range));
  const rows: ChangeReviewRow[] = [];
  let changedCount = 0;

  for (let rowIndex = 0; rowIndex < shape.rows; rowIndex += 1) {
    for (
      let columnIndex = 0;
      columnIndex < (input.before.values[rowIndex]?.length ?? shape.columns);
      columnIndex += 1
    ) {
      const executed =
        input.appliedRows == null ? true : rowIndex < input.appliedRows;
      const before = cellAt(input.before, rowIndex, columnIndex);
      const after = cellAt(input.after, rowIndex, columnIndex);
      const classified = classifyChange(before, after);
      const kind: ChangeKind = executed ? classified.kind : "not_executed";
      const changed = executed && classified.changed;
      if (changed) changedCount += 1;
      const address = rect
        ? `${columnLetterFromIndex(rect.startColumn - 1 + columnIndex)}${
            rect.startRow + rowIndex
          }`
        : `${input.range}#${rowIndex + 1}-${columnIndex + 1}`;
      rows.push({
        key: `${input.sheet}!${input.range}#${rowIndex}-${columnIndex}`,
        address,
        before,
        after,
        kind,
        changed,
        overwritesFormula: executed && classified.overwritesFormula
      });
    }
  }

  const totalCount = rows.length;
  const visible = rows.slice(0, Math.max(0, limit.maxRows));
  return {
    sheet: input.sheet,
    range: input.range,
    phase: input.phase,
    status: input.status,
    rows: visible,
    changedCount,
    totalCount,
    truncated: visible.length < totalCount || totalCount > limit.maxTotalCells
  };
}

/**
 * 摘掉所有单元格内容，只留计数与定位。
 * 任何要走聊天历史、日志或持久化的地方都只能用这个；
 * 完整前后矩阵只允许停留在会话内内存。
 */
export function changeReviewSummary(review: RangeChangeReview): ChangeReviewSummary {
  return {
    sheet: review.sheet,
    range: review.range,
    status: review.status,
    phase: review.phase,
    changedCount: review.changedCount,
    totalCount: review.totalCount,
    truncated: review.truncated
  };
}

/** 「只看改动」开关用：过滤掉没有变化的行（未执行的行保留，它们属于异常提示）。 */
export function onlyChangedRows(review: RangeChangeReview): ChangeReviewRow[] {
  return review.rows.filter((row) => row.changed || row.kind === "not_executed");
}
