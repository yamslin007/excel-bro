import type { CellValue } from "./contracts";

export { normalizeCellValue, displayCellValue } from "./cellNormalization";

/**
 * ``prepare_python_data`` 的区域切片契约（**两端同源**）。
 *
 * 为什么需要一份共享实现：A1 区域是**绝对**坐标，而从"已用区域"读回来的矩阵是
 * **相对**的。两者只有在已用区域从 A1 起时才相等；真宿主里数据从 B3 起是常态，
 * 直接拿绝对下标当矩阵下标会把整块数据错位 `(startRow-1, startColumn-1)`，而且
 * **不报错**（只是静默登记一片错数据）。
 *
 * 另一个必须共享的判定是"**真实空白 vs 人为补齐**"：缺行、短行、区域落在已读
 * 范围之外，都说明这一格**没有被真正读到**；用 null/"" 垫进去会让"读了但没读到"
 * 变成"读到了空值" —— 那是伪造。所以这三类一律拒绝，不补空白。
 *
 * 抛出的 `message` 是固定文案（不含单元格内容、不含路径），两端各自包装成自己的
 * 错误类型；**错误种类**是契约的一部分，文案不是。
 */
export const PYTHON_PREPARE_REFUSAL_MESSAGE = "数据登记区域超出已读取范围";

export class PythonPrepareRegionRefused extends Error {
  constructor() {
    super(PYTHON_PREPARE_REFUSAL_MESSAGE);
    this.name = "PythonPrepareRegionRefused";
  }
}

/** 0 基绝对矩形的四个边界（含）。 */
export interface PythonPrepareRect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

const RANGE = /^([A-Z]{1,3})([1-9][0-9]{0,6})(?::([A-Z]{1,3})([1-9][0-9]{0,6}))?$/;

function columnToIndex(column: string): number {
  let value = 0;
  for (const ch of column) value = value * 26 + (ch.charCodeAt(0) - 64);
  return value - 1;
}

/** 解析 ``A1`` / ``A1:B2`` 为 0 基绝对矩形。非法或反向区域返回 null。 */
export function parsePythonPrepareRect(range: string): PythonPrepareRect | null {
  const match = RANGE.exec(range.trim().toUpperCase());
  if (!match) return null;
  const c1 = columnToIndex(match[1]);
  const r1 = Number(match[2]) - 1;
  const c2 = match[3] ? columnToIndex(match[3]) : c1;
  const r2 = match[4] ? Number(match[4]) - 1 : r1;
  if (r2 < r1 || c2 < c1) return null;
  return { r1, c1, r2, c2 };
}

/** 已读整块区域的左上角（**1 基**，与宿主 `UsedRange.Row/Column` 一致）。 */
export interface PythonPrepareOrigin {
  startRow: number;
  startColumn: number;
}

/**
 * 在不读值的前提下判定区域是否完全落在已读范围内。
 *
 * Excel 侧走这条路径：它用 `getRange(绝对地址)` 直接读，不做切片，所以必须
 * 先拿已用区域的边界把越界请求挡掉 —— 否则"表外"会被静默当成空白返回，
 * 而 WPS 侧是拒绝，双端语义就分叉了。
 */
export function assertPythonPrepareRectWithin(
  rect: PythonPrepareRect,
  origin: PythonPrepareOrigin,
  rows: number,
  columns: number
): void {
  const startRow = Number(origin.startRow ?? 1);
  const startColumn = Number(origin.startColumn ?? 1);
  const mR1 = rect.r1 + 1 - startRow;
  const mC1 = rect.c1 + 1 - startColumn;
  const mR2 = rect.r2 + 1 - startRow;
  const mC2 = rect.c2 + 1 - startColumn;
  if (mR1 < 0 || mC1 < 0 || mR2 >= rows || mC2 >= columns) {
    throw new PythonPrepareRegionRefused();
  }
}

/**
 * 从已读矩阵里按 A1 绝对区域切片（WPS 侧走这条路径）。
 *
 * 逐行核宽度，缺行/短行一律拒绝，**不补空白**。
 */
export function slicePythonPrepareRegion(
  matrix: readonly CellValue[][],
  origin: PythonPrepareOrigin,
  rect: PythonPrepareRect
): CellValue[][] {
  const startRow = Number(origin.startRow ?? 1);
  const startColumn = Number(origin.startColumn ?? 1);
  const mR1 = rect.r1 + 1 - startRow;
  const mC1 = rect.c1 + 1 - startColumn;
  const mR2 = rect.r2 + 1 - startRow;
  const mC2 = rect.c2 + 1 - startColumn;
  if (mR1 < 0 || mC1 < 0) throw new PythonPrepareRegionRefused();
  if (mR2 >= matrix.length) throw new PythonPrepareRegionRefused();
  const values: CellValue[][] = [];
  for (let r = mR1; r <= mR2; r += 1) {
    const row = matrix[r];
    if (!Array.isArray(row) || row.length <= mC2) throw new PythonPrepareRegionRefused();
    values.push(row.slice(mC1, mC2 + 1));
  }
  return values;
}
