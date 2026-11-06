/**
 * `query_table` 的显式区域（只读）。
 *
 * 为什么需要它：`query_table` 原来只能读**整张**已授权工作表的已用区域。用户说
 * 「只读取来源!A1:B5」时，本端要么整表读（把区域外的数据也算进来），要么只能改用
 * 会写入的 `computeValues` 去表达。两条都不对：前者答非所问，后者把只读需求变成写入。
 *
 * 三件事在这里定死，宿主侧只按引用转发：
 *
 * 1. **写法**：`A1:B5` 或 `来源!A1:B5`（工作表名可带单引号）。只接受**单个矩形**
 *    A1 区域，列字母 + 行号都必须写；`A:A`、`1:5`、命名区域、多区域一律拒绝。
 * 2. **授权**：带工作表名时，那张表必须落在本次查询的范围里（由调用方核对），
 *    否则抛 `SCOPE_VIOLATION` —— 区域参数不能成为扩大授权的入口。
 * 3. **不越界读**：给定区域后，参与计算的矩阵**正好**是这个区域；区域外的单元格
 *    不参与，区域内的空格补 `null`（与宿主读一块空区域的行为一致）。
 *
 * 服务端协议侧（`server/app/models.py` 的 `QueryTableArguments.range`）按同一套规则
 * 校验；那里的注释与本文件逐条对应，改一处必须同步另一处。
 */

import type { CellValue } from "./contracts";
import { columnIndexFromLetter, columnLetterFromIndex } from "./range";

/** 1-based 的绝对行列矩形。 */
export interface AbsoluteRect {
  startRow: number;
  startColumn: number;
  rowCount: number;
  columnCount: number;
}

/** 区域参数：矩形 + 可选的所属工作表。 */
export interface QueryRect extends AbsoluteRect {
  /** 左上角在哪张表；`null` = 不限定，按 `scope` 决定。 */
  sheet: string | null;
}

const MAX_ROW = 1_048_576;
const MAX_COLUMN = 16_384;
/** 工作表名里不能出现的字符（宿主自己的限制＋我们用来分隔的 `!`）。 */
const FORBIDDEN_SHEET_CHARS = /[!:/\\?*[\]]/;
const A1_PATTERN =
  /^\$?([A-Za-z]{1,3})\$?([1-9]\d*)(?::\$?([A-Za-z]{1,3})\$?([1-9]\d*))?$/;

function splitSheet(value: string): { sheet: string | null; address: string } {
  const parts = value.trim().split("!");
  if (parts.length > 2) return { sheet: null, address: value };
  return {
    sheet: parts.length === 2 ? parts[0].trim().replace(/^'(.*)'$/, "$1") : null,
    address: (parts.length === 2 ? parts[1] : parts[0]).trim()
  };
}

/** 只是"写法不合法"的说明，供报错用；判定本身在 `parseQueryRange`。 */
export function queryRangeIssue(value: unknown): string | null {
  if (typeof value !== "string") return "区域必须是字符串，如 A1:B5 或 来源!A1:B5";
  const text = value.trim();
  if (!text) return "区域不能为空字符串；不限定区域时请省略 range";
  if (text.length > 40) return "区域写法过长（最多 40 个字符）";
  if (text.split("!").length > 2) return "区域只能是一个矩形，且只能带一个工作表名";
  const { sheet, address } = splitSheet(text);
  const match = A1_PATTERN.exec(address);
  if (!match) {
    return "区域必须是单个 A1 矩形（如 A1:B5）；整列 A:A、整行 1:5、命名区域和多区域都不支持";
  }
  const startRow = Number(match[2]);
  const endRow = Number(match[4] ?? match[2]);
  const startColumn = columnIndexFromLetter(match[1]) + 1;
  const endColumn = columnIndexFromLetter(match[3] ?? match[1]) + 1;
  if (endRow < startRow || endColumn < startColumn) {
    return "区域的结束行列不能小于起始行列";
  }
  if (endRow > MAX_ROW || endColumn > MAX_COLUMN) {
    return "区域超出工作表边界";
  }
  if (sheet !== null) {
    if (!sheet) return "区域里的工作表名为空";
    if (FORBIDDEN_SHEET_CHARS.test(sheet)) return "工作表名包含不允许的字符";
  }
  return null;
}

/** 解析区域；写法不合法返回 `null`（说明用 `queryRangeIssue`）。 */
export function parseQueryRange(value: unknown): QueryRect | null {
  if (queryRangeIssue(value) !== null) return null;
  const { sheet, address } = splitSheet(String(value));
  const match = A1_PATTERN.exec(address)!;
  const startRow = Number(match[2]);
  const startColumn = columnIndexFromLetter(match[1]) + 1;
  const endRow = Number(match[4] ?? match[2]);
  const endColumn = columnIndexFromLetter(match[3] ?? match[1]) + 1;
  return {
    sheet,
    startRow,
    startColumn,
    rowCount: endRow - startRow + 1,
    columnCount: endColumn - startColumn + 1
  };
}

/** 两个绝对矩形的交集；不相交返回 `null`。 */
export function intersectQueryRect(
  left: AbsoluteRect,
  right: AbsoluteRect
): AbsoluteRect | null {
  const startRow = Math.max(left.startRow, right.startRow);
  const startColumn = Math.max(left.startColumn, right.startColumn);
  const endRow = Math.min(
    left.startRow + left.rowCount - 1,
    right.startRow + right.rowCount - 1
  );
  const endColumn = Math.min(
    left.startColumn + left.columnCount - 1,
    right.startColumn + right.columnCount - 1
  );
  if (endRow < startRow || endColumn < startColumn) return null;
  return {
    startRow,
    startColumn,
    rowCount: endRow - startRow + 1,
    columnCount: endColumn - startColumn + 1
  };
}

/**
 * 按区域裁剪一块值矩阵。
 *
 * `origin` 是 `values[0][0]` 在表里的绝对位置（1-based）。结果**正好**是
 * `range.rowCount × range.columnCount`：落在 `origin` 之外的格补 `null`，
 * 与宿主"读一块空区域得到空格"的行为一致 —— 两端算出来的矩阵形状因此相同，
 * 不需要各自记住"区域是不是起点对齐"。
 */
export function sliceForQueryRange(
  values: readonly (readonly CellValue[])[],
  origin: { row: number; column: number },
  range: AbsoluteRect
): CellValue[][] {
  return Array.from({ length: range.rowCount }, (_unused, rowOffset) => {
    const source = values[range.startRow + rowOffset - origin.row];
    return Array.from({ length: range.columnCount }, (_cell, columnOffset) => {
      const sourceColumn = range.startColumn + columnOffset - origin.column;
      if (!source || sourceColumn < 0) return null;
      return source[sourceColumn] ?? null;
    });
  });
}

/** `A1:B5` 形态的地址字符串（写进错误信息与诊断用）。 */
export function queryRectAddress(range: AbsoluteRect): string {
  const start = `${columnLetterFromIndex(range.startColumn - 1)}${range.startRow}`;
  const end = `${columnLetterFromIndex(
    range.startColumn + range.columnCount - 2
  )}${range.startRow + range.rowCount - 1}`;
  return start === end ? start : `${start}:${end}`;
}
