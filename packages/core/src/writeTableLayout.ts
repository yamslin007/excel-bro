import type { CellValue } from "./contracts";
import { assertRectangularValues, parseExplicitA1Rect } from "./planCore";
import { columnLetterFromIndex } from "./range";

export const WRITE_TABLE_HEADER_FILL = "#DFF3E4";

export function writeTableValues(action: {headers: CellValue[]; rows: CellValue[][]}): CellValue[][] {
  const values = [action.headers, ...action.rows].map(row => [...row]);
  assertRectangularValues(values);
  return values;
}

/** writeTable writes a matrix and styles its header; it does not create a table object. */
export function writeTableLayout(action: {startCell: string; headers: CellValue[]; rows: CellValue[][]}) {
  const anchor = parseExplicitA1Rect(action.startCell);
  if (!anchor || anchor.rows !== 1 || anchor.columns !== 1) throw new Error("写表起点必须是单个 A1 单元格。");
  const values = writeTableValues(action);
  const endColumn = anchor.startColumn + action.headers.length - 1;
  const endRow = anchor.startRow + values.length - 1;
  if (endColumn > 16384 || endRow > 1048576) throw new Error("写表目标超出工作表边界。");
  const start = `${columnLetterFromIndex(anchor.startColumn - 1)}${anchor.startRow}`;
  const lastColumn = columnLetterFromIndex(endColumn - 1);
  return { values, range: `${start}:${lastColumn}${endRow}`, headerRange: `${start}:${lastColumn}${anchor.startRow}` };
}
