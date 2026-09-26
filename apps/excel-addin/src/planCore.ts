import type { CellValue } from "./contracts";

/** 宿主无关的 A1 矩形；Excel 和 WPS 的适配器均使用同一套边界与形状判断。 */
export type A1Rect = {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
  rows: number;
  columns: number;
};

function columnNumber(name: string): number {
  return [...name.toUpperCase()].reduce(
    (result, character) => result * 26 + character.charCodeAt(0) - 64,
    0
  );
}

export function parseExplicitA1Rect(address: string): A1Rect | null {
  const match = /^\$?([A-Z]{1,3})\$?([1-9]\d*)(?::\$?([A-Z]{1,3})\$?([1-9]\d*))?$/i.exec(address.trim());
  if (!match) return null;
  const startColumn = columnNumber(match[1]);
  const startRow = Number(match[2]);
  const endColumn = columnNumber(match[3] ?? match[1]);
  const endRow = Number(match[4] ?? match[2]);
  if (
    startColumn > 16_384 || endColumn > 16_384 ||
    startRow > 1_048_576 || endRow > 1_048_576 ||
    endColumn < startColumn || endRow < startRow
  ) return null;
  return {
    startRow, startColumn, endRow, endColumn,
    rows: endRow - startRow + 1,
    columns: endColumn - startColumn + 1
  };
}

export function assertRectangularValues(values: CellValue[][]): void {
  if (values.length === 0) return;
  const width = values[0].length;
  if (width === 0 || values.some((row) => row.length !== width)) {
    throw new Error("写入数据必须是非空的规则二维数组");
  }
}

/** 统一比较白名单写值矩阵与目标矩形；命名区域由具体宿主另行解析。 */
export function writeMatrixShapeIssue(
  address: string,
  values: CellValue[][],
  rect = parseExplicitA1Rect(address)
): string | null {
  if (!rect) return null;
  return rect.rows === values.length && rect.columns === values[0]?.length
    ? null
    : `目标区域 ${address} 是 ${rect.rows}×${rect.columns}，但写入矩阵是 ${values.length}×${values[0]?.length ?? 0}`;
}

export function blankCellMatrix(rows: number, columns: number): CellValue[][] {
  return Array.from({ length: rows }, () => Array<CellValue>(columns).fill(null));
}
