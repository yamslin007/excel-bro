import { describe, expect, it } from "vitest";
import {
  assertRectangularValues,
  blankCellMatrix,
  parseExplicitA1Rect,
  writeMatrixShapeIssue
} from "./planCore";

describe("shared Excel/WPS plan core", () => {
  it("parses the same bounded A1 rectangle for both hosts", () => {
    expect(parseExplicitA1Rect("$B$2:C3")).toEqual({
      startRow: 2, startColumn: 2, endRow: 3, endColumn: 3,
      rows: 2, columns: 2
    });
    expect(parseExplicitA1Rect("XFE1")).toBeNull();
    expect(parseExplicitA1Rect("A1048577")).toBeNull();
    expect(parseExplicitA1Rect("B2:A3")).toBeNull();
  });

  it("shares matrix shape checks and clear-content values", () => {
    expect(writeMatrixShapeIssue("B2:C3", [[1, 2], [3, 4]])).toBeNull();
    expect(writeMatrixShapeIssue("B2:C3", [[1, 2]])).toBe(
      "目标区域 B2:C3 是 2×2，但写入矩阵是 1×2"
    );
    expect(() => assertRectangularValues([[1, 2], [3]])).toThrow("规则二维数组");
    expect(blankCellMatrix(2, 2)).toEqual([[null, null], [null, null]]);
  });
});
