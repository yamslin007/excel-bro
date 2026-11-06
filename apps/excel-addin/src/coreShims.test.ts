import { describe, expect, it } from "vitest";
import * as core from "@excel-bro/core";
import * as planCoreShim from "./planCore";
import * as rangeShim from "./range";
import * as tableSchemaShim from "./tableSchema";

/**
 * 阶段 2 的 shim 必须是「同一个实现」而不是「行为相似的副本」。
 * 用函数引用相等来锁这一点：一旦有人把 shim 改回本地实现，
 * 引用不再相同，本测试立刻报红。阶段 6 删除 shim 时本文件一并删除。
 */
describe("阶段 2 兼容 shim", () => {
  it("planCore 的导出与 core 同一引用", () => {
    expect(planCoreShim.parseExplicitA1Rect).toBe(core.parseExplicitA1Rect);
    expect(planCoreShim.assertRectangularValues).toBe(core.assertRectangularValues);
    expect(planCoreShim.writeMatrixShapeIssue).toBe(core.writeMatrixShapeIssue);
    expect(planCoreShim.blankCellMatrix).toBe(core.blankCellMatrix);
  });

  it("range 的导出与 core 同一引用", () => {
    expect(rangeShim.firstCellOfRange).toBe(core.firstCellOfRange);
    expect(rangeShim.columnLetterFromIndex).toBe(core.columnLetterFromIndex);
    expect(rangeShim.columnIndexFromLetter).toBe(core.columnIndexFromLetter);
    expect(rangeShim.startColumnOfRange).toBe(core.startColumnOfRange);
    expect(rangeShim.suggestWriteTarget).toBe(core.suggestWriteTarget);
  });

  it("tableSchema 的导出与 core 同一引用", () => {
    expect(tableSchemaShim.normalizeField).toBe(core.normalizeField);
    expect(tableSchemaShim.detectHeaderIndex).toBe(core.detectHeaderIndex);
    expect(tableSchemaShim.detectSheetFields).toBe(core.detectSheetFields);
    expect(tableSchemaShim.looksLikeHeaderRow).toBe(core.looksLikeHeaderRow);
    expect(tableSchemaShim.isSyntheticField).toBe(core.isSyntheticField);
    expect(tableSchemaShim.syntheticFields).toBe(core.syntheticFields);
  });

  it("shim 没有遗漏 core 的对应导出", () => {
    expect(Object.keys(planCoreShim).sort()).toEqual([
      "assertRectangularValues",
      "blankCellMatrix",
      "parseExplicitA1Rect",
      "writeMatrixShapeIssue"
    ]);
    expect(Object.keys(rangeShim).sort()).toEqual([
      "columnIndexFromLetter",
      "columnLetterFromIndex",
      "firstCellOfRange",
      "startColumnOfRange",
      "suggestWriteTarget"
    ]);
    expect(Object.keys(tableSchemaShim).sort()).toEqual([
      "detectHeaderIndex",
      "detectSheetFields",
      "isSyntheticField",
      "looksLikeHeaderRow",
      "normalizeField",
      "syntheticFields"
    ]);
  });
});
