import { describe, expect, it } from "vitest";
import {
  compareSortValues,
  deduplicatedMatrix,
  duplicateRowPartition,
  reorderRows,
  sortedMatrix,
  sortRowOrder
} from "./tableTransforms";

const HEADER = ["名称", "数量", "备注"];

describe("排序：行顺序是确定性的", () => {
  it("单键升序：数字按大小，表头不动", () => {
    const values = [HEADER, ["甲", 30, ""], ["乙", 10, ""], ["丙", 20, ""]];
    const sorted = sortedMatrix(values, [{ column: 1, ascending: true }], true);
    expect(sorted[0]).toEqual(HEADER);
    expect(sorted.slice(1).map((row) => row[0])).toEqual(["乙", "丙", "甲"]);
  });

  it("单键降序：数字从大到小，表头仍不动", () => {
    const values = [HEADER, ["甲", 30, ""], ["乙", 10, ""], ["丙", 20, ""]];
    const sorted = sortedMatrix(values, [{ column: 1, ascending: false }], true);
    expect(sorted.slice(1).map((row) => row[0])).toEqual(["甲", "丙", "乙"]);
  });

  it("空白永远排在最后，升序降序都一样", () => {
    const values = [HEADER, ["甲", 10, ""], ["乙", null, ""], ["丙", 30, ""], ["丁", "", ""]];
    for (const ascending of [true, false]) {
      const sorted = sortedMatrix(values, [{ column: 1, ascending }], true);
      // 两条空白行（null / 空串）都在末尾，非空行按方向排。
      expect(sorted.slice(1).map((row) => row[0]).slice(-2).sort()).toEqual(["丁", "乙"].sort());
    }
  });

  it("多键：先按第一键，相等时按第二键", () => {
    const values = [
      HEADER,
      ["甲", 1, 2],
      ["乙", 1, 1],
      ["丙", 2, 9]
    ];
    const sorted = sortedMatrix(values, [
      { column: 1, ascending: true },
      { column: 2, ascending: true }
    ], true);
    expect(sorted.slice(1).map((row) => row[0])).toEqual(["乙", "甲", "丙"]);
  });

  it("完全相等时保持原始先后（稳定排序）", () => {
    const values = [HEADER, ["甲", 1, "x"], ["乙", 1, "y"], ["丙", 1, "z"]];
    const sorted = sortedMatrix(values, [{ column: 1, ascending: true }], true);
    expect(sorted.slice(1).map((row) => row[0])).toEqual(["甲", "乙", "丙"]);
  });

  it("hasHeaders=false：首行也是数据，要参与排序", () => {
    const values = [["b", 2], ["a", 1]];
    const sorted = sortedMatrix(values, [{ column: 0, ascending: true }], false);
    expect(sorted.map((row) => row[0])).toEqual(["a", "b"]);
  });

  it("类型档：数字 < 文本 < 逻辑值", () => {
    expect(compareSortValues(1, "a")).toBeLessThan(0);
    expect(compareSortValues("a", true)).toBeLessThan(0);
    expect(compareSortValues(true, null)).toBeLessThan(0);
  });

  it("reorderRows 只按给定顺序搬运，不重算", () => {
    const values = [HEADER, ["甲", 1, ""], ["乙", 2, ""]];
    expect(reorderRows(values, [2, 1], true)).toEqual([HEADER, ["乙", 2, ""], ["甲", 1, ""]]);
  });

  it("顺序数组就是原始行号，可用于回读比对", () => {
    const values = [HEADER, ["甲", 3, ""], ["乙", 1, ""]];
    expect(sortRowOrder(values, [{ column: 1, ascending: true }], true)).toEqual([2, 1]);
  });

  it("没有键就拒绝，而不是悄悄返回原顺序", () => {
    expect(() => sortRowOrder([HEADER], [], true)).toThrow(/至少需要一个键/);
  });
});

describe("去重：保留首次出现，删除的行按原先后给出", () => {
  it("单列去重：保留第一次出现", () => {
    const values = [HEADER, ["甲", 1, "首次"], ["乙", 2, ""], ["甲", 1, "重复"]];
    const { kept, duplicates } = duplicateRowPartition(values, [0], true);
    expect(kept).toEqual([1, 2]);
    expect(duplicates).toEqual([3]);
  });

  it("多列组合：任一分量不同就不算重复", () => {
    const values = [HEADER, ["甲", 1, ""], ["甲", 2, ""], ["甲", 1, ""]];
    const { duplicates } = duplicateRowPartition(values, [0, 1], true);
    expect(duplicates).toEqual([3]);
  });

  it("空单元格与非空不混淆", () => {
    const values = [HEADER, ["甲", null, ""], ["甲", 0, ""], ["甲", "", ""]];
    const { duplicates } = duplicateRowPartition(values, [1], true);
    // null 与空串视为同一种"空"，0 与它们不同。
    expect(duplicates).toEqual([3]);
  });

  it("hasHeaders=false：首行参与去重", () => {
    const values = [["甲", 1], ["甲", 1]];
    expect(duplicateRowPartition(values, [0], false).duplicates).toEqual([1]);
  });

  it("deduplicatedMatrix 给出保留后的矩阵，表头在前", () => {
    const values = [HEADER, ["甲", 1, ""], ["甲", 1, ""], ["乙", 2, ""]];
    expect(deduplicatedMatrix(values, [0], true)).toEqual([HEADER, ["甲", 1, ""], ["乙", 2, ""]]);
  });

  it("没有依据列就拒绝", () => {
    expect(() => duplicateRowPartition([HEADER], [], true)).toThrow(/至少需要一列依据/);
  });
});
