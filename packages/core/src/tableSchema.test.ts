import { describe, expect, it } from "vitest";
import {
  detectHeaderIndex,
  detectSheetFields,
  isSyntheticField,
  looksLikeHeaderRow
} from "./tableSchema";

describe("local table schema detection", () => {
  it("detects a header below report title rows without returning data rows", () => {
    const values = [
      ["月度经营报表", null, null],
      [null, null, null],
      ["门店", "类别", "数量"],
      ["一店", "甲", 12]
    ];

    expect(detectHeaderIndex(values)).toBe(2);
    expect(detectSheetFields(values)).toEqual(["门店", "类别", "数量"]);
  });

  it("prefers a row containing fields requested by the local tool", () => {
    const values = [
      ["报表", "日期"],
      ["名称", "金额"],
      ["甲", 10]
    ];

    expect(detectHeaderIndex(values, ["名称", "金额"])).toBe(1);
  });
});

/**
 * 现场（2026-11-01，与「刷新后重复确认范围」分开记录）：
 * 用户问「只读取 A1:B5，告诉我 B 列五个数的合计」，本地算出来是 14 —— 首行被
 * 当成表头吞掉了；模型侧同时被告知「B1 是表头」，于是追问 B1 是否参与。
 * 两条链路的共同起点都是这里的判据：纯数字行不是表头证据。
 */
describe("无表头区域的表头判据（2026-11-01 现场）", () => {
  it("只有一列的纯数字区域没有表头，首行是数据", () => {
    expect(detectHeaderIndex([[1], [2], [3], [4], [5]])).toBe(-1);
    expect(detectSheetFields([[1], [2], [3], [4], [5]])).toEqual(["列1"]);
  });

  it("多列纯数字区域同样没有表头", () => {
    const values = [
      [1, 10],
      [2, 20]
    ];
    expect(detectHeaderIndex(values)).toBe(-1);
    expect(detectSheetFields(values)).toEqual(["列1", "列2"]);
  });

  it("文本列 + 数字列的区域从内容上分不清，交给调用方声明", () => {
    /*
     * `[["甲", 1], ["乙", 2]]` 有两种都成立的解释：数据区（无表头），或
     * 「文本表头 + 数字列」的表。内容里没有证据可以推翻任一种，所以这里
     * **保持既有行为**（第 0 行是表头），由调用方用 `QueryTableArguments.hasHeaders`
     * 明说 —— 用户说过“五个数”“第一行也算”时传 false，执行侧才把首行计入。
     * 猜错在这里的代价是静默少算一行，所以宁可显式传参。
     */
    const values = [
      ["甲", 1],
      ["乙", 2]
    ];
    expect(detectHeaderIndex(values)).toBe(0);
    // 但点名合成列名时仍然判无表头：那是我们自己造的名字，只可能来自无表头区域。
    expect(isSyntheticField("列2")).toBe(true);
    expect(isSyntheticField("数量")).toBe(false);
    expect(detectHeaderIndex(values, ["列2"])).toBe(-1);
  });

  it("文本表头仍然是表头（既有行为不变）", () => {
    expect(detectSheetFields([["数量"], [1], [2]])).toEqual(["数量"]);
    expect(detectHeaderIndex([["数量"], [1], [2]])).toBe(0);
  });

  it("表头判据只看文本与非空，不看数值大小", () => {
    expect(looksLikeHeaderRow(["数量", null])).toBe(true);
    expect(looksLikeHeaderRow([1, 2, 3])).toBe(false);
    expect(looksLikeHeaderRow([null, ""])).toBe(false);
    // 混排仍算表头：只要有一个文本单元格就当表头证据。
    expect(looksLikeHeaderRow([2024, "收入"])).toBe(true);
  });
});
