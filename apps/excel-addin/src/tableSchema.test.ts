import { describe, expect, it } from "vitest";
import { detectHeaderIndex, detectSheetFields } from "./tableSchema";

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
