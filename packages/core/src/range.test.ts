import { describe, expect, it } from "vitest";
import {
  columnIndexFromLetter,
  columnLetterFromIndex,
  firstCellOfRange,
  startColumnOfRange,
  suggestWriteTarget
} from "./range";

// 迁入 core 前本模块没有测试。这里按迁移当时的实际行为立断言，
// 作为后续改动的回归网，而不是事后追认新行为。
describe("firstCellOfRange", () => {
  it("接受带表名、带 $ 与单格三种写法", () => {
    expect(firstCellOfRange("Sheet1!E2:E50")).toBe("E2");
    expect(firstCellOfRange("$E$2:$E$50")).toBe("E2");
    expect(firstCellOfRange("E2")).toBe("E2");
  });

  it("非单元格地址返回 null", () => {
    expect(firstCellOfRange("2E")).toBeNull();
    expect(firstCellOfRange("")).toBeNull();
    expect(firstCellOfRange("Sheet1!")).toBeNull();
  });
});

describe("列号与列字母互转", () => {
  it("在进位边界上闭合", () => {
    const cases: ReadonlyArray<[number, string]> = [
      [0, "A"],
      [25, "Z"],
      [26, "AA"],
      [51, "AZ"],
      [52, "BA"],
      [701, "ZZ"],
      [702, "AAA"]
    ];
    for (const [index, letters] of cases) {
      expect(columnLetterFromIndex(index)).toBe(letters);
      expect(columnIndexFromLetter(letters)).toBe(index);
    }
  });

  it("列字母不区分大小写", () => {
    expect(columnIndexFromLetter("aa")).toBe(columnIndexFromLetter("AA"));
  });
});

describe("startColumnOfRange", () => {
  it("取起始列并归一为大写", () => {
    expect(startColumnOfRange("A1:E20")).toBe("A");
    expect(startColumnOfRange("Sheet1!c2:e9")).toBe("C");
    expect(startColumnOfRange("$AB$3")).toBe("AB");
  });

  it("空值与非法地址返回 null", () => {
    expect(startColumnOfRange(null)).toBeNull();
    expect(startColumnOfRange("1A")).toBeNull();
  });
});

describe("suggestWriteTarget", () => {
  it("落在数据区右侧第一空列、首行下一行", () => {
    expect(suggestWriteTarget("A1:D20")).toBe("E2");
    expect(suggestWriteTarget("Sheet1!B3:C9")).toBe("D4");
  });

  it("单格区域按该格右侧一列", () => {
    expect(suggestWriteTarget("C5")).toBe("D6");
  });

  it("跨进位时列字母正确", () => {
    expect(suggestWriteTarget("A1:Z10")).toBe("AA2");
  });

  it("空值或解析失败退回 A2", () => {
    expect(suggestWriteTarget(null)).toBe("A2");
    expect(suggestWriteTarget("not-a-range")).toBe("A2");
  });
});
