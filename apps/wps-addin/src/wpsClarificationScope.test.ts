import { describe, expect, it } from "vitest";
import {
  wpsScopeChanged,
  wpsHeaderFingerprint,
  wpsSheetIdentity
} from "./wpsClarificationScope";

/** 测试工作簿名 —— 指纹里带它，所以每个用例都要显式给。 */
const WORKBOOK = "测试.xlsx";

const identity = (
  sheets: Array<{ index: number; name: string }>,
  workbookName = WORKBOOK
) => wpsSheetIdentity({ workbookName, sheets });

describe("WPS 授权身份指纹", () => {
  it("同一张表同一顺序得到同一个指纹", () => {
    expect(identity([{ index: 1, name: "订单" }, { index: 2, name: "明细" }])).toBe(
      identity([{ index: 1, name: "订单" }, { index: 2, name: "明细" }])
    );
  });

  it("同一张表**改名**后指纹变化 —— 不为了不断对话牺牲授权边界", () => {
    expect(identity([{ index: 1, name: "订单" }])).not.toBe(
      identity([{ index: 1, name: "订单（新）" }])
    );
  });

  it("工作表重排后指纹变化，即使工作表集合没变", () => {
    expect(identity([{ index: 1, name: "订单" }, { index: 2, name: "明细" }])).not.toBe(
      identity([{ index: 1, name: "明细" }, { index: 2, name: "订单" }])
    );
  });

  it("同索引换成另一张表时指纹变化", () => {
    expect(identity([{ index: 1, name: "订单" }])).not.toBe(
      identity([{ index: 1, name: "客户" }])
    );
  });

  it("工作表集合变化时指纹变化", () => {
    expect(identity([{ index: 1, name: "订单" }])).not.toBe(
      identity([{ index: 1, name: "订单" }, { index: 2, name: "明细" }])
    );
  });

  it("指纹里带宿主与工作簿名 —— 换工作簿不该沿用上一本的澄清", () => {
    expect(identity([{ index: 1, name: "订单" }], "A.xlsx")).not.toBe(
      identity([{ index: 1, name: "订单" }], "B.xlsx")
    );
  });

  it("输出可持久化：是字符串，且不含制表符与换行", () => {
    const value = identity([{ index: 1, name: "订单" }]);
    expect(typeof value).toBe("string");
    expect(value).not.toContain("\n");
    expect(value).not.toContain("\t");
  });
});

describe("WPS 表头指纹", () => {
  it("表头改名时指纹变化", () => {
    expect(wpsHeaderFingerprint([["类别", "归属"]])).not.toBe(
      wpsHeaderFingerprint([["类型", "归属"]])
    );
  });

  it("表头顺序变化时指纹变化", () => {
    expect(wpsHeaderFingerprint([["类别", "归属"]])).not.toBe(
      wpsHeaderFingerprint([["归属", "类别"]])
    );
  });

  it("表头行数变化时指纹变化", () => {
    expect(wpsHeaderFingerprint([["类别", "归属"]])).not.toBe(
      wpsHeaderFingerprint([["类别", "归属"], ["办公用品", "A"]])
    );
  });

  it("同样的表头得到同一个指纹，且只由传入的表头行构成", () => {
    const value = wpsHeaderFingerprint([["类别", "归属"]]);
    expect(value).toBe(wpsHeaderFingerprint([["类别", "归属"]]));
    // 数据行的值不在指纹里 —— 它会被持久化进对话历史，不能带单元格值。
    expect(value).not.toContain("办公用品");
  });

  it("空白与 null 都按空处理，不会因为取到 undefined 就抖动", () => {
    expect(wpsHeaderFingerprint([["类别", null]])).toBe(
      wpsHeaderFingerprint([["类别", undefined]])
    );
  });
});

describe("范围变化的判定", () => {
  it("没有记录指纹时不判为变化（兼容旧卡片）", () => {
    expect(wpsScopeChanged(undefined, "anything")).toBe(false);
  });

  it("指纹相同时不判为变化", () => {
    expect(wpsScopeChanged("same", "same")).toBe(false);
  });

  it("指纹不同时判为变化", () => {
    expect(wpsScopeChanged("before", "after")).toBe(true);
  });
});
