import { describe, expect, it } from "vitest";
import {
  assertWpsVerifiedSetAlignment,
  assertWpsVerifiedResizeRange,
  describeWpsFormatAction,
  wpsColorFromHex,
  wpsHorizontalAlignment,
  WPS_VERIFIED_FONT_PROPERTIES,
  WPS_VERIFIED_ALIGNMENT_PROPERTIES,
  WPS_VERIFIED_RESIZE_PROPERTIES
} from "./wpsFormatCapability";

describe("WPS 格式动作已验证属性白名单", () => {
  it("只开放实机已验证的属性集合，不因部分通过而全开", () => {
    // 字体：只有 bold/italic/name/size/color（underline 未验证，不在协议字段里）。
    expect([...WPS_VERIFIED_FONT_PROPERTIES].sort()).toEqual(
      ["bold", "color", "italic", "name", "size"].sort()
    );
    // 对齐：只有 horizontal/wrapText（vertical 未验证）。
    expect([...WPS_VERIFIED_ALIGNMENT_PROPERTIES].sort()).toEqual(
      ["horizontal", "wrapText"].sort()
    );
    // 尺寸：只有 rowHeight（columnWidth 未验证）。
    expect([...WPS_VERIFIED_RESIZE_PROPERTIES].sort()).toEqual(["rowHeight"].sort());
  });

  it("setAlignment 的垂直对齐支持合法枚举、拒绝未知枚举", () => {
    expect(() => assertWpsVerifiedSetAlignment({ horizontal: "center", vertical: "center" })).not.toThrow();
    expect(() => assertWpsVerifiedSetAlignment({ vertical: "unknown" })).toThrow("垂直对齐");
  });

  it("setAlignment 不带 vertical 时放行", () => {
    expect(() => assertWpsVerifiedSetAlignment({ horizontal: "center", wrapText: true }))
      .not.toThrow();
  });

  it("resizeRange 带 columnWidth 时明确拒绝（不静默忽略）", () => {
    expect(() => assertWpsVerifiedResizeRange({ rowHeight: 20, columnWidth: 12 }))
      .toThrow(/columnWidth/i);
  });

  it("resizeRange 只带 rowHeight 时放行", () => {
    expect(() => assertWpsVerifiedResizeRange({ rowHeight: 20 })).not.toThrow();
  });

  it("水平对齐字符串映射到 WPS 数字枚举", () => {
    expect(wpsHorizontalAlignment("center")).toBe(-4108);
    expect(wpsHorizontalAlignment("general")).toBe(1);
    expect(wpsHorizontalAlignment("left")).toBe(-4131);
  });

  it("未知水平对齐抛错", () => {
    expect(() => wpsHorizontalAlignment("middle")).toThrow();
  });

  it("#RRGGBB 颜色转 WPS 数字颜色", () => {
    expect(wpsColorFromHex("#FF0000")).toBe(0x0000ff);
    expect(wpsColorFromHex("#0000ff")).toBe(0xff0000);
  });

  it("非法颜色抛错", () => {
    expect(() => wpsColorFromHex("red")).toThrow();
    expect(() => wpsColorFromHex("#FFF")).toThrow();
  });

  it("describeWpsFormatAction 明确描述格式变化（不冒充写入值）", () => {
    expect(describeWpsFormatAction({ type: "setFont", bold: true })).toBe("设置字体（加粗）");
    expect(describeWpsFormatAction({ type: "setFont", name: "黑体", size: 14 }))
      .toBe("设置字体（字体改为「黑体」、字号改为 14）");
    expect(describeWpsFormatAction({ type: "setAlignment", horizontal: "center", wrapText: true }))
      .toBe("设置对齐（水平对齐改为「居中」、开启自动换行）");
  });

  it("describeWpsFormatAction 明示整行/整列影响", () => {
    expect(describeWpsFormatAction({ type: "resizeRange", rowHeight: 24 }))
      .toBe("调整尺寸（行高改为 24（影响整行））");
    expect(describeWpsFormatAction({ type: "resizeRange", columnWidth: 12 }))
      .toBe("调整尺寸（列宽改为 12（影响整列））");
  });
});
