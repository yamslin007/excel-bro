import { describe, expect, it } from "vitest";
import type { ExcelAction } from "./contracts";
import {
  EXCEL_DIRECT_FORMAT_REQUIREMENT,
  EXCEL_DISPLAYED_FORMAT_REQUIREMENT,
  EXCEL_FORMAT_ACTION_TYPES,
  blockedFormatActionIndexes,
  compatFormatActionIndexes,
  excelFormatActionFields,
  excelFormatCapabilityNotice,
  expectedDirectFormat,
  formatDirectValueMatches,
  isExcelFormatActionType,
  resolveExcelFormatCapability
} from "./excelFormatCapability";

/**
 * 格式能力模块的**纯函数**判据。
 *
 * 这一层不依赖 `Office`/`Excel` 全局：能力判定、受影响属性、期望值、取值比对都是纯计算。
 * 把它单独钉住，是因为**能力检测前移**与**副本回读核对**都建立在这些函数之上 ——
 * 它们一旦判错，用户在预览阶段看到的提示、副本里的核对结论就全错，而且错得静默。
 */

const FONT_ACTION = (overrides: Partial<Extract<ExcelAction, { type: "setFont" }>> = {}) =>
  ({ type: "setFont", sheet: "S", range: "A1", ...overrides }) as ExcelAction;

/** 宿主声明支持的 ExcelApi 版本集合 → `isSetSupported` 谓词。 */
function hostWith(maxMinor: number): (set: string, version: string) => boolean {
  const available = Array.from({ length: maxMinor }, (_, index) => `1.${index + 1}`);
  return (set, version) => {
    if (set !== "ExcelApi") return false;
    const [major, minor] = version.split(".").map(Number);
    return available.some((item) => {
      const [itemMajor, itemMinor] = item.split(".").map(Number);
      return itemMajor > major || (itemMajor === major && itemMinor >= minor);
    });
  };
}

describe("resolveExcelFormatCapability · 三条路径", () => {
  it("支持 1.19 ⇒ displayed，没有任何缺失或受堵动作", () => {
    const capability = resolveExcelFormatCapability(hostWith(19));
    expect(capability.path).toBe("displayed");
    expect(capability.missingApis).toEqual([]);
    expect(capability.compatActionTypes).toEqual([]);
    expect(capability.blockedActionTypes).toEqual([]);
  });

  it("最高 1.17（批量许可永久版）⇒ direct：缺 1.19、六个格式动作全走兼容路径、无一被堵", () => {
    const capability = resolveExcelFormatCapability(hostWith(17));
    expect(capability.path).toBe("direct");
    expect(capability.missingApis).toEqual([EXCEL_DISPLAYED_FORMAT_REQUIREMENT]);
    expect(capability.compatActionTypes).toEqual([...EXCEL_FORMAT_ACTION_TYPES]);
    expect(capability.blockedActionTypes).toEqual([]);
    // 兼容路径的说明必须把"怎么恢复"和"哪种区域会拒绝"讲清楚。
    expect(capability.note).toContain("直接格式");
    expect(capability.note).toContain("清空");
    expect(capability.note).toContain("条件格式");
    expect(capability.note).toContain("合并单元格");
  });

  it("连 1.9 都没有 ⇒ unavailable：两类 API 都缺，六个格式动作全被堵", () => {
    const capability = resolveExcelFormatCapability(hostWith(8));
    expect(capability.path).toBe("unavailable");
    expect(capability.missingApis).toEqual([
      EXCEL_DISPLAYED_FORMAT_REQUIREMENT,
      EXCEL_DIRECT_FORMAT_REQUIREMENT
    ]);
    expect(capability.blockedActionTypes).toEqual([...EXCEL_FORMAT_ACTION_TYPES]);
    expect(capability.compatActionTypes).toEqual([]);
  });

  it("判定用要求集、按传入谓词取版本：谓词收到的 set 恒为 ExcelApi", () => {
    const seen: Array<[string, string]> = [];
    resolveExcelFormatCapability((set, version) => {
      seen.push([set, version]);
      return false;
    });
    expect(seen).toEqual([
      ["ExcelApi", "1.19"],
      ["ExcelApi", "1.9"]
    ]);
  });
});

describe("excelFormatActionFields · 只列本次真要改的属性", () => {
  it("setFill / setBorders / setNumberFormat 各只有一个字段", () => {
    expect(excelFormatActionFields({ type: "setFill", sheet: "S", range: "A1", color: "#000000" })).toEqual(["fill"]);
    expect(
      excelFormatActionFields({ type: "setBorders", sheet: "S", range: "A1", sides: ["top"] })
    ).toEqual(["borders"]);
    expect(
      excelFormatActionFields({ type: "setNumberFormat", sheet: "S", range: "A1", formatCode: "0.00" })
    ).toEqual(["numberFormat"]);
  });

  it("setFont 只算给了的字段；bold:false 也算「要改」", () => {
    expect(excelFormatActionFields(FONT_ACTION({ bold: true }))).toEqual(["fontBold"]);
    expect(excelFormatActionFields(FONT_ACTION({ bold: false }))).toEqual(["fontBold"]);
    expect(excelFormatActionFields(FONT_ACTION({ name: "Arial", size: 12 }))).toEqual([
      "fontName",
      "fontSize"
    ]);
    // 一个都没给 ⇒ 空（调用方据此判"空动作"）。
    expect(excelFormatActionFields(FONT_ACTION({}))).toEqual([]);
  });

  it("setAlignment / resizeRange 各按实际字段取", () => {
    expect(
      excelFormatActionFields({ type: "setAlignment", sheet: "S", range: "A1", horizontal: "center" })
    ).toEqual(["horizontalAlignment"]);
    expect(
      excelFormatActionFields({ type: "setAlignment", sheet: "S", range: "A1", horizontal: "center", wrapText: false })
    ).toEqual(["horizontalAlignment", "wrapText"]);
    expect(excelFormatActionFields({ type: "resizeRange", sheet: "S", range: "A1", rowHeight: 30 })).toEqual([
      "rowHeight"
    ]);
    expect(
      excelFormatActionFields({ type: "resizeRange", sheet: "S", range: "A1", rowHeight: 30, columnWidth: 12 })
    ).toEqual(["rowHeight", "columnWidth"]);
  });

  it("非格式动作返回空数组（能力检查不该牵连值/公式动作）", () => {
    expect(
      excelFormatActionFields({ type: "writeValues", sheet: "S", range: "A1", values: [[1]] } as ExcelAction)
    ).toEqual([]);
    expect(isExcelFormatActionType("writeValues")).toBe(false);
    expect(isExcelFormatActionType("setFont")).toBe(true);
  });
});

describe("受影响动作定位与前移提示", () => {
  const plan: ExcelAction[] = [
    { type: "writeValues", sheet: "S", range: "A1", values: [[1]] } as ExcelAction,
    { type: "setFont", sheet: "S", range: "A2", bold: true } as ExcelAction,
    { type: "setFill", sheet: "S", range: "A3", color: "#FF0000" } as ExcelAction
  ];

  it("direct 宿主：兼容动作下标是格式动作的位置", () => {
    const capability = resolveExcelFormatCapability(hostWith(17));
    expect(compatFormatActionIndexes(plan, capability)).toEqual([1, 2]);
    expect(blockedFormatActionIndexes(plan, capability)).toEqual([]);
  });

  it("unavailable 宿主：被堵动作下标同样定位到具体步骤", () => {
    const capability = resolveExcelFormatCapability(hostWith(8));
    expect(blockedFormatActionIndexes(plan, capability)).toEqual([1, 2]);
    expect(compatFormatActionIndexes(plan, capability)).toEqual([]);
  });

  it("displayed 宿主：无话可说，提示为 null", () => {
    expect(excelFormatCapabilityNotice(plan, resolveExcelFormatCapability(hostWith(19)))).toBeNull();
  });

  it("direct 宿主：info 提示，首屏短说明 + 详情点缺 1.19 与受影响步骤", () => {
    const notice = excelFormatCapabilityNotice(plan, resolveExcelFormatCapability(hostWith(17)));
    expect(notice).not.toBeNull();
    expect(notice!.tone).toBe("info");
    // 首屏是短说明，技术原因（缺哪个 API、逐格恢复细节）收进 detail。
    expect(notice!.summary).toContain("第 2 步");
    expect(notice!.summary).toContain("第 3 步");
    expect(notice!.detail).toContain(EXCEL_DISPLAYED_FORMAT_REQUIREMENT);
    // 纯文本，不含 Markdown 标记。
    expect(notice!.summary).not.toContain("**");
    expect(notice!.detail).not.toContain("**");
  });

  it("unavailable 宿主：warning 提示；计划里没有格式动作时为 null", () => {
    const capability = resolveExcelFormatCapability(hostWith(8));
    const notice = excelFormatCapabilityNotice(plan, capability);
    expect(notice!.tone).toBe("warning");
    expect(notice!.detail).toContain(EXCEL_DIRECT_FORMAT_REQUIREMENT);
    expect(notice!.summary).toContain("第 2 步");
    // 只有值/公式动作 ⇒ 不受影响 ⇒ 不打扰用户。
    const valueOnly: ExcelAction[] = [
      { type: "writeValues", sheet: "S", range: "A1", values: [[1]] } as ExcelAction
    ];
    expect(excelFormatCapabilityNotice(valueOnly, capability)).toBeNull();
  });
});

describe("expectedDirectFormat · 逐格直接格式的期望值", () => {
  it("setFill 期望「有填充 + 该色」；resizeRange 不在逐格矩阵里", () => {
    expect(
      expectedDirectFormat({ type: "setFill", sheet: "S", range: "A1", color: "#FF0000" })
    ).toEqual([{ field: "fill", value: { pattern: "Solid", color: "#FF0000" } }]);
    expect(expectedDirectFormat({ type: "resizeRange", sheet: "S", range: "A1", rowHeight: 30 })).toEqual([]);
    expect(
      expectedDirectFormat({ type: "setNumberFormat", sheet: "S", range: "A1", formatCode: "0.00" })
    ).toEqual([]);
  });

  it("setFont / setAlignment 只含给了的属性", () => {
    expect(expectedDirectFormat(FONT_ACTION({ bold: true, size: 14 }))).toEqual([
      { field: "fontBold", value: true },
      { field: "fontSize", value: 14 }
    ]);
    expect(
      expectedDirectFormat({ type: "setAlignment", sheet: "S", range: "A1", horizontal: "center" })
    ).toEqual([{ field: "horizontalAlignment", value: "center" }]);
  });

  it("setBorders 用与执行器一致的默认补齐", () => {
    expect(
      expectedDirectFormat({ type: "setBorders", sheet: "S", range: "A1", sides: ["top"] })
    ).toEqual([
      { field: "borders", value: { top: { style: "continuous", color: "#000000", weight: "thin" } } }
    ]);
  });
});

describe("formatDirectValueMatches · 与执行器回读同口径", () => {
  it("枚举忽略大小写与分隔符，颜色忽略大小写", () => {
    expect(formatDirectValueMatches("horizontalAlignment", "Center", "center")).toBe(true);
    expect(formatDirectValueMatches("horizontalAlignment", "CENTER", "center")).toBe(true);
    expect(formatDirectValueMatches("fontUnderline", "single", "Single")).toBe(true);
    expect(
      formatDirectValueMatches("fill", { pattern: "SOLID", color: "#aabbcc" }, { pattern: "Solid", color: "#AABBCC" })
    ).toBe(true);
    expect(formatDirectValueMatches("fontColor", "#AABBCC", "#aabbcc")).toBe(true);
  });

  it("字号允许 0.1 容差", () => {
    expect(formatDirectValueMatches("fontSize", 12, 12.05)).toBe(true);
    expect(formatDirectValueMatches("fontSize", 12, 12.5)).toBe(false);
    expect(formatDirectValueMatches("fontSize", undefined, 12)).toBe(false);
  });

  it("**未直接设置（null）≠ 已设置成默认值**：期望是「直接设置过」时一律判不一致", () => {
    // bold 未直接设置（null）不该等于 bold:false；对齐未直接设置不该等于 "General"。
    expect(formatDirectValueMatches("fontBold", null, true)).toBe(false);
    expect(formatDirectValueMatches("fontBold", null, false)).toBe(false);
    expect(formatDirectValueMatches("horizontalAlignment", null, "General")).toBe(false);
    expect(formatDirectValueMatches("fill", null, { pattern: "Solid", color: "#FF0000" })).toBe(false);
    // 反向锚点：直接设置成 false 与期望 false 一致。
    expect(formatDirectValueMatches("fontBold", false, false)).toBe(true);
  });

  it("边框：期望的每条边都必须在实际值里对上", () => {
    const actual = { top: { style: "Continuous", color: "#000000", weight: "Thin" } };
    expect(
      formatDirectValueMatches("borders", actual, { top: { style: "continuous", color: "#000000", weight: "thin" } })
    ).toBe(true);
    // 少一条期望边 ⇒ 不一致。
    expect(
      formatDirectValueMatches("borders", actual, {
        top: { style: "continuous", color: "#000000", weight: "thin" },
        bottom: { style: "continuous", color: "#000000", weight: "thin" }
      })
    ).toBe(false);
  });
});
