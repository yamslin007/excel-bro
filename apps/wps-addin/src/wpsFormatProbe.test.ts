import { describe, expect, it } from "vitest";
import {
  runWpsFormatProbe,
  assertWpsFormatProbeTarget,
  readWpsFormatProbeTargetIdentity,
  sameProbeTarget,
  WPS_FORMAT_PROBE_GROUPS,
  XL_PATTERN,
  XL_H_ALIGN,
  XL_LINE_STYLE
} from "./wpsFormatProbe";
import type { WpsApplication, WpsRange, WpsWorksheet } from "./wpsHost";

/**
 * 替身宿主：一个带可读可写格式成员的单元格。
 *
 * 关键：替身必须按"真实 WPS 的语义"记账 —— 写进去的值要能被读回来，
 * 读不到（undefined/null）就真的是"没暴露"，而不是替身自己补一个默认值。
 */
function fakeRange(overrides: Partial<WpsRange> = {}): WpsRange {
  const font: Record<string, unknown> = { Bold: false, Italic: false, Name: "宋体", Size: 11, Color: 0, Underline: false };
  const interior: { Color: number; Pattern: unknown; PatternColor: unknown } = { Color: 0xffffff, Pattern: XL_PATTERN.none, PatternColor: 0 };
  const border: { LineStyle: unknown; Weight: unknown; ColorIndex: unknown } = { LineStyle: XL_LINE_STYLE.none, Weight: 1, ColorIndex: 1 };
  const state: Record<string, unknown> = {
    HorizontalAlignment: XL_H_ALIGN.left,
    VerticalAlignment: -4107, // xlVAlignBottom
    WrapText: false,
    RowHeight: 15,
    ColumnWidth: 8.43
  };
  return {
    Value2: "",
    Formula: undefined,
    Font: {
      get Bold() { return font.Bold; }, set Bold(v: unknown) { font.Bold = v; },
      get Italic() { return font.Italic; }, set Italic(v: unknown) { font.Italic = v; },
      get Name() { return font.Name; }, set Name(v: unknown) { font.Name = v; },
      get Size() { return font.Size; }, set Size(v: unknown) { font.Size = v; },
      get Color() { return font.Color; }, set Color(v: unknown) { font.Color = v; },
      get Underline() { return font.Underline; }, set Underline(v: unknown) { font.Underline = v; }
    },
    Interior: {
      get Color() { return interior.Color; }, set Color(v: number) { interior.Color = v; },
      get Pattern() { return interior.Pattern; }, set Pattern(v: unknown) { interior.Pattern = v; }
    },
    Borders: {
      Count: 6,
      Item: () => ({
        get LineStyle() { return border.LineStyle; }, set LineStyle(v: unknown) { border.LineStyle = v; },
        get Weight() { return border.Weight; }, set Weight(v: unknown) { border.Weight = v; },
        get ColorIndex() { return border.ColorIndex; }, set ColorIndex(v: unknown) { border.ColorIndex = v; }
      })
    },
    get HorizontalAlignment() { return state.HorizontalAlignment; }, set HorizontalAlignment(v: unknown) { state.HorizontalAlignment = v; },
    get VerticalAlignment() { return state.VerticalAlignment; }, set VerticalAlignment(v: unknown) { state.VerticalAlignment = v; },
    get WrapText() { return state.WrapText; }, set WrapText(v: unknown) { state.WrapText = v; },
    get RowHeight() { return state.RowHeight; }, set RowHeight(v: unknown) { state.RowHeight = v; },
    get ColumnWidth() { return state.ColumnWidth; }, set ColumnWidth(v: unknown) { state.ColumnWidth = v; },
    EntireRow: { AutoFit: () => undefined },
    EntireColumn: { AutoFit: () => undefined },
    ...overrides
  };
}

function fakeApplication(range: WpsRange = fakeRange()): WpsApplication {
  const sheet: WpsWorksheet = {
    Name: "探测表",
    Range: () => range
  };
  return {
    Version: "12.0",
    PluginStorage: { getItem: () => null, setItem: () => undefined },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined,
    ActiveWorkbook: {
      Name: "探测.xlsx",
      FullName: "C:\\Data\\探测.xlsx",
      Worksheets: { Count: 1, Item: () => sheet },
      ActiveSheet: sheet
    },
    ActiveSheet: sheet
  };
}

describe("WPS 格式能力受控探测", () => {
  it("在成员齐全的替身宿主上，可写成员全部 restore_verified，AutoFit 标 skipped，整轮基线恢复通过", () => {
    const range = fakeRange();
    const report = runWpsFormatProbe(fakeApplication(range));
    // 18 个成员 + 无整组基线失败（替身正确恢复），Interior/Borders 组基线都通过。
    expect(report.members.length).toBe(18);
    expect(report.baselineRestored).toBe(true);
    expect(report.aborted).toBe(false);
    for (const member of report.members) {
      if (member.member === "EntireRow.AutoFit" || member.member === "EntireColumn.AutoFit") {
        expect(member.status, member.member).toBe("skipped");
      } else {
        expect(member.status, member.member).toBe("restore_verified");
      }
    }
    // 探测后必须恢复原值，不留探测痕迹。
    expect(range.Font?.Bold).toBe(false);
    expect(range.Interior?.Color).toBe(0xffffff);
    expect(range.Interior?.Pattern).toBe(XL_PATTERN.none);
    expect(range.RowHeight).toBe(15);
    expect(range.HorizontalAlignment).toBe(XL_H_ALIGN.left);
  });

  it("runWpsFormatProbe 复用目标校验：非法目标（区域地址）在任何写入前抛错", () => {
    const range = fakeRange();
    const writes: string[] = [];
    // 记录所有 setter 调用，证明"抛错前零写入"。
    const rangeWithLog = new Proxy(range, {
      set(target, prop, value) {
        writes.push(String(prop));
        return Reflect.set(target, prop, value);
      }
    });
    expect(() => runWpsFormatProbe(fakeApplication(rangeWithLog), { probeCell: "A1:B2" })).toThrow(
      "不是单个单元格"
    );
    expect(writes.length).toBe(0);
  });

  it("直接调用入口拒绝有值的单元格（零写入）", () => {
    const range = fakeRange({ Value2: "已有数据" });
    expect(() => runWpsFormatProbe(fakeApplication(range))).toThrow("有值");
  });

  it("无法读取值（读取抛错）时拒绝，不把读取缺失解释为空白", () => {
    const range = fakeRange();
    Object.defineProperty(range, "Value2", { get: () => { throw new Error("读取值失败"); } });
    const check = assertWpsFormatProbeTarget(fakeApplication(range), "A1");
    expect(check.ok).toBe(false);
    if (!check.ok) expect(check.reason).toContain("无法读取");
  });

  it("无法读取公式（读取抛错）时拒绝", () => {
    const range = fakeRange();
    Object.defineProperty(range, "Formula", { get: () => { throw new Error("读取公式失败"); } });
    const check = assertWpsFormatProbeTarget(fakeApplication(range), "A1");
    expect(check.ok).toBe(false);
    if (!check.ok) expect(check.reason).toContain("无法读取");
  });

  it("宿主不暴露 Font 时，Font 成员判为 failed(读不到)", () => {
    const range = fakeRange({ Font: undefined });
    const report = runWpsFormatProbe(fakeApplication(range));
    const bold = report.members.find((m) => m.member === "Font.Bold");
    expect(bold?.status).toBe("failed");
    expect(bold?.failedAt).toBe("read");
  });

  it("哨兵与原值相同时，判为 skipped", () => {
    const range = fakeRange();
    Object.defineProperty(range, "RowHeight", {
      get: () => 20, // 与哨兵 20 相同
      set: () => undefined
    });
    const report = runWpsFormatProbe(fakeApplication(range));
    const row = report.members.find((m) => m.member === "RowHeight");
    expect(row?.status).toBe("skipped");
  });

  it("静默恢复失败（恢复写回不生效）判为 failed，且 aborted=true、后续成员标 skipped 零写入", () => {
    const range = fakeRange();
    // Font.Bold：写哨兵 true 生效；但恢复写回 false 时 setter 仍把值存成 true（写 false 不生效，
    // 模拟"恢复写回静默失败"——setter 不抛错，但读回来不是原值 false）。
    let stored = false;
    Object.defineProperty(range.Font, "Bold", {
      get: () => stored,
      set: () => { stored = true; } // 无论写什么，都存成 true
    });
    // 记录后续成员的写入，证明"恢复失败后零写入"。
    const writeLog: string[] = [];
    Object.defineProperty(range.Font, "Italic", {
      get: () => false,
      set: () => { writeLog.push("Italic"); }
    });
    const report = runWpsFormatProbe(fakeApplication(range));
    const bold = report.members.find((m) => m.member === "Font.Bold");
    expect(bold?.status).toBe("failed");
    expect(bold?.failedAt).toBe("restore_readback");
    // aborted 后后续成员（Font.Italic 等）标 skipped，且零写入。
    expect(report.aborted).toBe(true);
    expect(report.baselineRestored).toBe(false);
    const italic = report.members.find((m) => m.member === "Font.Italic");
    expect(italic?.status).toBe("skipped");
    expect(writeLog.length).toBe(0);
  });

  it("关联属性变化（恢复单个字段但连带改了同组其它字段）被发现，判整组基线失败", () => {
    const range = fakeRange();
    // Interior.Color 正常；Interior.Pattern 写哨兵时连带把 Color 改了（且恢复后 Color 没回去）。
    // 用有状态替身：设置 Pattern 会连带改 Color，恢复 Pattern 后 Color 仍残留哨兵色。
    let pattern: number = XL_PATTERN.none;
    let color = 0xffffff;
    Object.defineProperty(range.Interior, "Pattern", {
      get: () => pattern,
      set: (v: unknown) => {
        pattern = v as number;
        // 设 Pattern=solid 连带把 Color 改成 0xff0000；恢复 Pattern=none 时 Color 不还原（模拟连带残留）。
        if (v === XL_PATTERN.solid) color = 0xff0000;
      }
    });
    Object.defineProperty(range.Interior, "Color", {
      get: () => color,
      set: (v: number) => { color = v; }
    });
    const report = runWpsFormatProbe(fakeApplication(range));
    // 整组基线失败：Interior 组没有回到 [0xffffff, none]。
    const baselineFail = report.members.find((m) => m.failedAt === "baseline");
    expect(baselineFail).toBeTruthy();
    expect(baselineFail?.note).toContain("连带变化");
    expect(report.baselineRestored).toBe(false);
    expect(report.aborted).toBe(true);
  });

  it("恢复写回抛错时判为 failed(restore)，且 aborted", () => {
    const range = fakeRange();
    let writes = 0;
    let stored = 0xffffff;
    Object.defineProperty(range.Interior, "Color", {
      get: () => stored,
      set: (v: number) => {
        writes += 1;
        if (writes >= 2) throw new Error("恢复写入被拒");
        stored = v;
      }
    });
    const report = runWpsFormatProbe(fakeApplication(range));
    const color = report.members.find((m) => m.member === "Interior.Color");
    expect(color?.status).toBe("failed");
    expect(color?.failedAt).toBe("restore");
    expect(report.aborted).toBe(true);
  });

  it("无打开工作簿时抛错", () => {
    expect(() => runWpsFormatProbe({ ...fakeApplication(), ActiveWorkbook: undefined })).toThrow(
      "当前没有打开的工作簿"
    );
  });
});

describe("WPS 格式探测目标资格与身份", () => {
  it("合法目标：已保存 .xlsx 的空白单格通过", () => {
    const result = assertWpsFormatProbeTarget(fakeApplication(), "A1");
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.target.probeCell).toBe("A1");
      expect(result.target.workbookPath).toBe("C:\\Data\\探测.xlsx");
    }
  });

  it("区域地址（非单格）被拒绝", () => {
    const result = assertWpsFormatProbeTarget(fakeApplication(), "A1:B2");
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toContain("不是单个单元格");
  });

  it("未保存（非 .xlsx）工作簿被拒绝", () => {
    const app = fakeApplication();
    app.ActiveWorkbook!.FullName = "C:\\Data\\未保存.xlsm";
    const result = assertWpsFormatProbeTarget(app, "A1");
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toContain(".xlsx");
  });

  it("目标身份读取与比较：换簿后旧确认失效", () => {
    const app1 = fakeApplication();
    const id1 = readWpsFormatProbeTargetIdentity(app1, "A1");
    const app2 = fakeApplication();
    app2.ActiveWorkbook!.FullName = "C:\\Data\\另一个.xlsx";
    const id2 = readWpsFormatProbeTargetIdentity(app2, "A1");
    expect(id1).not.toBeNull();
    expect(id2).not.toBeNull();
    expect(sameProbeTarget(id1!, id2!)).toBe(false);
  });
});

describe("WPS 格式探测 · 能力组选择", () => {
  it("指定 groups 只探测选中组，未选中组不产生成员", () => {
    const range = fakeRange();
    const report = runWpsFormatProbe(fakeApplication(range), { groups: ["font"] });
    expect(report.groups).toEqual(["font"]);
    // 只含 font 组 6 个成员，不含 alignment/size/fill/borders/autofit。
    const names = report.members.map((m) => m.member);
    expect(names.every((n) => n.startsWith("Font."))).toBe(true);
    expect(names.length).toBe(6);
    expect(report.baselineRestored).toBe(true);
    expect(report.aborted).toBe(false);
  });

  it("指定多组时按 WPS_FORMAT_PROBE_GROUPS 顺序探测", () => {
    const range = fakeRange();
    const report = runWpsFormatProbe(fakeApplication(range), { groups: ["borders", "font"] });
    // groups 字段保留用户传入顺序，但成员按组内定义顺序产出（font 在 borders 之前）。
    expect(report.groups).toEqual(["borders", "font"]);
    const first = report.members[0].member;
    expect(first).toBe("Font.Bold");
  });

  it("只探测 size 组时，RowHeight/ColumnWidth 恢复后原值不变", () => {
    const range = fakeRange();
    const report = runWpsFormatProbe(fakeApplication(range), { groups: ["size"] });
    expect(report.members.map((m) => m.member)).toEqual(["RowHeight", "ColumnWidth"]);
    expect(range.RowHeight).toBe(15);
    expect(range.ColumnWidth).toBe(8.43);
    expect(report.baselineRestored).toBe(true);
  });

  it("组内恢复失败仍触发 abort：只选 size 组时 ColumnWidth 静默恢复失败，RowHeight 在后则 skipped", () => {
    const range = fakeRange();
    // RowHeight 在 ColumnWidth 之前（sizeMembers 顺序），所以让 RowHeight 恢复失败，
    // 后续 ColumnWidth 应 skipped 零写入。
    let stored = 15;
    let widthWrites = 0;
    Object.defineProperty(range, "RowHeight", {
      get: () => stored,
      set: () => { stored = 20; } // 恢复写回 15 不生效，静默失败
    });
    Object.defineProperty(range, "ColumnWidth", {
      get: () => 8.43,
      set: () => { widthWrites += 1; }
    });
    const report = runWpsFormatProbe(fakeApplication(range), { groups: ["size"] });
    const row = report.members.find((m) => m.member === "RowHeight");
    const col = report.members.find((m) => m.member === "ColumnWidth");
    expect(row?.status).toBe("failed");
    expect(col?.status).toBe("skipped");
    expect(widthWrites).toBe(0);
    expect(report.aborted).toBe(true);
    expect(report.baselineRestored).toBe(false);
  });

  it("groups 为全部组时等价于默认探测（18 成员 + AutoFit skipped）", () => {
    const range = fakeRange();
    const report = runWpsFormatProbe(fakeApplication(range), { groups: [...WPS_FORMAT_PROBE_GROUPS] });
    expect(report.groups).toEqual([...WPS_FORMAT_PROBE_GROUPS]);
    expect(report.members.length).toBe(18);
    expect(report.baselineRestored).toBe(true);
  });
});
