// @vitest-environment jsdom
/**
 * 审阅标记的生命周期（WPS 端）。
 *
 * 模拟范围：只替身 **WPS 的 FormatConditions 与工作簿对象**；标记模块是真的。
 * 这里**不**走"宿主不支持标记"的降级分支：降级只证明"没有标记"，证明不了
 * 标记存在时的归属判定、部分失败与刷新认领。
 *
 * WPS 没有可读的规则 id，归属凭据是本次随机 `Formula1`。所以"同一区域里
 * 另一条规则"必须活下来 —— 那是"不按区域/颜色批量删"的可观测后果。
 */
import { beforeEach, describe, expect, it } from "vitest";
import type { WpsApplication, WpsWorkbook } from "./wpsHost";
import {
  clearWpsResultMarks,
  markWpsResult,
  recoverWpsResultMarks,
  clearReviewMarksFromCopySheet
} from "./wpsResultReviewMarkers";

type RuleState = { formula: string; deleted: boolean; interior: { color: number; pattern: unknown } | null; borders: Record<number, { color: number; weight: unknown }> };

function installHost(options: { failOnAdd?: number; failOnDelete?: boolean; ranges?: string[] } = {}) {
  const byRange = new Map<string, RuleState[]>();
  const ranges = options.ranges ?? ["A1:B1"];
  for (const range of ranges) byRange.set(range, []);
  let seq = 0;

  const rangeObject = (range: string) => {
    const list = byRange.get(range) ?? [];
    const conditions = {
      get Count() { return list.filter(rule => !rule.deleted).length; },
      // 真宿主在 Add 时就带上公式；替身必须同样回填，否则归属判据永远对不上。
      Add: (_type: number, _operator?: unknown, formula?: string) => {
        seq += 1;
        if (options.failOnAdd === seq) throw new Error("第 2 处标记创建失败");
        const state: RuleState = { formula: formula ?? "", deleted: false, interior: null, borders: {} };
        list.push(state);
        return {
          get Formula1() { return state.formula; },
          set Formula1(next: string) { state.formula = next; },
          Borders: { Item: (edge: number) => ({
            get Color() { return state.borders[edge]?.color ?? 0; },
            set Color(v: number) { state.borders[edge] = { color: v, weight: state.borders[edge]?.weight }; },
            get Weight() { return state.borders[edge]?.weight; },
            set Weight(v: unknown) { state.borders[edge] = { color: state.borders[edge]?.color ?? 0, weight: v }; },
            get LineStyle() { return 1; },
            set LineStyle(_v: unknown) {}
          }) },
          Interior: {
            get Color() { return state.interior?.color ?? 0; },
            set Color(v: number) { state.interior = { color: v, pattern: state.interior?.pattern }; },
            get Pattern() { return state.interior?.pattern; },
            set Pattern(v: unknown) { state.interior = { color: state.interior?.color ?? 0, pattern: v }; }
          },
          Delete: () => {
            if (options.failOnDelete) throw new Error("删除被宿主拒绝");
            state.deleted = true;
          }
        };
      },
      Item: (index: number) => {
        const live = list.filter(rule => !rule.deleted)[index - 1];
        return {
          get Formula1() { return live?.formula; },
          Delete: () => {
            if (options.failOnDelete) throw new Error("删除被宿主拒绝");
            if (live) live.deleted = true;
          }
        };
      }
    };
    return { FormatConditions: conditions };
  };

  const sheet = { Name: "数据", Range: (range: string) => rangeObject(range) };
  const workbook = {
    Name: "测试.xlsx", FullName: "C:\\Data\\测试.xlsx", Path: "C:\\Data", ReadOnly: false,
    Worksheets: { Count: 1, Item: () => sheet },
    ActiveSheet: sheet
  } as unknown as WpsWorkbook;
  const application = { Version: "12.0", ActiveWorkbook: workbook, ActiveSheet: sheet } as unknown as WpsApplication;
  return {
    application,
    live: (range = ranges[0]) => (byRange.get(range) ?? []).filter(rule => !rule.deleted).map(rule => rule.formula),
    state: (range = ranges[0]) => (byRange.get(range) ?? []).filter(rule => !rule.deleted),
    pushForeign: (range: string, formula: string) => { (byRange.get(range) ?? []).push({ formula, deleted: false, interior: null, borders: {} }); }
  };
}

const prepared = {
  authorization: { workbookFullName: "C:\\Data\\测试.xlsx", sheets: [{ name: "数据", index: 1 }] },
  actions: [
    { sheetName: "数据", range: "A1:B1", type: "writeValues" },
    { sheetName: "数据", range: "C1:D1", type: "writeValues" }
  ]
} as never;

beforeEach(() => { localStorage.clear(); });

describe("WPS 审阅标记", () => {
  it("逐条落盘：创建部分成功后失败，已创建的可追踪并清理", () => {
    const host = installHost({ ranges: ["A1:B1", "C1:D1"], failOnAdd: 2 });
    expect(() => markWpsResult(host.application, prepared)).toThrow(/第 2 处标记创建失败/);
    // 第一条被清掉，记录不留。
    expect(host.live("A1:B1")).toHaveLength(0);
    expect(localStorage.getItem("excel-bro.wps.reviewMarks.v1")).toBeNull();
  });

  it("创建失败且清理也失败：保留归属记录（只有表名/区域/公式，没有业务值）", () => {
    const host = installHost({ ranges: ["A1:B1", "C1:D1"], failOnAdd: 2, failOnDelete: true });
    expect(() => markWpsResult(host.application, prepared)).toThrow(/暂未移除/);
    const stored = JSON.parse(localStorage.getItem("excel-bro.wps.reviewMarks.v1") ?? "null");
    expect(stored.ranges).toEqual([{ sheet: "数据", range: "A1:B1" }]);
    expect(JSON.stringify(stored)).not.toContain('"values"');
  });

  it("刷新后只按本次随机公式认领，同区域其它规则不碰", () => {
    const host = installHost({ ranges: ["A1:B1"] });
    const marks = markWpsResult(host.application, prepared);
    // 用户/别的任务在同一区域另建了一条规则。
    host.pushForeign("A1:B1", "=A1>0");
    expect(host.live("A1:B1")).toHaveLength(2);

    const note = recoverWpsResultMarks(host.application);
    expect(note).toContain("已清理");
    // 只剩别人的那条 —— 证明"不按区域批量删"。
    expect(host.live("A1:B1")).toEqual(["=A1>0"]);
    expect(localStorage.getItem("excel-bro.wps.reviewMarks.v1")).toBeNull();
    expect(marks.formula).toContain("EB_REVIEW_");
  });

  it("工作簿身份不符：不碰规则、不删记录", () => {
    const host = installHost();
    markWpsResult(host.application, prepared);
    const other = { ...host.application, ActiveWorkbook: { ...(host.application.ActiveWorkbook as object), FullName: "C:\\Data\\Other.xlsx", Name: "Other.xlsx" } } as unknown as WpsApplication;
    const note = recoverWpsResultMarks(other);
    expect(note).toContain("另一本工作簿");
    expect(host.live()).toHaveLength(1);
    expect(localStorage.getItem("excel-bro.wps.reviewMarks.v1")).not.toBeNull();
  });

  it("确认清理只删本次公式对应的规则", () => {
    const host = installHost();
    const marks = markWpsResult(host.application, prepared);
    host.pushForeign("A1:B1", "=B1>0");
    clearWpsResultMarks(host.application, marks);
    expect(host.live()).toEqual(["=B1>0"]);
  });
});
describe("WPS 审阅标记呈现：底色与边框", () => {
  it("不修改填充的动作：浅黄底色 + 醒目边框", () => {
    const host = installHost({ ranges: ["A1:B1"] });
    markWpsResult(host.application, prepared);
    const [rule] = host.state("A1:B1");
    // 底色 BGR：0xFFF3B0 → b=0xFF, g=0xF3, r=0xB0 → (0xFF<<16)|(0xF3<<8)|0xB0
    expect(rule.interior?.pattern).toBe(1);
    expect(rule.interior?.color).toBe((0xb0 << 16) | (0xf3 << 8) | 0xff);
    // 边框醒目金黄（0xB8860B → BGR 0x0B86B8）
    expect(rule.borders[7]?.color).toBe((0x0b << 16) | (0x86 << 8) | 0xb8);
    expect(rule.borders[7]?.weight).toBe(2);
  });

  it("涉及填充变化（setFill）的动作：只画边框，不叠底色", () => {
    const host = installHost({ ranges: ["A1:B1"] });
    const fillPlan = {
      authorization: { workbookFullName: "C:\\Data\\测试.xlsx", sheets: [{ name: "数据", index: 1 }] },
      actions: [{ sheetName: "数据", range: "A1:B1", type: "setFill", color: "#FF0000" }]
    } as never;
    markWpsResult(host.application, fillPlan);
    const [rule] = host.state("A1:B1");
    expect(rule.interior).toBeNull();
    expect(rule.borders[7]?.color).toBe((0x0b << 16) | (0x86 << 8) | 0xb8);
  });

  it("混合计划：写值区域叠底色、setFill 区域只边框，各按各的", () => {
    const host = installHost({ ranges: ["A1:B1", "C1:D1"] });
    const mixed = {
      authorization: { workbookFullName: "C:\\Data\\测试.xlsx", sheets: [{ name: "数据", index: 1 }] },
      actions: [
        { sheetName: "数据", range: "A1:B1", type: "writeValues" },
        { sheetName: "数据", range: "C1:D1", type: "setFill", color: "#00FF00" }
      ]
    } as never;
    markWpsResult(host.application, mixed);
    const [writeRule] = host.state("A1:B1");
    const [fillRule] = host.state("C1:D1");
    expect(writeRule.interior?.color).toBe((0xb0 << 16) | (0xf3 << 8) | 0xff);
    expect(fillRule.interior).toBeNull();
  });

  it("副本清理只删本次完整公式对应的规则：另一轮审阅的 EB_REVIEW_ 规则也保留", () => {
    const host = installHost({ ranges: ["A1:B1"] });
    const marks = markWpsResult(host.application, prepared);
    // 用户规则 + **另一轮审阅**的规则（同 EB_REVIEW_ 前缀、不同 nonce）。
    host.pushForeign("A1:B1", "=A1>0");
    host.pushForeign("A1:B1", '=N("EB_REVIEW_another_run_123")=0');
    expect(host.live("A1:B1")).toHaveLength(3);

    // 模拟一张通过 Copy() 带来的副本表：同一区域有本次规则、用户规则、另一轮审阅规则。
    const copySheet = {
      Range: (range: string) => {
        const obj = (host.application.ActiveWorkbook as unknown as {
          Worksheets: { Item(): { Range(r: string): unknown } }
        }).Worksheets.Item().Range(range);
        return obj as never;
      }
    };
    // 只给本次完整公式：另一轮审阅的规则即使同前缀也不删。
    clearReviewMarksFromCopySheet(copySheet as never, ["A1:B1"], marks.formula);
    expect(host.live("A1:B1")).toEqual(["=A1>0", '=N("EB_REVIEW_another_run_123")=0']);
  });
});

it("does not overlay business borders, and skips overlays when both border and fill are changed",()=>{
 const host=installHost({ranges:["A1:B1"]});
 const base={authorization:{workbookFullName:"C:\\Data\\测试.xlsx"}};
 const border={type:"setBorders",sheetName:"数据",range:"A1:B1"};
 const marks=markWpsResult(host.application,{...base,actions:[border]} as never);
 expect(host.state()[0].interior?.color).toBe(0xB0F3FF);
 expect(host.state()[0].borders).toEqual({});
 clearWpsResultMarks(host.application,marks);
 const combined=markWpsResult(host.application,{...base,actions:[border,{...border,type:"setFill"}]} as never);
 expect(host.live()).toHaveLength(0);expect(combined.rules).toHaveLength(0);
});
