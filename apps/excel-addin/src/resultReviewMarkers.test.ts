// @vitest-environment jsdom
/**
 * 审阅标记的生命周期（Excel 端）。
 *
 * 模拟范围：只替身 **Office.js 的条件格式 API 与工作簿身份**；标记模块本身是真的。
 * 这里刻意不走"宿主不支持标记"的降级分支 —— 降级只说明"没有标记"，不能证明
 * 标记存在时的归属、部分失败、刷新认领是对的。
 */
import { beforeEach, describe, expect, it, vi } from "vitest";

const identity = vi.hoisted(() => ({ value: "url:file:///E:/Book.xlsx" }));
vi.mock("./excel", () => ({ currentOfficeWorkbookIdentity: () => identity.value }));

import {
  clearExcelResultMarks,
  loadExcelResultMarks,
  markExcelResult,
  recoverExcelResultMarks
} from "./resultReviewMarkers";

type RuleState = { id: string; type: string; formula: string; deleted: boolean; borderColor: string; fillColor: string | null };

function installExcel(options: { failOnAdd?: number; failOnDelete?: boolean } = {}) {
  const rules: RuleState[] = [];
  let seq = 0;
  const makeRule = (state: RuleState) => ({
    id: state.id,
    type: state.type,
    custom: {
      rule: { formula: state.formula },
      format: {
        borders: { getItem: () => ({
          get color() { return state.borderColor; },
          set color(v: string) { state.borderColor = v; },
          get style() { return "Continuous"; },
          set style(_v: string) {}
        }) },
        fill: {
          get color() { return state.fillColor as string; },
          set color(v: string) { state.fillColor = v; }
        }
      },
      load: () => {}
    },
    load: () => {},
    delete: () => {
      if (options.failOnDelete) throw new Error("删除被宿主拒绝");
      state.deleted = true;
    }
  });
  const contexts = {
    sync: async () => {},
    workbook: {
      worksheets: {
        getItem: (sheet: string) => ({
          getRange: (range: string) => ({
            conditionalFormats: {
              add: () => {
                seq += 1;
                if (options.failOnAdd === seq) throw new Error("第 2 处标记创建失败");
                const state: RuleState = { id: `rule-${seq}`, type: "Custom", formula: "", deleted: false, borderColor: "", fillColor: null };
                rules.push(state);
                // 规则对象需要可写 formula（调用方随后设置）。
                const rule = makeRule(state);
                Object.defineProperty(rule.custom.rule, "formula", {
                  get: () => state.formula,
                  set: (next: string) => { state.formula = next; },
                  configurable: true
                });
                return rule;
              },
              get items() {
                return rules.filter(rule => !rule.deleted).map(makeRule);
              },
              load: () => {}
            },
            _sheet: sheet,
            _range: range
          })
        })
      }
    }
  };
  (globalThis as unknown as { Excel: unknown }).Excel = {
    run: async (job: (context: typeof contexts) => Promise<void>) => job(contexts)
  };
  return { rules, live: () => rules.filter(rule => !rule.deleted) };
}

const plan = {
  id: "plan", title: "t", summary: "s", assumptions: [], warnings: [],
  actions: [
    { type: "writeValues", sheet: "Sheet1", range: "A1", values: [[1]] },
    { type: "writeValues", sheet: "Sheet1", range: "B1", values: [[2]] }
  ]
} as never;

beforeEach(() => { localStorage.clear(); identity.value = "url:file:///E:/Book.xlsx"; });

describe("Excel 审阅标记", () => {
  it("只创建独立的边框规则，不碰原有格式；确认后只删本次规则", async () => {
    const host = installExcel();
    const marks = await markExcelResult(plan, identity.value);
    expect(host.live()).toHaveLength(2);
    // 归属记录只含表名/区域/规则 id 与随机公式，不含任何单元格内容。
    const stored = loadExcelResultMarks()!;
    expect(JSON.stringify(stored)).not.toMatch(/values|"A1":/);
    await clearExcelResultMarks(marks);
    expect(host.live()).toHaveLength(0);
    expect(loadExcelResultMarks()).toBeNull();
  });

  it("创建部分成功后失败：已创建的可追踪并清理，记录不残留", async () => {
    const host = installExcel({ failOnAdd: 2 });
    await expect(markExcelResult(plan, identity.value)).rejects.toThrow(/第 2 处标记创建失败/);
    expect(host.live()).toHaveLength(0);
    expect(loadExcelResultMarks()).toBeNull();
  });

  it("创建失败且清理也失败：保留归属记录，下次仍能认领", async () => {
    const host = installExcel({ failOnAdd: 2, failOnDelete: true });
    await expect(markExcelResult(plan, identity.value)).rejects.toThrow(/暂未移除/);
    expect(host.live()).toHaveLength(1);
    const stored = loadExcelResultMarks()!;
    expect(stored.entries).toHaveLength(1);
    // 记录只讲归属，不讲业务值。
    expect(JSON.stringify(stored)).not.toContain('"values"');
  });

  it("刷新后按规则 id + 随机公式认领遗留标记", async () => {
    const host = installExcel();
    await markExcelResult(plan, identity.value);
    expect(host.live()).toHaveLength(2);
    // 模拟刷新：内存里没有 marks 对象，只剩持久记录。
    const note = await recoverExcelResultMarks();
    expect(note).toContain("已清理");
    expect(host.live()).toHaveLength(0);
    expect(loadExcelResultMarks()).toBeNull();
  });

  it("工作簿身份不符：不碰任何规则，也不删记录", async () => {
    const host = installExcel();
    await markExcelResult(plan, identity.value);
    identity.value = "url:file:///E:/Other.xlsx";
    const note = await recoverExcelResultMarks();
    expect(note).toContain("另一本工作簿");
    expect(host.live()).toHaveLength(2);
    expect(loadExcelResultMarks()).not.toBeNull();
  });

  it("同一区域里别的规则不会被顺带删掉（不按区域批量清）", async () => {
    const host = installExcel();
    await markExcelResult(plan, identity.value);
    // 用户/别的任务在同一区域新建了一条规则。
    host.rules.push({ id: "user-rule", type: "Custom", formula: "=A1>0", deleted: false, borderColor: "", fillColor: null });
    await recoverExcelResultMarks();
    expect(host.live().map(rule => rule.id)).toEqual(["user-rule"]);
  });
});
describe("Excel 审阅标记呈现：底色与边框", () => {
  it("不修改填充的动作：浅黄底色 + 醒目金黄边框", async () => {
    const host = installExcel();
    await markExcelResult(plan, identity.value);
    const [rule] = host.live();
    expect(rule.borderColor).toBe("#B8860B");
    expect(rule.fillColor).toBe("#FFF3B0");
  });

  it("setFill 动作：只画边框，不叠底色，实际填充可见", async () => {
    const host = installExcel();
    const fillPlan = { id: "plan", title: "t", summary: "s", assumptions: [], warnings: [], actions: [{ type: "setFill", sheet: "Sheet1", range: "A1", color: "#FF0000" }] } as never;
    await markExcelResult(fillPlan, identity.value);
    const [rule] = host.live();
    expect(rule.borderColor).toBe("#B8860B");
    expect(rule.fillColor).toBeNull();
  });

  it("混合计划：写值区域叠底色、setFill 区域只边框", async () => {
    const host = installExcel();
    const mixed = {
      id: "plan", title: "t", summary: "s", assumptions: [], warnings: [],
      actions: [
        { type: "writeValues", sheet: "Sheet1", range: "A1", values: [[1]] },
        { type: "setFill", sheet: "Sheet1", range: "B1", color: "#00FF00" }
      ]
    } as never;
    await markExcelResult(mixed, identity.value);
    const rules = host.live();
    const writeRule = rules.find(rule => rule.id === "rule-1")!;
    const fillRule = rules.find(rule => rule.id === "rule-2")!;
    expect(writeRule.fillColor).toBe("#FFF3B0");
    expect(fillRule.fillColor).toBeNull();
  });
});

describe("Excel 审阅标记：按整份计划的区域重叠关系算样式", () => {
  it("setFont 与 setFill 同区域（两种顺序都算）⇒ 只边框，不叠底色", async () => {
    // 顺序 1：setFont 在前。
    let host = installExcel();
    await markExcelResult({ id: "p", title: "t", summary: "s", assumptions: [], warnings: [], actions: [
      { type: "setFont", sheet: "Sheet1", range: "A1", bold: true },
      { type: "setFill", sheet: "Sheet1", range: "A1", color: "#FF0000" }
    ] } as never, identity.value);
    expect(host.live()[0].fillColor).toBeNull();
    expect(host.live()[0].borderColor).toBe("#B8860B");

    // 顺序 2：setFill 在前。
    host = installExcel();
    await markExcelResult({ id: "p", title: "t", summary: "s", assumptions: [], warnings: [], actions: [
      { type: "setFill", sheet: "Sheet1", range: "A1", color: "#FF0000" },
      { type: "setFont", sheet: "Sheet1", range: "A1", bold: true }
    ] } as never, identity.value);
    expect(host.live()[0].fillColor).toBeNull();
  });

  it("写值区域与 setFill 区域部分重叠 ⇒ 写值那块也只边框", async () => {
    const host = installExcel();
    await markExcelResult({ id: "p", title: "t", summary: "s", assumptions: [], warnings: [], actions: [
      { type: "writeValues", sheet: "Sheet1", range: "A1:C3", values: [[1]] },
      { type: "setFill", sheet: "Sheet1", range: "B2:B3", color: "#FF0000" }
    ] } as never, identity.value);
    // 两块都只边框（写值块与填充块相交）。
    expect(host.live().map(rule => rule.fillColor)).toEqual([null, null]);
  });

  it("不相交：只有 setFill 那块只边框，写值那块照旧叠底色", async () => {
    const host = installExcel();
    await markExcelResult({ id: "p", title: "t", summary: "s", assumptions: [], warnings: [], actions: [
      { type: "writeValues", sheet: "Sheet1", range: "A1", values: [[1]] },
      { type: "setFill", sheet: "Sheet1", range: "D1", color: "#FF0000" }
    ] } as never, identity.value);
    const writeRule = host.live().find(rule => rule.id === "rule-1")!;
    const fillRule = host.live().find(rule => rule.id === "rule-2")!;
    expect(writeRule.fillColor).toBe("#FFF3B0");
    expect(fillRule.fillColor).toBeNull();
  });
});

it("does not overlay business borders, and skips overlays when border and fill overlap",async()=>{
 const host=installExcel();
 const border={type:"setBorders",sheet:"Sheet1",range:"A1:B1",sides:["top"]};
 const marks=await markExcelResult({actions:[border]} as never,identity.value);
 expect(host.live()[0].fillColor).toBe("#FFF3B0");expect(host.live()[0].borderColor).toBe("");
 await clearExcelResultMarks(marks);
 const combined=await markExcelResult({actions:[border,{type:"setFill",sheet:"Sheet1",range:"A1:B1",color:"#FF0000"}]} as never,identity.value);
 expect(host.live()).toHaveLength(0);expect(combined.entries).toHaveLength(0);
});
