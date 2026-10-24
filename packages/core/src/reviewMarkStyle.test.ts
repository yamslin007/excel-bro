import { describe, expect, it } from "vitest";
import {
  actionTouchesFill,
  planTouchesFill,
  reviewMarkStyleForAction,
  reviewMarkStyleForPlan,
  REVIEW_MARK_COLORS,
  REVIEW_MARK_LEGEND
} from "./reviewMarkStyle";

describe("审阅标记呈现方式判据（宿主无关）", () => {
  it("写值/公式/清空/字体/对齐/行高都不改填充 ⇒ 用底色 + 边框", () => {
    for (const type of ["writeValues", "writeFormulas", "clearRange", "setFont", "setAlignment", "resizeRange"]) {
      expect(reviewMarkStyleForAction({ type })).toBe("fill+border");
      expect(actionTouchesFill({ type })).toBe(false);
    }
  });

  it("setFill 涉及填充变化 ⇒ 只用边框，保证实际填充结果可见", () => {
    expect(actionTouchesFill({ type: "setFill" })).toBe(true);
    expect(reviewMarkStyleForAction({ type: "setFill" })).toBe("border-only");
  });

  it("整份计划只要有一处 setFill，planTouchesFill 为真", () => {
    expect(planTouchesFill([{ type: "writeValues" }, { type: "setFill" }])).toBe(true);
    expect(planTouchesFill([{ type: "writeValues" }, { type: "setFont" }])).toBe(false);
    expect(planTouchesFill([])).toBe(false);
  });

  it("配色与图例是固定常量，两端读同一份", () => {
    expect(REVIEW_MARK_COLORS.fill).toBe("#FFF3B0");
    expect(REVIEW_MARK_COLORS.border).toBe("#B8860B");
    expect(REVIEW_MARK_LEGEND.fill).toContain("浅黄色底色");
    expect(REVIEW_MARK_LEGEND.border).toContain("不叠底色");
  });
});

describe("按整份计划的区域重叠关系计算标记样式", () => {
  const styleOf = (actions: ReadonlyArray<Record<string, unknown>>, range: string) =>
    reviewMarkStyleForPlan(
      actions.map(action => ({ type: String(action.type), sheet: String(action.sheet), range: String(action.range) }))
    ).get(`Sheet1\u0000${range}`);

  it("setFont 与 setFill 同区域（两种顺序都算）⇒ 该区域只边框", () => {
    // 顺序 1：setFont 在前、setFill 在后。
    expect(styleOf([
      { type: "setFont", sheet: "Sheet1", range: "A1:B2", bold: true },
      { type: "setFill", sheet: "Sheet1", range: "A1:B2", color: "#FF0000" }
    ], "A1:B2")).toBe("border-only");
    // 顺序 2：setFill 在前、setFont 在后（覆盖顺序与动作先后无关）。
    expect(styleOf([
      { type: "setFill", sheet: "Sheet1", range: "A1:B2", color: "#FF0000" },
      { type: "setFont", sheet: "Sheet1", range: "A1:B2", bold: true }
    ], "A1:B2")).toBe("border-only");
  });

  it("写值区域与 setFill 区域部分重叠 ⇒ 重叠的那块只边框", () => {
    // 写值占 A1:C3，setFill 只占 B2:B3 —— 写值块与填充块相交，故只边框。
    expect(styleOf([
      { type: "writeValues", sheet: "Sheet1", range: "A1:C3", values: [] },
      { type: "setFill", sheet: "Sheet1", range: "B2:B3", color: "#FF0000" }
    ], "A1:C3")).toBe("border-only");
    // setFill 那块自身当然只边框。
    expect(styleOf([
      { type: "writeValues", sheet: "Sheet1", range: "A1:C3", values: [] },
      { type: "setFill", sheet: "Sheet1", range: "B2:B3", color: "#FF0000" }
    ], "B2:B3")).toBe("border-only");
  });

  it("不相交的两块：只有 setFill 那块只边框，另一块照旧叠底色", () => {
    const plan = [
      { type: "writeValues", sheet: "Sheet1", range: "A1", values: [] },
      { type: "setFill", sheet: "Sheet1", range: "D1", color: "#FF0000" }
    ];
    expect(styleOf(plan, "A1")).toBe("fill+border");
    expect(styleOf(plan, "D1")).toBe("border-only");
  });

  it("不同工作表互不牵连：Sheet2 的写值不受 Sheet1 的 setFill 影响", () => {
    const plan = [
      { type: "setFill", sheet: "Sheet1", range: "A1", color: "#FF0000" },
      { type: "writeValues", sheet: "Sheet2", range: "A1", values: [] }
    ];
    expect(reviewMarkStyleForPlan(plan).get("Sheet2\u0000A1")).toBe("fill+border");
    expect(reviewMarkStyleForPlan(plan).get("Sheet1\u0000A1")).toBe("border-only");
  });

  it("区域解析不出 ⇒ 保守只边框（证明不了不被 setFill 碰到，就不叠底色）", () => {
    expect(styleOf([
      { type: "writeValues", sheet: "Sheet1", range: "不是A1", values: [] },
      { type: "setFill", sheet: "Sheet1", range: "A1", color: "#FF0000" }
    ], "不是A1")).toBe("border-only");
  });
});

it("preserves changed borders and combined fill/border results in overlapping regions",()=>{
 const get=(actions:{type:string;sheet:string;range:string}[])=>reviewMarkStyleForPlan(actions).get("S\0A1:B2");
 const font={type:"setFont",sheet:"S",range:"A1:B2"};
 const border={type:"setBorders",sheet:"S",range:"B2:C3"};
 expect(get([font,border])).toBe("fill-only");
 expect(get([border,font,{type:"setFill",sheet:"S",range:"A1"}])).toBe("none");
 expect(get([font,{...border,sheet:"Other"}])).toBe("fill+border");
 expect(get([font,{...border,range:"Unknown"}])).toBe("fill-only");
 expect(reviewMarkStyleForAction(border)).toBe("fill-only");
});

it("does not mask a business border from the adjacent cell's review rule",()=>{
 const result=reviewMarkStyleForPlan([
  {type:"setBorders",sheet:"S",range:"A1"},
  {type:"writeValues",sheet:"S",range:"A2"},
  {type:"writeValues",sheet:"S",range:"C4"}
 ]);
 expect(result.get("S\0A2")).toBe("fill-only");
 expect(result.get("S\0C4")).toBe("fill+border");
});
