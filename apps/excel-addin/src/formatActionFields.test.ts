import { describe, expect, it } from "vitest";
import type { AnalysisPlan, ExcelAction } from "./contracts";
import { assertEditablePlan, assertFormatActionFields } from "./contracts";

/**
 * 格式动作的**字段级**校验。
 *
 * 动作名白名单只保证"这个动作被允许"，不保证"参数是宿主能接受的"：
 * 颜色写成 `red`、对齐写成 `middle`、列宽写成 900，白名单一条都不拦 ——
 * 它们会一路走到宿主才失败，而那一刻前面的动作可能已经落盘了。
 *
 * 这批用例同时钉住两件事：
 * 1. 非法值**在动任何单元格之前**被拒；
 * 2. 校验必须在**真实调用点**上被调用（`assertEditablePlan`）—— 一个没人调用的
 *    守卫等于没有守卫。
 */

function plan(actions: ExcelAction[]): AnalysisPlan {
  return {
    id: "plan-1",
    title: "格式",
    summary: "格式",
    assumptions: [],
    warnings: [],
    actions,
    acceptanceCriteria: []
  };
}

function font(overrides: Partial<Extract<ExcelAction, { type: "setFont" }>>) {
  return { type: "setFont", sheet: "Sheet1", range: "A1", ...overrides } as ExcelAction;
}

describe("格式动作字段校验 · 合法值放行", () => {
  it("七个格式动作的典型计划都能过", () => {
    expect(() =>
      assertEditablePlan(
        plan([
          { type: "setFill", sheet: "S", range: "A1:B1", color: "#DDEEFF" },
          font({ bold: true, italic: false, name: "微软雅黑", size: 14, color: "#112233" }),
          { type: "setNumberFormat", sheet: "S", range: "B2:B9", formatCode: "#,##0.00" },
          {
            type: "setBorders",
            sheet: "S",
            range: "A1:B9",
            sides: ["top", "bottom", "left", "right"],
            style: "dash",
            color: "#000000",
            weight: "medium"
          },
          { type: "setAlignment", sheet: "S", range: "A1:B1", horizontal: "center" },
          { type: "resizeRange", sheet: "S", range: "A1:B9", rowHeight: 24, columnWidth: 18 },
          { type: "autofit", sheet: "S", range: "A1:B9" }
        ])
      )
    ).not.toThrow();
  });

  it("setBorders 省略 style/color/weight 也放行（服务端有默认值）", () => {
    expect(() =>
      assertFormatActionFields({
        type: "setBorders",
        sheet: "S",
        range: "A1",
        sides: ["top"]
      })
    ).not.toThrow();
  });

  it("非格式动作直接跳过，不受这套字段规则约束", () => {
    expect(() =>
      assertFormatActionFields({
        type: "writeValues",
        sheet: "S",
        range: "A1",
        values: [[1]]
      } as ExcelAction)
    ).not.toThrow();
  });
});

describe("格式动作字段校验 · 非法值在动单元格之前就被拒", () => {
  it("颜色必须是 #RRGGBB", () => {
    for (const color of ["red", "#fff", "rgb(1,2,3)", "#GGGGGG"]) {
      expect(() =>
        assertFormatActionFields({ type: "setFill", sheet: "S", range: "A1", color } as ExcelAction)
      ).toThrow("填充色");
    }
  });

  it("对齐值必须落在闭合集合里", () => {
    expect(() =>
      assertFormatActionFields({
        type: "setAlignment",
        sheet: "S",
        range: "A1",
        horizontal: "middle"
      } as unknown as ExcelAction)
    ).toThrow("水平对齐");
    expect(() =>
      assertFormatActionFields({
        type: "setAlignment",
        sheet: "S",
        range: "A1",
        vertical: "middle"
      } as unknown as ExcelAction)
    ).toThrow("垂直对齐");
  });

  it("行高列宽字号必须落在宿主值域里", () => {
    expect(() =>
      assertFormatActionFields({ type: "resizeRange", sheet: "S", range: "A1", rowHeight: 410 } as ExcelAction)
    ).toThrow("行高");
    expect(() =>
      assertFormatActionFields({ type: "resizeRange", sheet: "S", range: "A1", columnWidth: 256 } as ExcelAction)
    ).toThrow("列宽");
    expect(() =>
      assertFormatActionFields({ type: "resizeRange", sheet: "S", range: "A1", rowHeight: -1 } as ExcelAction)
    ).toThrow("行高");
    expect(() => assertFormatActionFields(font({ size: 0 }))).toThrow("字号");
    expect(() => assertFormatActionFields(font({ size: 500 }))).toThrow("字号");
  });

  it("字体名不能为空也不能超长", () => {
    expect(() => assertFormatActionFields(font({ name: "" }))).toThrow("字体名");
    expect(() => assertFormatActionFields(font({ name: "   " }))).toThrow("字体名");
    expect(() => assertFormatActionFields(font({ name: "A".repeat(256) }))).toThrow("字体名");
  });

  it("数字格式码不能为空也不能超长", () => {
    expect(() =>
      assertFormatActionFields({
        type: "setNumberFormat",
        sheet: "S",
        range: "A1",
        formatCode: ""
      } as ExcelAction)
    ).toThrow("数字格式码");
    expect(() =>
      assertFormatActionFields({
        type: "setNumberFormat",
        sheet: "S",
        range: "A1",
        formatCode: "0".repeat(101)
      } as ExcelAction)
    ).toThrow("数字格式码");
  });

  it("边框必须有边、边名合法、且不重复", () => {
    expect(() =>
      assertFormatActionFields({
        type: "setBorders",
        sheet: "S",
        range: "A1",
        sides: []
      } as ExcelAction)
    ).toThrow("至少需要一条边");
    expect(() =>
      assertFormatActionFields({
        type: "setBorders",
        sheet: "S",
        range: "A1",
        sides: ["diagonal"]
      } as unknown as ExcelAction)
    ).toThrow("边框位置");
    expect(() =>
      assertFormatActionFields({
        type: "setBorders",
        sheet: "S",
        range: "A1",
        sides: ["top", "top"]
      } as ExcelAction)
    ).toThrow("重复");
  });

  it("线型/线宽/边框颜色也走闭合集合与十六进制校验", () => {
    expect(() =>
      assertFormatActionFields({
        type: "setBorders",
        sheet: "S",
        range: "A1",
        sides: ["top"],
        style: "wavy"
      } as unknown as ExcelAction)
    ).toThrow("线型");
    expect(() =>
      assertFormatActionFields({
        type: "setBorders",
        sheet: "S",
        range: "A1",
        sides: ["top"],
        weight: "fat"
      } as unknown as ExcelAction)
    ).toThrow("线宽");
    expect(() =>
      assertFormatActionFields({
        type: "setBorders",
        sheet: "S",
        range: "A1",
        sides: ["top"],
        color: "black"
      } as ExcelAction)
    ).toThrow("边框颜色");
  });
});

describe("格式动作字段校验 · 空动作必须被拒", () => {
  /**
   * 一个属性都没给 = 这个动作什么也不做。放过去会让计划"看起来成功、实际没改"，
   * 与"只有指定属性被改变"的承诺正好相反。Python 侧 `SetFontAction` /
   * `SetAlignmentAction` / `ResizeRangeAction` 的 `require_expected_property`
   * 是同一条规矩，两端必须一致（由 `server/tests/test_protocol_consistency.py` 锁住）。
   */
  it("setFont / setAlignment / resizeRange 全空时被拒", () => {
    expect(() => assertFormatActionFields(font({}))).toThrow("没有指定任何要修改的字体属性");
    expect(() =>
      assertFormatActionFields({ type: "setAlignment", sheet: "S", range: "A1" } as ExcelAction)
    ).toThrow("没有指定任何要修改的对齐属性");
    expect(() =>
      assertFormatActionFields({ type: "resizeRange", sheet: "S", range: "A1" } as ExcelAction)
    ).toThrow("没有指定行高或列宽");
  });

  it("每个单属性都能单独通过（可选 ≠ 可全空）", () => {
    for (const single of [
      { bold: true },
      { italic: true },
      { name: "Arial" },
      { size: 12 },
      { color: "#112233" }
    ]) {
      expect(() => assertFormatActionFields(font(single))).not.toThrow();
    }
    expect(() =>
      assertFormatActionFields({ type: "resizeRange", sheet: "S", range: "A1", rowHeight: 0 } as ExcelAction)
    ).not.toThrow();
    expect(() =>
      assertFormatActionFields({ type: "resizeRange", sheet: "S", range: "A1", columnWidth: 0 } as ExcelAction)
    ).not.toThrow();
  });

  it("bold:false / wrapText:false 是『真的改成否』，不是『没给』", () => {
    // 用真值判断会把 false 当成"没给"，于是 `setFont{bold:false}` 变成空动作被拒。
    expect(() => assertFormatActionFields(font({ bold: false }))).not.toThrow();
    expect(() =>
      assertFormatActionFields({ type: "setAlignment", sheet: "S", range: "A1", wrapText: false } as ExcelAction)
    ).not.toThrow();
  });
});

describe("格式动作字段校验 · 必须在真实调用点上生效", () => {
  it("assertEditablePlan 会拒绝字段非法的格式动作（不是只定义不调用）", () => {
    // 这条用例的判别点在**调用链**上：把 assertEditablePlan 里那行
    // `assertFormatActionFields(action)` 删掉，它必须变红。
    expect(() =>
      assertEditablePlan(plan([{ type: "setFill", sheet: "S", range: "A1", color: "red" } as ExcelAction]))
    ).toThrow("填充色");
    expect(() => assertEditablePlan(plan([font({})]))).toThrow(
      "没有指定任何要修改的字体属性"
    );
  });

  it("非法格式动作不会因为『前一个动作合法』而被漏检", () => {
    expect(() =>
      assertEditablePlan(
        plan([
          { type: "setFill", sheet: "S", range: "A1", color: "#DDEEFF" },
          { type: "setFill", sheet: "S", range: "A2", color: "red" } as ExcelAction
        ])
      )
    ).toThrow("填充色");
  });
});
