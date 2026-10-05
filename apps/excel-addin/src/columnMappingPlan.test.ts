/**
 * Excel 通道的本地列映射：复现并锁住"本端不支持该模式"的失败点。
 *
 * 分工与 WPS 完全一致（`wpsColumnMappingPlan.ts`），只是宿主读取方式不同：
 * 服务端在 `planningReadPolicy: "none"` 时只下发规则（它手上没有源值），
 * 客户端读出授权工作表里的源列、用共享 core 就地算成白名单 `writeValues` 计划。
 *
 * 本文件钉三件事：
 *
 * 1. **规则能算成计划**（旧行为直接抛"本端不支持该模式"）；
 * 2. **切片边界四个方向都拒**，不给没读到的行补空值；
 * 3. **语义不被兜底**：未知类别拒绝、显式"无问题→/"项不被当空白。
 *
 * 文件名与样本值都刻意不含业务语义：列位置、表名、标签由用例自己给，
 * 实现对它们一视同仁（生产代码里不得出现这些字符串）。
 */

import { describe, expect, it } from "vitest";

import { BlankPolicyNeededError } from "@excel-bro/core";
import {
  buildColumnMappingPlan,
  sliceMappingSource,
  type MappingSourceSheet
} from "./columnMappingPlan";
import { applyBlankPolicy } from "./localMappingResolution";
import type { AnalysisPlan, CellValue, ColumnMappingRule } from "./contracts";

/**
 * 取出计划里第一条 `writeValues` 动作的值矩阵。
 *
 * `AnalysisPlan.actions` 是 `ExcelAction` 联合体，直接读 `.values` 通不过类型检查
 * （vitest 只转译不做类型检查，所以这个错只有 `tsc --noEmit` 才暴露）。这里显式
 * 断言类型：列映射计划**恒定**只产出一条 `writeValues`，断言失败即测试失败。
 */
function writeValuesOf(plan: AnalysisPlan): CellValue[][] {
  const action = plan.actions[0];
  expect(action.type).toBe("writeValues");
  return (action as Extract<typeof action, { type: "writeValues" }>).values;
}

/** 一张 A1 起、两列宽的小表：第 1 列是来源列，第 2 列是空的目标列。 */
function sheetWith(
  rows: (string | number | boolean | null)[][],
  overrides: Partial<MappingSourceSheet> = {}
): MappingSourceSheet {
  return {
    name: "表一",
    startRow: 1,
    startColumn: 1,
    columnCount: 2,
    values: rows,
    ...overrides
  };
}

function rule(overrides: Partial<ColumnMappingRule> = {}): ColumnMappingRule {
  return {
    kind: "columnMapping",
    sourceSheet: "表一",
    targetSheet: "表一",
    sourceRange: "A2:A3",
    targetRange: "B2:B3",
    mapping: [
      { source: "甲类", target: "一类" },
      { source: "乙类", target: "二类" }
    ],
    blankLabel: "未分类",
    ...overrides
  };
}

describe("规则 → 可执行计划（Excel 通道）", () => {
  it("把规则算成单个白名单 writeValues 动作，而不是拒绝", () => {
    const sheet = sheetWith([
      ["来源", "目标"],
      ["甲类", null],
      ["乙类", null]
    ]);
    const plan = buildColumnMappingPlan(rule(), sheet, {
      summary: "按映射规则写入。",
      warnings: ["这是一张快照。"],
      assumptions: ["映射规则共 2 条。"]
    });

    expect(plan.actions).toHaveLength(1);
    expect(plan.actions[0]).toMatchObject({
      type: "writeValues",
      sheet: "表一",
      range: "B2:B3",
      values: [["一类"], ["二类"]]
    });
    // 服务端的告知不能因为换了计算位置就丢掉。
    expect(plan.warnings).toContain("这是一张快照。");
    expect(plan.assumptions).toContain("映射规则共 2 条。");
    // 本端生成 id，前缀与 WPS 区分，便于排查走的是哪条通道。
    expect(plan.id.startsWith("office-map-column-")).toBe(true);
  });

  it("源列不在第 1 列、目标列在别处时同样成立（不绑死列位置）", () => {
    // 4 列宽的表：来源在 C 列，目标在 D 列。
    const sheet: MappingSourceSheet = {
      name: "数据",
      startRow: 1,
      startColumn: 2,
      columnCount: 3,
      values: [
        ["备注", "来源", "目标"],
        ["x", "甲类", null],
        ["y", "乙类", null]
      ]
    };
    const plan = buildColumnMappingPlan(
      rule({
        sourceSheet: "数据",
        targetSheet: "数据",
        sourceRange: "C2:C3",
        targetRange: "D2:D3"
      }),
      sheet,
      { summary: "s", warnings: [], assumptions: [] }
    );
    expect(plan.actions[0]).toMatchObject({
      sheet: "数据",
      range: "D2:D3",
      values: [["一类"], ["二类"]]
    });
  });

  it("规则要求读的表与传入的表不一致时拒绝", () => {
    expect(() =>
      buildColumnMappingPlan(rule({ sourceSheet: "别的表" }), sheetWith([]), {
        summary: "s",
        warnings: [],
        assumptions: []
      })
    ).toThrow(/别的表/);
  });

  it("跨已授权工作表：源取 A 表、目标写 B 表，仍只读传入的授权源表", () => {
    // 来源表与目标表是两张不同的工作表，两者都在本次授权范围内。
    // 读取侧只拿到源表（目标表不需要读），计划仍按白名单写到目标表。
    const sourceSheet: MappingSourceSheet = {
      name: "来源表",
      startRow: 1,
      startColumn: 1,
      columnCount: 2,
      values: [
        ["来源", "备注"],
        ["甲类", "x"],
        ["乙类", "y"]
      ]
    };
    const plan = buildColumnMappingPlan(
      rule({
        sourceSheet: "来源表",
        targetSheet: "目标表",
        sourceRange: "A2:A3",
        targetRange: "C2:C3"
      }),
      sourceSheet,
      { summary: "s", warnings: [], assumptions: [] }
    );
    expect(plan.actions).toHaveLength(1);
    expect(plan.actions[0]).toMatchObject({
      type: "writeValues",
      // 写入落在目标表，不落在源表。
      sheet: "目标表",
      range: "C2:C3",
      values: [["一类"], ["二类"]]
    });
  });

  it("非等级文本映射：键值都是任意文本，不做数值/等级假设", () => {
    // 映射表两侧都是普通文本（无大小、无顺序语义），实现必须原样搬运，
    // 不得"推断"等级、排序或做任何归一。
    const sheet = sheetWith([
      ["来源", "目标"],
      ["待复核", null],
      ["已归档", null]
    ]);
    const plan = buildColumnMappingPlan(
      rule({
        sourceRange: "A2:A3",
        targetRange: "B2:B3",
        mapping: [
          { source: "待复核", target: "请人工确认" },
          { source: "已归档", target: "（已归档）" }
        ],
        blankLabel: "（空）"
      }),
      sheet,
      { summary: "s", warnings: [], assumptions: [] }
    );
    expect(writeValuesOf(plan)).toEqual([["请人工确认"], ["（已归档）"]]);
  });
});

describe("空白：四种处理在 Excel 通道同样成立（不靠 blankLabel 一个字符串）", () => {
  // 源列第 3 行是空的；目标列 B 提前放了一个公式位与一个数值，只为证明
  // "保留原值"不是读回来再写回 —— 计划里根本不该出现那一行。
  const blankSheet = sheetWith([
    ["来源", "目标"],
    ["甲类", null],
    [null, null]
  ]);
  const context = { summary: "按映射规则写入。", warnings: [], assumptions: [] };

  it("源区域一格空白都没有、也没给策略：不挡预览", () => {
    const plan = buildColumnMappingPlan(
      rule({ blankLabel: undefined, sourceRange: "A2:A2", targetRange: "B2:B2" }),
      blankSheet,
      context
    );
    expect(plan.actions).toEqual([{
      type: "writeValues", sheet: "表一", range: "B2:B2", values: [["一类"]]
    }]);
  });

  it("真读到空白又没给策略：抛结构化缺口并带上真实计数", () => {
    let caught: unknown;
    try {
      buildColumnMappingPlan(rule({ blankLabel: undefined }), blankSheet, context);
    } catch (error) {
      caught = error;
    }
    expect(caught).toBeInstanceOf(BlankPolicyNeededError);
    expect((caught as BlankPolicyNeededError).blankCount).toBe(1);
    // 不是一条"失败了"的文案：调用方拿这个计数去出澄清卡。
    expect((caught as Error).message).not.toContain("未分类");
  });

  it("选保留原值：那一格 0 次写入 —— 它不出现在任何动作范围里", () => {
    const plan = buildColumnMappingPlan(
      rule({ blankLabel: undefined, blankPolicy: { kind: "keep" } }),
      blankSheet,
      context
    );
    expect(plan.actions).toEqual([{
      type: "writeValues", sheet: "表一", range: "B2:B2", values: [["一类"]]
    }]);
    // B3 一次都不会被写：公式、值、数字格式、条件格式都不动。
    expect(JSON.stringify(plan.actions)).not.toContain("B3");
  });

  it("清空 / 写空串 / 写指定值 三条路各产不同的动作", () => {
    const cleared = buildColumnMappingPlan(
      rule({ blankLabel: undefined, blankPolicy: { kind: "clear" } }), blankSheet, context
    );
    expect(cleared.actions).toEqual([
      { type: "writeValues", sheet: "表一", range: "B2:B2", values: [["一类"]] },
      { type: "clearRange", sheet: "表一", range: "B3:B3", applyTo: "contents" }
    ]);

    const emptyString = buildColumnMappingPlan(
      rule({ blankLabel: undefined, blankPolicy: { kind: "value", value: "" } }), blankSheet, context
    );
    expect(emptyString.actions).toEqual([{
      type: "writeValues", sheet: "表一", range: "B2:B3", values: [["一类"], [""]]
    }]);

    const labelled = buildColumnMappingPlan(
      rule({ blankLabel: undefined, blankPolicy: { kind: "value", value: "待填" } }), blankSheet, context
    );
    expect(writeValuesOf(labelled)).toEqual([["一类"], ["待填"]]);
  });

  it("用户改了主意：新策略替换旧规则，不是把两条并在一起", () => {
    // 旧规则带着 blankLabel「未分类」，用户后来选了"清空"。若只是追加，
    // 两条策略并存会让下一次预览挑谁都不确定。
    const updated = applyBlankPolicy(rule({ blankLabel: "未分类" }), { kind: "clear" });
    expect(updated.blankPolicy).toEqual({ kind: "clear" });
    expect(updated.blankLabel).toBeUndefined();
    const plan = buildColumnMappingPlan(updated, blankSheet, context);
    expect(plan.actions.some((action) => action.type === "clearRange")).toBe(true);
    expect(JSON.stringify(plan.actions)).not.toContain("未分类");
  });
});

describe("源区域切片：越界拒绝，不补空值", () => {
  const sheet = sheetWith([
    ["来源", "目标"],
    ["甲类", null],
    ["乙类", null]
  ]);

  it("正常范围按声明宽度切出矩形", () => {
    expect(sliceMappingSource(sheet, "A2:A3")).toEqual([["甲类"], ["乙类"]]);
  });

  it("结束行超出已读入范围：拒绝", () => {
    expect(() => sliceMappingSource(sheet, "A2:A9")).toThrow(/超出已读取范围/);
  });

  it("起始行早于已读入起点：拒绝", () => {
    expect(() => sliceMappingSource(sheetWith([["甲类"]], { startRow: 5 }), "A2:A2")).toThrow(
      /早于/
    );
  });

  it("起始列早于已读入起点：拒绝", () => {
    expect(() =>
      sliceMappingSource(sheetWith([["甲类"]], { startColumn: 3 }), "A1:A1")
    ).toThrow(/早于/);
  });

  it("结束列超出声明宽度：拒绝", () => {
    expect(() => sliceMappingSource(sheet, "A1:C1")).toThrow(/超出已读取范围/);
  });

  it("无法解析的区域：拒绝", () => {
    expect(() => sliceMappingSource(sheet, "A:A")).toThrow(/无法解析/);
  });
});

describe("语义边界：不猜测、不兜底", () => {
  it("来源值不在映射表内时拒绝，并列出全部缺口", () => {
    const sheet = sheetWith([
      ["来源", "目标"],
      ["甲类", null],
      ["丙类", null]
    ]);
    expect(() =>
      buildColumnMappingPlan(rule(), sheet, {
        summary: "s",
        warnings: [],
        assumptions: []
      })
    ).toThrow(/不在映射表内/);
  });

  it('显式"无问题→/"项按显式项输出，真空白仍走 blankLabel', () => {
    const sheet = sheetWith([
      ["来源", "目标"],
      ["无问题", null],
      [null, null]
    ]);
    const plan = buildColumnMappingPlan(
      rule({
        sourceRange: "A2:A3",
        targetRange: "B2:B3",
        mapping: [{ source: "无问题", target: "/" }],
        blankLabel: "待填"
      }),
      sheet,
      { summary: "s", warnings: [], assumptions: [] }
    );
    // 真空白不等于"无问题"：前者写 blankLabel，后者按显式项写 "/"。
    expect(writeValuesOf(plan)).toEqual([["/"], ["待填"]]);
  });
});
