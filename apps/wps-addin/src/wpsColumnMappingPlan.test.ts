/**
 * WPS 侧「规则 → 本地计算 → 计划」这一步。
 *
 * 映射本身的逐格判定由 `packages/core/src/columnMapping.test.ts` 与
 * `config/column_mapping_cases.json` 覆盖，这里只测接线部分：区域切片、
 * 越界拒绝、以及"服务端下发的告知不被丢掉"。
 */

import { describe, expect, it } from "vitest";
import type { ColumnMappingRule } from "../../excel-addin/src/contracts";
import {
  buildWpsColumnMappingPlan,
  fixedValueFallbackFromFormula,
  sliceMappingSource,
  type MappingSourceSheet
} from "./wpsColumnMappingPlan";

const RULE: ColumnMappingRule = {
  kind: "columnMapping",
  sourceSheet: "订单",
  targetSheet: "订单",
  sourceRange: "B2:B4",
  targetRange: "C2:C4",
  mapping: [
    { source: "办公用品", target: "行政" },
    { source: "电子设备", target: "资产" }
  ],
  blankLabel: "未分类"
};

const CONTEXT = {
  summary: "按 2 条映射规则写入。",
  warnings: ["目标区域已有内容会被覆盖。"],
  assumptions: ["源列空白按「未分类」处理。"]
};

function sheet(overrides: Partial<MappingSourceSheet> = {}): MappingSourceSheet {
  return {
    name: "订单",
    startRow: 1,
    startColumn: 1,
    columnCount: 3,
    values: [
      ["单号", "类别", "归属"],
      ["X001", "办公用品", null],
      ["X002", "电子设备", null],
      ["X003", null, null]
    ],
    ...overrides
  };
}

describe("sliceMappingSource", () => {
  it("按 1 基行列号切出区域", () => {
    expect(sliceMappingSource(sheet(), "B2:B4")).toEqual([
      ["办公用品"],
      ["电子设备"],
      [null]
    ]);
  });

  it("读取范围不从 A1 开始时按偏移切片", () => {
    const shifted = sheet({
      startRow: 5,
      startColumn: 2,
      columnCount: 1,
      values: [["办公用品"], ["电子设备"]]
    });
    expect(sliceMappingSource(shifted, "B5:B6")).toEqual([
      ["办公用品"],
      ["电子设备"]
    ]);
  });

  it("行越界拒绝，不补 null", () => {
    // 补出来的 null 会被映射成空白标签，写进实际上没有数据的行。
    expect(() => sliceMappingSource(sheet(), "B2:B50")).toThrow(/超出已读取范围/);
  });

  it("列在读取起点之前也拒绝", () => {
    const shifted = sheet({ startColumn: 3, columnCount: 1, values: [["x"], ["y"]] });
    expect(() => sliceMappingSource(shifted, "B1:B2")).toThrow(/列，早于/);
  });

  it("列超出声明宽度也拒绝", () => {
    // 与"行内缺项折成 null"的区别：那是声明宽度之内的空格，这是宽度之外。
    const narrow = sheet({ columnCount: 2, values: [["单号", "类别"], ["X001", "办公用品"]] });
    expect(() => sliceMappingSource(narrow, "B1:D2")).toThrow(/列，超出已读取范围/);
  });

  it("行内缺项折成 null，不算越界", () => {
    // 读取侧会因尾部空值把行裁短，区域内缺的那几格是真实空单元格。
    const ragged = sheet({ columnCount: 2, values: [["单号", "类别"], ["X001"]] });
    expect(sliceMappingSource(ragged, "B2:B2")).toEqual([[null]]);
  });

  it("无法解析的区域拒绝", () => {
    expect(() => sliceMappingSource(sheet(), "B:B")).toThrow(/无法解析/);
  });
});

describe("buildWpsColumnMappingPlan", () => {
  it("算出目标矩阵，空白走 blankLabel", () => {
    const plan = buildWpsColumnMappingPlan(RULE, sheet(), CONTEXT);
    expect(plan.actions).toHaveLength(1);
    const action = plan.actions[0];
    expect(action.type).toBe("writeValues");
    expect(action).toMatchObject({ sheet: "订单", range: "C2:C4" });
    expect("values" in action ? action.values : null).toEqual([
      ["行政"],
      ["资产"],
      ["未分类"]
    ]);
  });

  it("纯文本同表映射优先给公式与逐格计算值验收，回退仍是确定性固定值", () => {
    const source = sheet({ values: [
      ["单号", "类别", "归属"],
      ["X001", "办公用品", null],
      ["X002", "电子设备", null],
      ["X003", "办公用品", null]
    ] });
    const plan = buildWpsColumnMappingPlan(RULE, source, CONTEXT);
    expect(plan.actions[0].type).toBe("writeFormulas");
    expect(plan.actions[0]).toMatchObject({ sheet: "订单", range: "C2:C4" });
    expect(plan.acceptanceCriteria).toEqual([{
      type: "rangeEquals", sheet: "订单", range: "C2:C4", expected: [["行政"], ["资产"], ["行政"]]
    }]);
    const fixed = fixedValueFallbackFromFormula(plan);
    expect(fixed?.actions).toEqual([{
      type: "writeValues", sheet: "订单", range: "C2:C4", values: [["行政"], ["资产"], ["行政"]]
    }]);
    expect(fixed?.warnings.join(" ")).toContain("试算或结果回读不符合计划");
    expect(fixedValueFallbackFromFormula(fixed!)).toBeNull();
  });

  it("保留服务端下发的告知与前提", () => {
    // 换了计算位置不改变"这是一张快照"的事实。
    const plan = buildWpsColumnMappingPlan(RULE, sheet(), CONTEXT);
    expect(plan.warnings).toEqual(expect.arrayContaining(CONTEXT.warnings));
    expect(plan.warnings.join(" ")).toContain("固定值");
    expect(plan.assumptions).toEqual(CONTEXT.assumptions);
    expect(plan.summary).toBe(CONTEXT.summary);
  });

  it("未命中的来源值拒绝，并报出全部缺口", () => {
    const unmapped = sheet({
      values: [
        ["单号", "类别"],
        ["X001", "差旅"],
        ["X002", "礼品"],
        ["X003", "办公用品"]
      ]
    });
    expect(() => buildWpsColumnMappingPlan(RULE, unmapped, CONTEXT)).toThrow(
      /差旅[\s\S]*礼品|礼品[\s\S]*差旅/
    );
  });

  it("传错工作表时拒绝，不改读别的表", () => {
    expect(() =>
      buildWpsColumnMappingPlan(RULE, sheet({ name: "别的表" }), CONTEXT)
    ).toThrow(/规则要求读取/);
  });

  it("源与目标重叠时拒绝", () => {
    const overlapping = { ...RULE, targetRange: "B2:B4" };
    expect(() => buildWpsColumnMappingPlan(overlapping, sheet(), CONTEXT)).toThrow(
      /重叠/
    );
  });

  it("计划 id 在本端生成且带 wps 前缀", () => {
    // 服务端没产计划，也就没有可沿用的 id；前缀便于在日志里区分来源。
    const plan = buildWpsColumnMappingPlan(RULE, sheet(), CONTEXT);
    expect(plan.id).toMatch(/^wps-map-column-/);
  });
});
