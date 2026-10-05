/**
 * 规划失败不得被标成"已完成"。
 *
 * 复现的失败点：Excel 收到 `local_plan_rule` 后在本端读源列、算计划，若这一步
 * 失败（源表未授权、区域越界、来源值不在映射表内……），错误会走 `catch` 分支
 * 并报给用户；但调用方此前无条件再补一次默认的 `finishActivity()`，把这一轮
 * 记成"已完成" —— 报错与"已完成"同时出现，用户误以为已经写入。
 *
 * 这里锁住判定：只有 `planned` 才允许补默认完成；失败/打断/跳过都不允许，
 * 且这轮**不会**产出任何计划对象（因此没有写入预览卡，也就没有写入）。
 */

import { describe, expect, it } from "vitest";

import { shouldFinishActivityAsCompleted } from "./turnPlanningOutcome";
import { buildColumnMappingPlan, type MappingSourceSheet } from "./columnMappingPlan";
import type { ColumnMappingRule } from "./contracts";

describe("规划终态：失败不得记成已完成", () => {
  it("只有产出计划的这一轮才允许以默认'已完成'收尾", () => {
    expect(shouldFinishActivityAsCompleted("planned")).toBe(true);
  });

  it("失败 / 打断 / 跳过都不允许补'已完成'", () => {
    expect(shouldFinishActivityAsCompleted("failed")).toBe(false);
    expect(shouldFinishActivityAsCompleted("aborted")).toBe(false);
    expect(shouldFinishActivityAsCompleted("skipped")).toBe(false);
  });

  it("本地映射失败时抛错、且不产出任何计划（无计划 → 无写入）", () => {
    const sheet: MappingSourceSheet = {
      name: "表一",
      startRow: 1,
      startColumn: 1,
      columnCount: 2,
      values: [
        ["来源", "目标"],
        // 这个来源值不在映射表里：必须拒绝，不得兜底成空白标签。
        ["未收录类别", null]
      ]
    };
    const rule: ColumnMappingRule = {
      kind: "columnMapping",
      sourceSheet: "表一",
      targetSheet: "表一",
      sourceRange: "A2:A2",
      targetRange: "B2:B2",
      mapping: [{ source: "甲类", target: "一类" }],
      blankLabel: "未分类"
    };

    let producedPlan: unknown = null;
    expect(() => {
      producedPlan = buildColumnMappingPlan(rule, sheet, {
        summary: "s",
        warnings: [],
        assumptions: []
      });
    }).toThrow(/不在映射表内/);
    // 关键：抛错路径没有产出计划对象，因此没有预览卡、也没有写入。
    expect(producedPlan).toBeNull();
  });
});
