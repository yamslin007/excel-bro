import { beforeEach, describe, expect, it } from "vitest";
import type { AnalysisPlan } from "./contracts";
import {
  attestSourceDependency,
  recordTrustedToolReads,
  resetTrustedSourceReads,
  trustedToolReads
} from "./planSourceDependency";

function plan(overrides: Partial<AnalysisPlan> = {}): AnalysisPlan {
  return {
    id: "plan-1",
    title: "t",
    summary: "s",
    assumptions: [],
    warnings: [],
    actions: [
      { type: "writeValues", sheet: "结果", range: "A1", values: [["ok"]] }
    ],
    ...overrides
  };
}

beforeEach(() => {
  resetTrustedSourceReads();
});

describe("可信来源记账", () => {
  it("同一轮多次读取按并集累积，不同轮互不串台", () => {
    recordTrustedToolReads("turn-1", ["数据"]);
    recordTrustedToolReads("turn-1", [" 数据 ", "辅助"]);
    recordTrustedToolReads("turn-2", ["别的表"]);

    expect(trustedToolReads("turn-1")).toEqual({
      origin: "tool-execution-record",
      sheets: ["数据", "辅助"]
    });
    expect(trustedToolReads("turn-2")?.sheets).toEqual(["别的表"]);
    expect(trustedToolReads("turn-3")).toBeNull();
  });

  it("空读与无名轮不产生记录 —— 空数组不等于「依赖为空」", () => {
    recordTrustedToolReads("turn-1", []);
    recordTrustedToolReads("turn-1", ["", "   "]);
    recordTrustedToolReads("", ["数据"]);

    expect(trustedToolReads("turn-1")).toBeNull();
    expect(trustedToolReads("")).toBeNull();
  });

  it("取出来的记录是副本，改它不会污染记账", () => {
    recordTrustedToolReads("turn-1", ["数据"]);
    const first = trustedToolReads("turn-1");
    (first!.sheets as string[]).push("伪造表");

    expect(trustedToolReads("turn-1")?.sheets).toEqual(["数据"]);
  });

  it("只保留最近若干轮：旧轮被淘汰，避免长会话无限增长", () => {
    for (let index = 0; index < 40; index += 1) {
      recordTrustedToolReads(`turn-${index}`, ["数据"]);
    }

    expect(trustedToolReads("turn-0")).toBeNull();
    expect(trustedToolReads("turn-39")).not.toBeNull();
  });
});

describe("计划 attestation", () => {
  it("有可信记录就写成实际来源依赖（去重、去空）", () => {
    const attested = attestSourceDependency(plan(), {
      origin: "deterministic-local-read",
      sheets: [" 数据 ", "数据", "", "辅助"]
    });

    expect(attested.sourceDependencySheets).toEqual(["数据", "辅助"]);
  });

  it("没有可信记录就**删掉**字段：模型/服务端自报的来源一律剥掉", () => {
    const forged = plan({ sourceDependencySheets: ["数据"] });

    const stripped = attestSourceDependency(forged, null);

    expect("sourceDependencySheets" in stripped).toBe(false);
    expect(stripped.actions).toEqual(forged.actions);
  });

  it("记录里只有空值时同样按「不确定」处理，不留空数组", () => {
    const stripped = attestSourceDependency(plan(), {
      origin: "tool-execution-record",
      sheets: ["", "  "]
    });

    expect("sourceDependencySheets" in stripped).toBe(false);
  });

  it("不改动原对象，也不从写入目标反推来源", () => {
    const original = plan({ sourceDependencySheets: ["伪造表"] });

    const attested = attestSourceDependency(original, {
      origin: "tool-execution-record",
      sheets: ["数据"]
    });

    expect(original.sourceDependencySheets).toEqual(["伪造表"]);
    expect(attested.sourceDependencySheets).toEqual(["数据"]);
    // 写入目标是「结果」，来源集合里没有它，也不是"授权范围减目标"推出来的。
    expect(attested.sourceDependencySheets).not.toContain("结果");
  });
});
