import { beforeEach, describe, expect, it } from "vitest";
import {
  attestResultSourceDependency,
  attestSourceDependency,
  recordTrustedToolReads,
  resetTrustedSourceReads,
  trustedToolReads
} from "./planSourceDependency";
import type { SourceDependencyPlan } from "./planSourceDependency";

/*
 * 用 core 自己的 `SourceDependencyPlan`（只有一个字段的最小结构化接口）。
 * **不要** import 加载项的 `AnalysisPlan`：那是宿主侧的完整计划形状，core 里
 * 没有也不该有 —— 引它会直接编译不过（`Module './contracts' has no exported
 * member 'AnalysisPlan'`），而且会把"计划必须有 actions"这类宿主约束拖进 core。
 */
function plan(overrides: SourceDependencyPlan = {}): SourceDependencyPlan {
  /*
   * **默认不给 `sourceDependencySheets`**：这是"旧登记项 / 服务端没声明来源"
   * 的场景，必须真的缺字段（默认塞 `["结果"]` 会让 `planOf({})` 变成"声明了
   * 来源"，把那条退回单侧记账的用例测成别的东西）。要来源时显式传。
   *
   * core 侧不做"计划必须有 id/actions"的校验，所以这里只保留这一个字段。
   */
  return { ...overrides };
}

/** 结果计划路径的别名：形状与 `plan` 相同，读起来更贴近"结果编译出的计划"。 */
const planOf = plan;

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
    expect(stripped.sourceDependencySheets).toBeUndefined();
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

describe("结果计划的来源依赖（服务端结果来源 ∩ 本端记账）", () => {
  it("两端都认的表才算依赖：只交给执行器的那些表", () => {
    /*
     * 本轮 prepare 了 A/B 两表（本端记账 = {A,B}），但只把 A 交给执行器
     * （服务端结果来源 = {A}）。诚实的依赖是 **A** —— 不是并集 {A,B}
     * （那会虚报 B 也是这份结果的来源），也不是服务端单方声明（本端无凭据）。
     */
    const plan = planOf({ sourceDependencySheets: ["A"] });
    const attested = attestResultSourceDependency(plan, [
      { origin: "tool-execution-record", sheets: ["A", "B"] }
    ]);
    expect(attested.sourceDependencySheets).toEqual(["A"]);
  });

  it("服务端声明了本端没读过的表 → 不采信那张表（无凭据）", () => {
    const plan = planOf({ sourceDependencySheets: ["A", "未读过的表"] });
    const attested = attestResultSourceDependency(plan, [
      { origin: "tool-execution-record", sheets: ["A"] }
    ]);
    expect(attested.sourceDependencySheets).toEqual(["A"]);
  });

  it("服务端没给来源（旧登记项）→ 退回单侧记账，仍不凭空造来源", () => {
    const plan = planOf({});
    const attested = attestResultSourceDependency(plan, [
      { origin: "tool-execution-record", sheets: ["A"] }
    ]);
    expect(attested.sourceDependencySheets).toEqual(["A"]);
  });

  it("本端没有任何记账 → 删字段（依赖不确定），即使服务端声明了来源", () => {
    const plan = planOf({ sourceDependencySheets: ["A"] });
    const attested = attestResultSourceDependency(plan, []);
    expect("sourceDependencySheets" in attested).toBe(false);
  });

  it("两侧都为真但无交集 → 删字段，不给出一个空的「依赖」", () => {
    const plan = planOf({ sourceDependencySheets: ["A"] });
    const attested = attestResultSourceDependency(plan, [
      { origin: "tool-execution-record", sheets: ["B"] }
    ]);
    expect("sourceDependencySheets" in attested).toBe(false);
  });

  it("不改动原计划对象", () => {
    const plan = planOf({ sourceDependencySheets: ["A"] });
    attestResultSourceDependency(plan, [
      { origin: "tool-execution-record", sheets: ["A"] }
    ]);
    expect(plan.sourceDependencySheets).toEqual(["A"]);
  });
});
