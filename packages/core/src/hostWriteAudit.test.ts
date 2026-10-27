import { describe, expect, it } from "vitest";
import type { HostWriteAudit, HostWriteObservation } from "./hostWriteAudit";
import {
  describeHostWriteObservation,
  normalizeHostWriteAudit,
  summarizeHostWrites
} from "./hostWriteAudit";

/**
 * 「没有观测」与「明确零写入」的分界（A1-E3 第三批）。
 *
 * 这些用例存在的理由：`/executions/{id}/host-writes` 是个 best-effort 上报口，
 * 一次都没上报时它返回的数组也是空的。过去验收说明把"空数组 + 0 次调用"读成
 * "明确零写入"，等于把**没收到**当成**没发生**。这里的判据是：
 *
 * - 只有 `complete + zero` 允许说"明确没有向宿主发起写调用"；
 * - `not_collected` / `partial` / 含 `unknown` 条目一律不许；
 * - "账本没有动作被认领"与"有动作可能写过但没观测"是两句不同的话，不许合并；
 * - 两端计量单位不同，文案里也不许互相折算。
 */

function observation(over: Partial<HostWriteObservation> = {}): HostWriteObservation {
  return {
    observation: "not_collected",
    callEvidence: "none",
    attemptedActions: 0,
    reportedActions: 0,
    missingActions: 0,
    unexpectedReports: 0,
    unit: "excel_sync_batches",
    lastReportedAt: null,
    ...over
  };
}

describe("观测覆盖 → 能不能说零写入", () => {
  it("空上报：没收到观测，既不下结论也不许说零", () => {
    const verdict = describeHostWriteObservation(
      observation({ observation: "not_collected", callEvidence: "none", attemptedActions: 1, missingActions: 1 })
    );
    expect(verdict.provesZeroWrites).toBe(false);
    expect(verdict.conclusive).toBe(false);
    expect(verdict.label).toContain("未采集");
    // 必须带出"账本认为可能写过"，否则读者会把"未采集"读成"没事"。
    expect(verdict.label).toContain("1 个动作可能写过");
  });

  it("未采集 + 账本里一个动作都没被认领：这是另一句话（不是同一句缺省）", () => {
    const verdict = describeHostWriteObservation(
      observation({ observation: "not_collected", callEvidence: "none", attemptedActions: 0 })
    );
    expect(verdict.label).toContain("未采集");
    expect(verdict.label).toContain("没有动作被认领");
    expect(verdict.label).not.toContain("可能写过");
    expect(verdict.provesZeroWrites).toBe(false);
  });

  it("部分覆盖 + 结果未知：不完整就不许说零", () => {
    const verdict = describeHostWriteObservation(
      observation({
        observation: "partial",
        callEvidence: "unknown",
        attemptedActions: 3,
        reportedActions: 1,
        missingActions: 2,
        lastReportedAt: 1_760_000_000
      })
    );
    expect(verdict.provesZeroWrites).toBe(false);
    expect(verdict.conclusive).toBe(false);
    expect(verdict.label).toContain("缺 2 个动作");
  });

  it("全覆盖且每次都是 0 次调用：这才是唯一的「明确零写入」", () => {
    const verdict = describeHostWriteObservation(
      observation({
        observation: "complete",
        callEvidence: "zero",
        attemptedActions: 2,
        reportedActions: 2,
        missingActions: 0,
        lastReportedAt: 1_760_000_000
      })
    );
    expect(verdict.provesZeroWrites).toBe(true);
    expect(verdict.conclusive).toBe(true);
    expect(verdict.label).toContain("零宿主刷新批次");
    expect(verdict.label).toContain("覆盖 2 个动作");
  });

  it("全覆盖但有条目结果未知：数字齐了也不能说零", () => {
    const verdict = describeHostWriteObservation(
      observation({
        observation: "complete",
        callEvidence: "unknown",
        attemptedActions: 1,
        reportedActions: 1,
        missingActions: 0
      })
    );
    expect(verdict.provesZeroWrites).toBe(false);
    expect(verdict.conclusive).toBe(false);
    expect(verdict.label).toContain("结果未知");
  });

  it("已经看到非零调用：即便覆盖不全也如实报出来，只是不许说结论", () => {
    const partial = describeHostWriteObservation(
      observation({ observation: "partial", callEvidence: "nonzero", attemptedActions: 3, reportedActions: 1, missingActions: 2 })
    );
    expect(partial.provesZeroWrites).toBe(false);
    expect(partial.conclusive).toBe(false);
    expect(partial.label).toContain("非零");
    expect(partial.label).toContain("缺 2 个动作");

    const full = describeHostWriteObservation(
      observation({ observation: "complete", callEvidence: "nonzero", attemptedActions: 2, reportedActions: 2 })
    );
    expect(full.conclusive).toBe(true);
    expect(full.provesZeroWrites).toBe(false);
  });

  it("WPS 的计量单位是「写 API 调用」，不冒充 Excel 的刷新批次", () => {
    const verdict = describeHostWriteObservation(
      observation({ observation: "complete", callEvidence: "zero", attemptedActions: 1, reportedActions: 1, unit: "wps_api_calls" })
    );
    expect(verdict.label).toContain("宿主写 API 调用");
    expect(verdict.label).not.toContain("批次");
  });

  it("矛盾的线上数据不许被升级：zero 但覆盖不全 / 有边界外上报", () => {
    // `partial + zero` 服务端不会产生；展示层不能假定它永不出现，更不能照着字面读。
    const partialZero = describeHostWriteObservation(
      observation({ observation: "partial", callEvidence: "zero", attemptedActions: 3, reportedActions: 1, missingActions: 2 })
    );
    expect(partialZero.provesZeroWrites).toBe(false);
    expect(partialZero.conclusive).toBe(false);
    expect(partialZero.label).toContain("只覆盖 1 个动作");
    expect(partialZero.label).toContain("不能说零写入");

    const strayReports = describeHostWriteObservation(
      observation({ observation: "complete", callEvidence: "zero", attemptedActions: 1, reportedActions: 2, unexpectedReports: 1 })
    );
    expect(strayReports.provesZeroWrites).toBe(false);
    expect(strayReports.label).toContain("执行边界之外");

    // 自洽的 complete+zero 仍然为真（负向判据要有正向锚点）。
    const consistent = describeHostWriteObservation(
      observation({ observation: "complete", callEvidence: "zero", attemptedActions: 1, reportedActions: 1 })
    );
    expect(consistent.provesZeroWrites).toBe(true);
  });

  it("只有 complete+zero 一种组合为真（其余组合逐个反查）", () => {
    const truths: string[] = [];
    for (const coverage of ["not_collected", "partial", "complete"] as const) {
      for (const evidence of ["none", "zero", "nonzero", "unknown"] as const) {
        const verdict = describeHostWriteObservation(
          observation({ observation: coverage, callEvidence: evidence, attemptedActions: 1, reportedActions: 1 })
        );
        if (verdict.provesZeroWrites) truths.push(`${coverage}+${evidence}`);
      }
    }
    expect(truths).toEqual(["complete+zero"]);
  });
});

describe("汇总与规范化不许替观测兜底", () => {
  it("写前失败的条目，即使调用数非零也不计入总数", () => {
    const entries: HostWriteAudit[] = [
      { actionIndex: 0, actionType: "writeValues", hostWriteCalls: 3, mutatingCalls: 3, outcome: "failed_before_write", readback: "not_applicable" },
      { actionIndex: 1, actionType: "writeValues", hostWriteCalls: 2, mutatingCalls: 1, outcome: "written", readback: "verified" },
      { actionIndex: 2, actionType: "writeValues", hostWriteCalls: 1, mutatingCalls: 1, outcome: "unknown", readback: "failed" }
    ];
    const summary = summarizeHostWrites(entries);
    expect(summary.reportedActions).toBe(3);
    expect(summary.hostWriteCalls).toBe(3);
    expect(summary.actionsWritten).toBe(1);
    expect(summary.actionsFailedBeforeWrite).toBe(1);
    expect(summary.actionsUnknown).toBe(1);
    expect(summary.readbackFailed).toBe(1);
  });

  it("不合法计数一律丢弃，不静默补 0", () => {
    const base = { actionIndex: 0, actionType: "writeValues", mutatingCalls: 1, outcome: "written", readback: "verified" };
    expect(normalizeHostWriteAudit({ ...base, hostWriteCalls: -1 })).toBeNull();
    expect(normalizeHostWriteAudit({ ...base, hostWriteCalls: 1.5 })).toBeNull();
    expect(normalizeHostWriteAudit({ ...base, hostWriteCalls: Number.NaN })).toBeNull();
    expect(normalizeHostWriteAudit({ ...base, hostWriteCalls: "2" })).toBeNull();
    // 合法条目原样过，错误码只是截断，不改判据。
    expect(normalizeHostWriteAudit({ ...base, hostWriteCalls: 0, errorCode: "X".repeat(80) }))
      .toMatchObject({ hostWriteCalls: 0, errorCode: "X".repeat(60) });
  });
});
