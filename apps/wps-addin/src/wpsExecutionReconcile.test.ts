import { describe, expect, it } from "vitest";
import type { WpsExecutionState } from "../../excel-addin/src/contracts";
import {
  wpsIdentityVerdict,
  wpsReconcileVerdict,
  wpsReconcileActions
} from "./wpsExecutionReconcile";

function state(overrides: Partial<WpsExecutionState>): WpsExecutionState {
  return {
    operationId: "wps-execution-0001",
    status: "running",
    revision: 1,
    workbookIdentity: "wps-path:c:\\data\\测试.xlsx",
    taskId: "task-keep",
    descriptor: { actions: [] },
    authorizedSheetNames: ["Sheet1"],
    nextActionIndex: 0,
    totalActions: 2,
    inFlightActionIndex: null,
    beforeSnapshot: null,
    checkpoint: null,
    result: null,
    error: null,
    canResume: true,
    ownerActive: false,
    requiresReconciliation: false,
    ownerToken: null,
    ...overrides
  };
}

describe("wpsIdentityVerdict", () => {
  it("完整字符串相等才判 match", () => {
    expect(
      wpsIdentityVerdict({
        currentIdentity: "wps-path:c:\\data\\测试.xlsx",
        recordedIdentity: "wps-path:c:\\data\\测试.xlsx"
      })
    ).toEqual({ kind: "match" });
  });

  it("同名异路径（不同目录）判 mismatch，不靠工作簿名放行", () => {
    expect(
      wpsIdentityVerdict({
        currentIdentity: "wps-path:c:\\other\\测试.xlsx",
        recordedIdentity: "wps-path:c:\\data\\测试.xlsx"
      }).kind
    ).toBe("mismatch");
  });

  it("当前未保存 / 读不到身份 → unavailable，不猜成同一个文件", () => {
    expect(
      wpsIdentityVerdict({
        currentIdentity: null,
        recordedIdentity: "wps-path:c:\\data\\测试.xlsx"
      }).kind
    ).toBe("unavailable");
  });
});

describe("wpsReconcileVerdict 收紧后的判定", () => {
  it.each([
    ["另一执行者仍持权", { ownerActive: true }, "completed", "blocked"],
    ["在途动作不可接管", { canResume: false, inFlightActionIndex: 0 }, "completed", "blocked"],
    ["无在途但现场被修改", {}, "conflict", "blocked"],
    ["剩余动作但现场未知", {}, "unknown", "insufficient"],
    ["零剩余动作但现场未知", { nextActionIndex: 2 }, "unknown", "insufficient"],
    ["动作游标越界", { nextActionIndex: 3 }, "completed", "insufficient"]
  ] as const)("%s 不提供执行或补记按钮", (_label, patch, localInspection, kind) => {
    const verdict = wpsReconcileVerdict({ state: state(patch), identityMatch: true, localInspection });
    expect(verdict.kind).toBe(kind);
    expect(wpsReconcileActions(verdict)).toEqual(["dismiss"]);
  });
  it("服务端 completed → 安全重放，不重写", () => {
    const verdict = wpsReconcileVerdict({
      state: state({ status: "completed" }),
      identityMatch: true,
      localInspection: null
    });
    expect(verdict.kind).toBe("completed");
    expect(wpsReconcileActions(verdict)).toEqual(["dismiss"]);
  });

  it("撤销过（disposition.undoneAt）→ 终态 undone，不给任何按钮", () => {
    const verdict = wpsReconcileVerdict({
      state: state({ disposition: { undoneAt: 1_790_000_700_000, abandonedAt: null } }),
      identityMatch: true,
      localInspection: "conflict"
    });
    expect(verdict.kind).toBe("undone");
    expect(wpsReconcileActions(verdict)).toEqual([]);
  });

  it("放弃过（disposition.abandonedAt）→ blocked，不冒充已撤销", () => {
    const verdict = wpsReconcileVerdict({
      state: state({ disposition: { undoneAt: null, abandonedAt: 1_790_000_700_000 } }),
      identityMatch: true,
      localInspection: "conflict"
    });
    expect(verdict.kind).toBe("blocked");
    expect(verdict).toMatchObject({ detail: expect.stringContaining("放弃变更") });
    expect(wpsReconcileActions(verdict)).toEqual(["dismiss"]);
  });

  it("disposition 终态优先于 status/现场：即使 status 仍 needs_reconciliation 也判 undone", () => {
    const verdict = wpsReconcileVerdict({
      state: state({
        status: "needs_reconciliation",
        inFlightActionIndex: 1,
        disposition: { undoneAt: 1_790_000_700_000, abandonedAt: null }
      }),
      identityMatch: true,
      localInspection: "not_started"
    });
    expect(verdict.kind).toBe("undone");
  });

  /*
   * 这一条是上一轮判定的**反向锚点**：那时 running 且无在途动作就返回
   * `continue`，不检查 canResume、身份、剩余动作，也不看现场。现在必须归成
   * `replan`（剩余动作非空 → 只能重新预览），绝不冒充"可以继续"。
   */
  it("running 且无在途动作、但剩余动作非空 → replan，不再等于可以继续", () => {
    const verdict = wpsReconcileVerdict({
      state: state({ status: "running", nextActionIndex: 1, totalActions: 2, inFlightActionIndex: null }),
      identityMatch: true,
      localInspection: "not_started"
    });
    expect(verdict).toEqual({ kind: "replan", remaining: 1 });
    expect(wpsReconcileActions(verdict)).toEqual(["replan", "dismiss"]);
  });

  it("failed 且剩余动作非空 → 同样只 replan，不自动放行", () => {
    const verdict = wpsReconcileVerdict({
      state: state({ status: "failed", nextActionIndex: 0, totalActions: 3 }),
      identityMatch: true,
      localInspection: "not_started"
    });
    expect(verdict).toEqual({ kind: "replan", remaining: 3 });
  });

  it("剩余动作为 0（只丢了最终回执）→ completed，只补记不重写", () => {
    const verdict = wpsReconcileVerdict({
      state: state({ status: "running", nextActionIndex: 2, totalActions: 2 }),
      identityMatch: true,
      localInspection: "completed"
    });
    expect(verdict.kind).toBe("completed");
    expect(wpsReconcileActions(verdict)).toEqual(["backfill", "dismiss"]);
  });

  it("剩余动作为 0 但本机现场仍是写前 → blocked，两者矛盾不自动补记", () => {
    const verdict = wpsReconcileVerdict({
      state: state({ status: "running", nextActionIndex: 2, totalActions: 2 }),
      identityMatch: true,
      localInspection: "not_started"
    });
    expect(verdict.kind).toBe("blocked");
    expect(wpsReconcileActions(verdict)).toEqual(["dismiss"]);
  });

  it("剩余动作为 0 但本机现场与写后冲突 → blocked", () => {
    const verdict = wpsReconcileVerdict({
      state: state({ status: "running", nextActionIndex: 2, totalActions: 2 }),
      identityMatch: true,
      localInspection: "conflict"
    });
    expect(verdict.kind).toBe("blocked");
  });

  it("canResume=false 一律不放行", () => {
    const verdict = wpsReconcileVerdict({
      state: state({ status: "running", canResume: false, nextActionIndex: 1, totalActions: 2 }),
      identityMatch: true,
      localInspection: "completed"
    });
    expect(verdict.kind).toBe("blocked");
    expect(wpsReconcileActions(verdict)).toEqual(["dismiss"]);
  });

  it("身份不一致 → insufficient，不给任何写入/补记按钮", () => {
    const verdict = wpsReconcileVerdict({
      state: state({ status: "running", nextActionIndex: 2, totalActions: 2 }),
      identityMatch: false,
      localInspection: "completed"
    });
    expect(verdict.kind).toBe("insufficient");
    expect(wpsReconcileActions(verdict)).toEqual(["dismiss"]);
  });

  it("有在途动作且本机现场与写后一致 → 补记为已应用", () => {
    const verdict = wpsReconcileVerdict({
      state: state({
        status: "needs_reconciliation",
        requiresReconciliation: true,
        inFlightActionIndex: 0
      }),
      identityMatch: true,
      localInspection: "completed"
    });
    // remaining = 在途这一格之后还有几个动作没跑：结论要能同时说"第 1 步已写入"和"第 2 步未执行"。
    expect(verdict).toEqual({ kind: "reconcile", actionIndex: 0, outcome: "applied", remaining: 1 });
    expect(wpsReconcileActions(verdict)).toEqual(["reconcile", "dismiss"]);
  });

  it("有在途动作且本机现场仍是写前 → 补记为未应用", () => {
    const verdict = wpsReconcileVerdict({
      state: state({
        status: "needs_reconciliation",
        requiresReconciliation: true,
        inFlightActionIndex: 1,
        nextActionIndex: 1
      }),
      identityMatch: true,
      localInspection: "not_started"
    });
    expect(verdict).toEqual({ kind: "reconcile", actionIndex: 1, outcome: "not_applied", remaining: 0 });
  });

  it("有在途动作但本机现场冲突 → blocked，不允许覆盖", () => {
    const verdict = wpsReconcileVerdict({
      state: state({
        status: "needs_reconciliation",
        requiresReconciliation: true,
        inFlightActionIndex: 0
      }),
      identityMatch: true,
      localInspection: "conflict"
    });
    expect(verdict.kind).toBe("blocked");
    expect(wpsReconcileActions(verdict)).toEqual(["dismiss"]);
  });

  it("有在途动作但本机恢复点读不到 → insufficient，不猜成未写入", () => {
    const verdict = wpsReconcileVerdict({
      state: state({
        status: "needs_reconciliation",
        requiresReconciliation: true,
        inFlightActionIndex: 0
      }),
      identityMatch: true,
      localInspection: null
    });
    expect(verdict.kind).toBe("insufficient");
    expect(wpsReconcileActions(verdict)).toEqual(["dismiss"]);
  });

  it("标记需要核对却没给在途动作号 → insufficient，不自己挑一个去补记", () => {
    const verdict = wpsReconcileVerdict({
      state: state({ status: "needs_reconciliation", requiresReconciliation: true, inFlightActionIndex: null }),
      identityMatch: true,
      localInspection: "completed"
    });
    expect(verdict.kind).toBe("insufficient");
  });

  it("服务端无记录 → 证据不足", () => {
    const verdict = wpsReconcileVerdict({ state: null, identityMatch: true, localInspection: null });
    expect(verdict.kind).toBe("insufficient");
    expect(wpsReconcileActions(verdict)).toEqual(["dismiss"]);
  });
});
