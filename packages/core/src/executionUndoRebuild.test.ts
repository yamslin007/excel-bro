import { describe, expect, it } from "vitest";
import {
  rebuildExecutionUndo,
  undoRecoveryDetail,
  type UndoRebuildCheckpoint,
  type UndoRebuildInput,
  type UndoRebuildTarget
} from "./executionUndoRebuild";

/**
 * 现场取自 2026-10-31 真机（Excel）：三个 `writeFormulas` 动作，第二步回读后中断 →
 * 用户核对在途动作（applied）→ 改 `来源!B5` → 续跑被拒 → 重开。
 *
 * 账本那时的事实：`status=failed`、`next=2`、`inFlight=null`、
 * `checkpoint` 覆盖第 1、2 步（各带写后指纹），第 3 步没写。
 */
const SOURCE = "来源";
const TARGETS = ["结果1", "结果2", "结果3"];

function checkpoint(overrides: Partial<UndoRebuildCheckpoint> = {}): UndoRebuildCheckpoint {
  return {
    planId: "plan-three-formulas",
    executionId: "op-three-formulas",
    workbookIdentity: "url:https://contoso.sharepoint.com/book.xlsx",
    coverage: "partial",
    completedActionCount: 2,
    postStateCoverage: "complete",
    ranges: [
      { actionIndex: 0, sheet: "结果1", range: "B2", postFingerprint: "v2:aaa" },
      { actionIndex: 1, sheet: "结果2", range: "B2", postFingerprint: "v2:bbb" }
    ],
    ...overrides
  };
}

const targets: (UndoRebuildTarget | null)[] = TARGETS.map((sheet, index) => ({
  actionIndex: index,
  sheet,
  range: "B2"
}));

function input(overrides: Partial<UndoRebuildInput<UndoRebuildCheckpoint>> = {}) {
  return {
    operationId: "op-three-formulas",
    status: "failed" as const,
    planId: "plan-three-formulas",
    workbookIdentity: "url:https://contoso.sharepoint.com/book.xlsx",
    nextActionIndex: 2,
    inFlightActionIndex: null,
    requiresReconciliation: false,
    checkpoint: checkpoint(),
    recordedAuthorization: [SOURCE, ...TARGETS],
    currentAuthorization: null,
    currentIdentity: "url:https://contoso.sharepoint.com/book.xlsx",
    targets,
    disposition: null,
    ...overrides
  };
}

describe("重开后撤销已写入内容：资格与续跑资格分开", () => {
  it("第二步中断、核对后来源已变化：撤销**仍然可用**（来源不是撤销的输入）", () => {
    /*
     * 这条是本轮的正面判据：续跑判据会因"来源已变化"拒绝，而撤销必须照旧可用。
     * 入参里**根本没有来源证据字段** —— 这不是"恰好没传"，而是判据的边界：
     * 撤销依赖的是写前恢复点与目标写后现场，不是数据来源是否还一致。
     */
    const outcome = rebuildExecutionUndo(input());
    expect(outcome.available).toBe(true);
    if (!outcome.available) return;
    expect(outcome.writtenActionCount).toBe(2);
    expect(outcome.rangeCount).toBe(2);
    // 原样返回账本里的写前快照：区域内容一个字都不改写。
    expect(outcome.checkpoint.ranges.map((range) => [range.sheet, range.range]))
      .toEqual([["结果1", "B2"], ["结果2", "B2"]]);
    expect(outcome.checkpoint.ranges.map((range) => range.postFingerprint))
      .toEqual(["v2:aaa", "v2:bbb"]);
  });

  it("在途动作没核对：明确阻断，并说清「先核对再撤销」（不撤销没写入的动作）", () => {
    const outcome = rebuildExecutionUndo(input({
      status: "needs_reconciliation",
      inFlightActionIndex: 1,
      requiresReconciliation: true,
      nextActionIndex: 1,
      checkpoint: checkpoint({
        completedActionCount: 1,
        ranges: [{ actionIndex: 0, sheet: "结果1", range: "B2", postFingerprint: "v2:aaa" }]
      })
    }));
    expect(outcome.available).toBe(false);
    if (outcome.available) return;
    expect(outcome.reason).toBe("in_flight_unknown");
    expect(outcome.detail).toContain("第 2 个动作可能在停止前已经写入");
    expect(outcome.detail).toContain("先核对在途动作");
  });

  it("服务端说需要核对却没给位置：同样阻断，不自己挑一个动作当在途", () => {
    const outcome = rebuildExecutionUndo(input({ requiresReconciliation: true }));
    expect(outcome.available).toBe(false);
    if (outcome.available) return;
    expect(outcome.reason).toBe("in_flight_unknown");
  });

  it("一个动作都没写成：没有可撤销的内容，界面不该画撤销说明", () => {
    const outcome = rebuildExecutionUndo(input({
      nextActionIndex: 0,
      checkpoint: checkpoint({ completedActionCount: 0, ranges: [] })
    }));
    expect(outcome.available).toBe(false);
    if (outcome.available) return;
    expect(outcome.reason).toBe("nothing_written");
  });

  it("账本没有写前快照：阻断，且明说不用当前单元格反推", () => {
    const outcome = rebuildExecutionUndo(input({ checkpoint: null }));
    expect(outcome.available).toBe(false);
    if (outcome.available) return;
    expect(outcome.reason).toBe("material_missing");
    expect(outcome.detail).toContain("不会用当前单元格反推");
    expect(outcome.detail).toContain("也不会直接清空目标");
  });

  it("写前快照缺执行后校验：阻断（目标是否被改过无从判断）", () => {
    const missing = rebuildExecutionUndo(input({
      checkpoint: checkpoint({
        ranges: [
          { actionIndex: 0, sheet: "结果1", range: "B2", postFingerprint: "v2:aaa" },
          { actionIndex: 1, sheet: "结果2", range: "B2" }
        ]
      })
    }));
    expect(missing.available).toBe(false);
    if (missing.available) return;
    expect(missing.reason).toBe("coverage_incomplete");
    expect(missing.detail).toContain("「结果2」B2 缺少执行后校验");

    const partialCoverage = rebuildExecutionUndo(input({
      checkpoint: checkpoint({ postStateCoverage: "partial" })
    }));
    expect(partialCoverage.available).toBe(false);
    if (partialCoverage.available) return;
    expect(partialCoverage.reason).toBe("coverage_incomplete");
  });

  it("进度与材料对不上 / 快照属于别的执行：阻断，不猜", () => {
    const progressGap = rebuildExecutionUndo(input({ nextActionIndex: 2,
      checkpoint: checkpoint({ completedActionCount: 1 }) }));
    expect(progressGap.available).toBe(false);
    if (progressGap.available) return;
    expect(progressGap.reason).toBe("plan_mismatch");
    expect(progressGap.detail).toContain("对不上");

    const foreign = rebuildExecutionUndo(input({ checkpoint: checkpoint({ executionId: "op-other" }) }));
    expect(foreign.available).toBe(false);
    if (foreign.available) return;
    expect(foreign.reason).toBe("plan_mismatch");
  });

  it("缺某一步的写前快照：阻断（撤销会漏掉它写过的单元格）", () => {
    const outcome = rebuildExecutionUndo(input({
      nextActionIndex: 2,
      checkpoint: checkpoint({
        completedActionCount: 2,
        ranges: [{ actionIndex: 0, sheet: "结果1", range: "B2", postFingerprint: "v2:aaa" }]
      })
    }));
    expect(outcome.available).toBe(false);
    if (outcome.available) return;
    expect(outcome.reason).toBe("coverage_incomplete");
    expect(outcome.detail).toContain("第 2 个动作没有留下写前快照");
  });

  it("快照区域不是该动作的目标：阻断（不能顺手恢复计划外的单元格）", () => {
    const outcome = rebuildExecutionUndo(input({
      checkpoint: checkpoint({
        ranges: [
          { actionIndex: 0, sheet: "结果1", range: "B2", postFingerprint: "v2:aaa" },
          { actionIndex: 1, sheet: "结果2", range: "C9", postFingerprint: "v2:bbb" }
        ]
      })
    }));
    expect(outcome.available).toBe(false);
    if (outcome.available) return;
    expect(outcome.reason).toBe("plan_mismatch");
    expect(outcome.detail).toContain("「结果2」C9 不是第 2 个动作的目标");
  });

  it("一个动作写多块区域（同一下标多条记录）：仍可用，按记录逐块恢复", () => {
    const outcome = rebuildExecutionUndo(input({
      checkpoint: checkpoint({
        completedActionCount: 1,
        ranges: [
          { actionIndex: 0, sheet: "结果1", range: "B2", postFingerprint: "v2:aaa" },
          { actionIndex: 0, sheet: "结果1", range: "B2", scopeDigest: "v2:aaa" }
        ]
      }),
      nextActionIndex: 1
    }));
    expect(outcome.available).toBe(true);
    if (!outcome.available) return;
    expect(outcome.rangeCount).toBe(2);
    expect(outcome.writtenActionCount).toBe(1);
  });

  it("授权收窄：阻断并指名少了哪张表（不自行扩大授权）", () => {
    const outcome = rebuildExecutionUndo(input({ currentAuthorization: [SOURCE, "结果1"] }));
    expect(outcome.available).toBe(false);
    if (outcome.available) return;
    expect(outcome.reason).toBe("plan_target_not_authorized_now");
    expect(outcome.detail).toContain("「结果2」");
  });

  it("登记授权里就没有这张表：记录自相矛盾，阻断", () => {
    const outcome = rebuildExecutionUndo(input({ recordedAuthorization: [SOURCE, "结果1"] }));
    expect(outcome.available).toBe(false);
    if (outcome.available) return;
    expect(outcome.reason).toBe("plan_target_never_authorized");
  });

  it("身份不符或读不到：阻断（撤销不跨工作簿）", () => {
    const mismatch = rebuildExecutionUndo(input({ currentIdentity: "url:https://contoso.sharepoint.com/other.xlsx" }));
    expect(mismatch.available).toBe(false);
    if (mismatch.available) return;
    expect(mismatch.reason).toBe("identity_mismatch");

    const unavailable = rebuildExecutionUndo(input({ currentIdentity: null }));
    expect(unavailable.available).toBe(false);
    if (unavailable.available) return;
    expect(unavailable.reason).toBe("identity_unavailable");
  });

  it("已经处置过（撤销/放弃）：不再重建入口，两个取值不互冒充", () => {
    const undone = rebuildExecutionUndo(input({ disposition: { undoneAt: 1_760_000_000, abandonedAt: null } }));
    expect(undone.available).toBe(false);
    if (undone.available) return;
    expect(undone.reason).toBe("disposition_recorded");
    expect(undone.detail).toContain("已经撤销过");

    const abandoned = rebuildExecutionUndo(input({ disposition: { undoneAt: null, abandonedAt: 1_760_000_000 } }));
    expect(abandoned.available).toBe(false);
    if (abandoned.available) return;
    expect(abandoned.reason).toBe("disposition_recorded");
    expect(abandoned.detail).toContain("「放弃」");
    expect(abandoned.detail).not.toContain("已经撤销过");
  });

  it("已完成：撤销跟随结果卡，不在恢复卡上重复给一个", () => {
    const outcome = rebuildExecutionUndo(input({ status: "completed" }));
    expect(outcome.available).toBe(false);
    if (outcome.available) return;
    expect(outcome.reason).toBe("already_completed");
  });

  it("地址比较忽略 $ 与大小写，但不接受「包含关系的近似」", () => {
    const absolute = rebuildExecutionUndo(input({
      checkpoint: checkpoint({
        ranges: [
          { actionIndex: 0, sheet: "结果1", range: "$B$2", postFingerprint: "v2:aaa" },
          { actionIndex: 1, sheet: "结果2", range: "b2", postFingerprint: "v2:bbb" }
        ]
      })
    }));
    expect(absolute.available).toBe(true);

    const wider = rebuildExecutionUndo(input({
      checkpoint: checkpoint({
        ranges: [
          { actionIndex: 0, sheet: "结果1", range: "B2:C3", postFingerprint: "v2:aaa" },
          { actionIndex: 1, sheet: "结果2", range: "B2", postFingerprint: "v2:bbb" }
        ]
      })
    }));
    expect(wider.available).toBe(false);
  });
});

describe("撤销说明的措辞（唯一出处）", () => {
  it("说清「只写回本次目标」和「来源保持你现在的修改」", () => {
    const detail = undoRecoveryDetail(2, 2, [SOURCE]);
    expect(detail).toContain("2 个动作");
    expect(detail).toContain("2 个区域");
    expect(detail).toContain(`来源工作表「${SOURCE}」保持你现在的修改`);
    expect(detail).toContain("撤销不会重新执行计划");
  });

  it("没有来源表时不编一句来源说明", () => {
    const detail = undoRecoveryDetail(1, 1, []);
    expect(detail).toContain("只写回本次执行碰过的目标区域");
    expect(detail).not.toContain("来源工作表");
  });
});
