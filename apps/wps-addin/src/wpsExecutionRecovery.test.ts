import { describe, expect, it } from "vitest";
import { SESSION_RECOVERY_LABEL, VERIFICATION_PASSED_LABEL } from "@excel-bro/core";
import {
  WPS_EXECUTION_OUTCOME_DETAIL_SUFFIX,
  WPS_EXECUTION_OUTCOME_LABEL,
  WPS_RECOVERY_CARD_LABEL,
  wpsExecutionCardTone,
  wpsExecutionHasRecoveryPoint,
  wpsRecoveryCardState
} from "./wpsExecutionRecovery";

/**
 * 撤销按钮的归属判定必须是**纯函数**，否则面板里那段 JSX 只能靠"看着对不对"
 * 来判断。这里只钉判定与文案，渲染由 `WpsAgentPanel.test.tsx` 从 DOM 上验。
 */
describe("WPS 恢复点卡片：按 executionId 归属", () => {
  const id = "exec-1";

  it("当前活跃恢复点与 executionId 匹配且允许撤销 → 可撤销", () => {
    expect(
      wpsRecoveryCardState({ executionId: id, undone: false, live: { executionId: id, state: "verified" } })
    ).toBe("undoable");
    expect(
      wpsRecoveryCardState({
        executionId: id,
        undone: false,
        live: { executionId: id, state: "partial_write_recoverable" }
      })
    ).toBe("undoable");
  });

  it("新执行取代旧恢复点（活跃点换了 executionId）→ 恢复点已不可用", () => {
    expect(
      wpsRecoveryCardState({ executionId: id, undone: false, live: { executionId: "exec-2", state: "verified" } })
    ).toBe("unavailable");
  });

  it("当前会话内没有恢复点（面板重开、加载项会话已结束）→ 恢复点已不可用", () => {
    expect(wpsRecoveryCardState({ executionId: id, undone: false, live: null })).toBe("unavailable");
  });

  it("宿主报告撤销冲突 → 撤销冲突（不可点，但仍如实展示）", () => {
    expect(
      wpsRecoveryCardState({ executionId: id, undone: false, live: { executionId: id, state: "undo_conflict" } })
    ).toBe("conflict");
  });

  it("没有留下恢复点的状态（restored / ready / applying / undoing）→ 不可用", () => {
    for (const state of ["restored", "ready", "applying", "undoing"]) {
      expect(wpsRecoveryCardState({ executionId: id, undone: false, live: { executionId: id, state } })).toBe(
        "unavailable"
      );
    }
  });

  it("已撤销是终端态：即使活跃点还在，也不许重新变回可点", () => {
    expect(
      wpsRecoveryCardState({ executionId: id, undone: true, live: { executionId: id, state: "verified" } })
    ).toBe("undone");
    expect(wpsRecoveryCardState({ executionId: id, undone: true, live: null })).toBe("undone");
  });

  it("四种状态各有自己的中文标签，互不重复、互不冒充", () => {
    // 裸的"可撤销"会让人以为恢复点是持久的 —— 必须带上"会话内"，
    // 且与两端共用的 `SESSION_RECOVERY_LABEL` 同源。
    expect(WPS_RECOVERY_CARD_LABEL.undoable).toBe(SESSION_RECOVERY_LABEL);
    expect(WPS_RECOVERY_CARD_LABEL.undoable).toContain("会话内");
    expect(WPS_RECOVERY_CARD_LABEL.undoable).toContain("可撤销");
    expect(WPS_RECOVERY_CARD_LABEL.conflict).toBe("撤销冲突");
    expect(WPS_RECOVERY_CARD_LABEL.undone).toBe("已撤销");
    expect(WPS_RECOVERY_CARD_LABEL.unavailable).toBe("恢复点已不可用");
    const labels = Object.values(WPS_RECOVERY_CARD_LABEL);
    expect(new Set(labels).size).toBe(labels.length);
    // "不可用"不能长得像"可撤销"：否则卡片降级时用户会以为还能点。
    expect(WPS_RECOVERY_CARD_LABEL.unavailable).not.toContain("可撤销");
    expect(WPS_RECOVERY_CARD_LABEL.undone).not.toContain("可撤销");
  });

  it("只有真的留下恢复点的执行状态才画卡", () => {
    expect(wpsExecutionHasRecoveryPoint("verified")).toBe(true);
    expect(wpsExecutionHasRecoveryPoint("partial_write_recoverable")).toBe(true);
    expect(wpsExecutionHasRecoveryPoint("undo_conflict")).toBe(true);
    expect(wpsExecutionHasRecoveryPoint("restored")).toBe(false);
    expect(wpsExecutionHasRecoveryPoint("ready")).toBe(false);
    expect(wpsExecutionHasRecoveryPoint("applying")).toBe(false);
    expect(wpsExecutionHasRecoveryPoint("undoing")).toBe(false);
    expect(wpsExecutionHasRecoveryPoint(null)).toBe(false);
    expect(wpsExecutionHasRecoveryPoint(undefined)).toBe(false);
  });
});

describe("WPS 执行结果卡的三种事实", () => {
  it("完全回读通过才允许出现共享的「验证通过」字样", () => {
    expect(WPS_EXECUTION_OUTCOME_LABEL.readback_passed).toContain(VERIFICATION_PASSED_LABEL);
    // 反例锚点：另两种结果**不得**借用这句。
    expect(WPS_EXECUTION_OUTCOME_LABEL.unverified).not.toContain(VERIFICATION_PASSED_LABEL);
    expect(WPS_EXECUTION_OUTCOME_LABEL.partial).not.toContain(VERIFICATION_PASSED_LABEL);
  });

  it("三种结果的标题两两不同", () => {
    const labels = Object.values(WPS_EXECUTION_OUTCOME_LABEL);
    expect(new Set(labels).size).toBe(labels.length);
  });

  it("只有完全回读通过才写「已回读核对」", () => {
    expect(WPS_EXECUTION_OUTCOME_DETAIL_SUFFIX.readback_passed).toContain("已回读核对");
    expect(WPS_EXECUTION_OUTCOME_DETAIL_SUFFIX.unverified).not.toContain("已回读核对");
    expect(WPS_EXECUTION_OUTCOME_DETAIL_SUFFIX.partial).not.toContain("已回读核对");
  });

  it("卡片的色调类名不把「未证实 / 部分写入」冒充成 verified", () => {
    expect(wpsExecutionCardTone({ outcome: "readback_passed", undone: false })).toBe("verified");
    expect(wpsExecutionCardTone({ outcome: "unverified", undone: false })).toBe("unverified");
    expect(wpsExecutionCardTone({ outcome: "partial", undone: false })).toBe("partial");
    expect(wpsExecutionCardTone({ outcome: "readback_passed", undone: true })).toBe("undone");
    // 已撤销优先于一切结果色调。
    expect(wpsExecutionCardTone({ outcome: "partial", undone: true })).toBe("undone");
  });
});
