import { describe, expect, it } from "vitest";
import {
  EXECUTION_RECOVERY_ACTION_LABEL,
  EXECUTION_RECOVERY_NOTE,
  EXECUTION_RECOVERY_TITLE,
  EXECUTION_RECOVERY_STATE_LABEL,
  EXECUTION_RECOVERY_ENTRY_HINT,
  EXECUTION_RECOVERY_UNDONE_DETAIL,
  EXECUTION_RECOVERY_UNDONE_NOTE,
  executionRecoveryActions,
  executionRecoveryBadge,
  executionRecoveryDetail,
  executionRecoveryNote,
  executionRecoveryProgress,
  executionRecoveryResumeVerdict,
  executionRecoveryIsTerminal,
  executionRecoveryTerminalReason,
  executionRecoveryTerminalSummary,
  selectExecutionRecoveryEntry,
  type ExecutionRecoveryVerdict
} from "./executionRecoveryView";

describe("执行恢复结论的共享呈现契约", () => {
  it("核对与续跑是两个结论，不是一个按钮里的两步", () => {
    const reconcile: ExecutionRecoveryVerdict = {
      kind: "reconcile", actionIndex: 1, outcome: "pending", remaining: 1
    };
    const resume: ExecutionRecoveryVerdict = { kind: "resume", remaining: 1 };
    expect(executionRecoveryActions(reconcile)).toEqual(["reconcile", "dismiss"]);
    expect(executionRecoveryActions(resume)).toEqual(["resume", "dismiss"]);
    // 待核对时**绝不**同时给续跑按钮：那会让"补记"这一下顺手把剩下的也写了。
    expect(executionRecoveryActions(reconcile)).not.toContain("resume");
  });

  it("待核对要说清第几步、还剩几步没跑，两句都在同一行正文里", () => {
    const detail = executionRecoveryDetail({
      kind: "reconcile", actionIndex: 1, outcome: "pending", remaining: 1
    });
    expect(detail).toContain("第 2 个动作");
    expect(detail).toContain("还有 1 个动作没执行");
    expect(detail).toContain("不写任何单元格");
    // 核对结论不宣称已经写入：现场还没读过。
    expect(detail).not.toContain("已写入");
  });

  it("核对结论按现场回读的结果分别措辞", () => {
    expect(executionRecoveryDetail({ kind: "reconcile", actionIndex: 0, outcome: "applied" }))
      .toContain("现场与写后一致");
    expect(executionRecoveryDetail({ kind: "reconcile", actionIndex: 0, outcome: "not_applied" }))
      .toContain("现场仍是写前状态");
  });

  it("只剩回执时给补记，且明确说补记不重写原表", () => {
    const verdict: ExecutionRecoveryVerdict = { kind: "completed" };
    expect(executionRecoveryActions(verdict)).toEqual(["backfill", "dismiss"]);
    expect(executionRecoveryDetail(verdict)).toContain("不会重写原表");
    expect(executionRecoveryBadge(verdict)).toBe("已完成，仅缺回执");
  });

  it("账本本来就有完成结果时不给补记按钮，也不说仅缺回执", () => {
    const verdict: ExecutionRecoveryVerdict = { kind: "completed", recorded: true };
    expect(executionRecoveryActions(verdict)).toEqual(["dismiss"]);
    expect(executionRecoveryBadge(verdict)).toBe("已记录完成");
  });

  it("证据不足与现场冲突一律不给写入或补记按钮", () => {
    for (const verdict of [
      { kind: "blocked", detail: "原执行者仍持有执行权。" },
      { kind: "insufficient", detail: "无法读取当前目标现场。" }
    ] as ExecutionRecoveryVerdict[]) {
      expect(executionRecoveryActions(verdict)).toEqual(["dismiss"]);
    }
  });

  it("七种结论都有状态名，五个动作都有按钮文案", () => {
    expect(Object.keys(EXECUTION_RECOVERY_STATE_LABEL).sort())
      .toEqual(["blocked", "completed", "insufficient", "reconcile", "replan", "resume", "undone"]);
    expect(Object.keys(EXECUTION_RECOVERY_ACTION_LABEL).sort())
      .toEqual(["backfill", "dismiss", "reconcile", "replan", "resume"]);
  });

  it("卡片标题与固定说明是唯一出处", () => {
    expect(EXECUTION_RECOVERY_TITLE).toBe("上次写入的恢复核对");
    expect(EXECUTION_RECOVERY_NOTE).toContain("只使用当前已授权工作表");
  });
});

describe("续跑结论必须被来源复核支撑（2026-10-31 真机）", () => {
  it("没复核过来源就不出 resume，也不许说“来源已核对”", () => {
    const verdict = executionRecoveryResumeVerdict(1, { state: "unverified" });
    expect(verdict.kind).toBe("blocked");
    expect(executionRecoveryActions(verdict)).toEqual(["dismiss"]);
    const detail = executionRecoveryDetail(verdict);
    expect(detail).toContain("还没有和登记时的来源核对过");
    expect(detail).toContain("重新核对（只读）");
    expect(detail).not.toContain("来源已核对");
    expect(detail).not.toContain("已核对，可继续");
  });

  it("复核过且一致才出 resume，正文才提来源", () => {
    const verdict = executionRecoveryResumeVerdict(2, { state: "matched" });
    expect(verdict.kind).toBe("resume");
    expect(executionRecoveryActions(verdict)).toEqual(["resume", "dismiss"]);
    expect(executionRecoveryDetail(verdict)).toContain("来源和目标已核对");
  });

  it("计划不读别的表时出 resume，但不冒称比过来源", () => {
    const verdict = executionRecoveryResumeVerdict(1, { state: "not-required" });
    expect(verdict.kind).toBe("resume");
    const detail = executionRecoveryDetail(verdict);
    expect(detail).toContain("不读其他工作表作为来源");
    expect(detail).not.toContain("来源和目标已核对");
  });

  it("来源变了仍然是阻断，并把“改回去再复核”这条路说出来", () => {
    const verdict = executionRecoveryResumeVerdict(1, {
      state: "refused",
      reason: "来源已变化，不能沿用旧结果续跑。"
    });
    expect(verdict.kind).toBe("blocked");
    expect(executionRecoveryActions(verdict)).toEqual(["dismiss"]);
    const detail = executionRecoveryDetail(verdict);
    expect(detail).toContain("来源已变化，不能沿用旧结果续跑。");
    expect(detail).toContain("改回登记时的值");
    expect(detail).toContain("重新核对（只读）");
  });

  it("一个动作都不剩时是完成，不是续跑（来源结论不再参与）", () => {
    expect(executionRecoveryResumeVerdict(0, { state: "unverified" }).kind).toBe("completed");
    expect(executionRecoveryResumeVerdict(-3, { state: "refused", reason: "x" }).kind)
      .toBe("completed");
  });

  it("旧调用点（没给 source）不许被读成“来源已核对”", () => {
    const detail = executionRecoveryDetail({ kind: "resume", remaining: 1 });
    expect(detail).not.toContain("来源和目标已核对");
    expect(executionRecoveryActions({ kind: "resume", remaining: 1 }))
      .toEqual(["resume", "dismiss"]);
  });
});

describe("撤销是终态（2026-10-31 真机：同卡同时“已撤销”与“可继续”）", () => {
  it("已撤销不给任何动作，连清除提示都不给", () => {
    const verdict: ExecutionRecoveryVerdict = { kind: "undone" };
    expect(executionRecoveryActions(verdict)).toEqual([]);
    // 终态与“需要人工核对”**不是**同一句话：后者还留一个清除入口。
    expect(executionRecoveryActions({ kind: "blocked", detail: "x" })).toEqual(["dismiss"]);
  });

  it("badge 就是“已撤销本次执行”，不会退成“可继续剩余动作”", () => {
    const badge = executionRecoveryBadge({ kind: "undone" });
    expect(badge).toBe("已撤销本次执行");
    expect(badge).not.toContain("继续");
  });

  it("正文说清三件事：回到执行前、来源保留、不再重做", () => {
    const detail = executionRecoveryDetail({ kind: "undone" });
    expect(detail).toBe(EXECUTION_RECOVERY_UNDONE_DETAIL);
    expect(detail).toContain("回到执行前");
    expect(detail).toContain("来源工作表保持你现在的修改");
    expect(detail).toContain("不会重做任何写入");
    expect(detail).not.toContain("可继续");
  });

  it("账本已有处置痕的 undone 用自己的措辞，不用默认的“已回到执行前”", () => {
    // 本机撤销成功 ⇒ 表格确实回退了，默认文案“回到执行前”成立。
    expect(executionRecoveryDetail({ kind: "undone" })).toContain("回到执行前");
    // 账本已留处置痕（常为另一台机器撤的）⇒ 本机没回退、现场未知，不能用“回到执行前”。
    const ledger = executionRecoveryDetail({
      kind: "undone",
      detail: "这次执行已经撤销过（账本已留处置痕），不会再执行或补记任何动作。"
    });
    expect(ledger).toContain("已经撤销过");
    expect(ledger).toContain("账本已留处置痕");
    expect(ledger).not.toContain("回到执行前");
  });

  it("进度行仍然只报步数，不因为终态而改口径", () => {
    expect(executionRecoveryProgress({ kind: "undone" }, { nextActionIndex: 2, totalActions: 3 }))
      .toBe("已完成 2 / 3 步");
  });
});

describe("卡片边界说明按结论派生（2026-10-31 真机：已撤销卡还写着续跑引导）", () => {
  it("已撤销状态不出现续跑引导或审阅措辞", () => {
    const note = executionRecoveryNote({ kind: "undone" });
    // 终态说明必须**替换**写死的续跑引导，而不是在它后面补一句。
    expect(note).toBe(EXECUTION_RECOVERY_UNDONE_NOTE);
    // 不再是那句引导用户"续跑 + 审阅"的固定说明。
    expect(note).not.toContain("续跑通过复核后执行剩余动作");
    expect(note).not.toContain("再审阅结果");
    expect(note).not.toBe(EXECUTION_RECOVERY_NOTE);
    expect(EXECUTION_RECOVERY_UNDONE_NOTE).toContain("已撤销");
  });

  it("未结束状态原句照旧，不因为派生函数而丢字", () => {
    for (const verdict of [
      { kind: "resume", remaining: 1 },
      { kind: "reconcile", actionIndex: 1, outcome: "pending", remaining: 1 },
      { kind: "replan", remaining: 1 },
      { kind: "blocked", detail: "x" },
      { kind: "insufficient", detail: "x" },
      { kind: "completed" }
    ] as ExecutionRecoveryVerdict[]) {
      expect(executionRecoveryNote(verdict)).toBe(EXECUTION_RECOVERY_NOTE);
    }
    // 原句仍是那唯一一句，含授权边界与续跑引导。
    expect(EXECUTION_RECOVERY_NOTE).toContain("只使用当前已授权工作表");
    expect(EXECUTION_RECOVERY_NOTE).toContain("续跑通过复核后执行剩余动作");
  });
});

describe("恢复入口的唯一性（卡片搬到输入区上方之后）", () => {
  type Row = { id: string; card: { operationId: string } | null };

  it("取最后一条带卡的记录：新的执行覆盖旧的，两次执行不合并", () => {
    const rows: Row[] = [
      { id: "m1", card: { operationId: "op-old" } },
      { id: "m2", card: null },
      { id: "m3", card: { operationId: "op-new" } },
      { id: "m4", card: null }
    ];
    expect(selectExecutionRecoveryEntry(rows, (row) => row.card)).toEqual({
      index: 2,
      row: rows[2],
      card: { operationId: "op-new" },
      operationId: "op-new"
    });
  });

  it("一条都没有卡时返回 null —— 宿主必须什么都不画，而不是画个空壳", () => {
    const rows: Row[] = [
      { id: "m1", card: null },
      { id: "m2", card: null }
    ];
    expect(selectExecutionRecoveryEntry(rows, (row) => row.card)).toBeNull();
    expect(selectExecutionRecoveryEntry([], (row: Row) => row.card)).toBeNull();
  });

  it("operationId 为空白的卡跳过：绑不上执行就不给按钮", () => {
    const rows: Row[] = [
      { id: "m1", card: { operationId: "op-real" } },
      { id: "m2", card: { operationId: "  " } }
    ];
    expect(selectExecutionRecoveryEntry(rows, (row) => row.card)?.operationId).toBe("op-real");
    const onlyBlank: Row[] = [{ id: "m1", card: { operationId: "" } }];
    expect(selectExecutionRecoveryEntry(onlyBlank, (row) => row.card)).toBeNull();
  });

  it("只挑一条：同一个 operationId 出现两次也不会变成两个入口", () => {
    const rows: Row[] = [
      { id: "m1", card: { operationId: "op-1" } },
      { id: "m2", card: { operationId: "op-1" } }
    ];
    const entry = selectExecutionRecoveryEntry(rows, (row) => row.card);
    expect(entry?.index).toBe(1);
    expect(rows.filter((row) => row.card === entry?.card)).toHaveLength(1);
  });

  it("进度一句话说清第几步与还剩几步；历史记录的指路句不许变空", () => {
    expect(executionRecoveryProgress(
      { kind: "reconcile", actionIndex: 1, outcome: "pending", remaining: 1 },
      { nextActionIndex: 1, totalActions: 3 }
    )).toBe("已完成 1 / 3 步 · 在途第 2 步");
    expect(executionRecoveryProgress(
      { kind: "resume", remaining: 1 },
      { nextActionIndex: 2, totalActions: 3 }
    )).toBe("已完成 2 / 3 步 · 剩余 1 步");
    expect(executionRecoveryProgress(
      { kind: "completed", recorded: true },
      { nextActionIndex: 3, totalActions: 3 }
    )).toBe("已完成 3 / 3 步");
    expect(EXECUTION_RECOVERY_ENTRY_HINT).toContain("输入框上方");
  });
});

describe("终态：只有撤销 / 放弃 / 收尾才退出固定区域", () => {
  it("三种终态各自命名，摘要说的是同一件事的不同版本", () => {
    expect(executionRecoveryTerminalReason({ kind: "undone" })).toBe("undone");
    expect(executionRecoveryTerminalReason({ kind: "blocked", detail: "x", abandoned: true }))
      .toBe("abandoned");
    expect(executionRecoveryTerminalReason({ kind: "completed", recorded: true })).toBe("settled");
    expect(executionRecoveryTerminalSummary({ kind: "undone" })).toContain("已撤销");
    expect(executionRecoveryTerminalSummary({ kind: "blocked", detail: "x", abandoned: true }))
      .toContain("放弃变更");
    expect(executionRecoveryTerminalSummary({ kind: "completed", recorded: true }))
      .toContain("已记录完成");
    for (const verdict of [
      { kind: "undone" } as const,
      { kind: "blocked", detail: "x", abandoned: true } as const,
      { kind: "completed", recorded: true } as const
    ]) {
      expect(executionRecoveryTerminalSummary(verdict)).toContain("展开");
    }
  });

  it("未知、阻断、证据不足、还要动手的结论**一律不是**终态", () => {
    const notTerminal: ExecutionRecoveryVerdict[] = [
      /* 待核对：现场还没查清 —— 隐藏它就是隐藏待办。 */
      { kind: "reconcile", actionIndex: 1, outcome: "pending", remaining: 1 },
      { kind: "reconcile", actionIndex: 1, outcome: "applied", remaining: 1 },
      { kind: "resume", remaining: 1, source: "matched" },
      { kind: "replan", remaining: 2 },
      /* 普通 blocked：现场冲突 / 执行权被占 / 来源变化 —— 与"放弃过"不是一回事。 */
      { kind: "blocked", detail: "现场与登记状态不一致" },
      { kind: "blocked", detail: "目标已变化", sceneChanged: true },
      { kind: "insufficient", detail: "读不到现场" },
      /* 缺回执：还有「补记完成」可点，不算收尾。 */
      { kind: "completed" }
    ];
    for (const verdict of notTerminal) {
      expect(executionRecoveryTerminalReason(verdict), JSON.stringify(verdict)).toBeNull();
      expect(executionRecoveryIsTerminal(verdict)).toBe(false);
      expect(executionRecoveryTerminalSummary(verdict)).toBeNull();
    }
  });

  it("终态不给写入按钮，也不因退出固定区域而丢掉结论正文", () => {
    const abandoned: ExecutionRecoveryVerdict = {
      kind: "blocked",
      detail: "这次执行已按「放弃变更」处置过，原表已按恢复点还原。",
      abandoned: true
    };
    expect(executionRecoveryActions(abandoned)).toEqual(["dismiss"]);
    expect(executionRecoveryDetail(abandoned)).toContain("放弃变更");
    expect(executionRecoveryActions({ kind: "undone" })).toEqual([]);
  });
});
