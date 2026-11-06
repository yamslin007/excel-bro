import { describe, expect, it } from "vitest";
import {
  RECOVERY_COMPOSER_ENTRY_LABEL,
  recoveryComposerAction,
  recoveryComposerEntry,
  recoveryComposerHeadline,
  recoveryComposerLines,
  recoveryComposerPrimaryLabel,
  recoveryComposerPrimaryTitle,
  recoveryComposerState,
  recoveryComposerTargetLabel,
  recoveryComposerTargets,
  recoveryComposerView,
  type RecoveryComposerFacts
} from "./recoveryComposer";
import { executionRecoveryDetail, type ExecutionRecoveryVerdict } from "./executionRecoveryView";

/** 这次计划的三条动作（真机 B 的形状：三张结果表各写一格公式）。 */
const ACTIONS = [
  { sheet: "结果1", range: "B2", type: "writeFormulas" },
  { sheet: "结果2", range: "B2", type: "writeFormulas" },
  { sheet: "结果3", range: "B2", type: "writeFormulas" }
];

/**
 * 一组基准事实：三条动作、账本已确认 1 步、第 2 步在途（还没核对）、下一条要写 `结果3!B2`。
 *
 * 与 2026-11-01/02 WPS 公式 B 真机同一形状（前两步写公式、第三步没跑），
 * 所以下面每条用例的断言都能对回那次真实现场。
 */
function facts(overrides: Partial<RecoveryComposerFacts> = {}): RecoveryComposerFacts {
  return {
    verdict: { kind: "reconcile", actionIndex: 1, outcome: "pending", remaining: 1 },
    nextActionIndex: 1,
    totalActions: 3,
    nextTarget: { sheet: "结果3", range: "B2" },
    actions: ACTIONS,
    ...overrides
  };
}

describe("恢复操作区：该不该占用输入框", () => {
  it("待核对：进入操作区，只有一颗主操作，且它说的是「检查」不是「继续」", () => {
    const state = recoveryComposerState(facts());
    expect(state.mode).toBe("recovery");
    if (state.mode !== "recovery") throw new Error("应当进入恢复操作区");
    expect(state.view.action).toBe("reconcile");
    expect(state.view.primaryLabel).toBe("检查写入结果");
    // 主操作这一刻只承接核对：正文里提到"再点一次才会继续"是**承诺下一步**，不是现在就给入口。
    expect(state.view.primaryLabel).not.toContain("继续");
    expect(state.view.lines.at(-1)).toContain("需要你再点一次");
  });

  it("核对完成：同一位置变成「继续剩余 N 步」，仍然只有一颗主操作", () => {
    const state = recoveryComposerState(facts({
      verdict: { kind: "resume", remaining: 1 },
      nextActionIndex: 2
    }));
    expect(state.mode).toBe("recovery");
    if (state.mode !== "recovery") throw new Error("应当进入恢复操作区");
    expect(state.view.action).toBe("resume");
    expect(state.view.primaryLabel).toBe("继续剩余 1 步");
    expect(state.view.headline).toBe("检查完成，可以继续");
  });

  it("终态（已撤销）：退回普通输入框，也不留待办入口", () => {
    for (const verdict of [
      { kind: "undone" },
      { kind: "blocked", detail: "已按「放弃变更」处置", abandoned: true },
      { kind: "completed", recorded: true }
    ] satisfies ExecutionRecoveryVerdict[]) {
      expect(recoveryComposerState(facts({ verdict }))).toEqual({ mode: "composer", entry: null });
    }
  });

  it("没有单一主操作的结论：不进入操作区（不编一颗点了没用的按钮）", () => {
    for (const verdict of [
      { kind: "blocked", detail: "来源数据已改变" },
      { kind: "insufficient", detail: "读不到记录" },
      { kind: "replan", remaining: 2 },
      { kind: "completed" }
    ] satisfies ExecutionRecoveryVerdict[]) {
      expect(recoveryComposerAction(verdict)).toBeNull();
      expect(recoveryComposerState(facts({ verdict }))).toEqual({ mode: "composer", entry: null });
    }
  });

  it("多条候选还没选定：不进入、也不给入口（不替用户挑一条）", () => {
    expect(recoveryComposerState(facts({ ambiguous: true }))).toEqual({ mode: "composer", entry: null });
  });

  it("用户按了「另提需求」：退成待办入口，任务与材料都还在", () => {
    const state = recoveryComposerState(facts({ choice: "composer" }));
    expect(state.mode).toBe("composer");
    if (state.mode !== "composer") throw new Error("应当退出操作区");
    expect(state.entry?.action).toBe("reconcile");
    expect(state.entry?.label).toContain(RECOVERY_COMPOSER_ENTRY_LABEL);
    expect(state.entry?.label).toContain("待核对");
    // 入口只说"可以切回去"，不承诺撤销/放弃。
    expect(state.entry?.title).toContain("回到恢复处理");
  });

  it("输入框里已有草稿：不抢占，只给入口", () => {
    const state = recoveryComposerState(facts({ hasDraft: true }));
    expect(state.mode).toBe("composer");
    if (state.mode !== "composer") throw new Error("不应抢占输入框");
    expect(state.entry).not.toBeNull();
  });

  it("输入框归别的正事（在跑 / 等审阅 / 在答题）：不进入，连入口都不给", () => {
    const state = recoveryComposerState(facts({ composerBusyElsewhere: true }));
    expect(state).toEqual({ mode: "composer", entry: null });
  });

  it("从待办入口点回来：即使有草稿也进入操作区（草稿留着，切回去还在）", () => {
    const state = recoveryComposerState(facts({ hasDraft: true, choice: "recovery" }));
    expect(state.mode).toBe("recovery");
  });

  it("显式选择压不过「输入框归别的正事」：不造一个点了没反应的入口", () => {
    for (const overrides of [
      { choice: "recovery" as const, composerBusyElsewhere: true },
      { choice: "recovery" as const, ambiguous: true }
    ]) {
      expect(recoveryComposerState(facts(overrides))).toEqual({ mode: "composer", entry: null });
    }
  });

  it("入口上的步数来自真实判据（核对完就变成「还剩 N 步」）", () => {
    const reconcile = recoveryComposerEntry(facts({ choice: "composer" }));
    expect(reconcile?.label).toContain("待核对");
    const resume = recoveryComposerEntry(facts({
      choice: "composer",
      verdict: { kind: "resume", remaining: 1 },
      nextActionIndex: 2
    }));
    expect(resume?.label).toContain("还剩 1 步");
  });
});

describe("恢复操作区：写了哪些目标（只能来自证据）", () => {
  it("账本游标之前算写入；在途那一步 pending 时两边都不算", () => {
    const split = recoveryComposerTargets(facts());
    expect(split.written.map(recoveryComposerTargetLabel)).toEqual(["结果1!B2"]);
    expect(split.unwritten.map(recoveryComposerTargetLabel)).toEqual(["结果3!B2"]);
    expect(split.inFlight?.outcome).toBe("pending");
    expect(recoveryComposerTargetLabel(split.inFlight?.target)).toBe("结果2!B2");
  });

  it("结论说 applied 才把在途那一步算进「已写入」", () => {
    const split = recoveryComposerTargets(facts({
      verdict: { kind: "reconcile", actionIndex: 1, outcome: "applied", remaining: 1 }
    }));
    expect(split.written.map(recoveryComposerTargetLabel)).toEqual(["结果1!B2", "结果2!B2"]);
    expect(split.unwritten.map(recoveryComposerTargetLabel)).toEqual(["结果3!B2"]);
  });

  it("结论说 not_applied：在途那一步落到「尚未写入」，不进「已写入」", () => {
    const split = recoveryComposerTargets(facts({
      verdict: { kind: "reconcile", actionIndex: 1, outcome: "not_applied", remaining: 1 }
    }));
    expect(split.written.map(recoveryComposerTargetLabel)).toEqual(["结果1!B2"]);
    expect(split.unwritten.map(recoveryComposerTargetLabel)).toEqual(["结果2!B2", "结果3!B2"]);
  });

  it("续跑态：游标之前是写入、之后是没写", () => {
    const split = recoveryComposerTargets(facts({
      verdict: { kind: "resume", remaining: 1 },
      nextActionIndex: 2
    }));
    expect(split.written.map(recoveryComposerTargetLabel)).toEqual(["结果1!B2", "结果2!B2"]);
    expect(split.unwritten.map(recoveryComposerTargetLabel)).toEqual(["结果3!B2"]);
    expect(split.inFlight).toBeNull();
  });

  it("宿主没给动作清单：一个位置都不编，退回只报步数", () => {
    const split = recoveryComposerTargets(facts({ actions: null }));
    expect(split).toEqual({ written: [], unwritten: [], inFlight: null });
    const lines = recoveryComposerLines(facts({ actions: null }));
    expect(lines[0]).toContain("第 2 步写没写完还没确认");
    expect(lines.join("\n")).not.toContain("结果1!B2");
    expect(lines.join("\n")).not.toContain("结果3");
  });

  it("位置只给到工作表也不补区域；超过 4 个折成「等 N 处」", () => {
    expect(recoveryComposerTargetLabel({ sheet: "结果3", range: "   " })).toBe("结果3");
    expect(recoveryComposerTargetLabel({ sheet: "   " })).toBeNull();
    const many = ["A", "B", "C", "D", "E", "F"].map((sheet) => ({ sheet, range: "A1" }));
    const lines = recoveryComposerLines(facts({
      verdict: { kind: "reconcile", actionIndex: 6, outcome: "pending", remaining: 0 },
      nextActionIndex: 6,
      totalActions: 7,
      actions: [...many, { sheet: "G", range: "A1" }]
    }));
    expect(lines.join("\n")).toContain("已写入：A!A1、B!A1、C!A1、D!A1 等 6 处");
  });
});

describe("恢复操作区：写什么", () => {
  it("第一行只说证据支持的事：三种 outcome 三种说法", () => {
    expect(recoveryComposerHeadline(facts({
      verdict: { kind: "reconcile", actionIndex: 1, outcome: "applied", remaining: 1 }
    }))).toBe("任务中断 · 还剩 1 步");
    expect(recoveryComposerHeadline(facts({
      verdict: { kind: "reconcile", actionIndex: 1, outcome: "not_applied", remaining: 1 }
    }))).toBe("任务中断 · 第 2 步没有写入");
    expect(recoveryComposerHeadline(facts({
      verdict: { kind: "reconcile", actionIndex: 1, outcome: "pending", remaining: 1 }
    }))).toBe("任务中断 · 第 2 步结果待核对");
  });

  it("状态未知时正文绝不宣称已写入（也不许把这一步算进「尚未写入」）", () => {
    const pending = recoveryComposerLines(facts());
    const text = pending.join("\n");
    expect(text).toContain("第 2 步（结果2!B2）是否已写入还没确认。");
    expect(text).not.toContain("已写入：结果2");
    expect(text).not.toContain("尚未写入：结果2");
    // 证据说写了才说"已写入"。
    const applied = recoveryComposerLines(facts({
      verdict: { kind: "reconcile", actionIndex: 1, outcome: "applied", remaining: 1 }
    }));
    expect(applied[0]).toBe("已写入：结果1!B2、结果2!B2");
  });

  it("已写入 / 尚未写入 两句都对着证据，不是对着步数算出来的", () => {
    const applied = recoveryComposerLines(facts({
      verdict: { kind: "reconcile", actionIndex: 1, outcome: "applied", remaining: 1 }
    }));
    expect(applied).toEqual([
      "已写入：结果1!B2、结果2!B2",
      "尚未写入：结果3!B2",
      "先检查已写入结果，不会修改表格。",
      "检查完成后还需要你再点一次，才会执行剩下的步骤。"
    ]);
    const nothingWritten = recoveryComposerLines(facts({
      verdict: { kind: "reconcile", actionIndex: 0, outcome: "pending", remaining: 2 },
      nextActionIndex: 0
    }));
    expect(nothingWritten[0]).toBe("第 1 步（结果1!B2）是否已写入还没确认。");
    expect(nothingWritten.join("\n")).toContain("尚未写入：结果2!B2、结果3!B2");
    // 一步都没写进去时不说"先检查已写入结果"。
    expect(nothingWritten.join("\n")).toContain("先检查当前结果，不会修改表格。");
  });

  it("续跑：下一步说清写哪儿、写什么；已完成步数不重复执行", () => {
    const lines = recoveryComposerLines(facts({
      verdict: { kind: "resume", remaining: 1 },
      nextActionIndex: 2
    }));
    expect(lines[0]).toBe("下一步：向结果3!B2写入公式。");
    expect(lines[1]).toBe("前 2 步不会重复执行。");
  });

  it("动作类型决定那句动词；认不出类型也不编内容", () => {
    const label = (type: string | null | undefined) =>
      recoveryComposerLines(facts({
        verdict: { kind: "resume", remaining: 1 },
        nextActionIndex: 0,
        actions: [{ sheet: "结果1", range: "B2", type }]
      }))[0];
    expect(label("writeValues")).toBe("下一步：向结果1!B2写入数值。");
    expect(label("clearRange")).toBe("下一步：清空结果1!B2。");
    expect(label("setFill")).toBe("下一步：调整结果1!B2的格式。");
    expect(label("somethingNew")).toBe("下一步：向结果1!B2写入。");
    expect(label(null)).toBe("下一步：向结果1!B2写入。");
  });

  it("续跑拿不到目标就只说步数，不编位置", () => {
    const lines = recoveryComposerLines(facts({
      verdict: { kind: "resume", remaining: 1 },
      nextActionIndex: 2,
      actions: null,
      nextTarget: null
    }));
    expect(lines[0]).toBe("继续后将执行剩余 1 步。");
    expect(lines.join("\n")).not.toContain("结果3");
  });

  it("正文里不出现术语：在途动作 / 现场 / 回执 / 补记", () => {
    const corpus = [
      ...recoveryComposerLines(facts()),
      ...recoveryComposerLines(facts({ verdict: { kind: "resume", remaining: 1 }, nextActionIndex: 2 })),
      recoveryComposerHeadline(facts()),
      recoveryComposerPrimaryLabel(facts()),
      recoveryComposerPrimaryLabel(facts({ verdict: { kind: "resume", remaining: 1 }, nextActionIndex: 2 })),
      recoveryComposerPrimaryTitle("reconcile", facts()),
      recoveryComposerPrimaryTitle("resume", facts({ verdict: { kind: "resume", remaining: 1 }, nextActionIndex: 2 }))
    ].join("\n");
    for (const jargon of ["在途动作", "现场", "回执", "补记"]) {
      expect(corpus).not.toContain(jargon);
    }
  });

  it("说清点下去会发生什么：核对不写、继续只写剩下的、都不会自动继续", () => {
    const reconcile = recoveryComposerView(facts());
    expect(reconcile.lines.at(-1)).toContain("需要你再点一次");
    expect(reconcile.primaryTitle).toContain("不修改任何单元格");
    expect(reconcile.primaryTitle).toContain("不会自动继续");
    const resume = recoveryComposerView(facts({
      verdict: { kind: "resume", remaining: 1 },
      nextActionIndex: 2
    }));
    expect(resume.primaryTitle).toContain("只执行剩下的 1 步");
    expect(resume.primaryTitle).toContain("已完成的");
  });

  it("主操作文案按证据分档：不说做不到的事", () => {
    const label = (verdict: ExecutionRecoveryVerdict, nextActionIndex = 1) =>
      recoveryComposerPrimaryLabel(facts({ verdict, nextActionIndex }));
    expect(label({ kind: "reconcile", actionIndex: 1, outcome: "applied", remaining: 1 }))
      .toBe("检查已写入结果");
    expect(label({ kind: "reconcile", actionIndex: 1, outcome: "pending", remaining: 1 }))
      .toBe("检查写入结果");
    expect(label({ kind: "reconcile", actionIndex: 1, outcome: "not_applied", remaining: 1 }))
      .toBe("检查当前结果");
    expect(label({ kind: "resume", remaining: 1 }, 2)).toBe("继续剩余 1 步");
    // 拿不到核对/继续的结论时给空串，而不是编一个按钮名。
    expect(label({ kind: "insufficient", detail: "读不到" })).toBe("");
  });

  it("详情里的三句技术口径原样来自共享层（不另写一套）", () => {
    const view = recoveryComposerView(facts());
    expect(view.detail).toBe(executionRecoveryDetail(facts().verdict));
    expect(view.progress).toBe("已完成 1 / 3 步 · 在途第 2 步");
    expect(view.note.length).toBeGreaterThan(0);
  });
});
