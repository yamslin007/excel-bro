import { describe, expect, it } from "vitest";
import {
  advanceWpsActivity,
  executionActivityUpdate,
  finishWpsActivity,
  readonlyActivityUpdate,
  resumeWpsActivity,
  startWpsActivity,
  updateWpsActivityMessage
} from "./wpsActivityLog";
import { wpsProgressFromTaskEvent } from "./wpsReadonlyAgent";
import type { TaskProgressEvent, TaskProgressStatus } from "../../excel-addin/src/contracts";
import type { WpsExecutionProgress } from "./wpsPlanExecutor";

describe("WPS shared activity timeline", () => {
  it("ignores replayed events and events belonging to another task", () => {
    const first = startWpsActivity({ id: "step-7", label: "步骤7" }, 100);
    const bound = advanceWpsActivity(first, { id: "step-7", label: "步骤7更新", operationId: "original", sequence: 7 }, 100, 200);
    expect(advanceWpsActivity(bound, { id: "duplicate", label: "重复", operationId: "original", sequence: 7 }, 100, 300)).toBe(bound);
    expect(advanceWpsActivity(bound, { id: "foreign", label: "其他任务", operationId: "other", sequence: 8 }, 100, 300)).toBe(bound);
  });
  it("updates one reading step instead of producing a card for every chunk", () => {
    const first = readonlyActivityUpdate({
      stage: "reading", stepId: "read-data:Sheet1", title: "读取 Sheet1",
      sheetName: "Sheet1", rowsRead: 10, totalRows: 20, message: "已读取 10 行"
    });
    const started = startWpsActivity(first, 100);
    const second = advanceWpsActivity(started, { ...first, current: 20 }, 100, 250);
    expect(second.steps).toHaveLength(1);
    expect(second.steps[0]).toMatchObject({ current: 20, total: 20, unit: "rows" });
    const completed = finishWpsActivity(second, "completed", 100, 300);
    expect(completed.outcome).toBe("completed");
    expect(completed.steps[0].status).toBe("completed");
    expect(completed.totalMs).toBe(200);
  });

  it("keeps action, target and backup/verification steps without cell values", () => {
    const backup = executionActivityUpdate({ stage: "snapshot", message: "保存恢复点" });
    const started = startWpsActivity(backup, 100);
    const write = executionActivityUpdate({
      stage: "writing", message: "写入中", actionIndex: 0,
      actionType: "writeValues", sheetName: "Sheet1", range: "A1:B2",
      completedRows: 1, totalRows: 2, actionCount: 2, cellCount: 4
    });
    const running = advanceWpsActivity(started, write, 100, 200);
    expect(running.steps[0].status).toBe("completed");
    expect(running.steps[1]).toMatchObject({
      actionType: "writeValues", target: { sheet: "Sheet1", range: "A1:B2" },
      current: 1, total: 2
    });
    expect(running.steps[1].facts).toEqual([
      { label: "进度", value: "1/2 步" },
      { label: "目标", value: "4 个单元格" }
    ]);
    const verified = advanceWpsActivity(running, executionActivityUpdate({
      stage: "verifying", message: "回读中", actionIndex: 0,
      actionType: "writeValues", sheetName: "Sheet1", range: "A1:B2"
    }), 100, 300);
    expect(verified.steps.map((step) => step.id)).toEqual(["snapshot", "writing:0", "verifying:0"]);
    expect(JSON.stringify(verified)).not.toContain("beforeValues");
  });

  it("never carries credentials, raw rows or the write matrix into the timeline", () => {
    // 过程时间线会随对话历史持久化。语言模型的凭据与源值都必须留在内存里：
    // 一旦进了时间线，就等于进了 localStorage 与用户可复制的消息正文。
    const hostile: WpsExecutionProgress & Record<string, unknown> = {
      stage: "writing",
      message: "写入中",
      actionIndex: 0,
      actionType: "writeValues",
      sheetName: "订单",
      range: "B2:B3",
      completedRows: 1,
      totalRows: 2,
      actionCount: 1,
      cellCount: 2,
      // 宿主回调若被误传这些字段，时间线不能照单收下。
      apiKey: "sk-live-should-never-persist",
      modelApiKey: "sk-live-should-never-persist",
      values: [["办公用品"], ["电子设备"]],
      beforeValues: [[""], ["x"]],
      dataRows: [["X001", "办公用品"]]
    };
    const log = advanceWpsActivity(
      startWpsActivity(executionActivityUpdate(hostile), 100),
      executionActivityUpdate(hostile),
      100,
      200
    );
    const serialized = JSON.stringify(log);
    expect(serialized).not.toContain("sk-live");
    expect(serialized).not.toContain("apiKey");
    expect(serialized).not.toContain("办公用品");
    expect(serialized).not.toContain("电子设备");
    expect(serialized).not.toContain("dataRows");
    expect(serialized).not.toContain("values");
    // 安全摘要仍要在：隐私断言不能把有用的过程信息一起删掉。
    expect(log.steps[0]).toMatchObject({
      actionType: "writeValues",
      target: { sheet: "订单", range: "B2:B3" }
    });
  });

  it("marks a never-finished step as interrupted rather than completed", () => {
    const started = startWpsActivity(
      executionActivityUpdate({ stage: "writing", message: "写入中", actionIndex: 0 }),
      100
    );
    // 刷新/断线：没有 outcome。调用方（restoreActivityLog）据此判定中断，
    // 这里钉住 finishWpsActivity 不会把它算成 completed。
    expect(started.outcome).toBeUndefined();
    const interrupted = finishWpsActivity(started, "interrupted", 100, 300);
    expect(interrupted.outcome).toBe("interrupted");
    expect(interrupted.steps[0].status).toBe("warning");
    expect(interrupted.steps[0].status).not.toBe("completed");
  });
});

/** 造一条真实形状的 `TaskProgressEvent`（与 `contracts.ts` 一致）。 */
function taskEvent(overrides: Partial<TaskProgressEvent> & {
  stepId: string;
  sequence: number;
  status: TaskProgressStatus;
}): TaskProgressEvent {
  return {
    schemaVersion: 1,
    operationId: "turn-1",
    executionId: null,
    phase: "planning",
    title: overrides.stepId,
    summary: null,
    detail: null,
    current: null,
    total: null,
    unit: null,
    actionIndex: null,
    actionCount: null,
    actionType: null,
    toolName: null,
    facts: null,
    target: null,
    trace: null,
    bridgeTool: null,
    pythonResult: null,
    timestamp: new Date(0).toISOString(),
    ...overrides
  };
}

describe("WPS 事件转换保住真实状态", () => {
  it("completed 事件走完整转换链后仍然是 completed", () => {
    // 证据 1：服务端说完成，界面就必须显示完成。
    // 这条链有三段（事件 → WpsReadonlyTurnProgress → WpsActivityUpdate → 时间线），
    // 以前第二段没有 status 字段、第三段强制写 progress，两头都对，中间把它丢了。
    const progress = wpsProgressFromTaskEvent(
      "turn-1",
      taskEvent({ stepId: "bridge.query_table:abc", sequence: 3, status: "completed", title: "本机数据查询已完成" })
    );
    expect(progress.status).toBe("completed");

    const update = readonlyActivityUpdate(progress);
    expect(update.status).toBe("completed");
    expect(update.id).toBe("bridge.query_table:abc");
    // 结构化字段整体透传：漏一个 bridgeTool 就够让续跑卡死一次。
    expect(progress.event).toBeTruthy();

    const started = startWpsActivity({ id: "answer", label: "生成回答或计划" }, 1_000);
    // 真实链路：started（1）→ progress（2）→ completed（3），同一 stepId。
    const begin = readonlyActivityUpdate(wpsProgressFromTaskEvent("turn-1",
      taskEvent({ stepId: "bridge.query_table:abc", sequence: 1, status: "started" })));
    const running = advanceWpsActivity(started, begin, 1_000, 1_200);
    expect(running.steps[1].status).toBe("started");
    const done = advanceWpsActivity(running, update, 1_000, 1_800);
    expect(done.steps.map((step) => step.id)).toEqual(["answer", "bridge.query_table:abc"]);
    expect(done.steps[1].status).toBe("completed");
    // 完成事件到达即冻结该步耗时：这一刻之后不再增长。
    // 起点是这一步**出现**的时刻（1200），不是本轮开始（1000）。
    expect(done.steps[1].elapsedMs).toBe(600);
  });

  it("同一工具连续两次调用是两个步骤，重放同一次调用不新增", () => {
    const first = readonlyActivityUpdate(wpsProgressFromTaskEvent("turn-1",
      taskEvent({ stepId: "bridge.query_table:aaa", sequence: 1, status: "started" })));
    const firstDone = readonlyActivityUpdate(wpsProgressFromTaskEvent("turn-1",
      taskEvent({ stepId: "bridge.query_table:aaa", sequence: 2, status: "completed" })));
    const second = readonlyActivityUpdate(wpsProgressFromTaskEvent("turn-1",
      taskEvent({ stepId: "bridge.query_table:bbb", sequence: 3, status: "started" })));

    let log = startWpsActivity(first, 1_000);
    log = advanceWpsActivity(log, firstDone, 1_000, 1_400);
    log = advanceWpsActivity(log, second, 1_000, 1_500);
    expect(log.steps.map((step) => step.id)).toEqual([
      "bridge.query_table:aaa",
      "bridge.query_table:bbb"
    ]);
    expect(log.steps[0].status).toBe("completed");

    // 同一次调用的重放（序号更新、内容一样）只更新不新增。
    const replay = readonlyActivityUpdate(wpsProgressFromTaskEvent("turn-1",
      taskEvent({ stepId: "bridge.query_table:bbb", sequence: 4, status: "progress" })));
    log = advanceWpsActivity(log, replay, 1_000, 1_600);
    expect(log.steps).toHaveLength(2);
  });

  it("迟到的'进行中'事件不会把已完成的步骤拉回进行中", () => {
    const late = readonlyActivityUpdate(wpsProgressFromTaskEvent("turn-1",
      taskEvent({ stepId: "bridge.read:a1", sequence: 99, status: "progress" })));
    // 先有一条本地起始步骤（真实链路的第一个事件就是本地阶段，没有服务端序号）。
    let log = startWpsActivity({ id: "answer", label: "生成回答或计划" }, 1_000);
    const begin = readonlyActivityUpdate(wpsProgressFromTaskEvent("turn-1",
      taskEvent({ stepId: "bridge.read:a1", sequence: 4, status: "started" })));
    log = advanceWpsActivity(log, begin, 1_000, 1_000);
    const done = readonlyActivityUpdate(wpsProgressFromTaskEvent("turn-1",
      taskEvent({ stepId: "bridge.read:a1", sequence: 5, status: "completed" })));
    log = advanceWpsActivity(log, done, 1_000, 1_500);
    expect(log.steps[1].status).toBe("completed");
    const afterLate = advanceWpsActivity(log, late, 1_000, 9_000);
    expect(afterLate.steps[1].status).toBe("completed");
    expect(afterLate.steps[1].elapsedMs).toBe(500);
  });
});

describe("WPS 收尾处理所有活动步骤", () => {
  it("当前步骤不在数组末位时也按身份收尾，未结束的其余步骤如实标注", () => {
    // 三条：a 已完成、b 卡在进行中、c 最后插入但其实早已开始。
    const early = readonlyActivityUpdate(wpsProgressFromTaskEvent("turn-1",
      taskEvent({ stepId: "b", sequence: 1, status: "started" })));
    const later = readonlyActivityUpdate(wpsProgressFromTaskEvent("turn-1",
      taskEvent({ stepId: "c", sequence: 2, status: "started" })));
    const done = readonlyActivityUpdate(wpsProgressFromTaskEvent("turn-1",
      taskEvent({ stepId: "a", sequence: 3, status: "completed" })));

    let log = startWpsActivity(early, 1_000);
    log = advanceWpsActivity(log, later, 1_000, 1_200);
    log = advanceWpsActivity(log, done, 1_000, 1_300);
    // c 是最近开始的未结束步骤 ⇒ 它拿活动终态；b 拿 warning。
    const finished = finishWpsActivity(log, "completed", 1_000, 2_000, "c");
    const status = new Map(finished.steps.map((step) => [step.id, step.status]));
    expect(status.get("c")).toBe("completed");
    expect(status.get("b")).toBe("warning");
    expect(status.get("a")).toBe("completed");
  });
});


it("本地步骤的完成重放保持耗时，暂停恢复保留独立等待桶", () => {
  let log = startWpsActivity({ id: "intent", label: "理解" }, 1000);
  log = advanceWpsActivity(log, { id: "intent", label: "理解", status: "completed" }, 1000, 1500);
  log = advanceWpsActivity(log, { id: "intent", label: "理解", status: "completed" }, 1000, 9000);
  expect(log.steps[0]).toMatchObject({ elapsedMs: 500, status: "completed" });
  log = advanceWpsActivity(log, { id: "next", label: "下一步", status: "started" }, 1000, 9000);
  log = finishWpsActivity(log, "awaiting_input", 1000, 10000);
  log = resumeWpsActivity(log, 20000);
  expect(log.pausedMs).toBe(10000);
  expect(log.steps[1]).toMatchObject({ id: "next", elapsedMs: 1000, status: "progress" });
  expect(log.timelineStartedAt).toBe(1000);
});

describe("WPS Python 代码工件", () => {
  const codeUpdate = (overrides: Record<string, unknown> = {}) => ({
    id: "python.analysis:py_0123456789ab",
    label: "正在校验 Python 代码",
    operationId: "turn-1",
    sequence: 1,
    status: "started" as const,
    pythonCode: {
      codeId: "code_py_0123456789ab",
      attemptId: "py_0123456789ab",
      stage: "validating" as const,
      version: 1,
      chars: 812,
      complete: true,
      ...overrides
    }
  });

  it("同一步骤的代码生命周期合并成一个工件，正文不随事件到达", () => {
    let log = startWpsActivity(codeUpdate(), 100);
    log = advanceWpsActivity(
      log,
      { ...codeUpdate(), sequence: 2, status: "progress", pythonCode: { ...codeUpdate().pythonCode, stage: "running", executedVersion: 1 } },
      100,
      300
    );
    expect(log.steps).toHaveLength(1);
    expect(log.steps[0].pythonCode?.stage).toBe("running");
    expect(log.steps[0].pythonCode?.code).toBe("");
    expect(log.steps[0].pythonCode?.executedVersion).toBe(1);
  });

  it("重放不会把同一段代码粘两遍", () => {
    let log = startWpsActivity(
      { ...codeUpdate(), pythonCode: { ...codeUpdate().pythonCode, stage: "generated", code: "x = 1" } },
      100
    );
    log = advanceWpsActivity(
      log,
      { ...codeUpdate(), sequence: 2, pythonCode: { ...codeUpdate().pythonCode, stage: "generated", code: "x = 1" } },
      100,
      200
    );
    expect(log.steps[0].pythonCode?.code).toBe("x = 1");
  });
});

describe("WPS 生产日志写回守卫", () => {
  it.each(["superseded", "closed"] as const)("%s 段拒绝迟到进度与迟到收尾", (status) => {
    const log = startWpsActivity({ id: "read", label: "暂停时的原文", operationId: "task-a" }, 100);
    const messages = [{ id: "old", replySegment: { id: "old", taskId: "task-a", status }, activityLog: log }];
    const snapshot = JSON.stringify(messages);
    let called = 0;
    const progressed = updateWpsActivityMessage(messages, "old", current => {
      called++;
      return advanceWpsActivity(current, { id: "late", label: "迟到进度" }, 100, 900);
    });
    const ended = updateWpsActivityMessage(progressed, "old", current => {
      called++;
      return finishWpsActivity(current, "completed", 100, 1000);
    });
    expect(called).toBe(0);
    expect(JSON.stringify(ended)).toBe(snapshot);
    expect(ended[0]).toBe(messages[0]);
  });
  it("活动段接收本任务进度，拒绝其他任务；无段的旧消息仍可写回", () => {
    const log = startWpsActivity({ id: "read", label: "读取", operationId: "task-a" }, 100);
    const messages = [{ id: "live", replySegment: { id: "live", taskId: "task-a", status: "live" as const }, activityLog: log }];
    const update = (current: typeof log) => finishWpsActivity(current, "completed", 100, 200);
    expect(updateWpsActivityMessage(messages, "live", update, "task-b")[0]).toBe(messages[0]);
    expect(updateWpsActivityMessage(messages, "live", update, "task-a")[0].activityLog.outcome).toBe("completed");
    expect(updateWpsActivityMessage([{ id: "legacy", activityLog: log }], "legacy", update)[0].activityLog.outcome).toBe("completed");
  });
});
