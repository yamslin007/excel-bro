import { describe, expect, it } from "vitest";
import {
  advanceWpsActivity,
  executionActivityUpdate,
  finishWpsActivity,
  readonlyActivityUpdate,
  startWpsActivity
} from "./wpsActivityLog";
import type { WpsExecutionProgress } from "./wpsPlanExecutor";

describe("WPS shared activity timeline", () => {
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
