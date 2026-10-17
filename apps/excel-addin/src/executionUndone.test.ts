// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import type { ChatMessage } from "./types/chat";
import {
  markFolderExecutionUndone,
  markWorkbookExecutionUndone
} from "./executionUndone";

/**
 * 「哪条消息该标成已撤销」的归属规则。
 *
 * 判据重点是**负向**：串卡 / 误标 / 跨模式误标都必须被挡住 —— 因为"把没撤销的执行
 * 标成已撤销"比"没标"更危险（用户会以为文件已经还原）。
 *
 * 身份匹配缺口（本轮修）：`planId` 回退**只在候选唯一时**才自动标；多条缺 ID 旧卡
 * 共用同一 planId ⇒ **一条都不标**，并返回 `ambiguous: true` 提示"无法确认恢复归属"。
 */

function message(overrides: Partial<ChatMessage> = {}): ChatMessage {
  return {
    id: overrides.id ?? "m1",
    role: "system",
    text: "",
    createdAt: "2026-10-14T00:00:00.000Z",
    ...overrides
  };
}

const undoneFlags = (messages: readonly ChatMessage[]) =>
  messages.map((item) => item.executionUndone === true);

describe("markWorkbookExecutionUndone · 精确匹配", () => {
  it("按 executionId 精确归属：同一 planId 多次执行只标真正恢复的那一次", () => {
    const first = message({ id: "a", executedPlanId: "same-plan", executedExecutionId: "exec-1" });
    const second = message({ id: "b", executedPlanId: "same-plan", executedExecutionId: "exec-2" });

    const result = markWorkbookExecutionUndone([first, second], {
      planId: "same-plan",
      executionId: "exec-2"
    });

    expect(undoneFlags(result.messages)).toEqual([false, true]);
    expect(result.ambiguous).toBe(false);
  });

  it("executionId 精确匹配不到任何卡：不标，且报告无法确认归属", () => {
    const run = message({ id: "a", executedPlanId: "p1", executedExecutionId: "exec-9" });

    const result = markWorkbookExecutionUndone([run], {
      planId: "p1",
      executionId: "exec-1"
    });

    expect(undoneFlags(result.messages)).toEqual([false]);
    expect(result.ambiguous).toBe(true);
  });

  it("现代快照（有 executionId）与缺 ID 旧卡混合：精确命中，旧卡不动", () => {
    const modern = message({ id: "modern", executedPlanId: "p1", executedExecutionId: "exec-1" });
    const legacy = message({ id: "legacy", executedPlanId: "p1" });

    const result = markWorkbookExecutionUndone([legacy, modern], {
      planId: "p1",
      executionId: "exec-1"
    });

    expect(undoneFlags(result.messages)).toEqual([false, true]);
    expect(result.ambiguous).toBe(false);
  });
});

describe("markWorkbookExecutionUndone · planId 回退（缺 executionId 的旧快照）", () => {
  it("候选唯一时才标：单张缺 ID 旧卡 + planId 匹配 ⇒ 标它", () => {
    const legacy = message({ id: "legacy", executedPlanId: "p1" });

    const result = markWorkbookExecutionUndone([legacy], {
      planId: "p1",
      executionId: undefined
    });

    expect(undoneFlags(result.messages)).toEqual([true]);
    expect(result.ambiguous).toBe(false);
  });

  it("多张缺 ID 旧卡同 planId：一条都不标，报告无法确认归属", () => {
    const a = message({ id: "a", executedPlanId: "same-plan" });
    const b = message({ id: "b", executedPlanId: "same-plan" });

    const result = markWorkbookExecutionUndone([a, b], {
      planId: "same-plan",
      executionId: undefined
    });

    expect(undoneFlags(result.messages)).toEqual([false, false]);
    expect(result.ambiguous).toBe(true);
  });

  it("planId 回退不误标文件夹模式的消息（有 executedUndoId）", () => {
    const legacy = message({ id: "legacy", executedPlanId: "plan-1" });
    const folderRun = message({
      id: "folder",
      executedPlanId: "plan-1",
      executedUndoId: "undo-1"
    });

    const result = markWorkbookExecutionUndone([legacy, folderRun], {
      planId: "plan-1",
      executionId: undefined
    });

    expect(undoneFlags(result.messages)).toEqual([true, false]);
    expect(result.ambiguous).toBe(false);
  });

  it("planId 不匹配：候选为 0 ⇒ 不标且报告无法确认归属", () => {
    const run = message({ id: "a", executedPlanId: "other-plan" });

    const result = markWorkbookExecutionUndone([run], {
      planId: "plan-1",
      executionId: undefined
    });

    expect(undoneFlags(result.messages)).toEqual([false]);
    expect(result.ambiguous).toBe(true);
  });
});

describe("markFolderExecutionUndone", () => {
  it("按 undoId 精确归属：同一 planId 多次文件夹执行只标被恢复的那些", () => {
    const first = message({ id: "a", executedPlanId: "same-plan", executedUndoId: "undo-1" });
    const second = message({ id: "b", executedPlanId: "same-plan", executedUndoId: "undo-2" });

    const result = markFolderExecutionUndone([first, second], [
      { planId: "same-plan", undoId: "undo-2" }
    ]);

    expect(undoneFlags(result.messages)).toEqual([false, true]);
    expect(result.ambiguous).toBe(false);
  });

  it("一次请求连带恢复多条：逐条标回各自消息，不漏标也不多标", () => {
    const first = message({ id: "a", executedPlanId: "p1", executedUndoId: "undo-1" });
    const second = message({ id: "b", executedPlanId: "p2", executedUndoId: "undo-2" });
    const untouched = message({ id: "c", executedPlanId: "p3", executedUndoId: "undo-3" });

    const result = markFolderExecutionUndone([first, second, untouched], [
      { planId: "p1", undoId: "undo-1" },
      { planId: "p2", undoId: "undo-2" }
    ]);

    expect(undoneFlags(result.messages)).toEqual([true, true, false]);
    expect(result.ambiguous).toBe(false);
  });

  it("多张缺 undoId 旧卡同 planId：一条都不标，报告无法确认归属", () => {
    const a = message({ id: "a", executedPlanId: "same-plan" });
    const b = message({ id: "b", executedPlanId: "same-plan" });

    const result = markFolderExecutionUndone([a, b], [
      { planId: "same-plan", undoId: undefined }
    ]);

    expect(undoneFlags(result.messages)).toEqual([false, false]);
    expect(result.ambiguous).toBe(true);
  });

  it("planId 回退不误标工作簿模式的消息（有 executedExecutionId）", () => {
    const legacy = message({ id: "legacy", executedPlanId: "plan-1" });
    const workbookRun = message({
      id: "workbook",
      executedPlanId: "plan-1",
      executedExecutionId: "exec-1"
    });

    const result = markFolderExecutionUndone([legacy, workbookRun], [
      { planId: "plan-1", undoId: undefined }
    ]);

    expect(undoneFlags(result.messages)).toEqual([true, false]);
    expect(result.ambiguous).toBe(false);
  });

  it("空条目列表：一条都不标（没有真正恢复成功就不标已撤销）", () => {
    const run = message({ id: "a", executedPlanId: "p1", executedUndoId: "undo-1" });

    const result = markFolderExecutionUndone([run], []);

    expect(undoneFlags(result.messages)).toEqual([false]);
    expect(result.ambiguous).toBe(false);
  });

  it("undoId 精确匹配不到任何卡：不标，且报告无法确认归属", () => {
    const run = message({ id: "a", executedPlanId: "p1", executedUndoId: "undo-9" });

    const result = markFolderExecutionUndone([run], [
      { planId: "p1", undoId: "undo-1" }
    ]);

    expect(undoneFlags(result.messages)).toEqual([false]);
    expect(result.ambiguous).toBe(true);
  });
});
