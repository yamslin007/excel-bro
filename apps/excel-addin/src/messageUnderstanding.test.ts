import { afterEach, describe, expect, it, vi } from "vitest";
import { understandMessage } from "./api";
import { assertMessageUnderstandingResponse } from "./contracts";
import { captureWorkbookCatalog } from "./excel";

afterEach(() => vi.unstubAllGlobals());

describe("message-only protocol", () => {
  it("projects real request bytes, excluding workbook, schema, results and image binaries", async () => {
    const fetch = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => new Response(JSON.stringify({ kind: "conversation", message: "不客气！", provider: "model" })));
    vi.stubGlobal("fetch", fetch);
    const controller = new AbortController();
    await understandMessage({
      prompt: "谢谢帮助", modelId: "model-1", imageCount: 0,
      conversation: [{ role: "user", text: "你好", private: "PRIVATE_SENTINEL" }],
      scope: { headers: ["PRIVATE_SENTINEL"] }, priorResult: { rows: [["PRIVATE_SENTINEL"]] },
      workbook: { name: "PRIVATE_SENTINEL" }, images: [{ base64: "PRIVATE_SENTINEL" }]
    } as never, controller.signal);
    expect(fetch.mock.calls[0][0]).toContain("/api/intent/message");
    const init = fetch.mock.calls[0][1] as RequestInit;
    expect(init.signal).toBe(controller.signal);
    expect(JSON.parse(String(init.body))).toEqual({
      prompt: "谢谢帮助", modelId: "model-1", imageCount: 0, conversation: [{ role: "user", text: "你好" }]
    });
    expect(init.body).not.toContain("PRIVATE_SENTINEL");
  });

  it.each([
    { kind: "proceed", confirmedPrompt: "读取整表", provider: "model" },
    { kind: "needs_workbook_context", summary: "", provider: "model" },
    { kind: "conversation", message: "你好", provider: "model", actions: [] },
    { kind: "task_recovery", turnId: "original", action: "execute", provider: "model" },
    { kind: "conversation", message: "你好", action: "continue", provider: "model" },
    { kind: "conversation", message: "你好", provider: "local" }
  ])("rejects invalid route without retrying another endpoint: %j", async (value) => {
    const fetch = vi.fn(async () => new Response(JSON.stringify(value)));
    vi.stubGlobal("fetch", fetch);
    await expect(understandMessage({ prompt: "你好啊" })).rejects.toThrow("消息理解响应格式无效");
    expect(fetch).toHaveBeenCalledOnce();
  });

  it("accepts explicit conversation, workbook and recovery routes", () => {
    expect(() => assertMessageUnderstandingResponse({ kind: "conversation", message: "不客气", provider: "model" })).not.toThrow();
    expect(() => assertMessageUnderstandingResponse({ kind: "needs_workbook_context", summary: "查询当前表", provider: "model" })).not.toThrow();
    expect(() => assertMessageUnderstandingResponse({ kind: "task_recovery", turnId: "original", provider: "model" })).not.toThrow();
    expect(() => assertMessageUnderstandingResponse({ kind: "task_recovery", turnId: "original", action: "inspect", provider: "model" })).not.toThrow();
    expect(() => assertMessageUnderstandingResponse({ kind: "task_recovery", turnId: "original", action: "continue", provider: "model" })).not.toThrow();
  });

  it("sends pending task IDs through the actual HTTP projection without workbook data", async () => {
    const fetch = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) =>
      new Response(JSON.stringify({ kind: "task_recovery", turnId: "original", provider: "model" })));
    vi.stubGlobal("fetch", fetch);
    const response = await understandMessage({ prompt: "继续", recoveryTasks: ["original"] });
    expect(response.kind).toBe("task_recovery");
    expect(JSON.parse(String(fetch.mock.calls[0][1]?.body)).recoveryTasks).toEqual(["original"]);
  });

  it("rejects unoffered recovery IDs even when the server response has a valid shape", async () => {
    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify({
      kind: "task_recovery", turnId: "other-conversation", provider: "model"
    }))));
    await expect(understandMessage({ prompt: "继续", recoveryTasks: ["original"] })).rejects.toThrow("未提供的恢复任务");
  });

  it("startup catalog never touches ranges or values", async () => {
    const touches = vi.fn(() => { throw new Error("Cells must not be read"); });
    const sheet = { name: "Sheet1", load: vi.fn(), getUsedRangeOrNullObject: touches, getRangeByIndexes: touches };
    Object.defineProperty(sheet, "values", { get: touches });
    vi.stubGlobal("Office", { context: { document: { url: "file:///C:/Data/book.xlsx" } } });
    vi.stubGlobal("Excel", { run: async (callback: (context: unknown) => unknown) => callback({
      workbook: { worksheets: { load: vi.fn(), items: [sheet], getActiveWorksheet: () => sheet } }, sync: async () => {}
    }) });
    const snapshot = await captureWorkbookCatalog();
    expect(snapshot.catalogOnly).toBe(true);
    expect(snapshot.worksheets[0].headers).toEqual([]);
    expect(snapshot.worksheets[0].dataRows).toEqual([]);
    expect(snapshot.worksheets[0].usedRange).toBeNull();
    expect(touches).not.toHaveBeenCalled();
  });
});
