import { describe, expect, it, vi } from "vitest";
import type { WpsApplication, WpsRange, WpsWorksheet } from "./wpsHost";
import { buildWpsReadonlyContexts, runWpsReadonlyTurn } from "./wpsReadonlyAgent";
import type { WpsSheetReadResult } from "./wpsWorkbookReader";

function makeApplication(): WpsApplication {
  const sheets: WpsWorksheet[] = [{ Name: "明细" }, { Name: "未授权" }];
  const workbook = {
    Name: "测试.xlsx",
    FullName: "C:\\Data\\测试.xlsx",
    Path: "C:\\Data",
    Worksheets: {
      Count: sheets.length,
      Item: (index: number | string) => sheets[Number(index) - 1]
    },
    ActiveSheet: sheets[0]
  };
  return {
    PluginStorage: { getItem: () => null, setItem: () => undefined },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined,
    ActiveWorkbook: workbook,
    ActiveSheet: sheets[0],
    Selection: { Address: () => "$A$1" }
  };
}

const readResult: WpsSheetReadResult = {
  sheet: { index: 1, name: "明细" },
  startRow: 1,
  startColumn: 1,
  rowCount: 3,
  columnCount: 2,
  chunksRead: 1,
  values: [["影院", "数量"], ["甲", 2], ["乙", 3]]
};

describe("WPS readonly agent", () => {
  it("builds structure-only contexts without raw data rows", () => {
    const { scope, workbook } = buildWpsReadonlyContexts(makeApplication(), [readResult]);
    expect(scope.sheets[0].headers).toEqual(["影院", "数量"]);
    expect(scope.worksheetNames).toEqual(["明细"]);
    expect(workbook.worksheets[0].dataRows).toEqual([]);
    expect(workbook.worksheets[0].usedRange).toBe("A1:B3");
  });

  it("runs the model-selected query locally and sends only compact results", async () => {
    const checkIntent = vi.fn(async () => ({
      kind: "tool_request" as const,
      summary: "统计记录",
      confirmedPrompt: "有多少条记录",
      provider: "model" as const,
      request: {
        id: "q1",
        tool: "query_table" as const,
        arguments: {
          mode: "aggregate" as const,
          metrics: [{ operation: "countRows" as const, outputName: "记录数" }]
        }
      }
    }));
    const executeQuery = vi.fn(() => ({
      requestId: "q1",
      tool: "query_table" as const,
      title: "统计",
      headers: ["记录数"],
      rows: [[2]],
      sourceSheets: ["明细"],
      scannedRows: 2,
      complete: true,
      calculation: "countRows",
      warnings: []
    }));
    const createAssistantResponse = vi.fn(async (request) => ({
      kind: "answer" as const,
      message: "共有 2 条记录。",
      provider: "model" as const,
      resultContext: {
        kind: "table" as const,
        title: "统计",
        headers: ["记录数"],
        rows: [[2]],
        sourceSheets: ["明细"],
        warnings: []
      }
    }));
    const result = await runWpsReadonlyTurn(
      {
        application: makeApplication(),
        selectedSheetIndexes: [1],
        prompt: "有多少条记录"
      },
      {
        checkIntent,
        executeQuery,
        createAssistantResponse,
        readSchemas: vi.fn(async () => [readResult]),
        readSheets: vi.fn(async () => [readResult])
      }
    );

    expect(result.kind).toBe("answer");
    expect(executeQuery).toHaveBeenCalledWith(
      expect.objectContaining({ id: "q1" }),
      [{ name: "明细", values: readResult.values }]
    );
    const plannerRequest = createAssistantResponse.mock.calls[0][0];
    expect(plannerRequest.workbook.worksheets[0].dataRows).toEqual([]);
    expect(plannerRequest.dataResults?.[0].rows).toEqual([[2]]);
  });

  it("returns a model-generated write plan for guarded preview and confirmation", async () => {
    const result = await runWpsReadonlyTurn(
      {
        application: makeApplication(),
        selectedSheetIndexes: [1],
        prompt: "删除当前工作表"
      },
      {
        checkIntent: vi.fn(async () => ({
          kind: "proceed" as const,
          summary: "删除工作表",
          confirmedPrompt: "删除当前工作表",
          provider: "model" as const
        })),
        executeQuery: vi.fn(),
        createAssistantResponse: vi.fn(async () => ({
          kind: "plan" as const,
          provider: "model" as const,
          plan: {
            id: "p1",
            title: "删除",
            summary: "删除",
            assumptions: [],
            warnings: [],
            actions: [{ type: "deleteWorksheet" as const, sheet: "明细" }],
            acceptanceCriteria: []
          }
        })),
        readSchemas: vi.fn(async () => [readResult]),
        readSheets: vi.fn(async () => [readResult])
      }
    );

    expect(result).toEqual(expect.objectContaining({
      kind: "plan",
      plan: expect.objectContaining({ id: "p1" })
    }));
  });

  it("answers a greeting without reading full rows or invoking the planner", async () => {
    const readSchemas = vi.fn(async () => [{ ...readResult, values: [readResult.values[0]] }]);
    const readSheets = vi.fn(async () => [readResult]);
    const createAssistantResponse = vi.fn();
    const result = await runWpsReadonlyTurn(
      { application: makeApplication(), selectedSheetIndexes: [1], prompt: "你好" },
      {
        checkIntent: vi.fn(async () => ({ kind: "conversation" as const, message: "你好，有什么需要我帮忙？", provider: "model" as const })),
        executeQuery: vi.fn(),
        createAssistantResponse,
        readSchemas,
        readSheets
      }
    );

    expect(result).toEqual({ kind: "answer", message: "你好，有什么需要我帮忙？", provider: "model", resultContext: null });
    expect(readSchemas).toHaveBeenCalledOnce();
    expect(readSheets).not.toHaveBeenCalled();
    expect(createAssistantResponse).not.toHaveBeenCalled();
  });
});
