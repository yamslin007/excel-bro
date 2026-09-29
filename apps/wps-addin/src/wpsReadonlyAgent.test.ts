import { describe, expect, it, vi } from "vitest";
import type { PlanRequest } from "../../excel-addin/src/contracts";
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

  describe("本地计算的列映射规则", () => {
    const mappingRule = {
      kind: "columnMapping" as const,
      sourceSheet: "明细",
      targetSheet: "明细",
      sourceRange: "A2:A3",
      targetRange: "C2:C3",
      mapping: [
        { source: "甲", target: "一号" },
        { source: "乙", target: "二号" }
      ],
      blankLabel: "未分类"
    };

    function ruleResponse() {
      return {
        kind: "local_plan_rule" as const,
        provider: "model" as const,
        rule: mappingRule,
        summary: "按 2 条映射规则写入。",
        warnings: ["源区域之后再改动，目标列不会自动跟着变。"],
        assumptions: ["源列空白按「未分类」处理。"]
      };
    }

    async function run(overrides: Record<string, unknown> = {}) {
      return runWpsReadonlyTurn(
        {
          application: makeApplication(),
          selectedSheetIndexes: [1],
          prompt: "把名称列翻译成编号"
        },
        {
          checkIntent: vi.fn(async () => ({
            kind: "proceed" as const,
            summary: "列映射",
            confirmedPrompt: "把名称列翻译成编号",
            provider: "model" as const
          })),
          executeQuery: vi.fn(),
          createAssistantResponse: vi.fn(async () => ruleResponse()),
          readSchemas: vi.fn(async () => [readResult]),
          readSheets: vi.fn(async () => [readResult]),
          ...overrides
        } as never
      );
    }

    it("在本端算出目标矩阵并产出计划", async () => {
      const result = await run();
      if (result.kind !== "plan") throw new Error(`期望计划，得到 ${result.kind}`);
      const action = result.plan.actions[0];
      expect(action).toMatchObject({ type: "writeValues", sheet: "明细", range: "C2:C3" });
      expect("values" in action ? action.values : null).toEqual([["一号"], ["二号"]]);
      expect(result.plan.warnings).toEqual(ruleResponse().warnings);
      expect(result.sourceSnapshots).toEqual([{
        sheetName: "明细",
        range: "A2:A3",
        values: [["甲"], ["乙"]]
      }]);
    });

    it("源值不上传：规划请求里既没有 dataRows 也没有算出的矩阵", async () => {
      // 单射映射可逆 —— 标签列加映射表足以还原源列，回传等于上传源列。
      const createAssistantResponse = vi.fn(async (request: PlanRequest) => {
        void request;
        return ruleResponse();
      });
      await run({ createAssistantResponse });
      expect(createAssistantResponse).toHaveBeenCalledOnce();
      const request = createAssistantResponse.mock.calls[0][0];
      expect(request.workbook.worksheets[0].dataRows).toEqual([]);
      expect(request.dataResults).toEqual([]);
      expect(request.planningReadPolicy).toBe("none");
      // 整个请求体里不应出现任何源值。
      expect(JSON.stringify(request)).not.toContain("甲");
    });

    it("规则要读的表不在授权范围内时拒绝", async () => {
      const outsideRule = {
        ...ruleResponse(),
        rule: { ...mappingRule, sourceSheet: "未授权" }
      };
      await expect(
        run({ createAssistantResponse: vi.fn(async () => outsideRule) })
      ).rejects.toThrow(/不在本次已授权的工作表里/);
    });

    it("只在这一支读取完整数据行", async () => {
      // 规则分支必须自己去读源列：上面的 schema 扫描只取字段名。
      const readSheets = vi.fn(async () => [readResult]);
      await run({ readSheets });
      expect(readSheets).toHaveBeenCalledOnce();
    });
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
