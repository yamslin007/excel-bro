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
  it("preserves a structured planner failure rather than returning a successful answer", async () => {
    const result = await runWpsReadonlyTurn({ application: makeApplication(), selectedSheetIndexes: [1], prompt: "生成预览" }, {
      checkIntent: vi.fn(async () => ({ kind: "proceed" as const, summary: "预览", confirmedPrompt: "生成预览", provider: "model" as const })),
      createAssistantResponse: vi.fn(async () => ({ kind: "answer" as const, provider: "system" as const, message: "规划已停止", outcome: "failed" as const, stopReason: "max_tokens" as const })),
      readSchemas: vi.fn(async () => [readResult]), readSheets: vi.fn(), executeQuery: vi.fn()
    });
    expect(result).toMatchObject({ kind: "answer", outcome: "failed" });
    expect(result).not.toHaveProperty("plan");
  });
  it("routes a planning-stage form to clarification before any data-row read or execution", async () => {
    const readSheets = vi.fn();
    const response = { kind: "clarification" as const, provider: "model" as const, clarification: {
      id: "planning-input", summary: "需要补充统计口径", question: "按什么单位？",
      reason: "结果取决于单位", scopeLabel: "明细", options: [],
      fields: [{ id: "unit", label: "统计单位" }]
    } };
    const result = await runWpsReadonlyTurn({ application: makeApplication(), selectedSheetIndexes: [1], prompt: "请生成统计预览" }, {
      checkIntent: vi.fn(async () => ({ kind: "proceed" as const, summary: "统计", confirmedPrompt: "请生成统计预览", provider: "model" as const })),
      createAssistantResponse: vi.fn(async () => response),
      readSchemas: vi.fn(async () => [readResult]), readSheets, executeQuery: vi.fn()
    });
    expect(result).toMatchObject({ kind: "clarification", clarification: { response }, headerRows: [["影院", "数量"]] });
    expect(readSheets).not.toHaveBeenCalled();
    expect(result).not.toHaveProperty("plan");
  });
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
          turnId: "turn-plan-1",
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
      plan: expect.objectContaining({ id: "p1" }),
      // 规划身份必须带出来：断线后先拿它查真实状态，而不是重发一遍需求。
      turnId: "turn-plan-1"
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
        warnings: ["目标区域已有内容会被覆盖。"],
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
      expect(action).toMatchObject({ type: "writeFormulas", sheet: "明细", range: "C2:C3" });
      expect("formulas" in action ? action.formulas : null).toEqual(expect.arrayContaining([
        [expect.stringContaining('EXACT(A2,"甲")')],
        [expect.stringContaining('EXACT(A3,"乙")')]
      ]));
      expect(result.plan.acceptanceCriteria).toEqual([{
        type: "rangeEquals", sheet: "明细", range: "C2:C3", expected: [["一号"], ["二号"]]
      }]);
      expect(result.plan.warnings).toEqual(expect.arrayContaining(ruleResponse().warnings));
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

    describe("空白：只有真读到空白才拦，四种处理互不冒充", () => {
      // 与 readResult 只差一格：A3 是空的（0 与 false 不算空白，这里用真空格）。
      const blankRead: WpsSheetReadResult = {
        ...readResult,
        values: [["影院", "数量"], ["甲", 2], ["", 3]]
      };

      async function runRule(rule: Record<string, unknown>, read: WpsSheetReadResult = blankRead) {
        return run({
          createAssistantResponse: vi.fn(async () => ({
            ...ruleResponse(),
            rule: { ...mappingRule, ...rule }
          })),
          readSheets: vi.fn(async () => [read]),
          readSchemas: vi.fn(async () => [read])
        });
      }

      /**
       * 计划里实际会落到目标区域的矩阵（公式通道时用试算期望值）。
       *
       * 这里刻意不引入 `AnalysisPlan` 的窄类型：`actions` 是联合体，逐个
       * 收窄会让测试代码比被测行为还长；断言值本身才是要钉的东西。
       */
      function targetMatrix(plan: { actions: readonly unknown[] }): unknown {
        const action = plan.actions[0] as { type?: string; values?: unknown } | undefined;
        if (action?.type === "writeValues") return action.values;
        const criterion = (
          plan as { acceptanceCriteria?: readonly { type?: string; expected?: unknown }[] }
        ).acceptanceCriteria?.[0];
        return criterion?.type === "rangeEquals" ? criterion.expected : null;
      }

      it("源区域真的有空白、没给策略：走结构化澄清，不是一条报错文案", async () => {
        const result = await runRule({ blankLabel: undefined });
        if (result.kind !== "clarification") throw new Error(`期望澄清，得到 ${result.kind}`);
        const local = result.clarification.localMapping;
        expect(local?.reason).toBe("blank");
        expect(local?.blankCount).toBe(1);
        // 澄清卡不带任何源值：源值只留本端，"有几格空白"足够判断规模。
        expect(local?.values).toEqual([]);
        expect(local?.rule).toMatchObject({ sourceRange: "A2:A3", targetRange: "C2:C3" });
        expect(result).not.toHaveProperty("plan");
      });

      it("同一条规则、源区域一格空白都没有：不挡预览", async () => {
        const result = await runRule({ blankLabel: undefined }, readResult);
        expect(result.kind).toBe("plan");
      });

      it("选了保留目标原值：那一格根本不进任何写入范围", async () => {
        const result = await runRule({ blankLabel: undefined, blankPolicy: { kind: "keep" } });
        if (result.kind !== "plan") throw new Error(`期望计划，得到 ${result.kind}`);
        expect(result.plan.actions.map((action) => (action as { range: string }).range))
          .toEqual(["C2:C2"]);
        // 不是"读回原值再写回"：C3 不出现在任何动作里。
        expect(JSON.stringify(result.plan.actions)).not.toContain("C3");
      });

      it("清空与写空串不是一回事", async () => {
        const cleared = await runRule({ blankLabel: undefined, blankPolicy: { kind: "clear" } });
        if (cleared.kind !== "plan") throw new Error(`期望计划，得到 ${cleared.kind}`);
        expect(cleared.plan.actions).toEqual([
          { type: "writeValues", sheet: "明细", range: "C2:C2", values: [["一号"]] },
          { type: "clearRange", sheet: "明细", range: "C3:C3", applyTo: "contents" }
        ]);

        const emptyString = await runRule({ blankLabel: undefined, blankPolicy: { kind: "value", value: "" } });
        if (emptyString.kind !== "plan") throw new Error(`期望计划，得到 ${emptyString.kind}`);
        expect(emptyString.plan.actions).toHaveLength(1);
        expect(emptyString.plan.actions[0]).toMatchObject({ sheet: "明细", range: "C2:C3" });
        expect(emptyString.plan.actions.some((action) => action.type === "clearRange")).toBe(false);
        expect(targetMatrix(emptyString.plan)).toEqual([["一号"], [""]]);
      });

      it("未覆盖类别 / 分隔符归因 / 空白 三条路各自归因，不互相冒充", async () => {
        const compositeRead: WpsSheetReadResult = {
          ...readResult,
          values: [["影院", "数量"], ["甲，乙", 2], ["乙", 3]]
        };
        const blank = await runRule({ blankLabel: undefined });
        const unmapped = await runRule(
          { blankLabel: undefined, mapping: [{ source: "甲", target: "一号" }] },
          readResult
        );
        // 声明了「、」，可这一格里是「，」：分隔符没声明全，还是这格本身就缺
        // 映射？归因不确定 —— 报上去让用户选，不替他猜成"补映射表"。
        const separator = await runRule(
          {
            blankLabel: undefined,
            delimiters: ["、"],
            undeclaredSeparators: ["，"],
            mapping: [{ source: "甲", target: "一号" }, { source: "乙", target: "二号" }]
          },
          compositeRead
        );
        for (const item of [blank, unmapped, separator]) {
          if (item.kind !== "clarification") throw new Error(`期望澄清，得到 ${item.kind}`);
        }
        const reasons = [blank, unmapped, separator].map(
          (item) => (item as { clarification: { localMapping?: { reason: string } } }).clarification.localMapping?.reason
        );
        expect(reasons).toEqual(["blank", "unmapped", "separator_or_mapping"]);
        // 三条路各自带自己的证据：空白带计数，另两条带具体值。
        expect(unmapped).toHaveProperty("clarification.localMapping.values", ["乙"]);
        expect(separator).toHaveProperty("clarification.localMapping.values", ["甲，乙"]);
        expect(blank).toHaveProperty("clarification.localMapping.blankCount", 1);
      });

      it("归因不确定的规则在读源值之前就分流：不白读一次源列", async () => {
        const readSheets = vi.fn(async () => [readResult]);
        const result = await run({
          createAssistantResponse: vi.fn(async () => ({
            ...ruleResponse(),
            rule: {
              ...mappingRule,
              clarification: { reason: "unresolvedSeparatorOrMapping", missingCategories: ["丙"] }
            }
          })),
          readSheets
        });
        expect(result.kind).toBe("clarification");
        expect(readSheets).not.toHaveBeenCalled();
      });

      it("没有读到源值就不许声称有空白", async () => {
        // 同一条规则，源区域一格都没读到：报的是"没读到"，**不是**"发现 N
        // 格空白"。事实只能来自真实读取 —— 反过来（没读就先说有空白）正是
        // 旧实现拦住预览的那条路。
        const outcome: unknown = await runRule(
          { blankLabel: undefined },
          { ...blankRead, values: [] }
        ).catch((thrown: unknown) => thrown);
        expect(outcome).toBeInstanceOf(Error);
        expect((outcome as Error).message).toContain("已读入到第 0 行");
        expect((outcome as Error).message).not.toContain("空白");
      });
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
