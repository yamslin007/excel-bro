import { describe, expect, it, vi } from "vitest";
import type { PlanRequest, IntentCheckRequest, MessageUnderstandingRequest } from "../../excel-addin/src/contracts";
import type { WpsApplication, WpsRange, WpsWorksheet } from "./wpsHost";
import {
  buildWpsReadonlyContexts,
  reinspectWpsPlanningPause,
  resumeWpsPlanning,
  runWpsReadonlyTurn,
  resumeWpsLocalMapping
} from "./wpsReadonlyAgent";
import { ApiRequestError } from "../../excel-addin/src/api";
import type { WpsSheetReadResult } from "./wpsWorkbookReader";
import { TargetConfirmationRequiredError } from "@excel-bro/core/authorizedTarget";

vi.mock("../../excel-addin/src/api", async (importOriginal) => ({
  ...await importOriginal<typeof import("../../excel-addin/src/api")>(),
  understandMessage: vi.fn(async () => ({ kind: "needs_workbook_context", summary: "需要字段结构", provider: "model" }))
}));

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
  it("passes one explicit task context through intent and planning without reading rows", async () => {
    const taskContext = { schemaVersion: 1 as const, originalPrompt: "根据类别填写等级", confirmations: [
      { questionId: "mode", label: "口径", value: "逐行填写" },
      { questionId: "blank", label: "空值", value: "不写入" }
    ] };
    const checkIntent = vi.fn(async (_request: IntentCheckRequest) => ({ kind: "proceed" as const, provider: "model" as const,
      summary: "填写", confirmedPrompt: "根据类别逐行填写，空值不写入" }));
    const createAssistantResponse = vi.fn(async (_request: PlanRequest) => ({ kind: "answer" as const,
      provider: "model" as const, message: "完整规则已收到" }));
    const readSheets = vi.fn();
    await runWpsReadonlyTurn({ application: makeApplication(), selectedSheetIndexes: [1], prompt: "根据类别填写等级", taskContext }, {
      checkIntent, createAssistantResponse, readSchemas: vi.fn(async () => [{ ...readResult, values: [readResult.values[0]] }]),
      readSheets, executeQuery: vi.fn(), streamPlanning: vi.fn(), getPlanningStatus: vi.fn(), continuePlanning: vi.fn()
    });
    expect(checkIntent.mock.calls[0][0].taskContext).toEqual(taskContext);
    expect(createAssistantResponse.mock.calls[0][0].taskContext).toEqual(taskContext);
    expect(readSheets).not.toHaveBeenCalled();
    expect(JSON.stringify(createAssistantResponse.mock.calls[0][0])).not.toContain('["甲",2]');
  });
  it("local mapping resume checks target before reading source rows", async () => {
    const readSheets = vi.fn(async () => [readResult]);
    await expect(resumeWpsLocalMapping({ application: makeApplication(), selectedSheetIndexes: [1],
      rule: { kind: "columnMapping", sourceSheet: "明细", targetSheet: "明细", sourceRange: "A2:A3", targetRange: "B2:B3",
        mapping: [{ source: "甲", target: "一" }, { source: "乙", target: "二" }], blankPolicy: { kind: "keep" } },
      context: { summary: "映射", warnings: [], assumptions: [] }, additions: [],
      beforeRead: async () => { throw new TargetConfirmationRequiredError(); }
    }, { readSheets, readSchemas: vi.fn(), checkIntent: vi.fn(), createAssistantResponse: vi.fn(), streamPlanning: vi.fn(), getPlanningStatus: vi.fn(), continuePlanning: vi.fn(), executeQuery: vi.fn() }))
      .rejects.toBeInstanceOf(TargetConfirmationRequiredError);
    expect(readSheets).not.toHaveBeenCalled();
  });
  it("local mapping resume checks again after host read; a switched target cannot become a plan", async () => {
    let switched = false;
    const readSheets = vi.fn(async () => { switched = true; return [readResult]; });
    const guard = vi.fn(async () => { if (switched) throw new TargetConfirmationRequiredError(); });
    await expect(resumeWpsLocalMapping({ application: makeApplication(), selectedSheetIndexes: [1],
      rule: { kind: "columnMapping", sourceSheet: "明细", targetSheet: "明细", sourceRange: "A2:A3", targetRange: "B2:B3",
        mapping: [{ source: "甲", target: "一" }, { source: "乙", target: "二" }], blankPolicy: { kind: "keep" } },
      context: { summary: "映射", warnings: [], assumptions: [] }, additions: [], beforeRead: guard
    }, { readSheets, readSchemas: vi.fn(), checkIntent: vi.fn(), createAssistantResponse: vi.fn(), streamPlanning: vi.fn(), getPlanningStatus: vi.fn(), continuePlanning: vi.fn(), executeQuery: vi.fn() }))
      .rejects.toBeInstanceOf(TargetConfirmationRequiredError);
    expect(readSheets).toHaveBeenCalledTimes(1);
    expect(guard).toHaveBeenCalledTimes(2);
  });
  it("pauses before any source rows are read when target changes during model intent", async () => {
    let changed = false;
    const readSchemas = vi.fn(async () => [readResult]);
    const readSheets = vi.fn();
    const createAssistantResponse = vi.fn();
    await expect(runWpsReadonlyTurn({ application: makeApplication(), selectedSheetIndexes: [1], prompt: "统计记录",
      beforeRead: async () => { if (changed) throw new TargetConfirmationRequiredError(); }
    }, {
      checkIntent: vi.fn(async () => { changed = true; return {
        kind: "tool_request" as const, summary: "统计", confirmedPrompt: "统计", provider: "model" as const,
        request: { id: "q", tool: "query_table" as const, arguments: { mode: "aggregate" as const,
          metrics: [{ operation: "countRows" as const, outputName: "记录数" }] } }
      }; }), readSchemas, readSheets, createAssistantResponse, streamPlanning: vi.fn(), getPlanningStatus: vi.fn(), continuePlanning: vi.fn(), executeQuery: vi.fn()
    })).rejects.toBeInstanceOf(TargetConfirmationRequiredError);
    expect(readSchemas).toHaveBeenCalledOnce();
    expect(readSheets).not.toHaveBeenCalled();
    expect(createAssistantResponse).not.toHaveBeenCalled();
  });
  it("preserves a structured planner failure rather than returning a successful answer", async () => {
    const result = await runWpsReadonlyTurn({ application: makeApplication(), selectedSheetIndexes: [1], prompt: "生成预览" }, {
      checkIntent: vi.fn(async () => ({ kind: "proceed" as const, summary: "预览", confirmedPrompt: "生成预览", provider: "model" as const })),
      createAssistantResponse: vi.fn(async () => ({ kind: "answer" as const, provider: "system" as const, message: "规划已停止", outcome: "failed" as const, stopReason: "max_tokens" as const })),
      readSchemas: vi.fn(async () => [readResult]), readSheets: vi.fn(), streamPlanning: vi.fn(), getPlanningStatus: vi.fn(), continuePlanning: vi.fn(), executeQuery: vi.fn()
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
      readSchemas: vi.fn(async () => [readResult]), readSheets, streamPlanning: vi.fn(), getPlanningStatus: vi.fn(), continuePlanning: vi.fn(), executeQuery: vi.fn()
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
        streamPlanning: vi.fn(),
        getPlanningStatus: vi.fn(),
        continuePlanning: vi.fn(),
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

  it("SDK 规划下 tool_request 不预置 dataResults（执行改由工具桥驱动）", async () => {
    // 缺陷复现：SDK 适配器只看表结构，模型要自己调 query_table 经工具桥读表。
    // 旧路径若把本地算好的 dataResults 塞进请求，会被服务端的"不上传数据行"
    // 守卫直接拒绝 —— 新工具桥永远拿不到执行权。
    const executeQuery = vi.fn();
    const readSheets = vi.fn(async () => [readResult]);
    const createAssistantResponse = vi.fn(async (_request: PlanRequest) => ({
      kind: "answer" as const,
      provider: "model" as const,
      message: "已按工具桥结果作答"
    }));
    await runWpsReadonlyTurn(
      {
        application: makeApplication(),
        selectedSheetIndexes: [1],
        prompt: "有多少条记录",
        modelId: "claude-agent-sdk"
      },
      {
        checkIntent: vi.fn(async () => ({
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
        })),
        executeQuery,
        streamPlanning: vi.fn(),
        getPlanningStatus: vi.fn(),
        continuePlanning: vi.fn(),
        createAssistantResponse,
        readSchemas: vi.fn(async () => [readResult]),
        readSheets
      }
    );
    const plannerRequest = createAssistantResponse.mock.calls[0][0];
    expect(plannerRequest.dataResults).toEqual([]);
    // 本地预置计算与全表读取都不该发生：读表要等到桥指针回来那一刻。
    expect(executeQuery).not.toHaveBeenCalled();
    expect(readSheets).not.toHaveBeenCalled();
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
        streamPlanning: vi.fn(),
        getPlanningStatus: vi.fn(),
        continuePlanning: vi.fn(),
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

  it("answers a greeting before authorization, schema reads or field-aware intent", async () => {
    const readSchemas = vi.fn(async () => [{ ...readResult, values: [readResult.values[0]] }]);
    const readSheets = vi.fn(async () => [readResult]);
    const createAssistantResponse = vi.fn();
    const checkIntent = vi.fn();
    const beforeRead = vi.fn();
    const result = await runWpsReadonlyTurn(
      { application: makeApplication(), selectedSheetIndexes: [], prompt: "你好", beforeRead },
      {
        understandMessage: vi.fn(async () => ({ kind: "conversation" as const, message: "你好，有什么需要我帮忙？", provider: "model" as const })),
        checkIntent,
        executeQuery: vi.fn(),
        streamPlanning: vi.fn(),
        getPlanningStatus: vi.fn(),
        continuePlanning: vi.fn(),
        createAssistantResponse,
        readSchemas,
        readSheets
      }
    );

    expect(result).toEqual({ kind: "answer", message: "你好，有什么需要我帮忙？", provider: "model", resultContext: null });
    expect(readSchemas).not.toHaveBeenCalled();
    expect(beforeRead).not.toHaveBeenCalled();
    expect(checkIntent).not.toHaveBeenCalled();
    expect(readSheets).not.toHaveBeenCalled();
    expect(createAssistantResponse).not.toHaveBeenCalled();
  });

  it("understands first, then checks authorization and reads fields before field-aware intent", async () => {
    const order: string[] = [];
    const progress: string[] = [];
    const readSchemas = vi.fn(async () => { order.push("schema"); return [readResult]; });
    const readSheets = vi.fn();
    const result = await runWpsReadonlyTurn({
      application: makeApplication(), selectedSheetIndexes: [1], prompt: "检查当前表",
      beforeRead: async () => { order.push("authorization"); },
      onProgress: (event) => progress.push(event.stepId ?? "")
    }, {
      understandMessage: vi.fn(async (request: MessageUnderstandingRequest) => {
        order.push("message");
        expect(request).not.toHaveProperty("scope");
        return { kind: "needs_workbook_context", summary: "检查当前表", provider: "model" } as const;
      }),
      checkIntent: vi.fn(async () => { order.push("field-intent"); return { kind: "conversation", message: "已理解字段", provider: "model" } as const; }),
      readSchemas, readSheets, executeQuery: vi.fn(), createAssistantResponse: vi.fn(),
      streamPlanning: vi.fn(), getPlanningStatus: vi.fn(), continuePlanning: vi.fn()
    });
    expect(result.kind).toBe("answer");
    expect(order).toEqual(["message", "authorization", "schema", "field-intent"]);
    expect(progress.slice(0, 3)).toEqual(["message-intent", "read-schemas", "intent"]);
    expect(readSheets).not.toHaveBeenCalled();
  });

  it("cannot read fields if message understanding fails or is cancelled", async () => {
    for (const cancel of [false, true]) {
      const controller = new AbortController();
      const beforeRead = vi.fn();
      const readSchemas = vi.fn();
      await expect(runWpsReadonlyTurn({
        application: makeApplication(), selectedSheetIndexes: [1], prompt: "检查当前表", beforeRead, signal: controller.signal
      }, {
        understandMessage: vi.fn(async () => {
          if (!cancel) throw new Error("消息理解失败");
          controller.abort();
          return { kind: "needs_workbook_context", summary: "检查当前表", provider: "model" } as const;
        }),
        readSchemas, readSheets: vi.fn(), executeQuery: vi.fn(), checkIntent: vi.fn(),
        createAssistantResponse: vi.fn(), streamPlanning: vi.fn(), getPlanningStatus: vi.fn(), continuePlanning: vi.fn()
      })).rejects.toThrow();
      expect(beforeRead).not.toHaveBeenCalled();
      expect(readSchemas).not.toHaveBeenCalled();
    }
  });

  it("a target change during message understanding is rejected before the first schema read", async () => {
    let changed = false;
    const readSchemas = vi.fn();
    await expect(runWpsReadonlyTurn({
      application: makeApplication(), selectedSheetIndexes: [1], prompt: "检查当前表",
      beforeRead: async () => { if (changed) throw new TargetConfirmationRequiredError(); }
    }, {
      understandMessage: vi.fn(async () => {
        changed = true;
        return { kind: "needs_workbook_context", summary: "检查当前表", provider: "model" } as const;
      }),
      readSchemas, readSheets: vi.fn(), checkIntent: vi.fn(), executeQuery: vi.fn(),
      createAssistantResponse: vi.fn(), streamPlanning: vi.fn(), getPlanningStatus: vi.fn(), continuePlanning: vi.fn()
    })).rejects.toBeInstanceOf(TargetConfirmationRequiredError);
    expect(readSchemas).not.toHaveBeenCalled();
  });

  /**
   * CAP-02：`runWpsTurn` 的**通道选择**纪律。
   *
   * 这一组测试保护的不是"能不能流式"，而是三条容易写错的分支：
   *  1. 有 turnId 时必须优先走流式（否则 CAP-02 等于没接）；
   *  2. **只**在端点不存在/连不上时回退一次性 —— 业务错误必须原样抛出；
   *  3. 暂停是结构化终态，不能被当成 answer 或失败。
   */
  describe("CAP-02 规划通道选择", () => {
    function baseDeps(overrides: Record<string, unknown> = {}) {
      return {
        checkIntent: vi.fn(async () => ({
          kind: "proceed" as const,
          summary: "预览",
          confirmedPrompt: "生成预览",
          provider: "model" as const,
          turnId: "turn-stream-1"
        })),
        readSchemas: vi.fn(async () => [readResult]),
        readSheets: vi.fn(async () => [readResult]),
        executeQuery: vi.fn(),
        createAssistantResponse: vi.fn(async () => ({
          kind: "answer" as const,
          provider: "model" as const,
          message: "一次性通道回答",
          resultContext: null
        })),
        streamPlanning: vi.fn(),
        getPlanningStatus: vi.fn(),
        continuePlanning: vi.fn(),
        ...overrides
      } as never;
    }

    async function run(deps: Record<string, unknown>) {
      return runWpsReadonlyTurn(
        { application: makeApplication(), selectedSheetIndexes: [1], prompt: "生成预览" },
        deps as never
      );
    }

    it("有 turnId 时优先走流式，且事件进增量进度（变异点：跳过流式直接一次性）", async () => {
      const progress: string[] = [];
      const received = vi.fn();
      const streamPlanning = vi.fn(async (_req: unknown, _turnId: string, hooks: { onProgress?: (e: { stepId: string; title: string; summary?: string; sequence: number }) => void }) => {
        hooks.onProgress?.({ stepId: "s1", title: "读取结构", summary: "正在读取", sequence: 7 });
        return { kind: "answer" as const, provider: "model" as const, message: "流式回答", resultContext: null };
      });
      const deps = baseDeps({ streamPlanning });
      const result = await runWpsReadonlyTurn(
        {
          application: makeApplication(),
          selectedSheetIndexes: [1],
          prompt: "生成预览",
          onProgress: (p) => { progress.push(p.title ?? ""); received(p); }
        },
        deps
      );
      expect(result).toMatchObject({ kind: "answer", message: "流式回答" });
      expect((deps as { createAssistantResponse: ReturnType<typeof vi.fn> }).createAssistantResponse)
        .not.toHaveBeenCalled();
      expect(progress).toContain("读取结构");
      expect(received).toHaveBeenCalledWith(expect.objectContaining({ turnId: "turn-stream-1", sequence: 7 }));
    });

    it("桥指针必须原样透传到面板（丢掉它会让服务端一直等到超时）", async () => {
      const seen: unknown[] = [];
      const cards: unknown[] = [];
      const pythonResult = { resultRef: "res_" + "c".repeat(32), title: "Python 分析结果",
        headers: ["列1"], rowCount: 1, columnCount: 1, sourceSheets: ["Sheet1"], complete: true };
      const streamPlanning = vi.fn(async (_req: unknown, _turnId: string, hooks: {
        onProgress?: (e: Record<string, unknown>) => void;
      }) => {
        hooks.onProgress?.({
          stepId: "bridge.query_table",
          title: "正在本机执行数据查询",
          sequence: 9,
          bridgeTool: { operationId: "op_" + "a".repeat(32), tool: "query_table" },
          pythonResult
        });
        return { kind: "answer" as const, provider: "model" as const, message: "ok", resultContext: null };
      });
      const deps = baseDeps({ streamPlanning });
      await runWpsReadonlyTurn(
        {
          application: makeApplication(),
          selectedSheetIndexes: [1],
          prompt: "生成预览",
          onProgress: (p) => { seen.push(p.bridgeTool); cards.push(p.pythonResult); }
        },
        deps
      );
      // 只有操作 ID 与工具名，没有任何请求参数/值。
      // 其余进度事件（意图、读结构…）本来就不带指针 —— 这里只挑出带指针的那些。
      expect(seen.filter((item) => item !== undefined && item !== null)).toEqual([
        { operationId: "op_" + "a".repeat(32), tool: "query_table" }
      ]);
      expect(cards.filter(Boolean)).toEqual([pythonResult]);
    });

    it("本地结果卡指针随回答一路上抛，且不夹带值", async () => {
      const card = {
        resultRef: "res_" + "b".repeat(32),
        title: "按类别汇总",
        headers: ["类别", "数量"],
        rowCount: 2,
        columnCount: 2,
        sourceSheets: ["明细"],
        complete: true
      };
      const deps = baseDeps({
        streamPlanning: vi.fn(async () => ({
          kind: "answer" as const,
          provider: "model" as const,
          message: "已在本机算出结果。",
          resultContext: null,
          localResult: card
        }))
      });
      const result = await run(deps);
      expect(result).toMatchObject({ kind: "answer", localResult: card });
      // 指针里只有引用与结构：一个数据行都没有。
      expect(JSON.stringify(result)).not.toContain("rows");
    });

    it("端点不存在（404）才回退一次性通道", async () => {
      const notFound = new ApiRequestError(404, "NOT_FOUND", "没有这个路由", false);
      const deps = baseDeps({ streamPlanning: vi.fn(async () => { throw notFound; }) });
      const result = await run(deps);
      expect(result).toMatchObject({ kind: "answer", message: "一次性通道回答" });
      expect((deps as { createAssistantResponse: ReturnType<typeof vi.fn> }).createAssistantResponse)
        .toHaveBeenCalledOnce();
    });

    it("本地服务连不上时回退一次性通道", async () => {
      const offline = new TypeError("Failed to fetch");
      const deps = baseDeps({ streamPlanning: vi.fn(async () => { throw offline; }) });
      const result = await run(deps);
      expect(result).toMatchObject({ kind: "answer", message: "一次性通道回答" });
    });

    it("业务错误必须原样抛出、绝不回退（变异点：把回退条件放宽成 any）", async () => {
      // 409 是"服务端收到了请求并明确拒绝"，回退只会掩盖原因、可能再花一次模型钱。
      const conflict = new ApiRequestError(409, "CAPABILITY_MISMATCH", "能力不匹配", false);
      const deps = baseDeps({ streamPlanning: vi.fn(async () => { throw conflict; }) });
      await expect(run(deps)).rejects.toBe(conflict);
      expect((deps as { createAssistantResponse: ReturnType<typeof vi.fn> }).createAssistantResponse)
        .not.toHaveBeenCalled();
    });

    it("暂停作为结构化终态返回，不落到 answer 分支（变异点：当成 answer 渲染）", async () => {
      const pause = {
        kind: "planning_pause" as const,
        turnId: "turn-stream-1",
        status: "waiting_user" as const,
        revision: 3,
        lastSequence: 7,
        reason: "本轮处理时间已达上限。",
        activeElapsedSeconds: 120,
        approvedWallTimeSeconds: 120,
        extensionSeconds: 60,
        extensionCount: 1,
        maximumExtensions: 3,
        canContinue: true,
        requiresExtension: true,
        pauseType: "budget" as const,
        currentModelId: null,
        modelSwitchCount: 0,
        canSwitchModel: false,
        availableModelConnections: []
      };
      const deps = baseDeps({ streamPlanning: vi.fn(async () => pause) });
      const result = await run(deps);
      expect(result).toMatchObject({ kind: "planning_pause", pause: { turnId: "turn-stream-1" } });
      expect(result).not.toHaveProperty("message");
      // 暂停不是失败路径：一次性通道根本不该被调用。
      expect((deps as { createAssistantResponse: ReturnType<typeof vi.fn> }).createAssistantResponse)
        .not.toHaveBeenCalled();
    });

    it("没有 turnId 时不走流式（服务端要靠它去重与恢复）", async () => {
      const deps = baseDeps({
        checkIntent: vi.fn(async () => ({
          kind: "proceed" as const,
          summary: "预览",
          confirmedPrompt: "生成预览",
          provider: "model" as const
        }))
      });
      await run(deps);
      expect((deps as { streamPlanning: ReturnType<typeof vi.fn> }).streamPlanning).not.toHaveBeenCalled();
      expect((deps as { createAssistantResponse: ReturnType<typeof vi.fn> }).createAssistantResponse)
        .toHaveBeenCalledOnce();
    });

    it("流式返回未知结果种类必须报错，不静默丢弃", async () => {
      const deps = baseDeps({
        streamPlanning: vi.fn(async () => ({ kind: "totally_new_kind", payload: 1 }))
      });
      await expect(run(deps)).rejects.toThrow(/尚不支持的结果类型/);
      expect((deps as { createAssistantResponse: ReturnType<typeof vi.fn> }).createAssistantResponse)
        .not.toHaveBeenCalled();
    });
  });

  /**
   * CAP-02：续跑入口的纪律。
   *
   * 这两个断言钉的是"续跑不会变成重跑"与"续跑不写单元格"。
   */
  describe("CAP-02 续跑入口", () => {
    it("先查状态再继续：revision 用服务端最新值，游标用客户端收到的位置", async () => {
      const order: string[] = [];
      const getPlanningStatus = vi.fn(async () => {
        order.push("status");
        return {
          turnId: "turn-r",
          status: "waiting_user" as const,
          revision: 8,
          lastSequence: 21,
          canContinue: true,
          pauseType: "budget" as const,
          reason: null
        };
      });
      const continuePlanning = vi.fn(async (body: { expectedRevision: number; afterSequence: number }) => {
        order.push("continue");
        expect(body.expectedRevision).toBe(8);
        expect(body.afterSequence).toBe(5);
        return { kind: "answer" as const, provider: "model" as const, message: "继续完成", resultContext: null };
      });
      const result = await resumeWpsPlanning(
        { turnId: "turn-r", afterSequence: 5 },
        { getPlanningStatus, continuePlanning } as never
      );
      expect(order).toEqual(["status", "continue"]);
      expect(result.outcome).toMatchObject({ kind: "answer", message: "继续完成" });
    });

    it("服务端说不能继续时不硬续（变异点：无视 canContinue 直接续）", async () => {
      const continuePlanning = vi.fn();
      // 刻意构造"有检查点但服务端明确说不能继续"：只凭 action 推导会误放行，
      // 必须由 canContinue 闸拦住 —— 这正是这条断言的唯一判别点。
      const result = await resumeWpsPlanning(
        { turnId: "turn-paused" },
        {
          getPlanningStatus: vi.fn(async () => ({
            turnId: "turn-paused",
            status: "waiting_user" as const,
            revision: 6,
            lastSequence: 12,
            canContinue: false,
            pauseType: "budget" as const,
            reason: "本轮预算已耗尽，需要重新发起。"
          })),
          continuePlanning
        } as never
      );
      expect(continuePlanning).not.toHaveBeenCalled();
      expect(result.outcome).toBeNull();
    });

    it("已完成的任务绝不续跑", async () => {
      const continuePlanning = vi.fn();
      const result = await resumeWpsPlanning(
        { turnId: "turn-done" },
        {
          getPlanningStatus: vi.fn(async () => ({
            turnId: "turn-done",
            status: "completed" as const,
            revision: 9,
            lastSequence: 30,
            canContinue: false,
            pauseType: null,
            reason: null
          })),
          continuePlanning
        } as never
      );
      expect(continuePlanning).not.toHaveBeenCalled();
      expect(result.outcome).toBeNull();
    });

    /**
     * 第 2 条：**调用方给了 requestId 就必须原样用**。
     *
     * 变异点：函数内部无视传入值、每次都 `newPlanningRequestId()` ——
     * 超时重试就变成"另一次继续"，服务端重复追加预算、重复跑规划。
     */
    it("沿用调用方传入的 requestId（幂等键不由内部另生成）", async () => {
      const continuePlanning = vi.fn(async (body: { requestId: string }) => ({
        kind: "answer" as const, provider: "model" as const, message: "好", resultContext: null
      }));
      await resumeWpsPlanning(
        { turnId: "turn-x", requestId: "plan-continue-fixed" },
        {
          getPlanningStatus: vi.fn(async () => ({
            turnId: "turn-x",
            status: "waiting_user" as const,
            revision: 3,
            lastSequence: 7,
            canContinue: true,
            pauseType: "budget" as const,
            reason: null
          })),
          continuePlanning
        } as never
      );
      expect(continuePlanning.mock.calls[0][0]).toMatchObject({
        requestId: "plan-continue-fixed"
      });
    });

    /**
     * 第 3 条：调用方给了**客户端游标**时必须用它，不得退回服务端最新序号。
     *
     * 变异点：忽略 `afterSequence` 直接用 `status.lastSequence` ——
     * 中间那些"已提交但客户端漏收"的事件会被永久跳过。
     */
    it("游标用客户端已确认位置，不用服务端最新序号（漏收补发）", async () => {
      const continuePlanning = vi.fn(async (_input: { afterSequence?: number }) => ({
        kind: "answer" as const, provider: "model" as const, message: "好", resultContext: null
      }));
      await resumeWpsPlanning(
        { turnId: "turn-y", requestId: "r-1", afterSequence: 4 },
        {
          getPlanningStatus: vi.fn(async () => ({
            turnId: "turn-y",
            status: "waiting_user" as const,
            revision: 5,
            lastSequence: 20,
            canContinue: true,
            pauseType: "budget" as const,
            reason: null
          })),
          continuePlanning
        } as never
      );
      // 服务端已到 20，但客户端只消费到 4 → 必须传 4，让 5..20 被补发。
      expect(continuePlanning.mock.calls[0][0]).toMatchObject({ afterSequence: 4 });
    });

    it("未给游标时从零补收，不能把服务端最新序号当作已收到", async () => {
      const continuePlanning = vi.fn(async (_input: { afterSequence?: number }) => ({
        kind: "answer" as const, provider: "model" as const, message: "好", resultContext: null
      }));
      await resumeWpsPlanning(
        { turnId: "turn-z", requestId: "r-2" },
        {
          getPlanningStatus: vi.fn(async () => ({
            turnId: "turn-z",
            status: "waiting_user" as const,
            revision: 5,
            lastSequence: 20,
            canContinue: true,
            pauseType: "budget" as const,
            reason: null
          })),
          continuePlanning
        } as never
      );
      expect(continuePlanning.mock.calls[0][0]).toMatchObject({ afterSequence: 0 });
    });

    it("续跑进度向面板转发原任务身份和实际事件序号", async () => {
      const onProgress = vi.fn();
      await resumeWpsPlanning({ turnId: "turn-events", afterSequence: 3, onProgress }, {
        getPlanningStatus: vi.fn(async () => ({ turnId: "turn-events", status: "interrupted", revision: 7,
          lastSequence: 20, canContinue: true, pauseType: "interrupted", reason: null })),
        continuePlanning: vi.fn(async (_body, hooks) => {
          hooks.onProgress({ stepId: "resume-1", title: "继续规划", sequence: 8 });
          return { kind: "answer", provider: "model", message: "完成" };
        })
      } as never);
      expect(onProgress).toHaveBeenCalledWith(expect.objectContaining({ turnId: "turn-events", sequence: 8 }));
    });
  });
/**
   * CAP-02 第 5 条：面板刷新后重新核对历史里的暂停。
   *
   * 这段不与上面的「续跑入口」合并 —— 它钉的是**重开那一刻**的行为：
   * 只核对、不续跑，且不能续跑的结论必须来自服务端而不是历史快照。
   */
  describe("CAP-02 刷新后重新核对暂停", () => {
    const pausedAt = {
      kind: "planning_pause" as const,
      turnId: "turn-reload",
      status: "waiting_user" as const,
      revision: 3,
      lastSequence: 9,
      reason: "本轮处理时间已达上限",
      activeElapsedSeconds: 120,
      approvedWallTimeSeconds: 600,
      extensionSeconds: 0,
      extensionCount: 0,
      maximumExtensions: 3,
      canContinue: true,
      requiresExtension: true,
      pauseType: "budget" as const,
      currentModelId: null,
      modelSwitchCount: 0,
      canSwitchModel: false,
      availableModelConnections: []
    };

    it("重开只核对、不自动续跑：continuePlanning 一次都不许被调用", async () => {
      const continuePlanning = vi.fn();
      const continueSpy = continuePlanning;
      const result = await reinspectWpsPlanningPause(
        { turnId: "turn-reload", pausedAt },
        undefined,
        {
          getPlanningStatus: vi.fn(async () => ({
            turnId: "turn-reload",
            status: "waiting_user" as const,
            revision: 7,
            lastSequence: 15,
            canContinue: true,
            pauseType: "budget" as const,
            reason: null
          })),
          continuePlanning: continueSpy
        } as never
      );
      expect(continueSpy).not.toHaveBeenCalled();
      expect(result.view.action).toBe("continue_checkpoint");
      expect(result.view.canContinue).toBe(true);
    });

    it("服务端已推进 revision 时，按钮用的是**服务端**的现场而不是历史快照", async () => {
      const result = await reinspectWpsPlanningPause(
        { turnId: "turn-reload", pausedAt },
        undefined,
        {
          getPlanningStatus: vi.fn(async () => ({
            turnId: "turn-reload",
            status: "waiting_user" as const,
            revision: 42,
            lastSequence: 88,
            canContinue: true,
            pauseType: "budget" as const,
            reason: null
          }))
        } as never
      );
      expect(result.payload.revision).toBe(42);
      expect(result.payload.lastSequence).toBe(88);
      // 历史快照的 revision=3 绝不能泄漏成续跑凭据（用旧值续会被 CAS 拒）。
      expect(result.payload.revision).not.toBe(pausedAt.revision);
    });

    it("服务端已完成时不给继续按钮（重开不是'任务还没跑完'的证据）", async () => {
      const result = await reinspectWpsPlanningPause(
        { turnId: "turn-reload", pausedAt },
        undefined,
        {
          getPlanningStatus: vi.fn(async () => ({
            turnId: "turn-reload",
            status: "completed" as const,
            revision: 50,
            lastSequence: 99,
            result: { kind: "answer" as const, provider: "model" as const, message: "已完成" },
            canContinue: false,
            pauseType: null,
            reason: null
          }))
        } as never
      );
      expect(result.view.canContinue).toBe(false);
      expect(result.view.action).toBe("none");
    });

    it("服务端没有暂停成因时退回历史快照**仅供显示**，revision 仍以服务端为准", async () => {
      const result = await reinspectWpsPlanningPause(
        { turnId: "turn-reload", pausedAt },
        undefined,
        {
          getPlanningStatus: vi.fn(async () => ({
            turnId: "turn-reload",
            status: "interrupted" as const,
            revision: 11,
            lastSequence: 20,
            canContinue: true,
            pauseType: null,
            reason: null
          }))
        } as never
      );
      expect(result.payload.pauseType).toBe("budget");
      expect(result.payload.revision).toBe(11);
    });

    it("历史里没有暂停快照时也能核对（不因缺快照就整条放弃）", async () => {
      const result = await reinspectWpsPlanningPause(
        { turnId: "turn-reload", pausedAt: null },
        undefined,
        {
          getPlanningStatus: vi.fn(async () => ({
            turnId: "turn-reload",
            status: "waiting_user" as const,
            revision: 5,
            lastSequence: 10,
            canContinue: true,
            pauseType: "model_error" as const,
            reason: "模型连接中断"
          }))
        } as never
      );
      expect(result.payload.pauseType).toBe("model_error");
      expect(result.payload.reason).toBe("模型连接中断");
    });
  });
});
