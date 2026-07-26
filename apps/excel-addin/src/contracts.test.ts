import { describe, expect, it } from "vitest";
import {
  assertAssistantResponse,
  assertIntentCheckResponse
} from "./contracts";

describe("assertAssistantResponse", () => {
  it("accepts a safe plan", () => {
    const response = {
      kind: "plan",
      provider: "local",
      plan: {
        id: "plan-1",
        title: "测试",
        summary: "测试计划",
        assumptions: [],
        warnings: [],
        actions: [{ type: "createWorksheet", sheet: "分析结果" }]
      }
    };

    expect(() => assertAssistantResponse(response)).not.toThrow();
  });

  it("accepts a deterministic split and aggregate plan", () => {
    const response = {
      kind: "plan",
      provider: "model",
      plan: {
        id: "split-plan",
        title: "按影片拆分",
        summary: "按影院汇总影片记录数和占比",
        assumptions: [],
        warnings: [],
        actions: [
          {
            type: "splitGroupAggregate",
            sheet: "sheet-dc1",
            splitBy: "影片名称",
            groupBy: ["影院名称", "影院编码"],
            metrics: [
              {
                operation: "countRows",
                outputName: "计数项:影厅",
                ratioOutputName: "占比"
              }
            ],
            includeBlankSplitValues: false,
            existingSheetPolicy: "rename",
            maxOutputSheets: 200
          }
        ]
      }
    };

    expect(() => assertAssistantResponse(response)).not.toThrow();
  });

  it("accepts a direct answer", () => {
    const response = {
      kind: "answer",
      provider: "local",
      message: "得分的平均值是 24.67。",
      resultContext: {
        kind: "table",
        title: "平均值",
        headers: ["字段", "指标", "结果", "有效数值数"],
        rows: [["得分", "平均值", 24.6667, 3]],
        primaryValueColumn: 2,
        sourceSheets: ["Sheet1"],
        warnings: []
      }
    };

    expect(() => assertAssistantResponse(response)).not.toThrow();
  });

  it("rejects an unknown action", () => {
    const response = {
      kind: "plan",
      provider: "model",
      plan: {
        actions: [{ type: "runMacro", sheet: "原始数据" }]
      }
    };

    expect(() => assertAssistantResponse(response)).toThrow("未授权动作");
  });

  it("rejects an unknown acceptance criterion", () => {
    const response = {
      kind: "plan",
      provider: "model",
      plan: {
        actions: [{ type: "createWorksheet", sheet: "结果" }],
        acceptanceCriteria: [{ type: "runMacro", sheet: "结果" }]
      }
    };

    expect(() => assertAssistantResponse(response)).toThrow("未知验收条件");
  });

  it("rejects a range check whose dimensions do not match", () => {
    const response = {
      kind: "plan",
      provider: "model",
      plan: {
        actions: [{ type: "createWorksheet", sheet: "结果" }],
        acceptanceCriteria: [
          {
            type: "rangeEquals",
            sheet: "结果",
            range: "A1:Z100",
            expected: [[1]]
          }
        ]
      }
    };

    expect(() => assertAssistantResponse(response)).toThrow("尺寸不一致");
  });
});

describe("assertIntentCheckResponse", () => {
  it("accepts a clarification with mutually exclusive options", () => {
    const response = {
      kind: "clarification",
      provider: "local",
      clarification: {
        id: "intent-1",
        summary: "比较多张表中的占比",
        question: "使用哪种口径？",
        reason: "不同口径结果不同",
        scopeLabel: "已选择 3 张工作表",
        options: [
          {
            id: "existing",
            label: "比较现有值",
            description: "直接比较占比列",
            resolution: "比较现有占比值"
          },
          {
            id: "aggregate",
            label: "重新汇总",
            description: "按电影重新计算",
            resolution: "按电影重新汇总"
          }
        ]
      }
    };

    expect(() => assertIntentCheckResponse(response)).not.toThrow();
  });

  it("rejects a clarification with fewer than two options", () => {
    const response = {
      kind: "clarification",
      provider: "model",
      clarification: {
        question: "使用哪种口径？",
        options: [
          {
            id: "only",
            label: "唯一选项",
            description: "没有选择意义",
            resolution: "继续"
          }
        ]
      }
    };

    expect(() => assertIntentCheckResponse(response)).toThrow(
      "确认问题无效"
    );
  });

  it("accepts a high-level local data tool request", () => {
    const response = {
      kind: "tool_request",
      provider: "model",
      summary: "按电影汇总影厅数并计算占比",
      confirmedPrompt: "在已选工作表中按电影汇总",
      request: {
        id: "query-1",
        tool: "query_table",
        arguments: {
          mode: "aggregate",
          scope: "selected",
          groupBy: ["影片名称"],
          metrics: [
            {
              operation: "sum",
              field: "计数项:影厅",
              outputName: "影厅数",
              ratioOutputName: "占比"
            }
          ],
          sortBy: "占比",
          sortDirection: "desc",
          limit: 10
        }
      }
    };

    expect(() => assertIntentCheckResponse(response)).not.toThrow();
  });
});
