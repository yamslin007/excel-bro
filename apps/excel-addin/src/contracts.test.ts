import { describe, expect, it } from "vitest";
import {
  assertAssistantResponse,
  assertIntentCheckResponse,
  parseTaskProgressEvent
} from "./contracts";

describe("parseTaskProgressEvent", () => {
  it("accepts structured progress and keeps auditable trace metadata", () => {
    const event = parseTaskProgressEvent({
      schemaVersion: 1,
      operationId: "execution-test-1",
      executionId: "0123456789abcdef0123456789abcdef",
      sequence: 3,
      phase: "verifying",
      stepId: "execution.verify",
      status: "warning",
      title: "验证发现不一致",
      summary: "2 / 3 项符合预期",
      current: 2,
      total: 3,
      unit: "checks",
      trace: {
        kind: "conflict",
        summary: "1 项结果与预期不一致",
        basis: "执行后重新读取"
      },
      timestamp: "2026-09-02T00:00:00Z"
    });

    expect(event).toMatchObject({
      operationId: "execution-test-1",
      sequence: 3,
      status: "warning",
      current: 2,
      total: 3,
      trace: { kind: "conflict" }
    });
  });

  it("converts legacy planning steps without breaking older servers", () => {
    expect(
      parseTaskProgressEvent(
        { phase: "planning", title: "正在查看工作簿结构" },
        "turn-legacy-test",
        7
      )
    ).toMatchObject({
      schemaVersion: 1,
      operationId: "turn-legacy-test",
      sequence: 7,
      phase: "planning"
    });
  });

  it("rejects impossible counters", () => {
    expect(
      parseTaskProgressEvent({
        schemaVersion: 1,
        operationId: "execution-test-2",
        sequence: 1,
        phase: "executing",
        stepId: "execution.action.1",
        status: "progress",
        title: "执行中",
        current: 2,
        total: 1,
        unit: "actions",
        timestamp: "2026-09-02T00:00:00Z"
      })
    ).toBeNull();
  });

  it("keeps safe tool facts and accepts the cancelled terminal status", () => {
    const event = parseTaskProgressEvent({
      schemaVersion: 1,
      operationId: "execution-cancelled",
      sequence: 4,
      phase: "planning",
      stepId: "planning.tool.1.1",
      status: "cancelled",
      title: "任务已取消",
      toolName: "read_range",
      facts: [
        { label: "范围", value: "A1:B3", style: "code" },
        { label: "结果", value: "3 行 × 2 列" }
      ],
      timestamp: "2026-09-02T00:00:00Z"
    });

    expect(event).toMatchObject({
      status: "cancelled",
      toolName: "read_range",
      facts: [
        { label: "范围", value: "A1:B3", style: "code" },
        { label: "结果", value: "3 行 × 2 列", style: "text" }
      ]
    });
  });

  it("limits facts and redacts suspicious values before rendering", () => {
    const event = parseTaskProgressEvent({
      schemaVersion: 1,
      operationId: "execution-safety",
      sequence: 1,
      phase: "planning",
      stepId: "planning.tool.safe",
      status: "completed",
      title: "已完成",
      facts: [
        { label: "结果", value: "Bearer very-secret-token-value" },
        { label: "apiKey", value: "plain-secret" },
        { label: "未授权标签", value: "不应展示" },
        ...Array.from({ length: 20 }, (_, index) => ({
          label: "结果",
          value: index === 0 ? "x".repeat(5000) : String(index)
        }))
      ],
      target: {
        file: "C:\\Users\\someone\\Desktop\\敏感目录\\报表.xlsx",
        sheet: "Sheet1"
      },
      timestamp: "2026-09-02T00:00:00Z"
    });

    expect(event?.facts).toHaveLength(11);
    expect(event?.facts?.[0].value).toBe("Bearer 已隐藏");
    expect(event?.facts?.[1].value).toBe("已隐藏");
    expect(event?.facts?.some((fact) => fact.value === "不应展示")).toBe(false);
    expect(event?.facts?.[2].value.length).toBeLessThanOrEqual(500);
    expect(event?.facts?.[2].value).toContain("内容过长，已截断");
    expect(event?.target?.file).toBe("报表.xlsx");
  });
});

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

  it("accepts a system degradation answer", () => {
    const response = {
      kind: "answer",
      provider: "system",
      message: "规划达到预算上限，请简化需求后重试。"
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
  it("accepts a direct conversational reply and rejects an empty one", () => {
    expect(() => assertIntentCheckResponse({ kind: "conversation", provider: "model", message: "你好！" })).not.toThrow();
    expect(() => assertIntentCheckResponse({ kind: "conversation", provider: "model", message: "  " })).toThrow("对话回复无效");
  });
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
            resolution: "按电影重新汇总",
            resolutionMode: "replace"
          }
        ]
      }
    };

    expect(() => assertIntentCheckResponse(response)).not.toThrow();
  });

  it("rejects an unknown clarification resolution mode", () => {
    const response = {
      kind: "clarification",
      provider: "model",
      clarification: {
        id: "intent-invalid-mode",
        summary: "确认拆分方式",
        question: "按什么建表？",
        reason: "维度冲突",
        scopeLabel: "当前工作表",
        options: [
          {
            id: "one",
            label: "按影院",
            description: "改为按影院建表",
            resolution: "按影院建表，表内按影片汇总",
            resolutionMode: "overwrite"
          },
          {
            id: "two",
            label: "按影片",
            description: "保留按影片建表",
            resolution: "按影片建表，表内按影院汇总"
          }
        ]
      }
    };

    expect(() => assertIntentCheckResponse(response)).toThrow("确认问题无效");
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
