import { describe, expect, it } from "vitest";
import type { IntentCheckResponse } from "./contracts";
import {
  describeIntentDecision,
  resolveConfirmedPrompt
} from "./intent";

describe("resolveConfirmedPrompt", () => {
  it("replaces the original instruction for a replacement choice", () => {
    const result = resolveConfirmedPrompt(
      "按影片名称建表，表内按影院汇总",
      "按影院名称和编码建表，表内按影片汇总",
      "replace"
    );

    expect(result).toBe("按影院名称和编码建表，表内按影片汇总");
    expect(result).not.toContain("按影片名称建表");
  });

  it("keeps background context for an additive choice with explicit precedence", () => {
    const result = resolveConfirmedPrompt(
      "按影片名称建表，表内按影院汇总",
      "占比分母使用全表中该影院的记录数",
      "append"
    );

    expect(result).toContain("用户最终确认（最高优先级");
    expect(result).toContain("占比分母使用全表中该影院的记录数");
    expect(result).toContain("原始需求背景");
  });
});

describe("describeIntentDecision", () => {
  it("labels a greeting as direct conversation without a workbook task", () => {
    const decision = describeIntentDecision({
      kind: "conversation", provider: "model", message: "你好！"
    });
    expect(decision.label).toContain("普通对话");
    expect(decision.note).toContain("无需读取工作表数据");
  });
  it("keeps clarification reasoning concise without duplicating the question or options", () => {
    const response: IntentCheckResponse = {
      kind: "clarification",
      provider: "model",
      clarification: {
        id: "intent-1",
        summary: "拆分维度与命名维度不一致",
        question: "应该按哪个字段创建工作表？",
        reason: "两个字段会产生不同的输出结构",
        scopeLabel: "当前工作表：Sheet1",
        options: [
          {
            id: "movie",
            label: "按影片建表",
            description: "每部影片一张表",
            resolution: "按影片建表"
          },
          {
            id: "cinema",
            label: "按影院建表",
            description: "每个影院一张表",
            resolution: "按影院建表"
          }
        ]
      }
    };

    const decision = describeIntentDecision(response);

    expect(decision.note).toBe("拆分维度与命名维度不一致");
    expect(decision.detail).toBe(
      "判断依据：两个字段会产生不同的输出结构"
    );
    expect(decision.detail).not.toContain("应该按哪个字段");
    expect(decision.detail).not.toContain("按影片建表");
  });
});
