/**
 * 「澄清的答案」与「授权动作」不得互相冒充。
 *
 * 这两条是"重授权死循环"的两个直接来源：
 * - 把"重新确认数据范围"写进 `answer`，用户的真实回答就没了；
 * - 改答案只追加不替换，提示词里留下两条互相冲突的确认。
 */

import { describe, expect, it } from "vitest";

import {
  REAUTHORIZATION_NOTICE,
  clarificationAfterReauthorization,
  mergeKeyedConfirmation,
  renderConfirmedPrompt
} from "./clarificationAnswers";
import { withTaskResumeConfirmation, type TaskResumeRecord } from "./taskResume";

type Card = {
  status: "pending" | "resolved" | "invalidated" | "unconfirmed";
  answer?: string;
};

describe("重新授权不是答案", () => {
  it("没有真实答案时保持原状 —— 不伪造'已回答'", () => {
    const card: Card = { status: "pending" };
    expect(clarificationAfterReauthorization(card)).toEqual({
      status: "pending"
    });
  });

  it("有真实答案时保留答案本身，只把状态推进到已作答", () => {
    const card: Card = { status: "unconfirmed", answer: "空白保留原值" };
    const next = clarificationAfterReauthorization(card);
    expect(next?.answer).toBe("空白保留原值");
    expect(next?.status).toBe("resolved");
  });

  it("任何情况下都不会往 answer 里塞'已重新确认数据范围'这类套话", () => {
    const before: Card = { status: "pending", answer: "补充类别映射：丙→费用" };
    const after = clarificationAfterReauthorization(before);
    expect(after?.answer).toBe(before.answer);
    expect(after?.answer ?? "").not.toContain("重新确认");
  });

  it("没有卡片时原样返回 undefined", () => {
    expect(clarificationAfterReauthorization(undefined)).toBeUndefined();
  });

  it("说明文案里明确写着它是一次授权动作，不是回答", () => {
    expect(REAUTHORIZATION_NOTICE).toContain("不是对那道题的回答");
  });
});

describe("改答案必须替换，不能追加", () => {
  it("同一 questionId 的后答替换先答", () => {
    const merged = mergeKeyedConfirmation(
      [
        { questionId: "q1", label: "旧", value: "按甲处理" },
        { questionId: "q2", label: "别的问题", value: "按乙处理" }
      ],
      { questionId: "q1", label: "新", value: "按丙处理" }
    );
    expect(merged).toEqual([
      { questionId: "q1", label: "新", value: "按丙处理" },
      { questionId: "q2", label: "别的问题", value: "按乙处理" }
    ]);
  });

  it("不同 questionId 正常追加", () => {
    const merged = mergeKeyedConfirmation(
      [{ questionId: "q1", label: "a", value: "A" }],
      { questionId: "q2", label: "b", value: "B" }
    );
    expect(merged).toHaveLength(2);
  });

  it("没有 questionId 时保守追加 —— 猜错会丢掉一条真实答案", () => {
    const merged = mergeKeyedConfirmation(
      [{ label: "a", value: "A" }],
      { label: "b", value: "B" }
    );
    expect(merged).toHaveLength(2);
  });

  it("渲染出来的提示词里不会同时出现同一道题的两条冲突确认", () => {
    const rendered = renderConfirmedPrompt("原需求", [
      { questionId: "q1", label: "新", value: "按丙处理" }
    ]);
    expect(rendered).toBe("原需求\n补充确认：按丙处理");
  });

  it("没有确认时就是原需求本身，不追加空行", () => {
    expect(renderConfirmedPrompt("原需求", [])).toBe("原需求");
  });
});

describe("恢复记录里的确认也按 questionId 替换", () => {
  const record: TaskResumeRecord = {
    schemaVersion: 1,
    taskId: "task-1",
    phase: "awaiting_clarification",
    blocker: "user_input",
    resumability: "reauthorize",
    writeState: "not_written",
    originalPrompt: "把类别列翻译成归属",
    confirmations: [
      { label: "旧", value: "空白写未分类", questionId: "column-mapping:blankPolicyNeeded" }
    ],
    authorizationRevision: 1,
    createdAt: "2026-09-30T00:00:00.000Z",
    updatedAt: "2026-09-30T00:00:00.000Z"
  };

  it("同一道题再确认一次：只留下新答案", () => {
    const next = withTaskResumeConfirmation(record, {
      label: "新",
      value: "空白保留原值",
      questionId: "column-mapping:blankPolicyNeeded"
    });
    expect(next.confirmations).toHaveLength(1);
    expect(next.confirmations[0].value).toBe("空白保留原值");
  });

  it("另一道题的确认不会挤掉已有答案", () => {
    const next = withTaskResumeConfirmation(record, {
      label: "另一道",
      value: "补上丙→费用",
      questionId: "column-mapping:matchedDelimiterMissingCategories"
    });
    expect(next.confirmations).toHaveLength(2);
  });

  it("旧记录没有 questionId 时保守追加，不猜", () => {
    const legacy: TaskResumeRecord = {
      ...record,
      confirmations: [{ label: "旧", value: "空白写未分类" }]
    };
    const next = withTaskResumeConfirmation(legacy, {
      label: "新",
      value: "空白保留原值",
      questionId: "column-mapping:blankPolicyNeeded"
    });
    expect(next.confirmations).toHaveLength(2);
  });

  it("questionId 能穿过白名单消毒（否则替换会静默退化成追加）", () => {
    const next = withTaskResumeConfirmation(record, {
      label: "新",
      value: "空白保留原值",
      questionId: "column-mapping:blankPolicyNeeded"
    });
    expect(next.confirmations[0].questionId).toBe(
      "column-mapping:blankPolicyNeeded"
    );
  });
});
