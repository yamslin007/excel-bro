import { describe, expect, it } from "vitest";
import {
  clarificationAllowsMultiple,
  clarificationBindingKey,
  clarificationSubmission,
  clarificationSubmissionAllowed,
  draftBelongsTo,
  emptyClarificationDraft,
  toggleClarificationOption
} from "./clarificationComposer";

const binding = { conversationId: "c1", messageId: "m1", clarificationId: "q1" };
const options = [
  { id: "o1", label: "按分类汇总", resolution: "按分类汇总后写出结果" },
  { id: "o2", label: "按月份汇总", resolution: "按月份汇总后写出结果", resolutionMode: "replace" as const }
];

describe("问题身份与草稿归属", () => {
  it("三段都在身份里：换对话/换消息/换问题任一变化都算另一道题", () => {
    const key = clarificationBindingKey(binding);
    expect(clarificationBindingKey({ ...binding })).toBe(key);
    expect(clarificationBindingKey({ ...binding, conversationId: "c2" })).not.toBe(key);
    expect(clarificationBindingKey({ ...binding, messageId: "m2" })).not.toBe(key);
    expect(clarificationBindingKey({ ...binding, clarificationId: "q2" })).not.toBe(key);
  });

  it("草稿只属于它自己的那道题；切任务/切对话后作废", () => {
    const draft = emptyClarificationDraft(binding);
    expect(draftBelongsTo(draft, { ...binding })).toBe(true);
    expect(draftBelongsTo(draft, { ...binding, conversationId: "other" })).toBe(false);
    expect(draftBelongsTo(draft, { ...binding, messageId: "another-message" })).toBe(false);
    expect(draftBelongsTo(draft, null)).toBe(false);
    expect(draftBelongsTo(null, binding)).toBe(false);
  });
});

describe("toggleClarificationOption", () => {
  it("单选：换成它，再点同一个就取消", () => {
    expect(toggleClarificationOption([], "o1")).toEqual(["o1"]);
    expect(toggleClarificationOption(["o1"], "o2")).toEqual(["o2"]);
    expect(toggleClarificationOption(["o1"], "o1")).toEqual([]);
  });

  it("多选：切换并保留点击顺序（不排序）", () => {
    let selected = toggleClarificationOption([], "o2", true);
    selected = toggleClarificationOption(selected, "o1", true);
    expect(selected).toEqual(["o2", "o1"]);
    expect(toggleClarificationOption(selected, "o2", true)).toEqual(["o1"]);
  });

  it("空 id 不改变任何东西", () => {
    expect(toggleClarificationOption(["o1"], "   ")).toEqual(["o1"]);
  });
});

describe("clarificationAllowsMultiple", () => {
  it("协议没有多选标记 ⇒ 恒为 false，不按选项数量猜", () => {
    expect(clarificationAllowsMultiple({})).toBe(false);
    expect(clarificationAllowsMultiple(null)).toBe(false);
    // 三个选项也不代表多选：点两个互斥选项被拼成一个答案是更糟的错误。
    expect(clarificationAllowsMultiple({ multiSelect: 3 as unknown })).toBe(false);
    // 只有协议显式给了标记才放开。
    expect(clarificationAllowsMultiple({ multiSelect: true })).toBe(true);
  });
});

describe("clarificationSubmission", () => {
  it("什么都没选、什么都没写 ⇒ 不许提交", () => {
    const result = clarificationSubmission({ options, selected: [], freeText: "  " });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toContain("请选择一项");
  });

  it("选了选项 ⇒ 结构化提交，答案用标签、resolution 原样带出去", () => {
    const result = clarificationSubmission({ options, selected: ["o2"], freeText: "" });
    expect(result.ok).toBe(true);
    if (result.ok && result.kind === "option") {
      expect(result.options.map((option) => option.id)).toEqual(["o2"]);
      expect(result.options[0].resolutionMode).toBe("replace");
      expect(result.answer).toBe("按月份汇总");
    }
  });

  it("选项 + 补充说明：两句都在，补充不被丢掉", () => {
    const result = clarificationSubmission({
      options, selected: ["o1"], freeText: "只要 2024 年的"
    });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.answer).toBe("按分类汇总\n补充：只要 2024 年的");
  });

  it("只写自由输入 ⇒ 原文提交，不当成选项", () => {
    const result = clarificationSubmission({ options, selected: [], freeText: "  按季度汇总  " });
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.kind).toBe("free");
      expect(result.answer).toBe("按季度汇总");
    }
  });

  it("选中了不存在的选项 id ⇒ 当成没选（不按标签猜）", () => {
    const result = clarificationSubmission({ options, selected: ["ghost"], freeText: "" });
    expect(result.ok).toBe(false);
  });

  it("多字段有必填没填 ⇒ 不许提交，并指名缺哪几项", () => {
    const fields = [{ id: "f1", label: "汇总口径" }, { id: "f2", label: "舍入方式" }];
    const result = clarificationSubmission({
      fields, fieldValues: { f1: "按分类" }
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.missing).toEqual(["f2"]);
      expect(result.reason).toContain("舍入方式");
    }
  });

  it("只打了空格 ⇒ 仍然算没填（空白不等于答案）", () => {
    const fields = [{ id: "f1", label: "汇总口径" }, { id: "f2", label: "舍入方式" }];
    const result = clarificationSubmission({ fields, fieldValues: { f1: " ", f2: "四舍五入" } });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.missing).toEqual(["f1"]);
  });

  it("多字段填完 ⇒ 走字段格式，不受选项影响", () => {
    const fields = [{ id: "f1", label: "汇总口径" }];
    const result = clarificationSubmission({
      fields, fieldValues: { f1: "按分类" }, options, selected: ["o1"]
    });
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.kind).toBe("fields");
      expect(result.answer).toBe("汇总口径：按分类");
    }
  });
});

describe("clarificationSubmissionAllowed", () => {
  const draft = emptyClarificationDraft(binding);

  it("绑定一致且仍待回答 ⇒ 允许", () => {
    expect(clarificationSubmissionAllowed({ draft, binding, status: "pending" })).toEqual({ ok: true });
  });

  it("已经处理过 ⇒ 拒绝（旧问题不得重复提交）", () => {
    for (const status of ["resolved", "cancelled", "resolving"]) {
      const result = clarificationSubmissionAllowed({ draft, binding, status });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.reason).toContain("已经处理过");
    }
  });

  it("草稿属于另一道题 ⇒ 拒绝，不刷新成新题继续用", () => {
    const result = clarificationSubmissionAllowed({
      draft, binding: { ...binding, messageId: "m2" }, status: "pending"
    });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toContain("不属于当前任务");
  });

  it("没有待答题 / 正在忙 ⇒ 拒绝", () => {
    expect(clarificationSubmissionAllowed({ draft, binding: null, status: "pending" }).ok).toBe(false);
    expect(clarificationSubmissionAllowed({ draft, binding, status: "pending", busy: true }).ok).toBe(false);
  });
});
