import { describe, expect, it } from "vitest";
import { modelUsageNotice, type ModelUsageErrorCode } from "./modelUsageNotice";

describe("shared model usage notices", () => {
  it.each([
    ["MODEL_QUOTA_EXHAUSTED", "余额不足", "模型设置"],
    ["MODEL_CONTEXT_LIMIT", "上下文长度", "缩短对话"],
    ["MODEL_OUTPUT_LIMIT", "未获得完整结果", "拆分任务"],
    ["MODEL_BUDGET_LIMIT", "模型费用上限", "预算设置"],
    ["MODEL_TURN_LIMIT", "调用轮数上限", "拆分任务"]
  ] as [ModelUsageErrorCode, string, string][])("%s explains cause, action and write evidence", (code, cause, action) => {
    const notice = modelUsageNotice(code);
    expect(notice).toContain(cause);
    expect(notice).toContain(action);
    expect(notice).toContain("以执行结果卡为准");
    expect(notice).not.toMatch(/已撤销|未修改工作表|写入失败/);
  });
  it.each(["MODEL_OVERLOADED", "MODEL_HTTP_ERROR", "toString"])("does not invent quota evidence for %s", (code) => {
    expect(modelUsageNotice(code)).toBeNull();
  });
});
