import { describe, expect, it, vi } from "vitest";
import {
  EXECUTION_DISPOSITION_LABEL,
  executionDispositionTarget,
  normalizeExecutionDisposition,
  recordExecutionDisposition,
  type ExecutionDisposition
} from "./executionDisposition";

describe("执行处置留痕（放弃 / 撤销）", () => {
  it("只有两种取值，且不互相冒充", () => {
    expect(normalizeExecutionDisposition("abandoned")).toBe("abandoned");
    expect(normalizeExecutionDisposition("undone")).toBe("undone");
    for (const bad of ["confirmed", "abandon", "UNDONE", "", null, undefined, 1]) {
      expect(normalizeExecutionDisposition(bad)).toBeNull();
    }
    // 名称与审阅卡上的按钮同名，不另起一套叫法。
    expect(EXECUTION_DISPOSITION_LABEL.abandoned).toBe("放弃变更");
    expect(EXECUTION_DISPOSITION_LABEL.undone).toBe("撤销本次执行");
  });

  it("执行身份只去首尾空白，不做大小写或前缀归一", () => {
    expect(executionDispositionTarget("  execution-abc  ")).toBe("execution-abc");
    expect(executionDispositionTarget("   ")).toBeNull();
    expect(executionDispositionTarget("")).toBeNull();
    expect(executionDispositionTarget(null)).toBeNull();
    expect(executionDispositionTarget(undefined)).toBeNull();
    // 大小写是身份的一部分：账本按完整字符串比，改写会让一条记录被当成另一次执行。
    expect(executionDispositionTarget("Execution-ABC")).toBe("Execution-ABC");
  });

  it("上报成功返回 recorded，并把归一后的身份与取值交给端点", async () => {
    const send = vi.fn(async () => undefined);
    await expect(
      recordExecutionDisposition(send, "  execution-1 ", "undone")
    ).resolves.toBe("recorded");
    expect(send).toHaveBeenCalledTimes(1);
    expect(send).toHaveBeenCalledWith("execution-1", "undone");
  });

  it("没有执行身份就是 skipped：不去编一个 id，也不调用端点", async () => {
    const send = vi.fn(async () => undefined);
    for (const bad of ["", "   ", null, undefined]) {
      await expect(recordExecutionDisposition(send, bad, "abandoned")).resolves.toBe("skipped");
    }
    expect(send).not.toHaveBeenCalled();
  });

  it("端点失败只报 failed，绝不抛错 —— 恢复已经发生，留痕不许反悔", async () => {
    const send = vi.fn(async () => {
      throw new Error("网络不可用");
    });
    await expect(
      recordExecutionDisposition(send, "execution-2", "abandoned")
    ).resolves.toBe("failed");
  });

  it("两种取值各发各的，不因为先后顺序被改写", async () => {
    const seen: Array<[string, ExecutionDisposition]> = [];
    const send = async (id: string, kind: ExecutionDisposition) => {
      seen.push([id, kind]);
    };
    await recordExecutionDisposition(send, "execution-3", "abandoned");
    await recordExecutionDisposition(send, "execution-3", "undone");
    expect(seen).toEqual([
      ["execution-3", "abandoned"],
      ["execution-3", "undone"]
    ]);
  });
});
