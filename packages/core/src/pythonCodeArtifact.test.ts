import { describe, expect, it } from "vitest";
import {
  applyPythonCodeEvent,
  DEFAULT_PYTHON_CODE_LIMITS,
  PYTHON_CODE_STAGE_LABEL,
  latestPythonCode,
  pythonCodeCopyText,
  pythonCodeView,
  type PythonCodeArtifact
} from "./pythonCodeArtifact";

const limits = DEFAULT_PYTHON_CODE_LIMITS;

function build(
  events: Parameters<typeof applyPythonCodeEvent>[1][]
): PythonCodeArtifact {
  let artifact: PythonCodeArtifact | null = null;
  for (const event of events) {
    artifact = applyPythonCodeEvent(artifact, event, limits).artifact;
  }
  return artifact!;
}

describe("Python 代码工件的状态", () => {
  it("完整代码一次到达时状态是「代码已生成」，不是「正在生成」", () => {
    const artifact = build([
      {
        codeId: "code-1",
        attemptId: "try-1",
        stage: "generated",
        version: 1,
        code: "import pandas as pd\n",
        complete: true,
        sequence: 10
      }
    ]);
    expect(artifact.stage).toBe("generated");
    expect(artifact.complete).toBe(true);
    expect(pythonCodeView(artifact, limits)?.label).toBe("代码已生成");
  });

  it("确实有增量证据时才允许说「正在生成」，并按顺序拼接", () => {
    const artifact = build([
      { codeId: "c", attemptId: "t", stage: "generating", version: 1, delta: "import", sequence: 1 },
      { codeId: "c", attemptId: "t", stage: "generating", version: 2, delta: " pandas", sequence: 2 },
      { codeId: "c", attemptId: "t", stage: "generated", version: 3, delta: " as pd", complete: true, sequence: 3 }
    ]);
    expect(artifact.code).toBe("import pandas as pd");
    expect(artifact.complete).toBe(true);
    expect(artifact.stage).toBe("generated");
  });

  it("重放的增量不会把同一段代码粘两遍", () => {
    const artifact = build([
      { codeId: "c", attemptId: "t", stage: "generating", version: 1, delta: "df = 1\n", sequence: 5 },
      { codeId: "c", attemptId: "t", stage: "generating", version: 1, delta: "df = 1\n", sequence: 5 },
      { codeId: "c", attemptId: "t", stage: "generating", version: 2, delta: "df = 2\n", sequence: 6 }
    ]);
    expect(artifact.code).toBe("df = 1\ndf = 2\n");
  });

  it("更小的事件序号整帧丢弃：既不追加内容也不回退阶段", () => {
    const current = build([
      { codeId: "c", attemptId: "t", stage: "running", version: 3, code: "x = 1", sequence: 9 }
    ]);
    const result = applyPythonCodeEvent(
      current,
      { codeId: "c", attemptId: "t", stage: "generated", version: 1, delta: "old", sequence: 4 },
      limits
    );
    expect(result.ignored).toBe("stale");
    expect(result.artifact).toBe(current);
  });

  it("同版本同阶段的重放判为重放，内容不变", () => {
    const current = build([
      { codeId: "c", attemptId: "t", stage: "generated", version: 1, code: "x = 1", complete: true, sequence: 3 }
    ]);
    const result = applyPythonCodeEvent(
      current,
      { codeId: "c", attemptId: "t", stage: "generated", version: 1, code: "x = 1", complete: true, sequence: 3 },
      limits
    );
    expect(result.ignored).toBe("replay");
    expect(result.artifact.code).toBe("x = 1");
    expect(result.artifact.version).toBe(1);
  });

  it("新的生成尝试是重建工件，不伪装成逐字续写", () => {
    const first = build([
      { codeId: "c1", attemptId: "try-1", stage: "paused", version: 4, code: "x = 1", complete: false, sequence: 10 }
    ]);
    const second = applyPythonCodeEvent(
      first,
      { codeId: "c2", attemptId: "try-2", stage: "generated", version: 1, code: "y = 2", complete: true, sequence: 20 },
      limits
    );
    expect(second.artifact.attemptId).toBe("try-2");
    expect(second.artifact.code).toBe("y = 2");
    expect(second.artifact.version).toBe(1);
    expect(second.ignored).toBeNull();
  });

  it("失败分支是终态：后续的「进行中」事件不能把它拉回来", () => {
    const artifact = build([
      { codeId: "c", attemptId: "t", stage: "generated", version: 1, code: "x", complete: true, sequence: 1 },
      { codeId: "c", attemptId: "t", stage: "validation_failed", version: 2, sequence: 2 },
      { codeId: "c", attemptId: "t", stage: "running", version: 3, sequence: 3 }
    ]);
    expect(artifact.stage).toBe("validation_failed");
  });
});

describe("「已运行」必须指向确切版本", () => {
  it("执行阶段之前出现的 executedVersion 不被接受", () => {
    const result = applyPythonCodeEvent(
      null,
      {
        codeId: "c",
        attemptId: "t",
        stage: "generated",
        version: 1,
        code: "x = 1",
        complete: true,
        executedVersion: 1
      },
      limits
    );
    expect(result.artifact.executedVersion).toBeNull();
    expect(pythonCodeView(result.artifact, limits)?.executedCodeMatches).toBe(false);
  });

  it("执行版本与当前版本一致时才显示「已运行的就是这份代码」", () => {
    const artifact = build([
      { codeId: "c", attemptId: "t", stage: "generated", version: 1, code: "x = 1", complete: true, sequence: 1 },
      { codeId: "c", attemptId: "t", stage: "running", version: 1, executedVersion: 1, sequence: 2 },
      { codeId: "c", attemptId: "t", stage: "analysis_completed", version: 2, executedVersion: 1, sequence: 3 }
    ]);
    const view = pythonCodeView(artifact, limits)!;
    expect(view.executedCodeMatches).toBe(false);
    expect(view.label).toBe("分析完成");
  });

  it("同一版本跑完时确实匹配", () => {
    const artifact = build([
      { codeId: "c", attemptId: "t", stage: "generated", version: 1, code: "x = 1", complete: true, sequence: 1 },
      { codeId: "c", attemptId: "t", stage: "running", version: 1, executedVersion: 1, sequence: 2 },
      { codeId: "c", attemptId: "t", stage: "analysis_completed", version: 1, executedVersion: 1, sequence: 3 }
    ]);
    expect(pythonCodeView(artifact, limits)?.executedCodeMatches).toBe(true);
  });
});

describe("草稿、截断与复制", () => {
  it("未完成的代码被标成草稿，且没有可复制全文", () => {
    const artifact = build([
      { codeId: "c", attemptId: "t", stage: "generating", version: 1, delta: "x = 1", sequence: 1 }
    ]);
    const view = pythonCodeView(artifact, limits)!;
    expect(view.draft).toBe(true);
    expect(view.note).toContain("未完成草稿");
    expect(view.canCopyAll).toBe(false);
    expect(pythonCodeCopyText(artifact, limits)).toBeNull();
  });

  it("被截断的代码不提供「复制全部」", () => {
    const artifact = build([
      {
        codeId: "c",
        attemptId: "t",
        stage: "generated",
        version: 1,
        code: "a".repeat(limits.maxChars + 10),
        complete: true,
        sequence: 1
      }
    ]);
    const view = pythonCodeView(artifact, limits)!;
    expect(view.truncated).toBe(true);
    expect(view.canCopyAll).toBe(false);
    expect(pythonCodeCopyText(artifact, limits)).toBeNull();
  });

  it("完整未截断时「复制全部」返回真实全文", () => {
    const artifact = build([
      { codeId: "c", attemptId: "t", stage: "generated", version: 1, code: "x = 1\n", complete: true, sequence: 1 }
    ]);
    expect(pythonCodeCopyText(artifact, limits)).toBe("x = 1\n");
  });

  it("长代码默认折叠，短代码不折叠", () => {
    const long = build([
      {
        codeId: "c",
        attemptId: "t",
        stage: "generated",
        version: 1,
        code: Array.from({ length: 40 }, (_, index) => `v${index} = ${index}`).join("\n"),
        complete: true,
        sequence: 1
      }
    ]);
    expect(pythonCodeView(long, limits)?.collapsible).toBe(true);
    const short = build([
      { codeId: "c", attemptId: "t", stage: "generated", version: 1, code: "x = 1", complete: true, sequence: 1 }
    ]);
    expect(pythonCodeView(short, limits)?.collapsible).toBe(false);
  });

  it("没有代码时什么都不展示", () => {
    expect(pythonCodeView(null, limits)).toBeNull();
  });

  it("正文没有随事件下发时仍然展示状态，但不预先断言正文已丢", () => {
    const artifact = build([
      {
        codeId: "c",
        attemptId: "t",
        stage: "generating",
        version: 1,
        chars: 812,
        codeRetained: false,
        sequence: 1
      }
    ]);
    const view = pythonCodeView(artifact, limits)!;
    expect(view.hasCode).toBe(false);
    expect(view.chars).toBe(812);
    // 正文不进账本 ⇒ "事件里没有正文"是常态。真正的结论要么是取回成功，
    // 要么是取不回来；在试过之前不许说"正文没有保留"。
    expect(view.retrievable).toBe(true);
    expect(view.note).toBeUndefined();
    expect(view.canCopyAll).toBe(false);
  });
});

describe("挑选用哪份代码", () => {
  it("最后一条带工件的步骤胜出（多次分析只看最近一次）", () => {
    const first = build([
      { codeId: "c1", attemptId: "t1", stage: "generated", version: 1, code: "first", complete: true }
    ]);
    const second = build([
      { codeId: "c2", attemptId: "t2", stage: "generated", version: 1, code: "second", complete: true }
    ]);
    expect(latestPythonCode([{ pythonCode: first }, { pythonCode: second }])?.code).toBe("second");
    expect(latestPythonCode([{ pythonCode: null }, {}])).toBeNull();
  });
});

describe("增量的判重按事件身份，不按文本", () => {
  const step = (
    version: number,
    sequence: number,
    delta: string,
    extra: Record<string, unknown> = {}
  ) => ({
    codeId: "c",
    attemptId: "t",
    stage: "generating" as const,
    version,
    sequence,
    delta,
    ...extra
  });

  it("连续右括号：第二个 ')' 是真实内容，不能被当重复丢掉", () => {
    const artifact = build([
      step(1, 1, "f(g(1"),
      step(2, 2, ")"),
      step(3, 3, ")", { complete: true, stage: "generated" })
    ]);
    expect(artifact.code).toBe("f(g(1))");
    expect(artifact.complete).toBe(true);
    expect(artifact.gap).toBe(false);
    expect(pythonCodeCopyText(artifact, limits)).toBe("f(g(1))");
  });

  it("连续换行与重复缩进都按片段保留，不按后缀去重", () => {
    const artifact = build([
      step(1, 1, "def f():\n"),
      step(2, 2, "\n"),
      step(3, 3, "    return 1\n"),
      step(4, 4, "    return 1\n", { complete: true, stage: "generated" })
    ]);
    expect(artifact.code).toBe("def f():\n\n    return 1\n    return 1\n");
    expect(artifact.complete).toBe(true);
  });

  it("内容相同但身份不同 = 两条真实片段，都要保留", () => {
    const artifact = build([step(1, 1, "pass\n"), step(2, 2, "pass\n")]);
    expect(artifact.code).toBe("pass\npass\n");
  });

  it("同一条事件重放（同序号）不重复追加", () => {
    const artifact = build([
      step(1, 1, "x = 1\n"),
      step(1, 1, "x = 1\n"),
      step(1, 1, "x = 1\n")
    ]);
    expect(artifact.code).toBe("x = 1\n");
    expect(pythonCodeView(artifact, limits)?.canCopyAll).toBe(false);
    expect(artifact.complete).toBe(false);
  });

  it("低序号的迟到帧整帧丢弃（内容与其后的阶段都不动）", () => {
    const artifact = build([
      step(1, 5, "a"),
      step(2, 6, "b"),
      step(3, 4, "SHOULD_NOT_APPEAR")
    ]);
    expect(artifact.code).toBe("ab");
  });

  it("缺片时标记 gap、压掉 complete、不提供复制全部", () => {
    const artifact = build([
      step(1, 1, "f(g(1"),
      // 版本从 1 跳到 3：中间那一版没收到。
      step(3, 2, ")", { complete: true, stage: "generated" })
    ]);
    expect(artifact.code).toBe("f(g(1)");
    expect(artifact.gap).toBe(true);
    expect(artifact.complete).toBe(false);
    const view = pythonCodeView(artifact, limits)!;
    expect(view.canCopyAll).toBe(false);
    expect(view.gap).toBe(true);
    expect(view.note).toContain("没有收齐");
    expect(pythonCodeCopyText(artifact, limits)).toBeNull();
  });

  it("完整快照能补齐缺片并恢复完整", () => {
    const artifact = build([
      step(1, 1, "f(g(1"),
      step(3, 2, ")"),
      { codeId: "c", attemptId: "t", stage: "generated", version: 4, sequence: 3,
        code: "f(g(1))", complete: true }
    ]);
    expect(artifact.code).toBe("f(g(1))");
    expect(artifact.gap).toBe(false);
    expect(artifact.complete).toBe(true);
    expect(pythonCodeCopyText(artifact, limits)).toBe("f(g(1))");
  });

  it("完整快照整体替换，不在旧文本后面再拼一遍", () => {
    const artifact = build([
      { codeId: "c", attemptId: "t", stage: "generating", version: 1, sequence: 1, delta: "old" },
      { codeId: "c", attemptId: "t", stage: "generated", version: 2, sequence: 2,
        code: "new", complete: true }
    ]);
    expect(artifact.code).toBe("new");
  });

  it("旧尝试的迟到事件不能覆盖当前尝试", () => {
    const artifact = build([
      { codeId: "c1", attemptId: "old", stage: "generated", version: 1, sequence: 10,
        code: "OLD_CODE", complete: true },
      { codeId: "c2", attemptId: "new", stage: "generating", version: 1, sequence: 20,
        delta: "new_code" },
      // 旧尝试迟到的帧：序号更小 ⇒ 整帧丢弃，不许把工件退回旧尝试。
      { codeId: "c1", attemptId: "old", stage: "generated", version: 2, sequence: 11,
        code: "OLD_CODE_V2", complete: true }
    ]);
    expect(artifact.attemptId).toBe("new");
    expect(artifact.code).toBe("new_code");
  });

  it("新尝试（序号更大）照常重建工件", () => {
    const artifact = build([
      { codeId: "c1", attemptId: "old", stage: "generated", version: 1, sequence: 10,
        code: "OLD", complete: true },
      { codeId: "c2", attemptId: "new", stage: "generated", version: 1, sequence: 30,
        code: "NEW", complete: true }
    ]);
    expect(artifact.attemptId).toBe("new");
    expect(artifact.code).toBe("NEW");
    expect(artifact.complete).toBe(true);
  });

  it("增量首帧就是版本 2（前面丢了）直接判缺片，不宣告完整", () => {
    const artifact = build([
      { codeId: "c", attemptId: "t", stage: "generating", version: 2, sequence: 1,
        delta: "rest", complete: true }
    ]);
    expect(artifact.gap).toBe(true);
    expect(artifact.complete).toBe(false);
    expect(pythonCodeView(artifact, limits)?.note).toContain("没有收齐");
  });
});

describe("未开始执行：不显示「已运行」，原因要落地", () => {
  it("generated → not_started（同版本）也能推进，并带出具体原因", () => {
    const artifact = build([
      { codeId: "c", attemptId: "t", stage: "generated", version: 1,
        complete: true, sequence: 1, chars: 20 },
      { codeId: "c", attemptId: "t", stage: "not_started", version: 1, sequence: 2,
        reason: "Python 隔离执行没有开启；这段代码没有运行。" }
    ]);
    expect(artifact.stage).toBe("not_started");
    expect(PYTHON_CODE_STAGE_LABEL[artifact.stage]).toBe("未确认执行");
    // 一次执行器都没起 ⇒ 没有执行版本。
    expect(artifact.executedVersion).toBeNull();
    const view = pythonCodeView(artifact, limits);
    expect(view?.note).toContain("没有开启");
    expect(view?.canCopyAll).toBe(false);
  });

  it("没有原因时也如实说「没有开始执行」，不套用「执行失败」", () => {
    const artifact = build([
      { codeId: "c", attemptId: "t", stage: "generated", version: 1,
        complete: true, sequence: 1, chars: 20 },
      { codeId: "c", attemptId: "t", stage: "not_started", version: 1, sequence: 2 }
    ]);
    const view = pythonCodeView(artifact, limits);
    expect(view?.note).toContain("未收到代码开始执行的确认");
    expect(view?.note).not.toContain("执行没有成功");
  });

  it("迟到的 running 不能把「未开始执行」改回「正在运行」", () => {
    const events: Parameters<typeof applyPythonCodeEvent>[1][] = [
      { codeId: "c", attemptId: "t", stage: "generated", version: 1,
        complete: true, sequence: 1, chars: 20 },
      { codeId: "c", attemptId: "t", stage: "not_started", version: 1, sequence: 2,
        reason: "运行器不可用" },
      // 时序错乱的一帧：它没有真实启动证据。
      { codeId: "c", attemptId: "t", stage: "running", version: 1, sequence: 3,
        executedVersion: 1 }
    ];
    const artifact = build(events);
    expect(artifact.stage).toBe("not_started");
    expect(artifact.executedVersion).toBeNull();
  });
});
