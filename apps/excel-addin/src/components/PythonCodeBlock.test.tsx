// @vitest-environment jsdom
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { PythonCodeArtifact } from "@excel-bro/core/pythonCodeArtifact";
import PythonCodeBlock from "./PythonCodeBlock";

afterEach(() => cleanup());

function artifact(overrides: Partial<PythonCodeArtifact> = {}): PythonCodeArtifact {
  return {
    codeId: "code_py_0123456789ab",
    attemptId: "py_0123456789ab",
    stage: "analysis_completed",
    code: "",
    complete: true,
    truncated: false,
    gap: false,
    executedVersion: 1,
    version: 1,
    chars: 42,
    codeRetained: false,
    ...overrides
  };
}

describe("代码块只读且不假装", () => {
  it("没有任何能执行代码的按钮", () => {
    const { container } = render(
      <PythonCodeBlock artifact={artifact({ code: "x = 1", codeRetained: true })} />
    );
    const labels = Array.from(container.querySelectorAll("button")).map(
      (button) => button.textContent?.trim() ?? ""
    );
    // 允许的按钮集合只有这两个：本仓库没有"绕过校验直接跑"的入口。
    expect(labels.every((label) => label === "查看代码" || label === "复制" || label === "收起代码")).toBe(true);
    expect(labels.some((label) => /运行|执行/.test(label))).toBe(false);
  });

  it("代码作为文本节点渲染，不解析成元素", () => {
    const dangerous = 'payload = "<img src=x onerror=alert(1)>"';
    const { container } = render(
      <PythonCodeBlock artifact={artifact({ code: dangerous, codeRetained: true })} />
    );
    expect(container.querySelector(".python-code-body")?.textContent).toBe(dangerous);
    expect(container.querySelector("img")).toBeNull();
  });

  it("「已运行」只在执行版本与当前版本一致时出现", () => {
    const matching = render(
      <PythonCodeBlock artifact={artifact({ code: "x = 1", codeRetained: true })} />
    );
    expect(matching.getByText("已运行")).toBeTruthy();
    cleanup();

    const drifted = render(
      <PythonCodeBlock
        artifact={artifact({ code: "x = 1", codeRetained: true, version: 3, executedVersion: 1 })}
      />
    );
    expect(drifted.queryByText("已运行")).toBeNull();
  });
});

describe("正文按 codeId 取回", () => {
  it("带任务身份取回正文后展示真实代码", async () => {
    const loadBody = vi.fn(async () => ({
      status: "ready" as const,
      code: "import pandas as pd\ndf = pd.read_excel('a.xlsx')",
      chars: 46
    }));
    const { container } = render(
      <PythonCodeBlock artifact={artifact()} taskId="turn-abc" loadBody={loadBody} />
    );

    await waitFor(() =>
      expect(container.querySelector(".python-code-body")?.textContent).toContain("import pandas")
    );
    // 只有 codeId、没有任务身份取不回正文，因此两者都必须出现在请求里。
    expect(loadBody).toHaveBeenCalledWith({
      taskId: "turn-abc",
      codeId: "code_py_0123456789ab"
    });
    // 拿到正文后不能再挂着"正在取回/没有保留"的说明。
    expect(container.querySelector(".python-code-note")).toBeNull();
  });

  it("取不回来时明说没有保留，且不提供复制", async () => {
    const loadBody = vi.fn(async () => ({
      status: "unavailable" as const,
      reason: "missing"
    }));
    const { container } = render(
      <PythonCodeBlock artifact={artifact()} taskId="turn-abc" loadBody={loadBody} />
    );

    await waitFor(() =>
      expect(container.querySelector(".python-code-note")?.textContent).toContain("正文没有保留")
    );
    expect(container.querySelector(".python-code-copy")).toBeNull();
    expect(container.querySelector(".python-code-body")).toBeNull();
    // 状态仍然是真实的：分析确实完成过。
    expect(container.querySelector(".python-code-state")?.textContent).toBe("分析完成");
  });

  it("没有任务身份就不发请求（不拿 codeId 当授权）", () => {
    const loadBody = vi.fn(async () => ({ status: "unavailable" as const, reason: "missing" }));
    const { container } = render(
      <PythonCodeBlock artifact={artifact()} taskId={null} loadBody={loadBody} />
    );
    expect(loadBody).not.toHaveBeenCalled();
    expect(container.querySelector(".python-code-note")?.textContent).toContain("正文没有保留");
  });

  it("取回途中说正在取回，不说已经丢了", () => {
    const loadBody = vi.fn(() => new Promise<never>(() => {}));
    const { container } = render(
      <PythonCodeBlock artifact={artifact()} taskId="turn-abc" loadBody={loadBody} />
    );
    const note = container.querySelector(".python-code-note")?.textContent ?? "";
    expect(note).toContain("正在取回");
    expect(note).not.toContain("没有保留");
  });
});

describe("长代码折叠", () => {
  it("超过阈值默认折叠，展开后给出全文，收起后回到预览", () => {
    const code = Array.from({ length: 40 }, (_, index) => `row_${index} = ${index}`).join("\n");
    const { container } = render(
      <PythonCodeBlock artifact={artifact({ code, codeRetained: true })} />
    );
    const body = container.querySelector(".python-code-body");
    expect(body?.textContent?.split("\n")).toHaveLength(6);
    expect(container.querySelector(".python-code-more")?.textContent).toContain("未展开");

    fireEvent.click(screen.getByText("查看代码"));
    expect(container.querySelector(".python-code-body")?.textContent?.split("\n")).toHaveLength(40);

    fireEvent.click(screen.getByText("收起代码"));
    expect(container.querySelector(".python-code-body")?.textContent?.split("\n")).toHaveLength(6);
  });
});
