// @vitest-environment jsdom
import { cleanup, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import type { ModelOption } from "../contracts";
import { ModelPickerLabel } from "./ModelPickerLabel";

afterEach(cleanup);
const option = (modelId?: string): ModelOption => ({
  id: "connection:demo", label: "我的连接 · deepseek-v4-flash", modelId,
  provider: "model", available: true, supportsVision: false,
});
describe("model toolbar label", () => {
  it("only uses 模型 for an unconfigured selection", () => {
    const { rerender } = render(<ModelPickerLabel model={null} />);
    expect(screen.getByText("模型")).toBeTruthy();
    rerender(<ModelPickerLabel model={option("deepseek-v4-flash")} />);
    expect(screen.getByText("deepseek-v4-flash").title).toBe("deepseek-v4-flash");
    expect(screen.queryByText("模型")).toBeNull();
  });
  it.each([
    ["deepseek-v4-flash", "DeepSeek"], ["deepseek-ai/DeepSeek-V4-Flash", "DeepSeek"],
    ["gpt-4.1", "OpenAI"], ["o3", "OpenAI"], ["claude-sonnet-4", "Anthropic"],
    ["gemini-2.5-pro", "Google"], ["Qwen/Qwen3-32B", "Qwen"],
    ["kimi-k2", "Moonshot"], ["glm-4.5", "Zhipu"],
    ["mistral-small", "Mistral"], ["meta-llama/Llama-3.3-70B", "Meta"], ["grok-3", "xAI"],
  ])("uses the vendor logo for %s", (id, brand) => {
    const { container } = render(<ModelPickerLabel model={option(id)} />);
    expect(container.querySelector("[data-model-brand]")?.getAttribute("data-model-brand")).toBe(brand);
  });
  it("does not infer a vendor from a nickname or hide a configured connection when metadata is unavailable", () => {
    const { container, rerender } = render(<ModelPickerLabel model={option()} />);
    expect(screen.getByText("我的连接 · deepseek-v4-flash")).toBeTruthy();
    expect(container.querySelector("[data-model-brand]")).toBeNull();
    rerender(<ModelPickerLabel model={option("custom-model")} />);
    expect(screen.getByText("custom-model")).toBeTruthy();
    expect(container.querySelector("[data-model-brand]")).toBeNull();
  });
});
