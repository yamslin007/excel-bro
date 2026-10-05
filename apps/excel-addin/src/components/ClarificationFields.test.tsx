// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";
import { assertAssistantResponse, assertIntentCheckResponse } from "../contracts";
import { ClarificationFields, clarificationFieldsAnswer } from "./ClarificationFields";

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
const fields = [{ id: "scope", label: "时间范围" }, { id: "unit", label: "统计单位", placeholder: "说明单位" }];
const response = { kind: "clarification", provider: "model", clarification: {
  id: "q1", summary: "补充统计口径", question: "请填写两个信息", reason: "口径尚未明确",
  scopeLabel: "当前已选表", options: [], fields
} };
const cleanups: (() => void)[] = [];
afterEach(() => { cleanups.splice(0).forEach((fn) => fn()); });

describe("structured clarification fields", () => {
  it("accepts the same form-only response in intent and planning contracts", () => {
    expect(() => assertIntentCheckResponse(response)).not.toThrow();
    expect(() => assertAssistantResponse(response)).not.toThrow();
  });
  it.each([
    { fields: [] },
    { fields: [{ id: "same", label: "a" }, { id: "same", label: "b" }] },
    { fields: [{ id: "x", label: " " }] }
  ])(
    "rejects empty, duplicate or invalid fields without usable options", ({ fields: invalid }) => {
      expect(() => assertAssistantResponse({ ...response, clarification: { ...response.clarification, fields: invalid } })).toThrow();
    }
  );
  it("does not submit a partial form; produces a labeled answer only after all fields are filled", async () => {
    const container = document.createElement("div");
    const root = createRoot(container);
    cleanups.push(() => act(() => root.unmount()));
    const onSubmit = vi.fn();
    const onChange = vi.fn();
    const render = async (values: Record<string, string>, busy = false) => act(async () => {
      root.render(<ClarificationFields fields={fields} values={values} busy={busy} onChange={onChange} onSubmit={onSubmit} />);
    });
    await render({ scope: "本月" });
    expect(container.querySelectorAll("textarea")).toHaveLength(2);
    expect(container.querySelector('textarea[aria-label="统计单位"]')).not.toBeNull();
    expect(container.querySelector("button")?.disabled).toBe(true);
    act(() => container.querySelector("button")?.click());
    expect(onSubmit).not.toHaveBeenCalled();
    await render({ scope: "本月", unit: "件" });
    act(() => container.querySelector("button")?.click());
    expect(onSubmit).toHaveBeenCalledWith("时间范围：本月\n统计单位：件");
    await render({ scope: "本月", unit: "件" }, true);
    expect([...container.querySelectorAll("textarea")].every((input) => input.disabled)).toBe(true);
    expect(container.querySelector("button")?.disabled).toBe(true);
    expect(clarificationFieldsAnswer(fields, { scope: " ", unit: "件" })).toBe("");
  });
});
