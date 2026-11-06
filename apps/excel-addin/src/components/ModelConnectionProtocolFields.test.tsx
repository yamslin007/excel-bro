// @vitest-environment jsdom
import { useState } from "react";
import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, expect, it } from "vitest";
import { ModelConnectionProtocolFields } from "./ModelConnectionProtocolFields";
import type { ModelConnectionDraft } from "../hooks/useModelManagement";
import { emptyModelConnectionDraft } from "../utils";

afterEach(cleanup);

it("offers one protocol field, keeps the draft identity and has no usage gate", () => {
  function Harness() {
    const [draft, setDraft] = useState<ModelConnectionDraft | null>({ ...emptyModelConnectionDraft(),
      id: "saved-connection", modelId: "same-model", baseUrl: "https://provider.test", apiKey: "synthetic-key" });
    return <><ModelConnectionProtocolFields draft={draft!} onChange={setDraft} />
      <output>{JSON.stringify(draft)}</output></>;
  }
  const view = render(<Harness />);
  expect(screen.getAllByRole("combobox")).toHaveLength(1);
  expect(screen.queryByText("测试用途")).toBeNull();
  for (const protocol of ["anthropic", "openai_compatible"]) {
    fireEvent.change(screen.getByRole("combobox"), { target: { value: protocol } });
    const draft = JSON.parse(view.container.querySelector("output")!.textContent!);
    expect(draft).toMatchObject({ protocol, id: "saved-connection", modelId: "same-model", apiKey: "synthetic-key" });
    expect(draft).not.toHaveProperty("purpose");
  }
});
