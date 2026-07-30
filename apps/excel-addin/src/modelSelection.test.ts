import { describe, expect, it } from "vitest";
import { chooseAvailableModel } from "./modelSelection";

const models = [
  { id: "local", available: true },
  { id: "connection:first", available: true },
  { id: "connection:new", available: true }
];

describe("model selection after settings changes", () => {
  it("selects a newly created connection ahead of local mode", () => {
    expect(
      chooseAvailableModel(
        models,
        "local",
        "connection:first",
        "connection:new"
      )
    ).toBe("connection:new");
  });

  it("keeps the current model when editing settings", () => {
    expect(
      chooseAvailableModel(models, "connection:first", "local")
    ).toBe("connection:first");
  });

  it("falls back when the current model was deleted", () => {
    expect(
      chooseAvailableModel(models, "connection:deleted", "local")
    ).toBe("local");
  });
});
