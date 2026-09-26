import { describe, expect, it } from "vitest";
import {
  chooseAvailableModel,
  connectionModelSuffix,
  isRealModelId,
  LEGACY_LOCAL_MODEL_ID,
  NO_MODEL_SELECTED,
  resolveDisplayedModelId,
  shouldPersistSelection
} from "./modelSelection";

describe("connectionModelSuffix", () => {
  it("numbers equal model IDs without changing their unique connection IDs", () => {
    const connections = [
      { id: "excel", modelId: "deepseek-chat" },
      { id: "other", modelId: "claude" },
      { id: "wps", modelId: "DeepSeek-Chat" }
    ];
    expect(connectionModelSuffix(connections, "excel")).toBe(" (1)");
    expect(connectionModelSuffix(connections, "wps")).toBe(" (2)");
    expect(connectionModelSuffix(connections, "other")).toBe("");
  });
});

const models = [
  { id: "connection:first", available: true },
  { id: "connection:new", available: true }
];

const withLegacyPlaceholder = [
  { id: LEGACY_LOCAL_MODEL_ID, available: true },
  ...models
];

describe("chooseAvailableModel", () => {
  it("prefers an explicitly requested valid model", () => {
    expect(
      chooseAvailableModel(models, "connection:first", "connection:first", "connection:new")
    ).toBe("connection:new");
  });

  it("keeps the current model when it is still valid", () => {
    expect(chooseAvailableModel(models, "connection:first", "connection:new")).toBe(
      "connection:first"
    );
  });

  it("falls back to the server default when the current model was deleted", () => {
    expect(
      chooseAvailableModel(models, "connection:deleted", "connection:new")
    ).toBe("connection:new");
  });

  it("falls back to the first usable model when no default is declared", () => {
    expect(chooseAvailableModel(models, "connection:deleted", null)).toBe(
      "connection:first"
    );
  });

  it("returns an explicit empty selection when nothing is configured", () => {
    expect(chooseAvailableModel([], "local", null)).toBe(NO_MODEL_SELECTED);
    expect(chooseAvailableModel([], "connection:gone", "connection:gone")).toBe(
      NO_MODEL_SELECTED
    );
  });

  it("never resolves to the legacy local placeholder", () => {
    expect(
      chooseAvailableModel(withLegacyPlaceholder, "local", "local")
    ).toBe("connection:first");
    expect(chooseAvailableModel([{ id: "local", available: true }], "local", "local")).toBe(
      NO_MODEL_SELECTED
    );
  });

  it("skips unavailable entries and the legacy id when picking the first model", () => {
    expect(
      chooseAvailableModel(
        [
          { id: "connection:offline", available: false },
          { id: "local", available: true },
          { id: "connection:live", available: true }
        ],
        "",
        null
      )
    ).toBe("connection:live");
  });

  it("ignores an unavailable preferred or current model", () => {
    expect(
      chooseAvailableModel(
        [
          { id: "connection:offline", available: false },
          { id: "connection:live", available: true }
        ],
        "connection:offline",
        null,
        "connection:offline"
      )
    ).toBe("connection:live");
  });
});

describe("isRealModelId", () => {
  it("rejects empty and the legacy placeholder", () => {
    expect(isRealModelId("")).toBe(false);
    expect(isRealModelId(null)).toBe(false);
    expect(isRealModelId(undefined)).toBe(false);
    expect(isRealModelId("local")).toBe(false);
    expect(isRealModelId("  ")).toBe(false);
  });

  it("accepts real catalog ids", () => {
    expect(isRealModelId("connection:abc")).toBe(true);
    expect(isRealModelId("gpt-4.1")).toBe(true);
  });
});

describe("shouldPersistSelection", () => {
  it("reports a migration when the stored value changes", () => {
    expect(shouldPersistSelection("local", "connection:first")).toBe(true);
    expect(shouldPersistSelection("connection:gone", "")).toBe(true);
  });

  it("stays quiet when nothing changed", () => {
    expect(shouldPersistSelection("connection:first", "connection:first")).toBe(
      false
    );
  });
});

describe("resolveDisplayedModelId", () => {
  it("returns exactly the id that would be sent with the request", () => {
    expect(resolveDisplayedModelId(models, "connection:new", "connection:first")).toBe(
      "connection:new"
    );
  });

  it("returns the empty state instead of the first list entry", () => {
    // 界面展示不得回退到列表首项，否则会出现“显示 A、请求发 B”。
    expect(resolveDisplayedModelId([], "connection:gone", null)).toBe("");
    expect(
      resolveDisplayedModelId([{ id: "local", available: true }], "local", "local")
    ).toBe("");
  });
});
