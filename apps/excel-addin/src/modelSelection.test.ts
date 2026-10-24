import { describe, expect, it } from "vitest";
import {
  chooseAvailableModel,
  planningReadPolicyForModel,
  connectionModelSuffix,
  isRealModelId,
  requiresStreamingToolDelivery,
  LEGACY_LOCAL_MODEL_ID,
  NO_MODEL_SELECTED,
  resolveDisplayedModelId,
  shouldPersistSelection
} from "./modelSelection";

describe("planningReadPolicyForModel", () => {
  it("keeps SDK free-text continuation header-only for standalone and saved connections", () => {
    expect(planningReadPolicyForModel("claude-agent-sdk", "ifNeeded")).toBe("none");
    expect(planningReadPolicyForModel("claude-agent-sdk:connection-id", "ifNeeded")).toBe("none");
  });
  it("uses SDK policy for existing connection IDs too", () => {
    expect(planningReadPolicyForModel("connection:id", "ifNeeded")).toBe("none");
    expect(planningReadPolicyForModel("connection:id", "none")).toBe("none");
    expect(planningReadPolicyForModel(null, "ifNeeded")).toBe("ifNeeded");
  });
});

describe("requiresStreamingToolDelivery", () => {
  it("SDK 连接需要流式工具投递", () => {
    expect(requiresStreamingToolDelivery("claude-agent-sdk")).toBe(true);
    expect(requiresStreamingToolDelivery("claude-agent-sdk:connection-0001")).toBe(true);
    expect(requiresStreamingToolDelivery("  Claude-Agent-SDK  ")).toBe(true);
  });

  it("所有有效连接都需要 SDK 流式投递；空选择仍交给模型选择校验", () => {
    expect(requiresStreamingToolDelivery("connection:regular")).toBe(true);
    expect(requiresStreamingToolDelivery("deepseek-chat")).toBe(true);
    expect(requiresStreamingToolDelivery(null)).toBe(false);
    expect(requiresStreamingToolDelivery("")).toBe(false);
    // 目录 ID 不能用来选择另一套 Agent；有效性由模型目录检查。
    expect(requiresStreamingToolDelivery("claude-agent-sdkx")).toBe(true);
  });
});

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


describe("SDK purpose-aware selection", () => {
  const catalogue = [
    { id: "connection:old", available: true, agentAvailable: false, formulaAvailable: true },
    { id: "connection:sdk", available: true, agentAvailable: true, formulaAvailable: false }
  ];
  it("does not silently replace an existing formula-only choice with another provider", () => {
    expect(chooseAvailableModel(catalogue, "connection:old", "connection:sdk")).toBe("");
  });
  it("allows an explicit compatible selection and excludes formula-only defaults", () => {
    expect(chooseAvailableModel(catalogue, "connection:old", null, "connection:sdk")).toBe("connection:sdk");
    expect(chooseAvailableModel(catalogue, "", "connection:old")).toBe("connection:sdk");
  });
  it("migrates the old SDK alias to the exact same connection", () => {
    expect(chooseAvailableModel(catalogue, "claude-agent-sdk:sdk", null)).toBe("connection:sdk");
  });
});
