import { afterEach, describe, expect, it, vi } from "vitest";
import { listModels } from "./api";

afterEach(() => vi.unstubAllGlobals());
const model = { id: "connection:a", label: "DS · deepseek-v4-flash", provider: "model", available: true, supportsVision: false };
const catalog = { defaultModelId: model.id, models: [model] };
const connection = { id: "a", catalogModelId: model.id, label: "DS", modelId: "deepseek-v4-flash", baseUrl: "https://example.test", supportsVision: false, apiKeyConfigured: true, apiKeyHint: null };
const settings = { baseUrl: null, defaultModel: null, apiKeyConfigured: false, apiKeyHint: null, connections: [connection] };
const json = (value: unknown) => new Response(JSON.stringify(value));

describe("legacy model catalog display metadata", () => {
  it("fills the upstream ID by the exact connection key and preserves routing", async () => {
    const fetcher = vi.fn().mockResolvedValueOnce(json(catalog)).mockResolvedValueOnce(json(settings));
    vi.stubGlobal("fetch", fetcher);
    const result = await listModels();
    expect(result.defaultModelId).toBe(model.id);
    expect(result.models[0]).toEqual({ ...model, modelId: "deepseek-v4-flash" });
    expect(fetcher).toHaveBeenCalledTimes(2);
  });
  it("does not overwrite a current catalog ID or fetch additional settings", async () => {
    const current = { ...catalog, models: [{ ...model, modelId: "deepseek-v4-flash" }] };
    const fetcher = vi.fn().mockResolvedValueOnce(json(current));
    vi.stubGlobal("fetch", fetcher);
    expect(await listModels()).toEqual(current);
    expect(fetcher).toHaveBeenCalledTimes(1);
  });
  it("does not borrow another connection's model even if its nickname matches", async () => {
    vi.stubGlobal("fetch", vi.fn().mockResolvedValueOnce(json(catalog)).mockResolvedValueOnce(json({ ...settings, connections: [{ ...connection, catalogModelId: "connection:b" }] })));
    expect(await listModels()).toEqual(catalog);
  });
  it("keeps the usable catalog when supplementary settings fail", async () => {
    vi.stubGlobal("fetch", vi.fn().mockResolvedValueOnce(json(catalog)).mockRejectedValueOnce(new Error("offline")));
    expect(await listModels()).toEqual(catalog);
  });
});
