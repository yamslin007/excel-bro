import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import capabilities from "../../../config/capabilities.json";
import { collectPythonData, type PythonDataHost } from "./pythonData";

const source = { sheetId: "1", sheetName: "Sheet1", range: "A1:B1" };
const options = { taskId: "test-task", authorizationVersion: "b".repeat(64), sources: [source], apiBaseUrl: "http://127.0.0.1:8765" };
const host = (): PythonDataHost => ({ hostId: "excel-office", identity: "PRIVATE_PATH",
  authorizedSources: [{ id: "1", name: "Sheet1" }], assertCurrent: vi.fn(async () => {}), read: vi.fn(async () => [[1, "PRIVATE_ROW"]]) });

let requests: { url: string; init: RequestInit }[];
beforeEach(() => {
  capabilities.pythonExecution.enabled = true;
  requests = [];
  vi.stubGlobal("fetch", vi.fn(async (url: string, init: RequestInit) => {
    requests.push({ url, init });
    const payload = url.endsWith("/scope") ? { scopeToken: "t".repeat(43) }
      : url.endsWith("/snapshot") ? { datasetHandle: "dataset://"+"d".repeat(43), contentFingerprint: "c".repeat(64) } : { revoked: true };
    return new Response(JSON.stringify(payload), { status: 200 });
  }));
});
afterEach(() => { capabilities.pythonExecution.enabled = false; vi.unstubAllGlobals(); });

describe("volatile local Python data", () => {
  it("off means no host access or fetch", async () => {
    capabilities.pythonExecution.enabled = false;
    const reader = host();
    await expect(collectPythonData(reader, options)).rejects.toThrow();
    expect(reader.assertCurrent).not.toHaveBeenCalled();
    expect(reader.read).not.toHaveBeenCalled();
    expect(fetch).not.toHaveBeenCalled();
  });
  it.each(["https://remote.example", "http://127.0.0.1.evil.example", "http://user@localhost:8765", "http://localhost:8765/plan", "http://localhost?x=1"])("rejects remote/ambiguous URL %s before cells", async (apiBaseUrl) => {
    const reader = host();
    await expect(collectPythonData(reader, { ...options, apiBaseUrl })).rejects.toThrow();
    expect(reader.read).not.toHaveBeenCalled(); expect(fetch).not.toHaveBeenCalled();
  });
  it.each(["excel-office", "wps-et"] as const)("same private snapshot protocol for %s", async (hostId) => {
    const reader = { ...host(), hostId };
    const session = await collectPythonData(reader, options);
    expect(reader.read).toHaveBeenCalledTimes(2);
    expect(requests.map((r) => r.url)).toEqual([`${options.apiBaseUrl}/api/python/data/scope`, `${options.apiBaseUrl}/api/python/data/snapshot`]);
    expect(requests[0].init.body).not.toContain("PRIVATE_ROW");
    expect(requests[0].init.body).not.toContain("PRIVATE_PATH");
    expect(requests[1].init.body).toContain("PRIVATE_ROW");
    expect(requests.every((r) => r.init.redirect === "error" && r.init.cache === "no-store")).toBe(true);
    expect(JSON.stringify(session)).not.toContain("PRIVATE_ROW");
    await session.dispose();
    await expect(session.revalidate()).rejects.toThrow();
    expect(requests[2].url).toContain("/revoke");
  });
  it("rejects scope widening before reading", async () => {
    const reader = host();
    await expect(collectPythonData(reader, { ...options, sources: [{ ...source, sheetId: "2" }] })).rejects.toThrow();
    expect(reader.read).not.toHaveBeenCalled(); expect(fetch).not.toHaveBeenCalled();
  });
  it("source drift aborts and revokes without sending rows", async () => {
    const reader = host();
    vi.mocked(reader.read).mockResolvedValueOnce([[1, "before"]]).mockResolvedValueOnce([[2, "after"]]);
    await expect(collectPythonData(reader, options)).rejects.toThrow();
    expect(requests.map((r) => r.url.split("/").pop())).toEqual(["scope", "revoke"]);
  });
  it("later data or identity changes invalidate the session", async () => {
    const reader = host();
    const session = await collectPythonData(reader, options);
    vi.mocked(reader.read).mockResolvedValue([[3, "changed"]]);
    await expect(session.revalidate()).rejects.toThrow();
    expect(requests.at(-1)?.url).toContain("/revoke");
  });
  it("identity drift revokes without a new read", async () => {
    const reader = host();
    const session = await collectPythonData(reader, options);
    vi.mocked(reader.assertCurrent).mockRejectedValue(new Error("PRIVATE_IDENTITY_ERROR"));
    await expect(session.revalidate()).rejects.toThrow("授权数据不可用");
    expect(reader.read).toHaveBeenCalledTimes(2);
    expect(requests.at(-1)?.url).toContain("/revoke");
  });
  it("rejects redirect and oversized reply", async () => {
    const reader = host();
    vi.stubGlobal("fetch", vi.fn(async () => new Response("x".repeat(5000), { status: 200 })));
    await expect(collectPythonData(reader, options)).rejects.toThrow();
    expect(reader.read).not.toHaveBeenCalled();
    vi.stubGlobal("fetch", vi.fn(async () => ({ ok: true, redirected: true })));
    await expect(collectPythonData(reader, options)).rejects.toThrow();
    expect(reader.read).not.toHaveBeenCalled();
  });
  it("cancellation revokes with an independent non-aborted request", async () => {
    const controller = new AbortController();
    const session = await collectPythonData(host(), { ...options, signal: controller.signal });
    controller.abort();
    await vi.waitFor(() => expect(requests.at(-1)?.url).toContain("/revoke"));
    expect(requests.at(-1)?.init.signal).not.toBe(controller.signal);
    expect(requests.at(-1)?.init.signal?.aborted).toBe(false);
    await expect(session.revalidate()).rejects.toThrow();
  });
});
