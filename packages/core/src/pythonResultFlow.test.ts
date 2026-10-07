import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import capabilities from "../../../config/capabilities.json";
import { collectPythonData, type PythonDataHost } from "./pythonData";
import { stagePythonResultPreview } from "./pythonResult";

const source = {sheetId:"1",sheetName:"Sheet1",range:"A1:B1"};
const target = {...source,range:"E1:F1"};
const options = {taskId:"task",authorizationVersion:"b".repeat(64),sources:[source],apiBaseUrl:"http://127.0.0.1:8765"};
const wire = '{"kind":"matrix","values":[[1.0,"PRIVATE_RESULT"]]}';
let requests: {url:string;body:string}[];
let resultReply: () => Response;
const digest = async (text:string) => Array.from(new Uint8Array(await crypto.subtle.digest("SHA-256",new TextEncoder().encode(text))))
  .map((byte) => byte.toString(16).padStart(2,"0")).join("");
const receipt = async () => ({resultHandle:"result://"+"r".repeat(43),contentFingerprint:await digest(wire),rows:1,columns:2});

beforeEach(() => {
  capabilities.pythonExecution.enabled = true;
  requests = []; resultReply = () => new Response(wire);
  vi.stubGlobal("fetch",vi.fn(async (url:string,init:RequestInit) => {
    requests.push({url,body:String(init.body)});
    if (url.endsWith("/result")) return resultReply();
    return new Response(JSON.stringify(url.endsWith("/scope") ? {scopeToken:"t".repeat(43)}
      : url.endsWith("/snapshot") ? {datasetHandle:"dataset://"+"d".repeat(43),contentFingerprint:"c".repeat(64)} : {revoked:true}));
  }));
});
afterEach(() => {capabilities.pythonExecution.enabled = false;vi.unstubAllGlobals();});

function host(hostId:PythonDataHost["hostId"]): PythonDataHost {
  return {hostId,identity:"PRIVATE_PATH",authorizedSources:[{id:"1",name:"Sheet1"}],
    assertCurrent:vi.fn(async () => {}),read:vi.fn(async () => [[1,"PRIVATE_SOURCE"]])};
}

describe.each(["excel-office","wps-et"] as const)("volatile result flow %s",(hostId) => {
  it("hashes actual Python JSON bytes; local preview, no external validation/storage",async () => {
    const storage = {setItem:vi.fn()};vi.stubGlobal("localStorage",storage);
    const reader = host(hostId),session = await collectPythonData(reader,options);
    const prepare = vi.fn(async () => ({previewOnly:true}));
    const preview = await stagePythonResultPreview(session,await receipt(),target,prepare);
    expect(preview.plan.actions[0].values).toEqual([[1,"PRIVATE_RESULT"]]);
    expect(prepare).toHaveBeenCalledOnce();
    expect(requests.every(({url}) => url.startsWith(options.apiBaseUrl+"/api/python/data/"))).toBe(true);
    expect(requests.filter(({url}) => url.endsWith("/result")).every(({body}) => !body.includes("PRIVATE_RESULT") && !body.includes("PRIVATE_SOURCE"))).toBe(true);
    expect(storage.setItem).not.toHaveBeenCalled();
    expect(JSON.stringify(session)).not.toContain("PRIVATE_RESULT");
    await session.dispose();await expect(preview.assertCurrent()).rejects.toThrow();
  });
  it("cannot prepare after source drift or file/authorization drift",async () => {
    const reader = host(hostId),session = await collectPythonData(reader,options),prepare = vi.fn(async () => 1);
    vi.mocked(reader.read).mockResolvedValue([[2,"CHANGED"]]);
    await expect(stagePythonResultPreview(session,await receipt(),target,prepare)).rejects.toThrow();
    expect(prepare).not.toHaveBeenCalled();
    expect(requests.some(({url}) => url.endsWith("/result"))).toBe(false);
    expect(requests.at(-1)?.url).toMatch(/revoke$/);
  });
  it("checks source again after fetching, before local preparation",async () => {
    const reader = host(hostId),session = await collectPythonData(reader,options),prepare = vi.fn(async () => 1);
    resultReply = () => {vi.mocked(reader.read).mockResolvedValue([[2,"CHANGED"]]);return new Response(wire);};
    await expect(stagePythonResultPreview(session,await receipt(),target,prepare)).rejects.toThrow();
    expect(prepare).not.toHaveBeenCalled();
  });
  it.each(["expiry","tamper","bom","oversized"])("rejects %s response and revokes",async (mode) => {
    const session = await collectPythonData(host(hostId),options),prepare = vi.fn(async () => 1);
    resultReply = () => mode === "expiry" ? new Response("{}",{status:409}) : new Response(mode === "tamper" ? wire.replace("1.0","9") : mode === "bom" ? "\uFEFF"+wire : "x".repeat(capabilities.pythonResult.maxBytes+1));
    await expect(stagePythonResultPreview(session,await receipt(),target,prepare)).rejects.toThrow("分析结果无法生成安全预览");
    expect(prepare).not.toHaveBeenCalled();expect(requests.at(-1)?.url).toMatch(/revoke$/);
  });
  it("cancellation and disabled flag invalidate already prepared preview",async () => {
    const controller = new AbortController();
    const session = await collectPythonData(host(hostId),{...options,signal:controller.signal});
    const preview = await stagePythonResultPreview(session,await receipt(),target,async () => 1);
    capabilities.pythonExecution.enabled = false;
    await expect(preview.assertCurrent()).rejects.toThrow();
    capabilities.pythonExecution.enabled = true;controller.abort();
    await expect(preview.assertCurrent()).rejects.toThrow();
  });
  it("does not let mutable public scope authorize another sheet or host",async () => {
    const reader = host(hostId),session = await collectPythonData(reader,options),prepare = vi.fn(async () => 1);
    await expect(session.assertHost(host(hostId))).rejects.toThrow();
    session.scope.sources = [{...source,sheetId:"2",sheetName:"Sheet2"}];
    await expect(stagePythonResultPreview(session,await receipt(),{...target,sheetId:"2",sheetName:"Sheet2"},prepare)).rejects.toThrow();
    expect(prepare).not.toHaveBeenCalled();
  });
  it("guards the immutable plan and rereads result at final preview boundary",async () => {
    const session = await collectPythonData(host(hostId),options);
    const preview = await stagePythonResultPreview(session,await receipt(),target,async () => 1);
    preview.plan.actions[0].values[0][0] = 9;
    await expect(preview.assertCurrent()).rejects.toThrow();
    preview.plan.actions[0].values[0][0] = 1;
    resultReply = () => new Response("{}",{status:409});
    await expect(preview.assertCurrent()).rejects.toThrow();
  });
});
