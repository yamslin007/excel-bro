import capabilities from "../../../config/capabilities.json";
import type { CellValue, PythonDataScope, PythonDataSource, PythonDatasetReceipt, PythonResultReceipt } from "./contracts";
import { checkedMatrixResult, checkedResultReceipt, PythonResultUnavailable } from "./pythonResult";
import { columnIndexFromLetter } from "./range";

export interface PythonDataHost {
  hostId: PythonDataScope["hostId"];
  /** In-memory identity only. A session nonce is required for unsaved files. */
  identity: string;
  authorizedSources: readonly { id: string; name: string }[];
  assertCurrent(): Promise<void>;
  read(source: PythonDataSource): Promise<CellValue[][]>;
}

export class PythonDataUnavailable extends Error {
  constructor() { super("授权数据不可用，请重新确认目标；原表未修改。"); this.name = "PythonDataUnavailable"; }
}

function localBase(raw: string): string {
  const url = new URL(raw);
  if (!["http:", "https:"].includes(url.protocol) || url.username || url.password ||
      url.search || url.hash || !["localhost", "127.0.0.1", "[::1]"].includes(url.hostname) ||
      url.pathname !== "/") throw new PythonDataUnavailable();
  return url.origin;
}

async function hash(value: string): Promise<string> {
  return hashBytes(new TextEncoder().encode(value));
}

async function hashBytes(value: Uint8Array): Promise<string> {
  const bytes = await crypto.subtle.digest("SHA-256", new Uint8Array(value));
  return Array.from(new Uint8Array(bytes), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

function shape(range: string) {
  const match = /^([A-Z]{1,3})([1-9][0-9]{0,6})(?::([A-Z]{1,3})([1-9][0-9]{0,6}))?$/.exec(range);
  if (!match) throw new PythonDataUnavailable();
  const r1 = Number(match[2]), r2 = Number(match[4] ?? match[2]);
  const c1 = columnIndexFromLetter(match[1]), c2 = columnIndexFromLetter(match[3] ?? match[1]);
  if (r2 < r1 || c2 < c1 || r2 > 1048576 || c2 >= 16384 ||
      (r2-r1+1)*(c2-c1+1) > capabilities.pythonData.maxCellsPerDataset) throw new PythonDataUnavailable();
  return { rows: r2-r1+1, columns: c2-c1+1 };
}

function snapshot(values: CellValue[][], source: PythonDataSource): string {
  const expected = shape(source.range);
  if (!Array.isArray(values) || values.length !== expected.rows || values.some((row) =>
    !Array.isArray(row) || row.length !== expected.columns || row.some((value) =>
      value !== null && (typeof value === "number" ? !Number.isFinite(value)
        : !["string", "boolean"].includes(typeof value))))) throw new PythonDataUnavailable();
  const serialized = JSON.stringify(values);
  if (new TextEncoder().encode(serialized).length > capabilities.pythonData.maxDatasetBytes) throw new PythonDataUnavailable();
  return serialized;
}

/** Read-only plumbing, not a UI grant or a model tool. Call only after consent.
 * Off/remote checks precede ANY host access. Values/tokens never use localStorage.
 */
export async function collectPythonData(host: PythonDataHost, options: {
  taskId: string; authorizationVersion: string; sources: readonly PythonDataSource[];
  apiBaseUrl: string; signal?: AbortSignal;
}) {
  if (!capabilities.pythonExecution.enabled) throw new PythonDataUnavailable();
  let base: string;
  try { base = localBase(options.apiBaseUrl); } catch { throw new PythonDataUnavailable(); }
  const sources = options.sources.map((source) => ({ ...source }));
  if (!/^[A-Za-z0-9_-]{1,80}$/.test(options.taskId) || !/^[0-9a-f]{64}$/.test(options.authorizationVersion) ||
      sources.length < 1 || sources.length > capabilities.pythonData.maxDatasetsPerScope ||
      new Set(sources.map((source) => JSON.stringify([source.sheetId, source.range]))).size !== sources.length) throw new PythonDataUnavailable();
  for (const source of sources) {
    shape(source.range);
    if (!host.authorizedSources.some((sheet) => sheet.id === source.sheetId && sheet.name === source.sheetName)) throw new PythonDataUnavailable();
  }
  const scope: PythonDataScope = { version: 1, taskId: options.taskId, hostId: host.hostId,
    workbookIdentityHash: await hash(host.identity), authorizationVersion: options.authorizationVersion, sources };
  let token: string | null = null;
  let disposed = false;
  const request = async (path: string, body: object, signal: AbortSignal | null | undefined = options.signal,
    expectedFingerprint?: string): Promise<Record<string, unknown>> => {
    const serialized = JSON.stringify(body);
    if (new TextEncoder().encode(serialized).length > capabilities.pythonData.maxRequestBytes) throw new PythonDataUnavailable();
    const controller = new AbortController();
    const abort = () => controller.abort();
    signal?.addEventListener("abort", abort, { once: true });
    if (signal?.aborted) controller.abort();
    const timeout = setTimeout(abort, capabilities.pythonData.requestTimeoutMs);
    try {
    const response = await fetch(`${base}/api/python/data/${path}`, { method: "POST",
      headers: { "Content-Type": "application/json" }, body: serialized, signal: controller.signal,
      redirect: "error", cache: "no-store", credentials: "omit" });
    if (!response.ok || response.redirected) throw new PythonDataUnavailable();
    if (!response.body) throw new PythonDataUnavailable();
    const reader = response.body.getReader();
    const chunks: Uint8Array[] = []; let length = 0;
    try {
      while (true) {
        const chunk = await reader.read();
        if (chunk.done) break;
        length += chunk.value.byteLength;
        if (length > (path === "result" ? capabilities.pythonResult.maxBytes : 4096)) { await reader.cancel(); throw new PythonDataUnavailable(); }
        chunks.push(chunk.value);
      }
    } finally { reader.releaseLock(); }
    const bytes = new Uint8Array(length); let offset = 0;
    for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length; }
    const text = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    // Hash actual response bytes, not JS reserialization (Python/JS float
    // encodings differ). The server response is the stored canonical JSON.
    if (expectedFingerprint && await hashBytes(bytes) !== expectedFingerprint) throw new PythonDataUnavailable();
    return JSON.parse(text) as Record<string, unknown>;
    } finally {
      clearTimeout(timeout);
      signal?.removeEventListener("abort", abort);
    }
  };
  const dispose = async () => {
    disposed = true;
    if (!token) return;
    const current = token; token = null;
    // Revocation is best-effort when offline; server TTL remains the fallback.
    try { await request("revoke", { scope, scopeToken: current }, null); } catch { /* no sensitive error echo */ }
  };
  const check = async () => {
    if (!capabilities.pythonExecution.enabled || disposed || options.signal?.aborted) throw new PythonDataUnavailable();
    await host.assertCurrent();
  };
  const receipts: PythonDatasetReceipt[] = [];
  const localFingerprints: string[] = [];
  const onAbort = () => { void dispose(); };
  options.signal?.addEventListener("abort", onAbort, { once: true });
  try {
    await check();
    const grant = await request("scope", { scope });
    if (typeof grant.scopeToken !== "string" || !/^[A-Za-z0-9_-]{43}$/.test(grant.scopeToken)) throw new PythonDataUnavailable();
    token = grant.scopeToken;
    for (let index = 0; index < sources.length; index++) {
      await check();
      const values = snapshot(await host.read(sources[index]), sources[index]);
      await check();
      if (snapshot(await host.read(sources[index]), sources[index]) !== values) throw new PythonDataUnavailable();
      await check();
      const receipt = await request("snapshot", { scope, scopeToken: token, sourceIndex: index, values: JSON.parse(values) });
      if (typeof receipt.datasetHandle !== "string" || !/^dataset:\/\/[A-Za-z0-9_-]{43}$/.test(receipt.datasetHandle) ||
          typeof receipt.contentFingerprint !== "string" || !/^[0-9a-f]{64}$/.test(receipt.contentFingerprint)) throw new PythonDataUnavailable();
      receipts.push({ datasetHandle: receipt.datasetHandle, contentFingerprint: receipt.contentFingerprint });
      localFingerprints.push(await hash(values));
    }
    await check();
    // Every future use must call revalidate; returning a handle is not a write grant.
    const captureSources = async () => {
      try {
        await check();
        const snapshots: {source:PythonDataSource;values:CellValue[][]}[] = [];
        for (let index = 0; index < sources.length; index++) {
          const serialized = snapshot(await host.read(sources[index]), sources[index]);
          if (await hash(serialized) !== localFingerprints[index]) throw new PythonDataUnavailable();
          await check();
          snapshots.push({source:{...sources[index]},values:JSON.parse(serialized) as CellValue[][]});
        }
        return snapshots;
      } catch { await dispose(); throw new PythonDataUnavailable(); }
    };
    const revalidate = async () => { await captureSources(); };
    const getResult = async (raw: PythonResultReceipt) => {
      const receipt = checkedResultReceipt(raw);
      try {
        await revalidate();
        const result = await request("result", {scope,scopeToken:token,datasets:receipts,result:receipt},options.signal,receipt.contentFingerprint);
        await revalidate();
        return checkedMatrixResult(result,receipt);
      } catch { await dispose(); throw new PythonResultUnavailable(); }
    };
    const assertHost = async (candidate: PythonDataHost) => {
      if (candidate !== host) throw new PythonDataUnavailable();
      await check();
    };
    return { scope: structuredClone(scope), receipts: receipts.map((receipt) => ({ ...receipt })), assertHost,
      captureSources, getResult, revalidate, dispose: async () => {
      options.signal?.removeEventListener("abort", onAbort);
      await dispose();
      localFingerprints.length = 0;
    } };
  } catch {
    options.signal?.removeEventListener("abort", onAbort);
    await dispose();
    throw new PythonDataUnavailable();
  }
}
