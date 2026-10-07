import capabilities from "../../../config/capabilities.json";
import type { CellValue, PythonDataScope, PythonDataSource, PythonMatrixResult, PythonResultReceipt } from "./contracts";
import { parseExplicitA1Rect } from "./planCore";
import type { PythonDataHost } from "./pythonData";

function validUnicode(value: string): boolean {
  // for-of combines valid surrogate pairs, without lookbehind support in WPS.
  for (const character of value) {
    const point = character.codePointAt(0)!;
    if (point >= 0xD800 && point <= 0xDFFF) return false;
  }
  return true;
}

export class PythonResultUnavailable extends Error {
  constructor() { super("分析结果无法生成安全预览，请重新读取并分析；原表未修改。"); this.name = "PythonResultUnavailable"; }
}

export function checkedResultReceipt(raw: PythonResultReceipt): PythonResultReceipt {
  if (!raw || Object.keys(raw).sort().join(",") !== "columns,contentFingerprint,resultHandle,rows" ||
      typeof raw.resultHandle !== "string" || !/^result:\/\/[A-Za-z0-9_-]{43}$/.test(raw.resultHandle) ||
      typeof raw.contentFingerprint !== "string" || !/^[0-9a-f]{64}$/.test(raw.contentFingerprint) ||
      !Number.isInteger(raw.rows) || raw.rows < 1 || raw.rows > capabilities.pythonResult.maxRows ||
      !Number.isInteger(raw.columns) || raw.columns < 1 || raw.columns > capabilities.pythonResult.maxColumns ||
      raw.rows*raw.columns > capabilities.pythonResult.maxCells) throw new PythonResultUnavailable();
  return { resultHandle:raw.resultHandle, contentFingerprint:raw.contentFingerprint, rows:raw.rows, columns:raw.columns };
}

export function checkedMatrixResult(raw: unknown, receipt: PythonResultReceipt): PythonMatrixResult {
  const metadata = checkedResultReceipt(receipt);
  const result = raw as PythonMatrixResult;
  if (!result || Object.keys(result).sort().join(",") !== "kind,values" || result.kind !== "matrix" ||
      !Array.isArray(result.values) || result.values.length !== metadata.rows) throw new PythonResultUnavailable();
  const values: CellValue[][] = [];
  for (const row of result.values) {
    if (!Array.isArray(row) || row.length !== metadata.columns) throw new PythonResultUnavailable();
    for (const value of row) {
      if (value === null || typeof value === "boolean") continue;
      if (typeof value === "number" && Number.isFinite(value) && (!Number.isInteger(value) || Number.isSafeInteger(value))) continue;
      if (typeof value === "string" && value.length <= capabilities.pythonResult.maxCellChars && validUnicode(value)) continue;
      throw new PythonResultUnavailable();
    }
    values.push([...row]);
  }
  if (new TextEncoder().encode(JSON.stringify(result)).length > capabilities.pythonResult.maxBytes) throw new PythonResultUnavailable();
  return { kind:"matrix",values };
}

/** Structural subset of AnalysisPlan, not a new executable action protocol. */
export interface PythonValuesPlan {
  id: string; title: string; summary: string; assumptions: string[]; warnings: string[];
  actions: { type:"writeValues"; sheet:string; range:string; values:CellValue[][] }[];
  acceptanceCriteria: { type:"rangeEquals"; sheet:string; range:string; expected:CellValue[][] }[];
}

/** Only the trusted caller selects an explicit target on an existing source
 * sheet. Output cannot supply coordinates, metadata, formulas or actions.
 * Text that hosts might reinterpret is rejected, never silently escaped.
 */
export function compilePythonValuesPlan(scope: PythonDataScope, receipt: PythonResultReceipt,
  raw: unknown, target: PythonDataSource): PythonValuesPlan {
  const result = checkedMatrixResult(raw,receipt);
  const rect = parseExplicitA1Rect(target.range);
  if (!scope.sources.some((source) => source.sheetId === target.sheetId && source.sheetName === target.sheetName) ||
      !rect || rect.rows !== receipt.rows || rect.columns !== receipt.columns ||
      target.range !== target.range.toUpperCase() || !/^[A-Z]{1,3}[1-9][0-9]{0,6}(?::[A-Z]{1,3}[1-9][0-9]{0,6})?$/.test(target.range)) throw new PythonResultUnavailable();
  for (const row of result.values) for (const value of row) {
    if (typeof value === "string" && (/^[\s]*[=+\-@']/.test(value) || value.includes("\0"))) throw new PythonResultUnavailable();
  }
  return { id:`python-values-${crypto.randomUUID()}`, title:"Python 分析结果写入预览",
    summary:"把本机隔离分析得到的固定值写入指定区域；尚未写入原表。", assumptions:[],
    warnings:["结果仅在当前授权会话内有效；源数据变化需重新分析。", "固定值不会随来源变化自动更新。"],
    actions:[{type:"writeValues",sheet:target.sheetName,range:target.range,values:result.values.map((row) => [...row])}],
    acceptanceCriteria:[{type:"rangeEquals",sheet:target.sheetName,range:target.range,expected:result.values.map((row) => [...row])}] };
}

export interface PythonResultSession {
  scope: PythonDataScope;
  assertHost(host: PythonDataHost): Promise<void>;
  getResult(receipt: PythonResultReceipt): Promise<PythonMatrixResult>;
  captureSources(): Promise<{ source:PythonDataSource; values:CellValue[][] }[]>;
  revalidate(): Promise<void>;
}

/** Preview only. The caller must bind assertCurrent to the last pre-write
 * guard if a later product UI enables execution. Nothing is serialized into
 * history, planning, remote validation or execution ledgers by this helper.
 */
export async function stagePythonResultPreview<T>(session: PythonResultSession, receipt: PythonResultReceipt,
  target: PythonDataSource, prepareLocal: (plan:PythonValuesPlan,
    sources:{source:PythonDataSource;values:CellValue[][]}[]) => Promise<T>) {
  const scope = structuredClone(session.scope), resultReceipt = checkedResultReceipt(receipt), destination = {...target};
  if (!capabilities.pythonExecution.enabled) throw new PythonResultUnavailable();
  const result = await session.getResult(resultReceipt);
  const plan = compilePythonValuesPlan(scope,resultReceipt,result,destination);
  const fingerprint = JSON.stringify(plan);
  const sources = await session.captureSources();
  if (JSON.stringify(sources.map((snapshot) => snapshot.source)) !== JSON.stringify(scope.sources)) throw new PythonResultUnavailable();
  const prepared = await prepareLocal(plan,sources);
  const assertCurrent = async () => {
    if (!capabilities.pythonExecution.enabled || JSON.stringify(plan) !== fingerprint || JSON.stringify(session.scope) !== JSON.stringify(scope)) throw new PythonResultUnavailable();
    await session.getResult(resultReceipt);
  };
  await assertCurrent();
  return { plan, prepared, assertCurrent };
}
