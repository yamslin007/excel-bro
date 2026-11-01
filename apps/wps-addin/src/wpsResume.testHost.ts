import type {AnalysisPlan, CellValue, WpsExecutionState} from "../../excel-addin/src/contracts";
import type {WpsApplication, WpsRange, WpsWorksheet} from "./wpsHost";
import {parseExplicitA1Rect} from "../../excel-addin/src/planCore";
import {prepareWpsPlan} from "./wpsPlanExecutor";
import {persistWpsRecoveryPackage} from "./wpsRecoveryPackage";
import * as ledger from "./wpsExecutionLedger";
export const operationId = "wps-resume-0001";
/**
 * 登记执行（带 `taskId`）的持久化通道。
 *
 * 走的是**生产函数** `persistWpsRecoveryPackage`（构建 → 落盘 → 逐字节读回校验），
 * 只是它内部的两个网络调用在测试文件里被替成了本机后端替身
 * （`wpsRecoveryStore.testHost.ts`）。不要在这里塞一个"直接返回成功"的假函数 ——
 * 那样"写前落盘失败必须零宿主写入"这条判据就永远抓不到东西。
 */
export const recoveryChannel = { persistMaterial: persistWpsRecoveryPackage };
/**
 * 恢复现场替身。
 *
 * `extraSheet`：额外放一张从未参与执行的工作表（默认没有）。
 *
 * 用途是复现真机 2026-10-28 的"授权比原执行登记多/少一张表"：登记身份与材料只覆盖
 * Source/Target，而用户当下多勾了一张表。授权集合是续跑判据的一部分，所以这种现场
 * 必须能被测到 —— 但**不能**把它变成默认形状，否则一大批只关心两表现场的用例
 * 会一起改变前提。
 */
/**
 * 替身里的一格：要么是常量，要么是"{公式文本 + 计算值}"。
 *
 * 为什么要分开：真机里 `Value2` 给的是**计算值**、`Formula` 给的是**公式文本**，
 * 两者可以同时"数对公式错"（`=1+1` 与 `=2` 都算 2）。替身如果让两者共用一个存储，
 * "拿计算值冒充公式"的缺陷就永远测不出来 —— 那种替身比真实现宽容。
 */
export type FixtureCell = CellValue | { formula: string; value: CellValue };

export function fixture(options: {
  extraSheet?: boolean;
  /** 公式文本 → 计算值。没登记的公式算 `null`（替身不做真实求值，也不猜）。 */
  formulaValues?: Record<string, CellValue>;
} = {}) {
  const storage = new Map<string, unknown>(), writes: string[] = [];
  const cells = new Map<string, FixtureCell>([["Source!A1", "private-source"], ["Target!A1", "old1"], ["Target!A2", "old2"], ["Target!A3", "old3"]]);
  const computed = (formula: string): CellValue => options.formulaValues?.[formula] ?? null;
  const physical = (cell: FixtureCell | undefined): { formula: string | null; value: CellValue } => {
    if (cell && typeof cell === "object") return { formula: cell.formula, value: cell.value };
    const plain = (cell ?? null) as CellValue;
    return typeof plain === "string" && plain.startsWith("=")
      ? { formula: plain, value: computed(plain) }
      : { formula: null, value: plain };
  };
  /*
   * 替身必须支持**任意矩形**（不只 A 列），并且 `UsedRange` 由单元格自己算出来。
   *
   * 为什么非改不可：这一版要证明"公式引用了登记快照没覆盖的那一格"能被拦下。
   * 旧替身只会拼 `A<row>`、`UsedRange` 死认 `A1:A3`，于是 `来源!B10` 这种形状
   * **在测试里根本表达不出来** —— 判据再对也没有东西能撞它（比真实现更宽容的替身
   * 会让缺陷永远测不出来）。
   */
  const columnLetters = (column: number): string => {
    let value = column, text = "";
    while (value > 0) {
      text = String.fromCharCode(65 + ((value - 1) % 26)) + text;
      value = Math.floor((value - 1) / 26);
    }
    return text;
  };
  const columnIndex = (letters: string): number =>
    [...letters.toUpperCase()].reduce((index, char) => index * 26 + char.charCodeAt(0) - 64, 0);
  type CellRect = {startRow: number; startColumn: number; endRow: number; endColumn: number};
  const cellKey = (Name: string, row: number, column: number) => `${Name}!${columnLetters(column)}${row}`;
  /** 地址按真宿主的写法归一：单格 `A3`、多格 `A3:B10`。 */
  const rectAddress = (rect: CellRect): string => {
    const first = `${columnLetters(rect.startColumn)}${rect.startRow}`;
    const last = `${columnLetters(rect.endColumn)}${rect.endRow}`;
    return first === last ? first : `${first}:${last}`;
  };
  const matrixAt = (Name: string, rect: CellRect, formulas: boolean): CellValue[][] =>
    Array.from({length: rect.endRow - rect.startRow + 1}, (_, row) =>
      Array.from({length: rect.endColumn - rect.startColumn + 1}, (_, column) => {
        const cell = physical(cells.get(cellKey(Name, rect.startRow + row, rect.startColumn + column)));
        return formulas ? cell.formula ?? cell.value : cell.value;
      }));
  /**
   * 已用区域 = 该表所有已存在单元格的**包围盒**（与真宿主一致：中间的空格也算在内）。
   * 空表按 `A1:A1` 报 —— 与真宿主返回一张"只有 A1"的已用区域同形。
   */
  const usedRect = (Name: string): CellRect => {
    let minRow = Number.POSITIVE_INFINITY, minColumn = Number.POSITIVE_INFINITY;
    let maxRow = 0, maxColumn = 0;
    for (const key of cells.keys()) {
      const match = /^(.+)!([A-Z]{1,3})([1-9]\d*)$/.exec(key);
      if (!match || match[1] !== Name) continue;
      const row = Number(match[3]), column = columnIndex(match[2]);
      minRow = Math.min(minRow, row); maxRow = Math.max(maxRow, row);
      minColumn = Math.min(minColumn, column); maxColumn = Math.max(maxColumn, column);
    }
    if (!Number.isFinite(minRow)) return {startRow: 1, startColumn: 1, endRow: 1, endColumn: 1};
    return {startRow: minRow, startColumn: minColumn, endRow: maxRow, endColumn: maxColumn};
  };
  const makeRange = (Name: string, rect: CellRect): WpsRange => {
    const rows = rect.endRow - rect.startRow + 1, columns = rect.endColumn - rect.startColumn + 1;
    const range: WpsRange = {
      Row: rect.startRow, Column: rect.startColumn, MergeCells: false,
      Rows: {Count: rows}, Columns: {Count: columns},
      Item: (row: number, column = 1) => makeRange(Name, {
        startRow: rect.startRow + row - 1, startColumn: rect.startColumn + column - 1,
        endRow: rect.startRow + row - 1, endColumn: rect.startColumn + column - 1})
    };
    const write = (value: unknown) => {
      writes.push(`${Name}!${rectAddress(rect)}`);
      const matrix = (Array.isArray(value) ? value : [[value]]) as CellValue[][];
      matrix.forEach((row, rowOffset) => row.forEach((cell, columnOffset) => {
        // 写公式就跟真机一样只落"公式文本 + 它此刻算出来的值"，两者分开存。
        cells.set(cellKey(Name, rect.startRow + rowOffset, rect.startColumn + columnOffset),
          typeof cell === "string" && cell.startsWith("=") ? {formula: cell, value: computed(cell)} : cell);
      }));
    };
    Object.defineProperties(range, {
      Value2: {get: () => matrixAt(Name, rect, false), set: write},
      Formula: {get: () => matrixAt(Name, rect, true), set: write}
    });
    return range;
  };
  const rectOf = (first: WpsRange | string, last?: WpsRange | string): CellRect => {
    if (typeof first === "string") {
      const rect = parseExplicitA1Rect(first);
      if (!rect) throw new Error("fixture range");
      return rect;
    }
    const startRow = Number(first.Row), startColumn = Number(first.Column);
    if (!Number.isInteger(startRow) || !Number.isInteger(startColumn)) throw new Error("fixture range");
    const endRow = typeof last === "object" ? Number(last.Row) : startRow + Number(first.Rows?.Count ?? 1) - 1;
    const endColumn = typeof last === "object" ? Number(last.Column) : startColumn + Number(first.Columns?.Count ?? 1) - 1;
    return {startRow, startColumn, endRow: Math.max(startRow, endRow), endColumn: Math.max(startColumn, endColumn)};
  };
  const sheetNames = options.extraSheet ? ["Source", "Target", "Extra"] : ["Source", "Target"];
  const sheets: WpsWorksheet[] = sheetNames.map(Name => {
    const sheet: WpsWorksheet = {Name, ProtectContents: false,
      Range: (first: WpsRange | string, last?: WpsRange | string) => makeRange(Name, rectOf(first, last))};
    Object.defineProperty(sheet, "UsedRange", {get: () => makeRange(Name, usedRect(Name))});
    return sheet;
  });
  const application: WpsApplication = {Version:"12.0", ActiveWorkbook: {Name:"test.xlsx", FullName:"C:\\Data\\test.xlsx",
    Path:"C:\\Data", ReadOnly:false, Worksheets:{Count:sheets.length,Item:i=>sheets[Number(i)-1]}, ActiveSheet:sheets[1]},
    ActiveSheet:sheets[1], PluginStorage:{getItem:k=>storage.get(k),setItem:(k,v)=>{storage.set(k,v);}},
    CreateTaskPane:()=>({ID:1,Visible:true}),GetTaskPane:()=>undefined};
  const plan: AnalysisPlan = {id:"plan-resume",title:"Update",summary:"Update",assumptions:[],warnings:[],acceptanceCriteria:[],
    sourceDependencySheets:["Source"], actions:[1,2,3].map(i=>({type:"writeValues",sheet:"Target",range:`A${i}:A${i}`,values:[[i*10]]}))};
  const prepared = prepareWpsPlan(application,[1,2],plan,[{sheetName:"Source",range:"A1:A1",values:[["private-source"]]}]);
  const state: WpsExecutionState = {operationId,taskId:"task-1",status:"running",revision:5,
    workbookIdentity:"wps-path:c:\\data\\test.xlsx", descriptor:ledger.wpsRecoveryDescriptor(prepared),
    authorizedSheetNames:["Source","Target"],totalActions:3,nextActionIndex:2,inFlightActionIndex:null,
    ownerActive:false,canResume:true,requiresReconciliation:false};
  return {application,prepared,state,cells,writes,storage};
}
