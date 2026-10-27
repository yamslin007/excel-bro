import type {AnalysisPlan, CellValue, WpsExecutionState} from "../../excel-addin/src/contracts";
import type {WpsApplication, WpsRange, WpsWorksheet} from "./wpsHost";
import {parseExplicitA1Rect} from "../../excel-addin/src/planCore";
import {prepareWpsPlan} from "./wpsPlanExecutor";
import * as ledger from "./wpsExecutionLedger";
export const operationId = "wps-resume-0001";
export function fixture() {
  const storage = new Map<string, unknown>(), writes: string[] = [];
  const cells = new Map<string, CellValue>([["Source!A1", "private-source"], ["Target!A1", "old1"], ["Target!A2", "old2"], ["Target!A3", "old3"]]);
  const sheets: WpsWorksheet[] = ["Source", "Target"].map(Name => ({Name, ProtectContents: false, Range(address: string | WpsRange, end?: WpsRange) {
    const rect = parseExplicitA1Rect(typeof address === "string" ? address : `A${address.Row}:A${end?.Row ?? address.Row}`)!;
    if (!rect || rect.startColumn !== 1 || rect.endColumn !== 1) throw new Error("fixture range");
    const read = () => Array.from({length: rect.endRow - rect.startRow + 1}, (_, i) => [cells.get(`${Name}!A${rect.startRow+i}`) ?? null]);
    const range: WpsRange = {Row: rect.startRow, Column: 1, Item: (row: number) => ({Row:rect.startRow+row-1, Column:1}), MergeCells: false, Rows: {Count: rect.endRow-rect.startRow+1}, Columns:{Count:1}};
    const write = (value: CellValue[][] | CellValue) => {
      writes.push(`${Name}!${address}`);
      const matrix = Array.isArray(value) ? value as CellValue[][] : [[value]];
      matrix.forEach((row,i) => cells.set(`${Name}!A${rect.startRow+i}`, row[0]));
    };
    Object.defineProperties(range, {Value2: {get:read,set:write}, Formula:{get:read,set:write}});
    return range;
  }}));
  for (const sheet of sheets) Object.defineProperty(sheet, "UsedRange", {get: () => sheet.Range!("A1:A3")});
  const application: WpsApplication = {Version:"12.0", ActiveWorkbook: {Name:"test.xlsx", FullName:"C:\\Data\\test.xlsx",
    Path:"C:\\Data", ReadOnly:false, Worksheets:{Count:2,Item:i=>sheets[Number(i)-1]}, ActiveSheet:sheets[1]},
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
