import type {CellValue} from "./contracts";
import {assertRectangularValues,parseExplicitA1Rect} from "./planCore";
import {columnLetterFromIndex} from "./range";
export type CopyRangeInput={sheet:string;sourceSheet:string;sourceRange:string;targetRange:string;copyType:string;skipBlanks:boolean;transpose:boolean;filters?:unknown[]};
export type CopyRangeWrite={type:"writeValues";sheet:string;range:string;values:CellValue[][]}|{type:"writeFormulas";sheet:string;range:string;formulas:string[][]};
/** Deterministic local lowering: values or links only; no workbook access or model call. */
export function copyRangePlan(action:CopyRangeInput,sourceValues:CellValue[][], selectedRows?: readonly number[]):CopyRangeWrite[]{
  if (!["values","link"].includes(action.copyType) || (action.filters?.length && !selectedRows)) throw new Error("此复制路径支持值或链接；条件复制必须提供已解析的来源行，完整格式与相对公式复制不适用。");
  const source=parseExplicitA1Rect(action.sourceRange),target=parseExplicitA1Rect(action.targetRange);
  if(!source||!target) throw new Error("复制来源与目标必须是明确的 A1 区域。");
  assertRectangularValues(sourceValues);
  if(sourceValues.length!==source.rows||sourceValues[0]?.length!==source.columns) throw new Error("复制来源矩阵与区域尺寸不符。");
  const sourceRows = selectedRows ?? Array.from({length:source.rows}, (_, index) => index);
  if (!sourceRows.length || new Set(sourceRows).size !== sourceRows.length || sourceRows.some(row => !Number.isInteger(row) || row < 0 || row >= source.rows)) throw new Error("复制筛选行无效或没有命中数据。");
  const rows=action.transpose?source.columns:sourceRows.length,columns=action.transpose?sourceRows.length:source.columns;
  if((target.rows!==1||target.columns!==1)&&(target.rows!==rows||target.columns!==columns)) throw new Error("复制目标尺寸与来源不符。");
  if(target.startRow+rows-1>1048576||target.startColumn+columns-1>16384) throw new Error("复制目标超出工作表边界。");
  const writes:CopyRangeWrite[]=[];
  for(let r=0;r<rows;r++){
    let start=0,values:CellValue[]=[],formulas:string[]=[];
    const flush=()=>{
      if(!values.length)return;
      const first=`${columnLetterFromIndex(target.startColumn+start-1)}${target.startRow+r}`;
      const last=`${columnLetterFromIndex(target.startColumn+start+values.length-2)}${target.startRow+r}`;
      const range=values.length===1?first:`${first}:${last}`;
      writes.push(action.copyType==="link"?{type:"writeFormulas",sheet:action.sheet,range,formulas:[formulas]}:{type:"writeValues",sheet:action.sheet,range,values:[values]});
      values=[];formulas=[];
    };
    for(let c=0;c<columns;c++){
      const sr=sourceRows[action.transpose?c:r],sc=action.transpose?r:c,value=sourceValues[sr][sc];
      if(action.skipBlanks&&(value===null||value==="")){flush();continue;}
      if(!values.length)start=c;
      values.push(value);
      formulas.push(`='${action.sourceSheet.replace(/'/g,"''")}'!${columnLetterFromIndex(source.startColumn+sc-1)}${source.startRow+sr}`);
    }
    flush();
  }
  if(!writes.length) throw new Error("来源全部为空且设置了跳过空白，无需写入。");
  return writes;
}
