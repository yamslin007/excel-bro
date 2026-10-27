import { borderEdges, type CellBorderSide } from "@excel-bro/core/borderEdges";
import { columnLetterFromIndex } from "@excel-bro/core/range";
import { normalizeBorderAction, type ExcelAction } from "../../excel-addin/src/contracts";
import { wpsOleColorFromHex } from "./wpsFormatCapability";
import type { WpsRange } from "./wpsHost";
import { countHostWrite } from "./wpsHostWriteCounter";
export type BorderAction = Extract<ExcelAction,{type:"setBorders"}>;
export const BORDER_FIELDS = ["borderTop","borderBottom","borderLeft","borderRight"] as const;
export type BorderField = typeof BORDER_FIELDS[number];
const sides = ["top","bottom","left","right"] as const;
const indexes = {top:8,bottom:9,left:7,right:10} as const;
const styles = {continuous:1,dash:-4115,dashDot:4,dot:-4118,double:-4119,none:-4142} as const;
const weights = {hairline:1,thin:2,medium:-4138,thick:4} as const;
export const borderField = (side:CellBorderSide):BorderField => BORDER_FIELDS[sides.indexOf(side)];
function edge(range:WpsRange,field:BorderField) {
  const value = range.Borders?.Item?.(indexes[sides[BORDER_FIELDS.indexOf(field)]]);
  if(!value) throw new Error("宿主没有可逐边读写的边框接口。");
  return value;
}
/** Fixed tuple: style, weight, resolved OLE color, color mode, theme, tint. */
export function readWpsBorder(range:WpsRange,field:BorderField):string|null {
  try {
    const b=edge(range,field), style=b.LineStyle, weight=b.Weight, color=b.Color, index=b.ColorIndex;
    if(![style,weight,color,index].every(v=>typeof v === "number" && Number.isFinite(v))) return null;
    const theme=b.ThemeColor, tint=b.TintAndShade;
    // Unknown color provenance must not be flattened to an RGB value on undo.
    if(typeof theme !== "number" || !Number.isInteger(theme) || theme<0 || theme>12 || typeof tint!=="number" || !Number.isFinite(tint) || tint < -1 || tint > 1) return null;
    const mode=theme>0?2:index===-4105?1:0;
    return JSON.stringify([style,weight,color,mode,theme,tint]);
  } catch { return null; }
}
export function borderTarget(action:BorderAction):string {
  const a=normalizeBorderAction(action);
  return JSON.stringify([styles[a.style],weights[a.weight],wpsOleColorFromHex(a.color),0,0,0]);
}
export function restoreWpsBorder(range:WpsRange,field:BorderField,saved:string):void {
  const values:unknown=JSON.parse(saved);
  if(!Array.isArray(values)||values.length!==6||!values.every(v=>typeof v==="number"&&Number.isFinite(v))) throw new Error("边框恢复快照无效。");
  const [style,weight,color,mode,theme,tint]=values as number[];
  // 计数口径：一条边算**一次**宿主写调用（同一个 Border 对象的属性组）。
  // 撤销用的格式回滚也会走到这里，但那时没有活动作用域 ⇒ 不计。
  countHostWrite(true);
  const b=edge(range,field);
  b.Weight=weight;
  if(mode===2) b.ThemeColor=theme;
  else if(mode===1) b.ColorIndex=-4105;
  else b.Color=color;
  b.TintAndShade=tint;
  // Weight and Color can implicitly enable a line: restore LineStyle last.
  b.LineStyle=style;
}
export function applyWpsBorders(range:WpsRange,action:BorderAction):void {
  for(const side of action.sides) {
    if(!sides.includes(side as CellBorderSide)) throw new Error("内部边框必须先展开为逐格边界。");
    restoreWpsBorder(range,borderField(side as CellBorderSide),borderTarget(action));
  }
}
/** Explicit single-cell edges make existing property-level snapshots/recovery exact. */
export function expandWpsBorders(actions:readonly ExcelAction[],maxCells:number,maxActions:number):ExcelAction[] {
  const owners=new Map<string,string>();
  return actions.flatMap((action):ExcelAction[]=>{
    if(action.type!=="setBorders") return [action];
    const grouped=new Map<string,CellBorderSide[]>();
    for(const e of borderEdges(action.range,action.sides,maxCells)) {
      const range=`${columnLetterFromIndex(e.column-1)}${e.row}`, physical=`${action.sheet}\0${e.key}`, owner=`${range}:${e.side}`;
      const previous=owners.get(physical);
      if(previous && previous!==owner) throw new Error("计划从相邻单元格重复设置同一边界，请合并边框范围后重新生成。");
      owners.set(physical,owner);
      const cellSides=grouped.get(range)??[];
      if(!cellSides.includes(e.side)) cellSides.push(e.side);
      grouped.set(range,cellSides);
      if(grouped.size>maxActions) throw new Error("展开后的边框动作超过恢复记录限额。");
    }
    if(!grouped.size) throw new Error("所选区域没有可设置的内部边框。");
    return [...grouped].map(([range,sides])=>({...action,...normalizeBorderAction(action),range,sides}));
  });
}
