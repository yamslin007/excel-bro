import { parseExplicitA1Rect } from "@excel-bro/core/planCore";
import type { CellValue } from "../../excel-addin/src/contracts";
import type { WpsApplication, WpsRange, WpsWorksheet } from "./wpsHost";

/** Stateful host for capability tests: ranges alias one grid, like the actual host. */
export function capabilityHost(names = ["明细"], rows = 8, columns = 8) {
  const grids = names.map(() => Array.from({length:rows}, (_,r) => Array.from({length:columns}, (_,c) => ({
    value: `old-${r}-${c}` as CellValue, formula: `old-${r}-${c}` as CellValue,
    bold:false, italic:false, name:"宋体", size:11, color:0,
    horizontalAlignment:1, verticalAlignment:-4107, wrapText:false,
    rowHeight:13.5, fillColor:0x123456, fillPattern:-4142, fillPatternColor:0,
    numberFormat:"General"
  }))));
  let writes = 0;
  const sheets: WpsWorksheet[] = names.map((Name, index) => {
    const grid = grids[index];
    const cell = (r:number,c:number): WpsRange => {
      const state = grid[r][c];
      const bind = (pairs: Record<string,string>) => Object.defineProperties({}, Object.fromEntries(
        Object.entries(pairs).map(([api,key]) => [api,{
          get:()=>state[key as keyof typeof state],
          set:(value:unknown)=>{(state as unknown as Record<string,unknown>)[key]=value;writes++;}, enumerable:true
        }])));
      return Object.assign(bind({NumberFormat:"numberFormat", HorizontalAlignment:"horizontalAlignment",VerticalAlignment:"verticalAlignment",WrapText:"wrapText",RowHeight:"rowHeight"}), {
        Font:bind({Bold:"bold",Italic:"italic",Name:"name",Size:"size",Color:"color"}),
        Interior:bind({Color:"fillColor",Pattern:"fillPattern",PatternColor:"fillPatternColor"}),
        Rows:{Count:1},Columns:{Count:1}
      }) as WpsRange;
    };
    const Range = (address:string): WpsRange => {
      const rect = parseExplicitA1Rect(address);
      if (!rect || rect.endRow > rows || rect.endColumn > columns) throw new Error(`Invalid test range: ${address}`);
      const r0=rect.startRow-1,c0=rect.startColumn-1;
      const read=(key:"value"|"formula")=>Array.from({length:rect.rows},(_,r)=>Array.from({length:rect.columns},(_,c)=>grid[r0+r][c0+c][key]));
      const write=(key:"value"|"formula", input:unknown)=>{
        const matrix=Array.isArray(input) ? input : [[input]];
        for(let r=0;r<rect.rows;r++) for(let c=0;c<rect.columns;c++) {
          const v=matrix[r][c] as CellValue;
          grid[r0+r][c0+c].value=v;
          grid[r0+r][c0+c].formula=v;
          writes++;
        }
      };
      const result = rect.rows===1&&rect.columns===1 ? cell(r0,c0) : {};
      Object.assign(result,{MergeCells:false,Rows:{Count:rect.rows},Columns:{Count:rect.columns},Item:(r:number,c:number)=>cell(r0+r-1,c0+c-1)});
      Object.defineProperties(result,{
        Value2:{get:()=>read("value"),set:(v:unknown)=>write("value",v),enumerable:true},
        Formula:{get:()=>read("formula"),set:(v:unknown)=>write("formula",v),enumerable:true}
      });
      return result;
    };
    return {Name,Index:index+1,ProtectContents:false,Range};
  });
  const storage=new Map<string,unknown>();
  const book={Name:"test.xlsx",FullName:"C:\\Data\\test.xlsx",Path:"C:\\Data",ReadOnly:false,
    Worksheets:{Count:sheets.length,Item:(i:number|string)=>typeof i === "number"?sheets[i-1]:sheets.find(s=>s.Name===i)!},ActiveSheet:sheets[0]};
  const application:WpsApplication={Version:"12.0",ActiveWorkbook:book,ActiveSheet:sheets[0],
    PluginStorage:{getItem:key=>storage.get(key),setItem:(key,value)=>{storage.set(key,value);}},
    CreateTaskPane:()=>({ID:1,Visible:true}),GetTaskPane:()=>undefined};
  return {application,grids,sheets,storage,writes:()=>writes};
}
