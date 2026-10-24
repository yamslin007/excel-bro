import { parseExplicitA1Rect } from "./planCore";
export type CellBorderSide = "top" | "bottom" | "left" | "right";
export type RegionBorderSide = CellBorderSide | "insideHorizontal" | "insideVertical";
/** One physical grid edge has two possible cell owners. Keep the key for alias checks. */
export function borderEdges(range: string, sides: readonly RegionBorderSide[], maxCells: number) {
  const rect = parseExplicitA1Rect(range);
  if (!rect || rect.rows * rect.columns > maxCells) throw new Error("边框区域无效或超过单次限额。");
  const edges: Array<{row:number; column:number; side:CellBorderSide; key:string}> = [];
  const add = (row:number,column:number,side:CellBorderSide) => {
    const key = side === "top" || side === "bottom"
      ? `h:${row + (side === "bottom" ? 1 : 0)}:${column}`
      : `v:${row}:${column + (side === "right" ? 1 : 0)}`;
    edges.push({row,column,side,key});
  };
  for (const side of sides) {
    if (side === "top" || side === "bottom") for(let c=rect.startColumn;c<=rect.endColumn;c++) add(side === "top" ? rect.startRow : rect.endRow,c,side);
    else if(side === "left" || side === "right") for(let r=rect.startRow;r<=rect.endRow;r++) add(r,side === "left" ? rect.startColumn : rect.endColumn,side);
    else if(side === "insideHorizontal") for(let r=rect.startRow;r<rect.endRow;r++) for(let c=rect.startColumn;c<=rect.endColumn;c++) add(r,c,"bottom");
    else if(side === "insideVertical") for(let r=rect.startRow;r<=rect.endRow;r++) for(let c=rect.startColumn;c<rect.endColumn;c++) add(r,c,"right");
    else throw new Error("边框方向无效。");
  }
  return edges;
}
