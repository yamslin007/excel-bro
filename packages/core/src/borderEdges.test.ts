import { describe,it,expect } from "vitest";
import { borderEdges } from "./borderEdges";
describe("physical border edges",()=>{
 it("distinguishes the perimeter from inner grid lines",()=>{
  const outer=borderEdges("B2:C3",["top","bottom","left","right"],4);
  const inner=borderEdges("B2:C3",["insideHorizontal","insideVertical"],4);
  expect(outer).toHaveLength(8);expect(inner.map(e=>e.key)).toEqual(["h:3:2","h:3:3","v:2:3","v:3:3"]);
  expect(new Set([...outer,...inner].map(e=>e.key)).size).toBe(12);
 });
 it("identifies opposing sides as the same edge and bounds expansion",()=>{
  expect(borderEdges("A1",["bottom"],1)[0].key).toBe(borderEdges("A2",["top"],1)[0].key);
  expect(()=>borderEdges("A1:B2",["top"],3)).toThrow();expect(()=>borderEdges("invalid",["top"],10)).toThrow();
  expect(borderEdges("A1",["insideHorizontal","insideVertical"],1)).toEqual([]);
 });
});
