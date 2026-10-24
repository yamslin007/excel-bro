import { describe, expect, it } from "vitest";
import { copyRangePlan, type CopyRangeInput } from "./copyRangePlan";
const action: CopyRangeInput = {sheet:"目标",sourceSheet:"O'Brien",sourceRange:"B2:C5",targetRange:"D3",copyType:"link",skipBlanks:false,transpose:true,filters:[{}]};
describe("filtered range lowering", () => {
  it("maps compacted transposed links to the original source rows", () => {
    const output = copyRangePlan(action, [["key","value"],["a",1],["b",2],["a",3]], [0,1,3]);
    expect(output).toEqual([
      {type:"writeFormulas",sheet:"目标",range:"D3:F3",formulas:[["='O''Brien'!B2","='O''Brien'!B3","='O''Brien'!B5"]]},
      {type:"writeFormulas",sheet:"目标",range:"D4:F4",formulas:[["='O''Brien'!C2","='O''Brien'!C3","='O''Brien'!C5"]]}
    ]);
  });
  it.each([{rows:[1,1]},{rows:[-1]},{rows:[4]},{rows:[0.5]},{rows:[]}])("rejects invalid source row selection $rows", ({rows}) => {
    expect(() => copyRangePlan(action,[["k","v"],["a",1],["b",2],["a",3]], rows)).toThrow();
  });
});
