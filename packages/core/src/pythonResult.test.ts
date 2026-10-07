import { describe, expect, it } from "vitest";
import { compilePythonValuesPlan, checkedMatrixResult } from "./pythonResult";
import type { PythonDataScope } from "./contracts";

const scope: PythonDataScope = { version:1, taskId:"task", hostId:"excel-office",
  workbookIdentityHash:"a".repeat(64), authorizationVersion:"b".repeat(64),
  sources:[{sheetId:"1",sheetName:"Sheet1",range:"A1:B2"}] };
const receipt = { resultHandle:"result://"+"r".repeat(43), contentFingerprint:"c".repeat(64), rows:1, columns:2 };
const target = { sheetId:"1",sheetName:"Sheet1",range:"E2:F2" };
const result = { kind:"matrix" as const,values:[[3,"PRIVATE_RESULT"]] };

describe("local-only Python values compiler", () => {
  it.each(["excel-office","wps-et"] as const)("same whitelist for %s", (hostId) => {
    const plan = compilePythonValuesPlan({...scope,hostId},receipt,result,target);
    expect(plan.actions).toEqual([{type:"writeValues",sheet:"Sheet1",range:"E2:F2",values:result.values}]);
    expect(plan.acceptanceCriteria).toEqual([{type:"rangeEquals",sheet:"Sheet1",range:"E2:F2",expected:result.values}]);
    expect(plan.title+plan.summary+plan.warnings.join("")).not.toContain("PRIVATE_RESULT");
    expect(JSON.stringify(plan)).not.toContain("result://");
  });
  it("does not let output choose targets/actions/formulas", () => {
    for (const value of ["=1", "+SUM(1)", "-cmd|x", "@SUM(1)", " \t=HYPERLINK(1)"]) {
      expect(() => compilePythonValuesPlan(scope,receipt,{kind:"matrix",values:[[1,value]]},target)).toThrow();
    }
    expect(() => compilePythonValuesPlan(scope,receipt,{...result,actions:[]} as never,target)).toThrow();
    expect(() => compilePythonValuesPlan(scope,receipt,result,{...target,sheetId:"2"})).toThrow();
  });
  it("requires exact shape and creates independent matrices", () => {
    expect(() => compilePythonValuesPlan(scope,receipt,result,{...target,range:"E2:E2"})).toThrow();
    const plan = compilePythonValuesPlan(scope,receipt,result,target);
    plan.actions[0].values[0][0] = 99;
    expect(result.values[0][0]).toBe(3);
    expect(plan.acceptanceCriteria[0].expected[0][0]).toBe(3);
  });
  it.each([NaN,Infinity,2**53,"\ud800",{},"x".repeat(32768)])("rejects unsafe scalar independently: %#",(value) => {
    expect(() => checkedMatrixResult({kind:"matrix",values:[[1,value]]},receipt)).toThrow();
  });
  it("supports valid surrogate pairs, booleans/null and numeric negatives",() => {
    expect(checkedMatrixResult({kind:"matrix",values:[[null,true,-1,0.25,"😀"]]}, {...receipt,columns:5}).values)
      .toEqual([[null,true,-1,0.25,"😀"]]);
  });
  it.each(["E0:F0","XFE1:XFF1","E1048577:F1048577","E2:D2","e2:f2","Sheet1!E2:F2","E:F"]) ("rejects ambiguous/out-of-bounds target %s",(range) => {
    expect(() => compilePythonValuesPlan(scope,receipt,result,{...target,range})).toThrow();
  });
});
