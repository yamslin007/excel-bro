import { describe, expect, it } from "vitest";
import { compileControlledScript } from "@excel-bro/core/controlledScript";
import { assertEditablePlan } from "./contracts";
import { executePlan, undoExecution } from "./excel";
import { createOfficeTestHost, installOfficeTestHost } from "./officeTestHost";

describe("controlled JS to Excel executor", () => {
  it("writes and restores both authorized sheets; never writes the unselected sheet", async () => {
    const host = createOfficeTestHost({sheets:[{name:"One",values:[["before1"]]},{name:"Two",values:[["before2"]]},{name:"Secret",values:[["private"]]}]});
    installOfficeTestHost(host);
    const plan: unknown = {id:"script-test",title:"多表",summary:"多表",assumptions:[],warnings:[],
      actions:compileControlledScript('for (const sheet of authorizedSheets) emit({type:"setFont",sheet,range:"A1",bold:true});',["One","Two"])};
    assertEditablePlan(plan);
    expect(host.writes()).toHaveLength(0);
    const result = await executePlan(plan, {authorizedSheetNames:["One","Two"]});
    expect(host.cellStyle("One","A1")?.fontBold).toBe(true);
    expect(host.cellStyle("Two","A1")?.fontBold).toBe(true);
    expect(host.writes().every(write => write.sheet !== "Secret")).toBe(true);
    await undoExecution(result.undoSnapshot!);
    expect(host.cellStyle("One","A1")?.fontBold).not.toBe(true);
    expect(host.cellStyle("Two","A1")?.fontBold).not.toBe(true);
  });
  it("still rejects scope revoked after compiling", async () => {
    const host = createOfficeTestHost({sheets:[{name:"One",values:[[1]]},{name:"Two",values:[[2]]}]});
    installOfficeTestHost(host);
    const plan: unknown = {id:"script-revoke",title:"多表",summary:"",assumptions:[],warnings:[],
      actions:compileControlledScript('for (const sheet of authorizedSheets) emit({type:"setFont",sheet,range:"A1",bold:true});',["One","Two"])};
    assertEditablePlan(plan);
    await expect(executePlan(plan, {authorizedSheetNames:["One"]})).rejects.toThrow();
    expect(host.writes()).toHaveLength(0);
  });
});
