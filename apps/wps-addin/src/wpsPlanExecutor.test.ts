import { describe, expect, it } from "vitest";
import type { AnalysisPlan, CellValue } from "../../excel-addin/src/contracts";
import type { WpsApplication, WpsRange, WpsWorksheet } from "./wpsHost";
import {
  executeWpsPlan,
  latestWpsUndo,
  prepareWpsPlan,
  undoWpsExecution
} from "./wpsPlanExecutor";

function clone(matrix: CellValue[][]): CellValue[][] {
  return matrix.map((row) => [...row]);
}

function host(initial: CellValue[][]) {
  let values = clone(initial);
  let formulas = clone(initial);
  const storage = new Map<string, unknown>();
  const range: WpsRange = {
    MergeCells: false,
    Rows: { Count: initial.length },
    Columns: { Count: initial[0].length }
  };
  Object.defineProperties(range, {
    Value2: {
      get: () => clone(values),
      set: (next: CellValue[][] | CellValue) => {
        values = Array.isArray(next) ? clone(next as CellValue[][]) : [[next]];
        formulas = clone(values);
      },
      enumerable: true
    },
    Formula: {
      get: () => clone(formulas),
      set: (next: CellValue[][] | CellValue) => {
        formulas = Array.isArray(next) ? clone(next as CellValue[][]) : [[next]];
        values = clone(formulas);
      },
      enumerable: true
    }
  });
  const sheet: WpsWorksheet = {
    Name: "明细",
    ProtectContents: false,
    Range: () => range
  };
  const workbook = {
    Name: "测试.xlsx",
    FullName: "C:\\Data\\测试.xlsx",
    Path: "C:\\Data",
    ReadOnly: false,
    Worksheets: { Count: 1, Item: () => sheet },
    ActiveSheet: sheet
  };
  const application: WpsApplication = {
    Version: "12.1.0.28488",
    ActiveWorkbook: workbook,
    ActiveSheet: sheet,
    PluginStorage: {
      getItem: (key) => storage.get(key),
      setItem: (key, value) => { storage.set(key, value); }
    },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  };
  return {
    application,
    values: () => clone(values),
    overwrite: (next: CellValue[][]) => { values = clone(next); formulas = clone(next); }
  };
}

function plan(actions: AnalysisPlan["actions"]): AnalysisPlan {
  return {
    id: "plan-1",
    title: "更新明细",
    summary: "更新已授权区域",
    assumptions: [],
    warnings: [],
    actions,
    acceptanceCriteria: []
  };
}

describe("WPS guarded plan executor", () => {
  it("does not write before confirmation, then verifies and restores the snapshot", async () => {
    const target = host([["旧值", 1], ["保留", 2]]);
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "writeValues", sheet: "明细", range: "B2:C3", values: [["新值", 10], ["更新", 20]] }
    ]));

    expect(target.values()).toEqual([["旧值", 1], ["保留", 2]]);
    const record = await executeWpsPlan(target.application, prepared);
    expect(record.state).toBe("verified");
    expect(target.values()).toEqual([["新值", 10], ["更新", 20]]);
    expect(latestWpsUndo(target.application)?.executionId).toBe(record.executionId);

    const restored = await undoWpsExecution(target.application, record.executionId);
    expect(restored.state).toBe("restored");
    expect(target.values()).toEqual([["旧值", 1], ["保留", 2]]);
  });

  it("supports deleting cell contents through the same snapshot-safe primitive", async () => {
    const target = host([["需要清空", 1], ["仍需清空", 2]]);
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "clearRange", sheet: "明细", range: "B2:C3", applyTo: "contents" }
    ]));
    const record = await executeWpsPlan(target.application, prepared);
    expect(target.values()).toEqual([[null, null], [null, null]]);
    await undoWpsExecution(target.application, record.executionId);
    expect(target.values()).toEqual([["需要清空", 1], ["仍需清空", 2]]);
  });

  it("rejects structural deletion and unauthorized actions before touching the workbook", () => {
    const target = host([["原值"]]);
    expect(() => prepareWpsPlan(target.application, [1], plan([
      { type: "deleteWorksheet", sheet: "明细" }
    ]))).toThrow("当前不执行");
    expect(target.values()).toEqual([["原值"]]);
  });

  it("refuses undo when the verified result was modified again", async () => {
    const target = host([["旧值"]]);
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "writeValues", sheet: "明细", range: "A1", values: [["新值"]] }
    ]));
    const record = await executeWpsPlan(target.application, prepared);
    target.overwrite([["用户后来修改"]]);

    await expect(undoWpsExecution(target.application, record.executionId)).rejects.toThrow("又被修改");
    expect(target.values()).toEqual([["用户后来修改"]]);
  });
});
