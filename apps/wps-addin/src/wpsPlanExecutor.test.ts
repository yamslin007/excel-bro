import { describe, expect, it } from "vitest";
import type { AnalysisPlan, CellValue } from "../../excel-addin/src/contracts";
import type { WpsApplication, WpsRange, WpsWorksheet } from "./wpsHost";
import type { WpsReadAuthorization } from "./wpsWorkbookReader";
import {
  executeWpsPlan,
  latestWpsUndo,
  matrixFingerprintV2,
  prepareWpsPlan,
  undoWpsExecution,
  wpsWorkbookLocator
} from "./wpsPlanExecutor";

function clone(matrix: CellValue[][]): CellValue[][] {
  return matrix.map((row) => [...row]);
}

function host(
  initial: CellValue[][],
  evaluateFormula?: (formula: CellValue) => CellValue,
  normalizeWrittenFormula?: (formula: CellValue) => CellValue
) {
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
        const requested = Array.isArray(next) ? clone(next as CellValue[][]) : [[next]];
        formulas = requested.map((row) => row.map((formula) => normalizeWrittenFormula?.(formula) ?? formula));
        values = formulas.map((row) => row.map((formula) => evaluateFormula?.(formula) ?? formula));
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
    Version: "12.0",
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
    overwrite: (next: CellValue[][]) => { values = clone(next); formulas = clone(next); },
    storedRecords: () => JSON.parse(String(storage.get("excel-bro.wps.undo.v1") ?? "[]")),
    replaceRecords: (records: unknown) => { storage.set("excel-bro.wps.undo.v1", JSON.stringify(records)); }
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

function hostWithSource() {
  const target = host([["旧目标"]]);
  let source = "原类别";
  const sheet = target.application.ActiveWorkbook!.Worksheets.Item(1);
  const targetRange = sheet.Range!("A1");
  const sourceRange: WpsRange = { MergeCells: false };
  Object.defineProperties(sourceRange, {
    Value2: { get: () => source, enumerable: true },
    Formula: { get: () => source, enumerable: true }
  });
  sheet.Range = (address) => address === "B1" ? sourceRange : targetRange;
  return { ...target, changeSource: (value: string) => { source = value; } };
}

describe("WPS guarded plan executor", () => {
  it("allows the observed 12.0 JSAPI in development but rechecks it before execution", async () => {
    const target = host([["旧值"]]);
    target.application.Version = "12.0";
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "writeValues", sheet: "明细", range: "A1", values: [["新值"]] }
    ]));
    expect(target.values()).toEqual([["旧值"]]);
    target.application.Version = "11.0";
    await expect(executeWpsPlan(target.application, prepared)).rejects.toThrow("宿主版本已变化");
    expect(target.values()).toEqual([["旧值"]]);
    target.application.Version = "12.0";
    const record = await executeWpsPlan(target.application, prepared);
    expect(record.state).toBe("verified");
    expect(target.values()).toEqual([["新值"]]);
  });

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

  it("previews and verifies A1 formulas by Formula rather than their calculated Value2, then undoes", async () => {
    const target = host([["旧值"]], (formula) => formula === "=1+1" ? 2 : formula);
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "writeFormulas", sheet: "明细", range: "A1", formulas: [["=1+1"]] }
    ]));
    expect(prepared.actions[0].preview).toEqual([["=1+1"]]);
    expect(target.values()).toEqual([["旧值"]]);

    const record = await executeWpsPlan(target.application, prepared);
    expect(record.state).toBe("verified");
    expect(target.values()).toEqual([[2]]);
    expect(record.actions[0].afterFingerprint).toMatch(/^v2:/);

    const restored = await undoWpsExecution(target.application, record.executionId);
    expect(restored.state).toBe("restored");
    expect(target.values()).toEqual([["旧值"]]);
  });

  it("does not mark a formula write verified when Formula readback differs", async () => {
    const target = host([["旧值"]], undefined, (formula) => formula === "=1+1" ? "=1+2" : formula);
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "writeFormulas", sheet: "明细", range: "A1", formulas: [["=1+1"]] }
    ]));

    await expect(executeWpsPlan(target.application, prepared)).rejects.toThrow("回读结果与计划不一致");
    expect(target.storedRecords()[0].state).toBe("partial_write_recoverable");
    const restored = await undoWpsExecution(target.application, target.storedRecords()[0].executionId);
    expect(restored.state).toBe("restored");
    expect(target.values()).toEqual([["旧值"]]);
  });

  it("rejects unsafe or unsupported formula plans before creating a recovery record", () => {
    const target = host([["旧值"]]);
    expect(() => prepareWpsPlan(target.application, [1], plan([
      { type: "writeFormulas", sheet: "明细", range: "A1", formulas: [["=WEBSERVICE(\"https://example.com\")"]] }
    ]))).toThrow("WEBSERVICE");
    expect(() => prepareWpsPlan(target.application, [1], plan([
      { type: "writeFormulas", sheet: "明细", range: "A1:B1", formulas: [["=1+1"]] }
    ]))).toThrow("写入矩阵");
    expect(() => prepareWpsPlan(target.application, [1], plan([
      { type: "writeFormulas", sheet: "明细", range: "A1", formulas: [["=1+1"]], formulaR1C1: "=R[0]C[-1]" }
    ]))).toThrow("R1C1");
    expect(target.values()).toEqual([["旧值"]]);
    expect(target.storedRecords()).toEqual([]);
  });

  it("rejects a changed target after preview, before creating a recovery record or writing", async () => {
    const target = host([["旧值"]]);
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "writeValues", sheet: "明细", range: "A1", values: [["新值"]] }
    ]));
    target.overwrite([["用户改过"]]);

    await expect(executeWpsPlan(target.application, prepared)).rejects.toThrow("重新预览");
    expect(target.values()).toEqual([["用户改过"]]);
    expect(target.storedRecords()).toEqual([]);
  });

  it("rejects a plan changed after preview instead of executing stale confirmation", async () => {
    const target = host([["旧值"]]);
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "writeValues", sheet: "明细", range: "A1", values: [["新值"]] }
    ]));
    prepared.actions[0].values[0][0] = "未预览的值";

    await expect(executeWpsPlan(target.application, prepared)).rejects.toThrow("重新预览");
    expect(target.values()).toEqual([["旧值"]]);
    expect(target.storedRecords()).toEqual([]);
  });

  it("binds a locally computed mapping to its source values from computation through execution", async () => {
    const target = hostWithSource();
    const mappingPlan = plan([
      { type: "writeValues", sheet: "明细", range: "A1", values: [["新目标"]] }
    ]);
    const sourceSnapshot = [{ sheetName: "明细", range: "B1", values: [["原类别"]] }];
    target.changeSource("计算后变化");
    expect(() => prepareWpsPlan(target.application, [1], mappingPlan, sourceSnapshot)).toThrow("重新生成计划");
    target.changeSource("原类别");
    const prepared = prepareWpsPlan(target.application, [1], mappingPlan, sourceSnapshot);
    target.changeSource("预览后变化");
    await expect(executeWpsPlan(target.application, prepared)).rejects.toThrow("源区域");
    expect(target.values()).toEqual([["旧目标"]]);
    expect(target.storedRecords()).toEqual([]);
  });

  it("new records use typed v2 fingerprints and distinguish empty string from null on undo", async () => {
    const target = host([[null]]);
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "writeValues", sheet: "明细", range: "A1", values: [[""]] }
    ]));
    const record = await executeWpsPlan(target.application, prepared);
    expect(record.version).toBe(2);
    expect(record.actions[0].beforeFingerprint).toMatch(/^v2:/);
    expect(record.actions[0].afterFingerprint).toMatch(/^v2:/);

    target.overwrite([[null]]);
    await expect(undoWpsExecution(target.application, record.executionId)).rejects.toThrow("又被修改");
    expect(target.values()).toEqual([[null]]);
  });

  it("keeps v1 records readable but refuses ambiguous null/empty undo without deleting recovery data", async () => {
    const target = host([["原值"]]);
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "writeValues", sheet: "明细", range: "A1", values: [["新值"]] }
    ]));
    const record = await executeWpsPlan(target.application, prepared);
    const records = target.storedRecords();
    records[0].version = 1;
    records[0].actions[0].beforeFingerprint = legacyFingerprint([["原值"]]);
    records[0].actions[0].afterFingerprint = legacyFingerprint([["新值"]]);
    target.replaceRecords(records);
    expect((await undoWpsExecution(target.application, record.executionId)).state).toBe("restored");

    const second = await executeWpsPlan(target.application, prepareWpsPlan(target.application, [1], plan([
      { type: "writeValues", sheet: "明细", range: "A1", values: [[null]] }
    ])));
    const ambiguous = target.storedRecords();
    ambiguous[0].version = 1;
    ambiguous[0].actions[0].beforeFingerprint = legacyFingerprint([["原值"]]);
    ambiguous[0].actions[0].afterFingerprint = legacyFingerprint([[null]]);
    target.replaceRecords(ambiguous);
    target.overwrite([[""]]);
    await expect(undoWpsExecution(target.application, second.executionId)).rejects.toThrow("旧版恢复点");
    expect(target.values()).toEqual([[""]]);
    expect(latestWpsUndo(target.application)?.state).toBe("undo_conflict");
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

describe("WPS versioned undo fingerprints", () => {
  it("locks typed v2 vectors and prioritizes formulas", () => {
    expect(matrixFingerprintV2([[null], [""]], [[null], [null]])).toBe("v2:28:c7b970ff");
    expect(matrixFingerprintV2([[2]], [["=1+1"]])).toBe("v2:22:503c7d21");
    expect(matrixFingerprintV2([[2]], [["=1+1"]])).toBe(matrixFingerprintV2([[3]], [["=1+1"]]));
    expect(matrixFingerprintV2([[null]], [[null]])).not.toBe(matrixFingerprintV2([[""]], [[null]]));
  });
});

function legacyFingerprint(values: CellValue[][]): string {
  const text = JSON.stringify(values.map((row) => row.map((value) => ({ value: value === "" ? null : value }))));
  let hash = 0x811c9dc5;
  for (let index = 0; index < text.length; index += 1) {
    hash ^= text.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return `${text.length}:${(hash >>> 0).toString(16).padStart(8, "0")}`;
}

describe("wpsWorkbookLocator", () => {
  // locator 的 value 必须与已落盘 WpsUndoRecord.workbookIdentity 逐字相同，
  // 否则现有撤销记录会因身份对不上而失效。这里把两者锁在一起。
  function authorization(fullName: string | null): WpsReadAuthorization {
    return {
      workbookName: "book.xlsx",
      workbookFullName: fullName,
      workbookPath: null,
      sheets: [],
      workbookRef: {} as WpsReadAuthorization["workbookRef"]
    };
  }

  it("给出 wps-path 判别式，值沿用既有小写归一且保留反斜杠", () => {
    expect(wpsWorkbookLocator(authorization("C:\\Users\\Jia\\Book1.xlsx"))).toEqual({
      kind: "wps-path",
      value: "c:\\users\\jia\\book1.xlsx"
    });
  });

  it("与执行时写入撤销记录的身份逐字一致", async () => {
    const target = host([["旧值"]]);
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "writeValues", sheet: "明细", range: "A1", values: [["新值"]] }
    ]));
    const record = await executeWpsPlan(target.application, prepared);
    // host() 夹具的 FullName 常量，与执行路径读到的是同一个值。
    const locator = wpsWorkbookLocator(authorization("C:\\Data\\测试.xlsx"));
    expect(locator.value).toBe(record.workbookIdentity);
  });

  it("未保存或非 .xlsx 时拒绝，不生成占位身份", () => {
    expect(() => wpsWorkbookLocator(authorization(null))).toThrow("已保存到本地磁盘");
    expect(() => wpsWorkbookLocator(authorization("C:\\a\\b.et"))).toThrow(
      "已保存到本地磁盘"
    );
  });
});
