import { describe, expect, it } from "vitest";
import { compileControlledScript } from "@excel-bro/core/controlledScript";
import { assertEditablePlan } from "../../excel-addin/src/contracts";
import { captureWpsBeforeResult, stageWpsBeforeComparison } from "./wpsNativeComparison";
import {
  canClaimWriteSuccess,
  cellDisplayText,
  type RangeChangeReview
} from "@excel-bro/core/changeReview";
import type { AnalysisPlan, CellValue } from "../../excel-addin/src/contracts";
import type { WpsApplication, WpsRange, WpsWorksheet } from "./wpsHost";
import type { WpsReadAuthorization } from "./wpsWorkbookReader";
import {
  WPS_SUPPORTED_ACTION_KEYS,
  executeWpsPlan,
  latestWpsUndo,
  matrixFingerprintV2,
  prepareWpsPlan,
  toChangeFormulas,
  undoWpsExecution,
  wpsActionChangeReview,
  wpsWorkbookLocator,
  type StoredAction
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
  it("controlled JS executes and restores two selected sheets; optional comparison does not block capture", async () => {
    const one = host([["before1"]]), two = host([["before2"]]), secret = host([["private"]]);
    const sheets = [one, two, secret].map((fixture, i) => {
      const sheet = fixture.application.ActiveWorkbook!.Worksheets.Item(1);
      sheet.Name = ["One", "Two", "Secret"][i];
      return sheet;
    });
    one.application.ActiveWorkbook!.Worksheets = {Count:3, Item:(index: number) => sheets[index - 1]};
    const scriptPlan: unknown = plan([]);
    Object.assign(scriptPlan!, {actions:compileControlledScript('for(const sheet of authorizedSheets) emit({type:"writeValues",sheet,range:"A1",values:[[sheet]]});',["One","Two"])});
    assertEditablePlan(scriptPlan);
    const prepared = prepareWpsPlan(one.application, [1,2], scriptPlan);
    let before!: ReturnType<typeof captureWpsBeforeResult>;
    expect(() => { before = captureWpsBeforeResult(one.application, prepared); }).not.toThrow();
    expect(before.values.map(value => value.values)).toEqual([[["before1"]], [["before2"]]]);
    expect(one.values()).toEqual([["before1"]]);
    const record = await executeWpsPlan(one.application, prepared);
    expect(one.values()).toEqual([["One"]]);
    expect(two.values()).toEqual([["Two"]]);
    expect(secret.values()).toEqual([["private"]]);
    expect(() => stageWpsBeforeComparison(one.application, before)).toThrow("单张工作表");
    await undoWpsExecution(one.application, record.executionId);
    expect(one.values()).toEqual([["before1"]]);
    expect(two.values()).toEqual([["before2"]]);
    expect(() => prepareWpsPlan(one.application, [1], scriptPlan)).toThrow();
  });
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

  it("空的 formulaR1C1 等于没有 R1C1：与 Excel、服务端同一口径，照常写 A1 公式", async () => {
    /*
     * 同一份计划只差 `formulaR1C1` 这一个字段的值：非空 ⇒ 拒绝（上面那条），
     * `""` / `null` / 省略 ⇒ 三种都必须照常写 A1 公式。以前这里是
     * `!== undefined`，`""` 会被当成 R1C1 整份拒掉，而 Excel 与服务端都按真值判。
     */
    for (const formulaR1C1 of ["", null, undefined] as const) {
      const target = host([["旧值"]]);
      const prepared = prepareWpsPlan(target.application, [1], plan([
        { type: "writeFormulas", sheet: "明细", range: "A1", formulas: [["=1+1"]], ...(formulaR1C1 === undefined ? {} : { formulaR1C1 }) }
      ]));
      await executeWpsPlan(target.application, prepared);
      expect(target.values()).toEqual([["=1+1"]]);
    }
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

  it("同一份计划执行过一次后，再点一次确认也不会二次写入（预览指纹守卫拦在写入之前）", async () => {
    const target = host([["旧值"]]);
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "writeValues", sheet: "明细", range: "A1", values: [["新值"]] }
    ]));
    const first = await executeWpsPlan(target.application, prepared);
    expect(target.values()).toEqual([["新值"]]);

    // 计划对象本身没被动过（confirmationSnapshot 依然匹配），能挡住第二次的是
    // 「目标区域已不是预览那一刻的样子」这条复核 —— 见 wpsPlanExecutor 保存恢复点前
    // 对每个动作 previewFingerprint 的比对。
    await expect(executeWpsPlan(target.application, prepared)).rejects.toThrow("重新预览");
    expect(target.values()).toEqual([["新值"]]);
    // 被拦下来的第二次没有留下新的恢复点，撤销入口仍指向真正写入的那一次。
    expect(target.storedRecords()).toHaveLength(1);
    expect(latestWpsUndo(target.application)?.executionId).toBe(first.executionId);
  });

  it("回读不一致失败后，同一份计划同样不允许直接重跑", async () => {
    const target = host([["旧值"]], undefined, (formula) => (formula === "=1+1" ? "=1+2" : formula));
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "writeFormulas", sheet: "明细", range: "A1", formulas: [["=1+1"]] }
    ]));
    await expect(executeWpsPlan(target.application, prepared)).rejects.toThrow("回读结果与计划不一致");
    const afterFirstFailure = target.values();

    // 第一次已经写进去了（只是与计划不符），区域指纹因此不再是预览时的样子，
    // 第二次执行必须被拒绝，而不是"再写一遍盖上去"。
    await expect(executeWpsPlan(target.application, prepared)).rejects.toThrow("重新预览");
    expect(target.values()).toEqual(afterFirstFailure);
    expect(target.storedRecords()).toHaveLength(1);
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

  /**
   * 常用格式操作在 WPS 端的接通状态（Task #53）：
   * - setFont / setAlignment / resizeRange **已接通**（但只开放已实机验证的属性子集，
   *   未验证属性在预检里明确拒绝，见 `wpsFormatCapability.test.ts`）。
   * - setFill / setNumberFormat 缺少快照时拒绝；setBorders 缺少逐边颜色来源快照时拒绝，autofit 仍未接通，必须保持 fail-closed：
   *   在动工作簿**之前**整份计划被拒，且不留下任何恢复点。
   *
   * 这条用例的作用是让"顺手把配置放开了"这件事立刻变红 —— 而不是等到实机上
   * 出现"计划显示成功、表格格式没变"。
   */
  it("refuses missing format metadata and still-unwired actions before touching the workbook", () => {
    const target = host([["表头", 1200]]);
    const unwiredActions: AnalysisPlan["actions"] = [
      { type: "setFill", sheet: "明细", range: "A1:B1", color: "#DDEEFF" },
      { type: "setNumberFormat", sheet: "明细", range: "B2", formatCode: "#,##0.00" },
      { type: "setBorders", sheet: "明细", range: "A1:B1", sides: ["top", "bottom"] },
      { type: "autofit", sheet: "明细", range: "A1:B1" }
    ];
    for (const action of unwiredActions) {
      expect(() => prepareWpsPlan(target.application, [1], plan([action]))).toThrow(
        action.type === "setFill" || action.type === "setNumberFormat" || action.type === "setBorders" ? /无法完整逐格采集/ : /当前不执行/
      );
    }
    expect(target.values()).toEqual([["表头", 1200]]);
    // 没有任何写入 ⇒ 不允许留下恢复点；否则撤销列表里会出现一份"什么也没改"的记录。
    expect(target.storedRecords()).toEqual([]);
  });

  /**
   * 配置与执行的漂移检测。
   *
   * `capabilities.wpsWrite.allowedActions` 是配置，`WPS_SUPPORTED_ACTION_KEYS`
   * 是执行器真实实现的集合。两边任意一侧单独改动都会让这条用例变红 ——
   * 这正是"配置与模型能力目录只声明实际已接通的能力"的机器判据。
   */
  it("keeps the write allowlist and the implemented action set in lockstep", async () => {
    const capabilities = (await import("../../../config/capabilities.json")).default;
    const allowed = capabilities.wpsWrite.allowedActions as readonly string[];
    expect([...allowed].sort()).toEqual([...WPS_SUPPORTED_ACTION_KEYS].sort());
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

/** 直接把某个格子写成公式，用来造"修改前是公式"的现场。 */
function seedFormula(target: ReturnType<typeof host>, address: string, formula: string): void {
  const sheet = target.application.ActiveWorkbook!.Worksheets.Item(1);
  const range = sheet.Range!(address);
  range.Formula = formula;
}

function storedAction(partial: Partial<StoredAction> & Pick<StoredAction, "range">): StoredAction {
  const rows = partial.beforeValues?.length ?? 1;
  const columns = partial.beforeValues?.[0]?.length ?? 1;
  return {
    actionIndex: 0,
    type: "writeValues",
    sheetIndex: 1,
    sheetName: "明细",
    rowCount: rows,
    columnCount: columns,
    cellCount: rows * columns,
    values: [],
    preview: [],
    previewFingerprint: "preview",
    rect: { startRow: 1, startColumn: 1, endRow: rows, endColumn: columns },
    beforeValues: [[null]],
    beforeValuesCaptured: true,
    beforeFormulas: [[null]],
    beforeFingerprint: "before",
    afterFingerprint: null,
    appliedRows: rows,
    ...partial
  };
}

describe("WPS 双栏变更对照（实际变更）", () => {
  it("写入核对通过后给出实际变更：phase=actual、status=applied、前后取自本次读写", async () => {
    const target = host([["旧值"]]);
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "writeValues", sheet: "明细", range: "A1", values: [["新值"]] }
    ]));
    let delivered: { reviews: RangeChangeReview[]; executionId: string } | null = null;

    const record = await executeWpsPlan(target.application, prepared, undefined, (reviews, executionId) => {
      delivered = { reviews, executionId };
    });

    expect(delivered).not.toBeNull();
    const { reviews, executionId } = delivered!;
    expect(executionId).toBe(record.executionId);
    expect(reviews).toHaveLength(1);
    const [review] = reviews;
    expect(review.phase).toBe("actual");
    expect(review.status).toBe("applied");
    expect(canClaimWriteSuccess(review.status)).toBe(true);
    expect(review.rows[0].address).toBe("A1");
    expect(cellDisplayText(review.rows[0].before).text).toBe("旧值");
    expect(cellDisplayText(review.rows[0].after).text).toBe("新值");
    expect(review.changedCount).toBe(1);
  });

  it("空单元格与空字符串是两个状态，互转都要算一次改动", async () => {
    const target = host([[null]]);
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "writeValues", sheet: "明细", range: "A1", values: [[""]] }
    ]));
    let delivered: { reviews: RangeChangeReview[] } | null = null;
    await executeWpsPlan(target.application, prepared, undefined, (reviews) => {
      delivered = { reviews };
    });

    const row = delivered!.reviews[0].rows[0];
    expect(row.changed).toBe(true);
    expect(cellDisplayText(row.before).text).toBe("（空）");
    expect(cellDisplayText(row.after).text).toBe("（空字符串）");
  });

  it("修改前是公式且被本次覆盖时，标记覆盖并保留原公式与计算值", async () => {
    const target = host([["旧值"]], (formula) => (formula === "=1+1" ? 2 : formula));
    seedFormula(target, "A1", "=1+1");
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "writeValues", sheet: "明细", range: "A1", values: [["新值"]] }
    ]));
    let delivered: { reviews: RangeChangeReview[] } | null = null;
    await executeWpsPlan(target.application, prepared, undefined, (reviews) => {
      delivered = { reviews };
    });

    const row = delivered!.reviews[0].rows[0];
    expect(row.overwritesFormula).toBe(true);
    expect(row.kind).toBe("formula");
    // 值栏是计算值，公式原文另起一行 —— 不许拿公式顶替值。
    expect(cellDisplayText(row.before).text).toBe("2");
    expect(cellDisplayText(row.before).formula).toBe("=1+1");
    expect(cellDisplayText(row.after).formula).toBeNull();
  });

  it("回读与计划不一致时照样交付对照，且不得宣称写入成功；回读矩阵不落 PluginStorage", async () => {
    const target = host([["旧值"]], undefined, (formula) => (formula === "=1+1" ? "=1+2" : formula));
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "writeFormulas", sheet: "明细", range: "A1", formulas: [["=1+1"]] }
    ]));
    let delivered: { reviews: RangeChangeReview[]; executionId: string } | null = null;

    await expect(
      executeWpsPlan(target.application, prepared, undefined, (reviews, executionId) => {
        delivered = { reviews, executionId };
      })
    ).rejects.toThrow("回读结果与计划不一致");

    // §6.5：失败也要先给对照，不能只丢一句错误。
    expect(delivered).not.toBeNull();
    const review = delivered!.reviews[0];
    expect(review.status).toBe("readback_mismatch");
    expect(canClaimWriteSuccess(review.status)).toBe(false);
    // 现场如实展示：回读到的就是 =1+2，不是计划里的 =1+1。
    expect(cellDisplayText(review.rows[0].after).formula).toBe("=1+2");
    // 回读矩阵只活在内存里，落盘记录里不该出现它（计划值 =1+1 本来就在记录里）。
    expect(JSON.stringify(target.storedRecords())).not.toContain("=1+2");
  });

  it("清空也要渲染成一行改动，不被当成无变化过滤掉", async () => {
    const target = host([["需要清空"]]);
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "clearRange", sheet: "明细", range: "A1", applyTo: "contents" }
    ]));
    let delivered: { reviews: RangeChangeReview[] } | null = null;
    await executeWpsPlan(target.application, prepared, undefined, (reviews) => {
      delivered = { reviews };
    });

    const row = delivered!.reviews[0].rows[0];
    expect(row.kind).toBe("empty");
    expect(row.changed).toBe(true);
    expect(cellDisplayText(row.before).text).toBe("需要清空");
    expect(cellDisplayText(row.after).text).toBe("（空）");
  });

  it("一个动作都没写时整块标未执行，修改后不给假内容", () => {
    const action = storedAction({
      range: "A1:A2",
      beforeValues: [["甲"], ["乙"]],
      beforeFormulas: [[null], [null]],
      appliedRows: 0
    });
    const review = wpsActionChangeReview(action, null, null);
    expect(review.status).toBe("not_executed");
    expect(review.rows.every((row) => row.kind === "not_executed")).toBe(true);
    expect(review.changedCount).toBe(0);
    // 没回读过就说没回读过，不能拿空矩阵冒充"写成了空"。
    expect(cellDisplayText(review.rows[0].after).unresolved).toBe(true);
  });

  it("部分写入：已写的行照常比对，未写的行标未执行且不进改动计数", () => {
    const action = storedAction({
      range: "A1:A2",
      beforeValues: [["甲"], ["乙"]],
      beforeFormulas: [[null], [null]],
      appliedRows: 1
    });
    const review = wpsActionChangeReview(action, { values: [["丙"], ["乙"]], formulas: [[null], [null]], valuesCaptured: true }, true);
    expect(review.rows[0].kind).toBe("value");
    expect(review.rows[1].kind).toBe("not_executed");
    expect(review.changedCount).toBe(1);
  });

  it("宿主没采到显示值时只给占位，公式原文照常展示，不冒充值", () => {
    const action = storedAction({
      range: "A1",
      beforeValues: [["甲"]],
      beforeFormulas: [[null]]
    });
    const review = wpsActionChangeReview(
      action,
      { values: [[null]], formulas: [["=SUM(B:B)"]], valuesCaptured: false },
      true
    );
    const display = cellDisplayText(review.rows[0].after);
    expect(display.unresolved).toBe(true);
    expect(display.text).toBe("显示值未采集");
    expect(display.formula).toBe("=SUM(B:B)");
    expect(display.text).not.toContain("=SUM");
  });

  it("range.Formula 对常量返回常量本身，只有 = 开头才算公式", () => {
    expect(toChangeFormulas([["=A1+B1", "123", "abc", null, 7]])).toEqual([["=A1+B1", null, null, null, null]]);
  });

  it("明细条数超过配置上限时截断，但改动计数仍是全量", () => {
    const total = 260;
    const action = storedAction({
      range: `A1:A${total}`,
      rowCount: total,
      beforeValues: Array.from({ length: total }, () => [null]),
      beforeFormulas: Array.from({ length: total }, () => [null]),
      appliedRows: total
    });
    const review = wpsActionChangeReview(
      action,
      {
        values: Array.from({ length: total }, () => ["新值"]),
        formulas: Array.from({ length: total }, () => [null]),
        valuesCaptured: true
      },
      true
    );
    expect(review.truncated).toBe(true);
    expect(review.rows.length).toBeLessThan(total);
    // 计数走全量，不因为截断就少报。
    expect(review.changedCount).toBe(total);
    expect(review.totalCount).toBe(total);
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

/**
 * 旧计划不能直接写入：预览之后**身份**换了，计划就该作废。
 *
 * 与"源区域变了所以拒绝"是两件事 —— 后者是数据层面的指纹比对，这里钉的是
 * 更靠外的一层：换了个工作簿、或同一索引上的表被换掉，即使单元格内容一字
 * 未动，也不允许拿旧计划去写。少了这一层，"重新确认范围"就成了走过场。
 */
describe("预览之后的身份变化：旧计划不得写入", () => {
  const writePlan = () => plan([
    { type: "writeValues", sheet: "明细", range: "A1", values: [["新值"]] }
  ]);

  it("切换到另一个工作簿：拒绝执行，原表一格都不动", async () => {
    const target = host([["旧值"]]);
    const prepared = prepareWpsPlan(target.application, [1], writePlan());
    // 内容一字未动，只有工作簿身份变了 —— 指纹比对救不了这种情况。
    (target.application.ActiveWorkbook as { FullName: string }).FullName =
      "C:\\Data\\另一个.xlsx";
    (target.application.ActiveWorkbook as { Name: string }).Name = "另一个.xlsx";
    await expect(executeWpsPlan(target.application, prepared)).rejects.toThrow(
      "工作簿身份发生变化"
    );
    expect(target.values()).toEqual([["旧值"]]);
  });

  it("同一索引上的表被换掉：拒绝执行，原表一格都不动", async () => {
    const target = host([["旧值"]]);
    const prepared = prepareWpsPlan(target.application, [1], writePlan());
    const sheet = target.application.ActiveWorkbook!.Worksheets.Item(1);
    (sheet as { Name: string }).Name = "换过来的表";
    await expect(executeWpsPlan(target.application, prepared)).rejects.toThrow(
      "被移动、重命名或删除"
    );
    expect(target.values()).toEqual([["旧值"]]);
  });

  it("工作表被保护：拒绝执行，不先写后报", async () => {
    const target = host([["旧值"]]);
    const prepared = prepareWpsPlan(target.application, [1], writePlan());
    const sheet = target.application.ActiveWorkbook!.Worksheets.Item(1);
    (sheet as { ProtectContents: boolean }).ProtectContents = true;
    await expect(executeWpsPlan(target.application, prepared)).rejects.toThrow(
      "已被保护"
    );
    expect(target.values()).toEqual([["旧值"]]);
  });
});

describe("SDK fixed value provenance at the host boundary", () => {
  it.each([{}, {sourceFingerprint: ""}, {sourceFingerprint: "forged", sourceFingerprintSheets: []}])(
    "rejects missing/incomplete SDK anchors before any mutation: %j", async (fields) => {
      const fixture = host([["before"]]);
      const request = {...plan([{type: "writeValues", sheet: "明细", range: "A1", values: [["after"]]}]), ...fields};
      const prepared = prepareWpsPlan(fixture.application, [1], request, [], {requireSourceAnchor: true});
      await expect(executeWpsPlan(fixture.application, prepared)).rejects.toThrow("没有记录");
      expect(fixture.values()).toEqual([["before"]]);
      expect(fixture.storedRecords()).toEqual([]);
    }
  );
  it("cannot disable the SDK gate after preparation", async () => {
    const fixture = host([["before"]]);
    const prepared = prepareWpsPlan(fixture.application, [1], plan([
      {type: "writeValues", sheet: "明细", range: "A1", values: [["after"]]}
    ]), [], {requireSourceAnchor: true});
    prepared.sourceAnchorRequired = false;
    await expect(executeWpsPlan(fixture.application, prepared)).rejects.toThrow("计划或预览已变化");
    expect(fixture.values()).toEqual([["before"]]);
  });
});


describe("SDK anchor source scope is checked before host writes", () => {
  async function setup() {
    const {capabilityHost} = await import("./wpsCapabilityTestHost");
    const {wpsAuthorizedSourceFingerprint} = await import("./wpsWorkbookReader");
    const fixture = capabilityHost(["Source", "Target"]);
    const request = {...plan([{type:"writeValues",sheet:"Target",range:"A1",values:[[123]]}]),
      sourceFingerprint:wpsAuthorizedSourceFingerprint(fixture.application,[1]), sourceFingerprintSheets:["Source"]};
    return {fixture,request};
  }
  it("a source subset is compared as that subset, not all authorized sheets", async () => {
    const {fixture,request} = await setup();
    const prepared = prepareWpsPlan(fixture.application,[1,2],request,[],{requireSourceAnchor:true});
    await executeWpsPlan(fixture.application,prepared);
    expect(fixture.grids[1][0][0].value).toBe(123);
    expect(fixture.grids[0][0][0].value).toBe("old-0-0");
  });
  it("a source outside current authorization is refused without reading it", async () => {
    const {fixture,request} = await setup();
    const prepared = prepareWpsPlan(fixture.application,[2],request,[],{requireSourceAnchor:true});
    Object.defineProperty(fixture.sheets[0], "UsedRange", {get:()=>{throw new Error("unauthorized read");}});
    await expect(executeWpsPlan(fixture.application,prepared)).rejects.toThrow("来源范围与当前授权范围不一致");
    expect(fixture.writes()).toBe(0);
  });
  it("unreadable live source is not accepted as unchanged", async () => {
    const {fixture,request} = await setup();
    const prepared = prepareWpsPlan(fixture.application,[1,2],request,[],{requireSourceAnchor:true});
    Object.defineProperty(fixture.sheets[0], "UsedRange", {get:()=>{throw new Error("temporarily unavailable");}});
    await expect(executeWpsPlan(fixture.application,prepared)).rejects.toThrow("无法核对当前数据来源");
    expect(fixture.writes()).toBe(0);
  });
  it("local source snapshots do not exempt SDK fixed values from provenance", async () => {
    const {fixture,request} = await setup();
    const prepared = prepareWpsPlan(fixture.application,[1,2], {...request,sourceFingerprint:null},
      [{sheetName:"Source",range:"A1",values:[["old-0-0"]]}], {requireSourceAnchor:true});
    expect(prepared.sources).toHaveLength(1);
    await expect(executeWpsPlan(fixture.application,prepared)).rejects.toThrow("没有记录");
    expect(fixture.writes()).toBe(0);
  });
  it("a local result reference requires provenance even without caller options", async () => {
    const {fixture,request} = await setup();
    const prepared = prepareWpsPlan(fixture.application,[1,2], {...request,sourceFingerprint:null,localResultRef:"res_"+"a".repeat(32)});
    await expect(executeWpsPlan(fixture.application,prepared)).rejects.toThrow("没有记录");
    expect(fixture.writes()).toBe(0);
  });
});

it("recomputes multi-sheet anchors in their original order", async () => {
  const {capabilityHost} = await import("./wpsCapabilityTestHost");
  const {wpsAuthorizedSourceFingerprint} = await import("./wpsWorkbookReader");
  const fixture = capabilityHost(["SourceA", "SourceB", "Target"]);
  const request = {...plan([{type:"writeValues",sheet:"Target",range:"A1",values:[[123]]}]),
    sourceFingerprint:wpsAuthorizedSourceFingerprint(fixture.application,[2,1]),
    sourceFingerprintSheets:["SourceB","SourceA"]};
  expect(request.sourceFingerprint).not.toBeNull();
  const prepared = prepareWpsPlan(fixture.application,[1,2,3],request,[],{requireSourceAnchor:true});
  await executeWpsPlan(fixture.application,prepared);
  expect(fixture.grids[2][0][0].value).toBe(123);
});
