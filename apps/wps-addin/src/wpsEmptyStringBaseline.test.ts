/**
 * 阶段 G 的空串负向基线在阶段 4 转为真实行为断言。
 *
 * 指纹为旧撤销记录继续折叠空串，但计划值与独立回读必须区分 "" 和 null。
 */
import { describe, expect, it } from "vitest";
import type { AnalysisPlan, CellValue } from "../../excel-addin/src/contracts";
import type { WpsApplication, WpsRange, WpsWorksheet } from "./wpsHost";
import { executeWpsPlan, latestWpsUndo, prepareWpsPlan } from "./wpsPlanExecutor";

function clone(matrix: CellValue[][]): CellValue[][] {
  return matrix.map((row) => [...row]);
}

/**
 * @param collapseEmptyOnWrite 模拟无法保真空串的宿主：写入 "" 后回读得到 null。
 *   openpyxl 保存 .xlsx 后重新加载同样会把 "" 读回 None，所以这是真实存在的宿主行为。
 */
function host(initial: CellValue[][], collapseEmptyOnWrite = false) {
  let values = clone(initial);
  let formulas = clone(initial);
  const storage = new Map<string, unknown>();
  const persist = (next: CellValue[][]): CellValue[][] =>
    collapseEmptyOnWrite
      ? next.map((row) => row.map((cell) => (cell === "" ? null : cell)))
      : clone(next);
  const range: WpsRange = {
    MergeCells: false,
    Rows: { Count: initial.length },
    Columns: { Count: initial[0].length }
  };
  Object.defineProperties(range, {
    Value2: {
      get: () => clone(values),
      set: (next: CellValue[][] | CellValue) => {
        values = persist(Array.isArray(next) ? (next as CellValue[][]) : [[next]]);
        formulas = clone(values);
      },
      enumerable: true
    },
    Formula: {
      get: () => clone(formulas),
      set: (next: CellValue[][] | CellValue) => {
        formulas = persist(Array.isArray(next) ? (next as CellValue[][]) : [[next]]);
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
  const application: WpsApplication = {
    Version: "12.0",
    ActiveWorkbook: {
      Name: "测试.xlsx",
      FullName: "C:\\Data\\测试.xlsx",
      Path: "C:\\Data",
      ReadOnly: false,
      Worksheets: { Count: 1, Item: () => sheet },
      ActiveSheet: sheet
    },
    ActiveSheet: sheet,
    PluginStorage: {
      getItem: (key) => storage.get(key),
      setItem: (key, value) => { storage.set(key, value); }
    },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  };
  return { application, values: () => clone(values) };
}

function plan(actions: AnalysisPlan["actions"]): AnalysisPlan {
  return {
    id: "plan-empty-baseline",
    title: "写入空串",
    summary: "验证空串写入语义",
    assumptions: [],
    warnings: [],
    actions,
    acceptanceCriteria: []
  };
}

const emptyStringPlan = plan([
  { type: "writeValues", sheet: "明细", range: "A1:A2", values: [["未分类"], [""]] }
]);

describe("WPS 空串写入与独立回读", () => {
  it("准备阶段保留计划中的空串", () => {
    const target = host([["旧值"], ["旧值"]]);
    const prepared = prepareWpsPlan(target.application, [1], emptyStringPlan);
    expect(prepared.actions[0].values).toEqual([["未分类"], [""]]);
  });

  it("宿主把空串折成 null 时拒绝报告 verified，仍保留可恢复记录", async () => {
    const target = host([["旧值"], ["旧值"]], true);
    const prepared = prepareWpsPlan(target.application, [1], emptyStringPlan);
    await expect(executeWpsPlan(target.application, prepared)).rejects.toThrow("回读结果与计划不一致");
    const record = latestWpsUndo(target.application);
    expect(record?.state).toBe("partial_write_recoverable");
    expect(target.values()).toEqual([["未分类"], [null]]);
  });

  it("宿主保留空串时回读可以准确通过", async () => {
    const target = host([["旧值"], ["旧值"]]);
    const prepared = prepareWpsPlan(target.application, [1], emptyStringPlan);
    const record = await executeWpsPlan(target.application, prepared);
    expect(record.state).toBe("verified");
    expect(target.values()).toEqual([["未分类"], [""]]);
  });

  it("参照组：不含空串的写入在同一宿主上正常验证通过", async () => {
    const target = host([["旧值"], ["旧值"]], true);
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "writeValues", sheet: "明细", range: "A1:A2", values: [["未分类"], ["费用"]] }
    ]));
    const record = await executeWpsPlan(target.application, prepared);
    expect(record.state).toBe("verified");
    expect(target.values()).toEqual([["未分类"], ["费用"]]);
  });
});
