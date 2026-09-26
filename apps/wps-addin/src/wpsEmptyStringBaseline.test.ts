/**
 * 阶段 G 基线：空串写入语义的负向证据。
 *
 * 本文件只建立可重现的失败证据，不修复缺陷。修复在阶段 4
 * （见 docs/DUAL_HOST_REFACTOR_PROPOSAL_2026-09-26.md 第 1.1 节与第 3 节）。
 *
 * 用 `it.fails` 固定预期失败：断言一旦变为通过，vitest 会报错，
 * 迫使重新审查而不是静默转绿。这对应方案 §3.4 的 XPASS 要求。
 */
import { describe, expect, it } from "vitest";
import type { AnalysisPlan, CellValue } from "../../excel-addin/src/contracts";
import type { WpsApplication, WpsRange, WpsWorksheet } from "./wpsHost";
import { executeWpsPlan, prepareWpsPlan } from "./wpsPlanExecutor";

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
    Version: "12.1.0.28488",
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

describe("阶段 G 基线：空串写入语义（已知缺陷，预期失败）", () => {
  it.fails(
    "缺陷 1：prepareWpsPlan 应保留计划中的空串，实际在准备阶段就折成 null",
    () => {
      const target = host([["旧值"], ["旧值"]]);
      const prepared = prepareWpsPlan(target.application, [1], emptyStringPlan);
      // 失败原因锁定在 wpsPlanExecutor.ts:195 —— valuesForAction 对计划值调用
      // canonicalCell，"" 在任何宿主交互之前就变成 null，写入意图已丢失。
      expect(prepared.actions[0].values).toEqual([["未分类"], [""]]);
    }
  );

  it.fails(
    "缺陷 2：无法保真空串的宿主不应报告 verified，实际归一后判定通过",
    async () => {
      const target = host([["旧值"], ["旧值"]], true);
      const prepared = prepareWpsPlan(target.application, [1], emptyStringPlan);
      const record = await executeWpsPlan(target.application, prepared);
      // 失败原因锁定在 wpsPlanExecutor.ts:147-149 —— matricesEqual 通过
      // canonicalMatrix 对 actual 与 expected 双侧折叠，expected="" 与
      // actual=null 归一后都是 null，回读核对这道最后防线对该差异失明。
      // 期望行为：拒绝写入，或标记为未独立验证，而不是 verified。
      expect(record.state).not.toBe("verified");
    }
  );

  // 以下两条断言当前实际行为（而非期望行为），用于锁定失败原因。
  // 阶段 4 修复后它们会开始失败，届时应连同上面的 it.fails 一起更新。
  it("锁定缺陷 1 的确切表现：空串在准备阶段变成 null", () => {
    const target = host([["旧值"], ["旧值"]]);
    const prepared = prepareWpsPlan(target.application, [1], emptyStringPlan);
    expect(prepared.actions[0].values).toEqual([["未分类"], [null]]);
  });

  it("锁定缺陷 2 的确切表现：宿主折叠空串后仍报告 verified", async () => {
    const target = host([["旧值"], ["旧值"]], true);
    const prepared = prepareWpsPlan(target.application, [1], emptyStringPlan);
    const record = await executeWpsPlan(target.application, prepared);
    expect(record.state).toBe("verified");
    expect(target.values()).toEqual([["未分类"], [null]]);
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
