import { describe, expect, it, vi } from "vitest";
import type { AnalysisPlan } from "../../excel-addin/src/contracts";
import type { WpsApplication, WpsWorkbook } from "./wpsHost";
import { executeWpsPlan, latestWpsUndo, prepareWpsPlan } from "./wpsPlanExecutor";
import {
  changedWpsCellAddresses,
  compareCommittedWpsResultToCopy,
  discardWpsNativeComparison,
  stageWpsNativeComparison,
  verifyWpsNativeComparison,
  WpsFormulaTrialError
} from "./wpsNativeComparison";

function fakeHost(options: { formulaDisplayOnOriginal?: string; formulaDisplayOnCopy?: string; freshWorkbookWrappers?: boolean } = {}) {
  const originalCell = { value: "原值", formula: "原值", color: undefined as number | undefined };
  const copiedCell = { value: "原值", formula: "原值", color: undefined as number | undefined };
  let active: WpsWorkbook;
  let copyClosed = false;
  const originalRange = () => ({
    get Value2() { return originalCell.value; },
    set Value2(value: string) { originalCell.value = value; originalCell.formula = value; },
    get Formula() { return originalCell.formula; },
    set Formula(value: string) {
      originalCell.formula = value;
      originalCell.value = value.startsWith("=") ? options.formulaDisplayOnOriginal ?? "/" : value;
    },
    MergeCells: false,
    Interior: {
      get Color() { return originalCell.color; },
      set Color(value: number | undefined) { originalCell.color = value; }
    }
  });
  const copyRange = () => ({
    get Value2() { return copiedCell.value; },
    set Value2(value: string) { copiedCell.value = value; copiedCell.formula = value; },
    get Formula() { return copiedCell.formula; },
    set Formula(value: string) { copiedCell.formula = value; copiedCell.value = value.startsWith("=") ? options.formulaDisplayOnCopy ?? "/" : value; },
    MergeCells: false,
    Interior: {
      get Color() { return copiedCell.color; },
      set Color(value: number | undefined) { copiedCell.color = value; }
    }
  });
  const copySheet = {
    Name: "Sheet1", ProtectContents: false,
    Range: vi.fn((...args: unknown[]) => {
      if (args.length !== 1) throw new Error("Too many parameters.");
      return copyRange();
    })
  };
  const copy: WpsWorkbook = {
    Name: "Book1", ReadOnly: false,
    Worksheets: { Count: 1, Item: () => copySheet },
    Windows: { Count: 1, Item: () => ({ Caption: "Book1" }) },
    Activate: () => { active = copy; },
    Close: vi.fn((saveChanges?: boolean) => {
      if (saveChanges !== false) throw new Error("副本不得保存");
      copyClosed = true;
      active = original;
      return true;
    })
  };
  const originalSheet = {
    Name: "Sheet1", ProtectContents: false,
    Range: vi.fn((...args: unknown[]) => {
      if (args.length !== 1) throw new Error("Too many parameters.");
      return originalRange();
    }),
    Copy: vi.fn((...args: unknown[]) => {
      if (args.length !== 0) throw new Error("Too many parameters.");
      copiedCell.value = originalCell.value;
      copiedCell.formula = originalCell.formula;
      active = copy;
    })
  };
  const original: WpsWorkbook = {
    Name: "测试.xlsx", FullName: "C:\\Data\\测试.xlsx", Path: "C:\\Data", ReadOnly: false,
    Worksheets: { Count: 1, Item: () => originalSheet },
    ActiveSheet: originalSheet,
    Activate: () => { active = original; }
  };
  active = original;
  const storage = new Map<string, unknown>();
  const compare = vi.fn((...args: unknown[]) => {
    if (args.length !== 1 || args[0] !== "Book1") throw new Error("并排参数不符");
    if (active !== original) throw new Error("必须从原表发起并排");
    return true;
  });
  const application = {
    Version: "12.0",
    get ActiveWorkbook() { return options.freshWorkbookWrappers ? { ...active } : active; },
    Windows: { Count: 2, Item: () => ({ Caption: "Book1" }), CompareSideBySideWith: compare },
    PluginStorage: {
      getItem: (key: string) => storage.get(key),
      setItem: (key: string, value: unknown) => { storage.set(key, value); }
    },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  } as WpsApplication;
  const plan: AnalysisPlan = {
    id: "native-compare", title: "更新单元格", summary: "测试副本与原表分离",
    assumptions: [], warnings: [], acceptanceCriteria: [],
    actions: [{ type: "writeValues", sheet: "Sheet1", range: "A1", values: [["结果"]] }]
  };
  return {
    application, plan, original, originalSheet, copy, compare,
    setActiveWorkbook: (workbook: WpsWorkbook) => { active = workbook; },
    originalCell, copiedCell, copyClosed: () => copyClosed
  };
}

describe("WPS 原生双窗口的隔离审阅", () => {
  it("只标真正变化的同坐标单元格，公式变化即使显示值相同也标记", () => {
    const host = fakeHost();
    const action = prepareWpsPlan(host.application, [1], host.plan).actions[0];
    const twoRows = {
      ...action,
      range: "A1:A2", rowCount: 2, cellCount: 2,
      rect: { startRow: 1, endRow: 2, startColumn: 1, endColumn: 1 },
      beforeValues: [["旧"], ["相同"]],
      beforeFormulas: [["旧"], ["相同"]],
      beforeFingerprint: "", afterFingerprint: null, appliedRows: 2
    };
    expect(changedWpsCellAddresses(twoRows, {
      values: [["新"], ["相同"]], formulas: [["新"], ["相同"]], valuesCaptured: true
    })).toEqual(["A1"]);
    expect(changedWpsCellAddresses(twoRows, {
      values: [["旧"], ["相同"]], formulas: [["=1+1"], ["相同"]], valuesCaptured: true
    })).toEqual(["A1"]);
  });
  it("左原表不写、右副本回读后只标记变化位置，保存前仍待确认", () => {
    const host = fakeHost();
    const prepared = prepareWpsPlan(host.application, [1], host.plan);
    const staged = stageWpsNativeComparison(host.application, prepared);
    expect(host.originalSheet.Copy).toHaveBeenCalledWith();
    expect(host.originalCell.value).toBe("原值");
    expect(host.originalCell.color).toBeUndefined();
    expect(latestWpsUndo(host.application)).toBeNull();
    expect(host.copiedCell.value).toBe("结果");
    expect(host.copiedCell.color).toBe(0xE8F5EA);
    expect(staged.changedCells).toEqual(["A1"]);
    expect(staged.reviews[0].status).toBe("applied");
    expect(host.compare).toHaveBeenCalledWith("Book1");
    verifyWpsNativeComparison(host.application, staged);
    expect(host.originalCell.value).toBe("原值");
    discardWpsNativeComparison(staged);
    expect(host.copyClosed()).toBe(true);
    expect(host.copy.Close).toHaveBeenCalledWith(false);
    expect(host.originalCell.value).toBe("原值");
  });

  it("同一本已保存工作簿返回新的 JSAPI 包装对象时仍能生成、复核并提交", async () => {
    const host = fakeHost({ freshWorkbookWrappers: true });
    const prepared = prepareWpsPlan(host.application, [1], host.plan);
    const staged = stageWpsNativeComparison(host.application, prepared);
    expect(host.originalSheet.Copy).toHaveBeenCalledOnce();
    expect(host.originalCell.value).toBe("原值");
    verifyWpsNativeComparison(host.application, staged);
    const record = await executeWpsPlan(host.application, prepared);
    expect(record.state).toBe("verified");
    expect(host.originalCell.value).toBe("结果");
    expect(compareCommittedWpsResultToCopy(staged)).toBe("matched");
  });

  it("同名同内容但不同路径的工作簿不应被误当成原表", () => {
    const host = fakeHost({ freshWorkbookWrappers: true });
    const prepared = prepareWpsPlan(host.application, [1], host.plan);
    host.setActiveWorkbook({
      ...host.original,
      FullName: "D:\\Other\\测试.xlsx",
      Path: "D:\\Other"
    });
    expect(() => stageWpsNativeComparison(host.application, prepared))
      .toThrow("当前活动工作簿已不是预览时的原表");
    expect(host.originalSheet.Copy).not.toHaveBeenCalled();
    expect(host.originalCell.value).toBe("原值");
  });

  it("副本被保存为磁盘文件后拒绝当作未保存的审阅副本写回", () => {
    const host = fakeHost({ freshWorkbookWrappers: true });
    const staged = stageWpsNativeComparison(host.application, prepareWpsPlan(host.application, [1], host.plan));
    host.copy.Path = "C:\\Data";
    host.copy.FullName = "C:\\Data\\Book1.xlsx";
    expect(() => verifyWpsNativeComparison(host.application, staged)).toThrow("审阅副本身份已变化");
    expect(host.originalCell.value).toBe("原值");
  });

  it("原表在并排审阅期间改变时拒绝写回", () => {
    const host = fakeHost();
    const staged = stageWpsNativeComparison(host.application, prepareWpsPlan(host.application, [1], host.plan));
    host.originalCell.value = "他人新值";
    expect(() => verifyWpsNativeComparison(host.application, staged)).toThrow("原表或授权范围在预览后发生变化");
    expect(host.originalCell.value).toBe("他人新值");
  });

  it("原表在生成副本前已变或计划被篡改时，连副本都不生成", () => {
    const host = fakeHost();
    const prepared = prepareWpsPlan(host.application, [1], host.plan);
    host.originalCell.value = "预览后外部修改";
    expect(() => stageWpsNativeComparison(host.application, prepared)).toThrow("原表或授权范围在预览后发生变化");
    expect(host.originalSheet.Copy).not.toHaveBeenCalled();
    host.originalCell.value = "原值";
    prepared.actions[0].values[0][0] = "未经预览的值";
    expect(() => stageWpsNativeComparison(host.application, prepared)).toThrow("写入计划或预览已变化");
    expect(host.originalSheet.Copy).not.toHaveBeenCalled();
  });

  it("多表授权不会退化为复制整本工作簿", () => {
    const host = fakeHost();
    const prepared = prepareWpsPlan(host.application, [1], host.plan);
    prepared.authorization.sheets.push({ index: 2, name: "另一张表" });
    expect(() => stageWpsNativeComparison(host.application, prepared)).toThrow("只支持一张已授权工作表");
    expect(host.originalSheet.Copy).not.toHaveBeenCalled();
    expect(host.originalCell.value).toBe("原值");
  });

  it("右侧副本被手动改动后拒绝把未确认内容写回", () => {
    const host = fakeHost();
    const staged = stageWpsNativeComparison(host.application, prepareWpsPlan(host.application, [1], host.plan));
    host.copiedCell.value = "手工改动";
    expect(() => verifyWpsNativeComparison(host.application, staged)).toThrow("审阅副本在生成对照后发生变化");
    expect(host.originalCell.value).toBe("原值");
  });

  it("公式文本未变但计算显示值漂移时，两边都不能冒充仍是已确认现场", () => {
    const host = fakeHost();
    host.originalCell.formula = "=1+1";
    host.originalCell.value = "2";
    const prepared = prepareWpsPlan(host.application, [1], host.plan);
    const staged = stageWpsNativeComparison(host.application, prepared);
    host.originalCell.value = "3";
    expect(() => verifyWpsNativeComparison(host.application, staged)).toThrow("原表在并排审阅后发生变化");
    const formulaHost = fakeHost();
    formulaHost.plan.actions = [{ type: "writeFormulas", sheet: "Sheet1", range: "A1", formulas: [["=1+1"]] }];
    const formulaStaged = stageWpsNativeComparison(
      formulaHost.application,
      prepareWpsPlan(formulaHost.application, [1], formulaHost.plan)
    );
    formulaHost.copiedCell.value = "3"; // formula text remains =1+1
    expect(() => verifyWpsNativeComparison(formulaHost.application, formulaStaged))
      .toThrow("审阅副本在生成对照后发生变化");
  });

  it("明确保存后才用原执行器写回，并逐格核对副本与原表结果", async () => {
    const host = fakeHost();
    const prepared = prepareWpsPlan(host.application, [1], host.plan);
    const staged = stageWpsNativeComparison(host.application, prepared);
    expect(host.originalCell.value).toBe("原值");
    verifyWpsNativeComparison(host.application, staged);
    const record = await executeWpsPlan(host.application, prepared);
    expect(record.state).toBe("verified");
    expect(host.originalCell.value).toBe("结果");
    expect(host.originalCell.color).toBeUndefined();
    expect(compareCommittedWpsResultToCopy(staged)).toBe("matched");
    host.originalCell.value = "计算结果发生偏差";
    expect(compareCommittedWpsResultToCopy(staged)).toBe("different");
  });

  it("公式文本相同但原表计算值不同，明确报告副本与提交结果不一致", async () => {
    const host = fakeHost({ formulaDisplayOnOriginal: "另一计算值" });
    host.plan.actions = [{ type: "writeFormulas", sheet: "Sheet1", range: "A1", formulas: [["=1+1"]] }];
    const prepared = prepareWpsPlan(host.application, [1], host.plan);
    const staged = stageWpsNativeComparison(host.application, prepared);
    expect(host.originalCell.value).toBe("原值");
    expect(host.copiedCell.formula).toBe("=1+1");
    expect(host.copiedCell.color).toBe(0xE8F5EA);
    verifyWpsNativeComparison(host.application, staged);
    const record = await executeWpsPlan(host.application, prepared);
    expect(record.state).toBe("verified");
    expect(compareCommittedWpsResultToCopy(staged)).toBe("different");
  });

  it("公式文本回读相同但副本计算值错误时拒绝审阅与写回，并关闭试写副本", () => {
    const host = fakeHost();
    host.plan.id = "wps-map-column-test";
    host.plan.actions = [{ type: "writeFormulas", sheet: "Sheet1", range: "A1", formulas: [["=IF(EXACT(B1,\"甲\"),\"正确\",NA())"]] }];
    host.plan.acceptanceCriteria = [{ type: "rangeEquals", sheet: "Sheet1", range: "A1", expected: [["正确"]] }];
    expect(() => stageWpsNativeComparison(host.application, prepareWpsPlan(host.application, [1], host.plan)))
      .toThrow(WpsFormulaTrialError);
    expect(host.copyClosed()).toBe(true);
    expect(host.originalCell.value).toBe("原值");
    expect(host.compare).not.toHaveBeenCalled();
  });

  it("公式与本地预期计算值都相同才开放右侧副本审阅", () => {
    const host = fakeHost({ formulaDisplayOnCopy: "正确" });
    host.plan.id = "wps-map-column-test";
    host.plan.actions = [{ type: "writeFormulas", sheet: "Sheet1", range: "A1", formulas: [["=IF(EXACT(B1,\"甲\"),\"正确\",NA())"]] }];
    host.plan.acceptanceCriteria = [{ type: "rangeEquals", sheet: "Sheet1", range: "A1", expected: [["正确"]] }];
    const staged = stageWpsNativeComparison(host.application, prepareWpsPlan(host.application, [1], host.plan));
    expect(staged.copyAfter[0].values).toEqual([["正确"]]);
    expect(staged.reviews[0].status).toBe("applied");
    expect(host.originalCell.value).toBe("原值");
    expect(host.compare).toHaveBeenCalledOnce();
  });

  it("宿主并排失败时关闭自己创建的副本且不碰原表", () => {
    const host = fakeHost();
    host.compare.mockReturnValueOnce(false);
    expect(() => stageWpsNativeComparison(host.application, prepareWpsPlan(host.application, [1], host.plan)))
      .toThrow("WPS 未能打开原生并排比较");
    expect(host.copyClosed()).toBe(true);
    expect(host.originalCell.value).toBe("原值");
  });
});
