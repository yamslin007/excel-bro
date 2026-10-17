import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { AnalysisPlan, ExcelAction } from "./contracts";
import { createOfficeTestHost, installOfficeTestHost, type TestOfficeHost } from "./officeTestHost";
import {
  assertExcelNativeComparisonUnchanged,
  closeExcelNativeComparison,
  compareExcelNativeCommittedResult,
  excelNativeReviewMode,
  nativeComparisonPlanIssue,
  stageExcelNativeComparison
} from "./excelNativeComparison";
import { executePlan } from "./excel";

/**
 * **格式动作的副本审阅闭环**（原生模式）。
 *
 * 之前格式动作进不了副本审阅：`nativeComparisonPlanIssue` 一律拒绝，用户只能在确认后
 * 直接写原表，看不到"加粗/居中/行高"会变成什么样。现在格式动作走与执行器**同源**的
 * `applyFormatAction`，并在副本上逐格回读核对。
 *
 * 本文件用**真实产品代码** `stageExcelNativeComparison` + 有真实状态的替身宿主证明：
 *  1. 格式动作真的写进了副本（且只写副本，原表零写入）；
 *  2. 副本回读不符时**拒绝开放写回**；
 *  3. 副本里的格式被改动会在写回前被发现（`copyFormatEvidence`）；
 *  4. 提交后对照把**格式**一并比对（`compareExcelNativeCommittedResult`）。
 *
 * 证据边界：替身宿主自动化证据，非真实 Excel 实机证据。
 */

const layout = vi.hoisted(() => ({ prepare: vi.fn(), arrange: vi.fn(), close: vi.fn() }));
vi.mock("./excelWindowLayout", () => ({
  prepareLocalExcelWindows: layout.prepare,
  arrangeLocalExcelWindows: layout.arrange,
  closeLocalExcelWindows: layout.close
}));

beforeEach(() => {
  // 并排窗口由本机桥提供；桥不可用时只影响**摆放**，不影响"能不能写回"。
  layout.prepare.mockReset().mockRejectedValue(new Error("unavailable"));
  layout.arrange.mockReset().mockResolvedValue({
    status: "arranged",
    layoutVersion: 3,
    copySide: "left",
    originalSide: "right",
    synchronousScrolling: "enabled"
  });
  layout.close.mockReset().mockResolvedValue({ status: "closed", layoutVersion: 4, recoveryCopyPath: "E:/recovery/review-copy.xlsx", copyClosed: true });
});

const SHEET = "数据";
const BOOK_URL = "https://contoso.sharepoint.com/format.xlsx";

function plan(actions: ExcelAction[]): AnalysisPlan {
  return {
    id: "plan-copy-format",
    title: "格式",
    summary: "格式",
    assumptions: [],
    warnings: [],
    actions
  };
}

let host: TestOfficeHost;

function createHost(): TestOfficeHost {
  return createOfficeTestHost({
    url: BOOK_URL,
    activeSheet: SHEET,
    selectedRange: "A1:B2",
    sheets: [
      {
        name: SHEET,
        values: [
          ["甲", "乙"],
          ["丙", "丁"]
        ]
      }
    ]
  });
}

beforeEach(() => {
  host = createHost();
  installOfficeTestHost(host);
});

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("格式动作在副本审阅里的计划准入", () => {
  it("六个格式动作都放行；空格式动作被拒", () => {
    expect(
      nativeComparisonPlanIssue(plan([{ type: "setFont", sheet: SHEET, range: "A1:B2", bold: true }]), [SHEET])
    ).toBeNull();
    expect(
      nativeComparisonPlanIssue(plan([{ type: "setAlignment", sheet: SHEET, range: "A1:B2", horizontal: "center" }]), [SHEET])
    ).toBeNull();
    expect(
      nativeComparisonPlanIssue(plan([{ type: "resizeRange", sheet: SHEET, range: "A1:B2", rowHeight: 30 }]), [SHEET])
    ).toBeNull();
    // 一个属性都没给 ⇒ 什么也不做，不能为它开放写回。
    expect(
      nativeComparisonPlanIssue(plan([{ type: "setFont", sheet: SHEET, range: "A1:B2" } as ExcelAction]), [SHEET])
    ).toContain("没有指定任何要修改的属性");
  });
});

describe("格式动作在副本里被真实试算", () => {
  it("同区域加粗与居中能进入副本，原表零写入，变化格不重复计数", async () => {
    const p = plan([
      { type: "setFont", sheet: SHEET, range: "A1:B2", bold: true },
      { type: "setAlignment", sheet: SHEET, range: "A1:B2", horizontal: "center" }
    ]);
    expect(nativeComparisonPlanIssue(p, [SHEET])).toBeNull();
    const session = await stageExcelNativeComparison(p, [SHEET]);
    expect(host.originalWrites()).toEqual([]);
    expect(session.changedCells).toBe(4);
    const style = host.cellStyle(session.copySheetName, "A1", host.copyStores()[0]);
    expect(style?.fontBold).toBe(true);
    expect(style?.horizontalAlignment).toBe("Center");
    await executePlan(p, { authorizedSheetNames: [SHEET], beforeFirstWrite: context =>
      assertExcelNativeComparisonUnchanged(context, session, [SHEET], p) });
    expect(await compareExcelNativeCommittedResult(session)).toBe("matched");
  });

  it("字号依次覆盖按最终状态核对，不用第一个动作的中间状态误拒", async () => {
    const p = plan([
      { type: "setFont", sheet: SHEET, range: "A1:B2", size: 14 },
      { type: "setFont", sheet: SHEET, range: "B2", size: 16 }
    ]);
    const session = await stageExcelNativeComparison(p, [SHEET]);
    const store = host.copyStores()[0];
    expect(host.cellStyle(session.copySheetName, "A1", store)?.fontSize).toBe(14);
    expect(host.cellStyle(session.copySheetName, "B2", store)?.fontSize).toBe(16);
    await Excel.run(context => assertExcelNativeComparisonUnchanged(context, session, [SHEET], p));
    expect(host.originalWrites()).toEqual([]);
  });

  it("行高证据区分未执行和执行，副本同行其他列改行高后拒绝写回", async () => {
    const p = plan([{ type: "resizeRange", sheet: SHEET, range: "A1:A2", rowHeight: 30 }]);
    const session = await stageExcelNativeComparison(p, [SHEET]);
    expect(await compareExcelNativeCommittedResult(session)).toBe("different");
    await executePlan(p, { authorizedSheetNames: [SHEET] });
    expect(await compareExcelNativeCommittedResult(session)).toBe("matched");
    // Separate review to test pre-write validation while the original stays unchanged.
    const next = await stageExcelNativeComparison(p, [SHEET]);
    host.setRowHeight(next.copySheetName, 0, 40, host.copyStores()[1]);
    await expect(Excel.run(context => assertExcelNativeComparisonUnchanged(context, next, [SHEET], p)))
      .rejects.toThrow("副本中的格式已变化");
  });

  it("格式副本不得用绿色高亮覆盖计划填充色", async () => {
    const p = plan([{ type: "setFill", sheet: SHEET, range: "A1", color: "#FF0000" }]);
    const session = await stageExcelNativeComparison(p, [SHEET]);
    expect(host.cellStyle(session.copySheetName, "A1", host.copyStores()[0])?.fillColor).toBe("#FF0000");
    expect(host.originalWrites()).toEqual([]);
  });

  it("格式和值写入的重叠仍拒绝", () => {
    expect(nativeComparisonPlanIssue(plan([
      { type: "setFont", sheet: SHEET, range: "A1:B2", bold: true },
      { type: "writeValues", sheet: SHEET, range: "A1", values: [[1]] }
    ]), [SHEET])).toContain("重叠");
  });

  it("setFont{bold} 写进副本（逐格加粗）并在副本上标出改动；原表零写入", async () => {
    expect(excelNativeReviewMode()).toBe("native");
    const formatPlan = plan([{ type: "setFont", sheet: SHEET, range: "A1:B2", bold: true }]);

    const session = await stageExcelNativeComparison(formatPlan, [SHEET]);

    // 只写了副本：原表一个写都没有。
    expect(host.originalWrites()).toEqual([]);
    expect(host.copyWrites().length).toBeGreaterThan(0);

    // 副本里四格确实被加粗了（格式动作真的在副本上生效）。
    const copyStore = host.copyStores()[0];
    expect(copyStore).toBeTruthy();
    for (const address of ["A1", "B1", "A2", "B2"]) {
      expect(host.cellStyle(session.copySheetName, address, copyStore)?.fontBold).toBe(true);
    }
    // 四格都是"会改动"的格。
    expect(session.changedCells).toBe(4);

    // 关闭副本，且原表始终未被写。
    expect(await closeExcelNativeComparison(session)).toBe("closed");
    expect(host.originalWrites()).toEqual([]);
  });

  it("提交后对照把格式一并比对：原表未执行时 different，执行后 matched", async () => {
    const formatPlan = plan([{ type: "setFont", sheet: SHEET, range: "A1:B2", bold: true }]);
    const session = await stageExcelNativeComparison(formatPlan, [SHEET]);

    // 原表还没执行 ⇒ 格式对不上。
    expect(await compareExcelNativeCommittedResult(session)).toBe("different");

    // 真实执行器把同一份计划写进原表。
    await executePlan(formatPlan, {
      authorizedSheetNames: [SHEET],
      beforeFirstWrite: async (context) =>
        assertExcelNativeComparisonUnchanged(context, session, [SHEET], formatPlan)
    });

    // 提交后对照：值/公式/格式都一致 ⇒ matched（格式也进了判据）。
    expect(await compareExcelNativeCommittedResult(session)).toBe("matched");
  });

  it("副本里的格式被改动 ⇒ 写回前被拦下（copyFormatEvidence）", async () => {
    const formatPlan = plan([{ type: "setFont", sheet: SHEET, range: "A1:B2", bold: true }]);
    const session = await stageExcelNativeComparison(formatPlan, [SHEET]);
    const copyStore = host.copyStores()[0];
    // 模拟副本被外部编辑：把 A1 的加粗取消。
    host.setCellStyle(session.copySheetName, "A1", { fontBold: false }, { store: copyStore });

    await Excel.run(async (context: Excel.RequestContext) => {
      await expect(
        assertExcelNativeComparisonUnchanged(context, session, [SHEET], formatPlan)
      ).rejects.toThrow("副本中的格式已变化");
    });
  });

  it("副本里的格式没被改动时，写前复核放行且自身零写入", async () => {
    const formatPlan = plan([{ type: "setFont", sheet: SHEET, range: "A1:B2", bold: true }]);
    const session = await stageExcelNativeComparison(formatPlan, [SHEET]);
    const before = host.writes().length;

    await Excel.run(async (context: Excel.RequestContext) => {
      await expect(
        assertExcelNativeComparisonUnchanged(context, session, [SHEET], formatPlan)
      ).resolves.toBeUndefined();
    });

    // 复核本身是只读的。
    expect(host.writes()).toHaveLength(before);
  });
});
