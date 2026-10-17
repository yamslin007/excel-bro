// @vitest-environment jsdom
import { useState } from "react";
import { cleanup, fireEvent, render, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { ExcelNativeReviewActions } from "./ExcelNativeReviewActions";
import { useExcelNativeComparison } from "../hooks/useExcelNativeComparison";
import { assertExcelNativeComparisonUnchanged, nativeComparisonPlanIssue } from "../excelNativeComparison";
import { closeLocalExcelWindows } from "../excelWindowLayout";
import { executePlan, undoExecution } from "../excel";
import { createOfficeTestHost, installOfficeTestHost } from "../officeTestHost";
import type { AnalysisPlan } from "../contracts";

// Real review component, controller, copy generation, execution and undo.
// Fixed plan + Office.js stateful double + window-layout bridge double.
// No model, server ledger, full App routing or real-client evidence is claimed.
vi.mock("../excelWindowLayout", () => ({
  prepareLocalExcelWindows: vi.fn(async () => { throw new Error("layout unavailable"); }),
  arrangeLocalExcelWindows: vi.fn(),
  closeLocalExcelWindows: vi.fn(async () => ({ status: "closed", layoutVersion: 4, recoveryCopyPath: "E:/recovery/review-copy.xlsx", copyClosed: true }))
}));
const SHEET = "验收";
const plan: AnalysisPlan = { id: "format-review", title: "加粗并居中", summary: "格式审阅", assumptions: [], warnings: [], actions: [
  { type: "setFont", sheet: SHEET, range: "A1:B2", bold: true },
  { type: "setAlignment", sheet: SHEET, range: "A1:B2", horizontal: "center" }
] };
function Harness({ plan }: { plan: AnalysisPlan }) {
  const controller = useExcelNativeComparison();
  const [execution, setExecution] = useState<Awaited<ReturnType<typeof executePlan>>>();
  const [undone, setUndone] = useState(false);
  return <>
    <ExcelNativeReviewActions messageId="format" plan={plan} authorization={[SHEET]}
      controller={controller} disabled={controller.pending} commit={async session => {
        // 本替身只走副本比对路径；直接执行不经过 `assertExcelNativeComparisonUnchanged`。
        if (!session) throw new Error("直接执行路径没有副本会话");
        const result = await executePlan(plan, { authorizedSheetNames: [SHEET],
          beforeFirstWrite: context => assertExcelNativeComparisonUnchanged(context, session, [SHEET], plan) });
        setExecution(result);
        return true;
      }} />
    {execution?.undoSnapshot && <button onClick={() => void undoExecution(execution.undoSnapshot!).then(() => setUndone(true))}>撤销测试执行</button>}
    {undone && <p>测试恢复完成</p>}
  </>;
}
beforeEach(() => vi.clearAllMocks());
afterEach(() => { cleanup(); vi.unstubAllGlobals(); });

it("放弃兼容副本调用关闭桥、展示恢复文件，原表零写入且写回入口撤下", async () => {
  const host = createOfficeTestHost({ url: "file:///E:/format-review.xlsx", activeSheet: SHEET,
    excelApiVersions: Array.from({ length: 17 }, (_, i) => `1.${i + 1}`), excelApiDesktopVersions: [],
    sheets: [{ name: SHEET, values: [["甲", 100], ["乙", 200]] }] });
  installOfficeTestHost(host);
  const ui = render(<Harness plan={plan} />);
  fireEvent.click(ui.getByRole("button", { name: "生成副本并排审阅" }));
  await waitFor(() => expect(ui.getByRole("button", { name: "放弃变更" })).toBeTruthy());
  expect(ui.queryByRole("button", { name: "重新排列" })).toBeNull();
  fireEvent.click(ui.getByRole("button", { name: "放弃变更" }));
  await waitFor(() => expect(ui.getByText("审阅副本已关闭。")).toBeTruthy());
  expect(closeLocalExcelWindows).toHaveBeenCalledOnce();
  expect(closeLocalExcelWindows).toHaveBeenCalledWith('url:file:///e:/format-review.xlsx', SHEET,
    expect.stringMatching(/^EB_[a-f0-9]{20}$/));
  expect(ui.getByText('E:/recovery/review-copy.xlsx', { selector: 'code' })).toBeTruthy();
  expect(ui.queryByRole("button", { name: "确认写回原表" })).toBeNull();
  expect(ui.queryByRole("checkbox")).toBeNull();
  expect(host.originalWrites()).toEqual([]);
  expect(host.cellStyle(SHEET, "A1")?.fontBold).toBeUndefined();
});

describe.each(["native", "manual"] as const)("格式审阅组件真实链路 · %s", mode => {
  it("同区域组合先生成副本，确认才写回，撤销恢复混合字号与原有数据", async () => {
    const host = createOfficeTestHost({ url: "file:///E:/format-review.xlsx", activeSheet: SHEET,
      excelApiVersions: Array.from({ length: 17 }, (_, i) => `1.${i + 1}`),
      excelApiDesktopVersions: mode === "native" ? ["1.1"] : [],
      sheets: [{ name: SHEET, values: [["甲", 100], ["乙", 200]] }] });
    installOfficeTestHost(host);
    host.setCellStyle(SHEET, "A1", { fontSize: 12 });
    host.setCellStyle(SHEET, "A2", { fontSize: 16 });
    expect(nativeComparisonPlanIssue(plan, [SHEET])).toBeNull();
    const ui = render(<Harness plan={plan} />);
    fireEvent.click(ui.getByRole("button", { name: "生成副本并排审阅" }));
    await waitFor(() => expect(ui.getByRole("button", { name: "确认写回原表" })).toBeTruthy());
    expect(host.createWorkbookCalls()).toHaveLength(1);
    expect(host.originalWrites()).toEqual([]);
    if (mode === "manual") {
      // Verify the actual generated OOXML, not the double's cloned copy state.
      const styles = host.seedEntries(0)["xl/styles.xml"];
      expect(styles).toContain("<b/>");
      expect(styles).toContain('horizontal="center"');
      expect((ui.getByRole("button", { name: "确认写回原表" }) as HTMLButtonElement).disabled).toBe(true);
      fireEvent.click(ui.getByRole("checkbox", { name: "我已核对副本，确认按此计划修改原表。" }));
    }
    fireEvent.click(ui.getByRole("button", { name: "确认写回原表" }));
    await waitFor(() => expect(ui.getByRole("button", { name: "撤销测试执行" })).toBeTruthy());
    expect(host.cellStyle(SHEET, "A1")?.fontBold).toBe(true);
    expect(host.cellStyle(SHEET, "A1")?.horizontalAlignment).toBe("Center");
    fireEvent.click(ui.getByRole("button", { name: "撤销测试执行" }));
    await waitFor(() => expect(ui.getByText("测试恢复完成")).toBeTruthy());
    expect(host.cellStyle(SHEET, "A1")?.fontBold).toBeUndefined();
    expect(host.cellStyle(SHEET, "A1")?.horizontalAlignment).toBeUndefined();
    expect(host.cellStyle(SHEET, "A1")?.fontSize).toBe(12);
    expect(host.cellStyle(SHEET, "A2")?.fontSize).toBe(16);
    await Excel.run(async context => {
      const range = context.workbook.worksheets.getItem(SHEET).getRange("A1:B2");
      range.load("values"); await context.sync();
      expect(range.values).toEqual([["甲", 100], ["乙", 200]]);
    });
  });

  // 行高是**行级**属性：`A1:B2` 的两行整体受影响（同行 C 列虽未被本动作写值，
  // 行高一样变）。这条链路单独验，因为它不落任何单元格格式证据 —— 若把行高
  // 混进"逐格格式"里，就会出现"没有变化格 ⇒ 判定无改动"的假通过。
  it("行高动作：副本内按整行生效、确认才写回、撤销把每一行行高逐行还原", async () => {
    const host = createOfficeTestHost({ url: "file:///E:/format-rowheight.xlsx", activeSheet: SHEET,
      excelApiVersions: Array.from({ length: 17 }, (_, i) => `1.${i + 1}`),
      excelApiDesktopVersions: mode === "native" ? ["1.1"] : [],
      sheets: [{ name: SHEET, values: [["甲", 100], ["乙", 200], ["丙", 300]] }] });
    installOfficeTestHost(host);
    // 原始行高各不相同（第 1 行 20、第 2 行 30）；第 3 行保持默认。
    host.setRowHeight(SHEET, 0, 20);
    host.setRowHeight(SHEET, 1, 30);
    const rowPlan: AnalysisPlan = { id: "format-rowheight", title: "调整行高", summary: "行高审阅",
      assumptions: [], warnings: [], actions: [
        { type: "resizeRange", sheet: SHEET, range: "A1:B2", rowHeight: 44 }
      ] };
    expect(nativeComparisonPlanIssue(rowPlan, [SHEET])).toBeNull();
    const ui = render(<Harness plan={rowPlan} />);
    fireEvent.click(ui.getByRole("button", { name: "生成副本并排审阅" }));
    await waitFor(() => expect(ui.getByRole("button", { name: "确认写回原表" })).toBeTruthy());
    expect(host.createWorkbookCalls()).toHaveLength(1);
    // 确认前原表零写入：行高与值都没动。
    expect(host.originalWrites()).toEqual([]);
    expect(host.rowHeight(SHEET, 0)).toBe(20);
    expect(host.rowHeight(SHEET, 1)).toBe(30);
    if (mode === "manual") {
      fireEvent.click(ui.getByRole("checkbox", { name: "我已核对副本，确认按此计划修改原表。" }));
    }
    fireEvent.click(ui.getByRole("button", { name: "确认写回原表" }));
    await waitFor(() => expect(ui.getByRole("button", { name: "撤销测试执行" })).toBeTruthy());
    // 整行生效：两行都到 44，未在范围内的第 3 行不动，且未声明的列宽保持默认。
    expect(host.rowHeight(SHEET, 0)).toBe(44);
    expect(host.rowHeight(SHEET, 1)).toBe(44);
    expect(host.rowHeight(SHEET, 2)).toBe(15);
    expect(host.columnWidth(SHEET, 0)).toBe(8.43);
    fireEvent.click(ui.getByRole("button", { name: "撤销测试执行" }));
    await waitFor(() => expect(ui.getByText("测试恢复完成")).toBeTruthy());
    // 撤销逐行还原（不是"统一刷成同一个值"），未声明的列宽/值始终未动。
    expect(host.rowHeight(SHEET, 0)).toBe(20);
    expect(host.rowHeight(SHEET, 1)).toBe(30);
    expect(host.columnWidth(SHEET, 0)).toBe(8.43);
    await Excel.run(async context => {
      const range = context.workbook.worksheets.getItem(SHEET).getRange("A1:B3");
      range.load("values"); await context.sync();
      expect(range.values).toEqual([["甲", 100], ["乙", 200], ["丙", 300]]);
    });
  });
});

/** 关闭回执不得改写业务结果或清除撤销记录。宿主与本机桥替身。 */
describe.each([
  "closed", "already_closed", "close_unconfirmed", "failed", "unsupported", "identity_lost", "ambiguous", "user_saved"
] as const)("兼容副本自动关闭（桥回执 %s 独立于写回结果）", status => {
  it("业务结论仍是「已写回原表」并排在前，副本状态独立显示，且原表仍可撤销", async () => {
    const host = createOfficeTestHost({ url: "file:///E:/format-review.xlsx", activeSheet: SHEET,
      excelApiVersions: Array.from({ length: 17 }, (_, i) => `1.${i + 1}`),
      excelApiDesktopVersions: [],
      sheets: [{ name: SHEET, values: [["甲", 100], ["乙", 200]] }] });
    installOfficeTestHost(host);
    host.setCellStyle(SHEET, "A1", { fontSize: 12 });
    vi.mocked(closeLocalExcelWindows).mockResolvedValueOnce({ status, layoutVersion: 4, recoveryCopyPath: "E:/recovery/review-copy.xlsx",
      copyClosed: status === "closed" || status === "already_closed" });
    const ui = render(<Harness plan={plan} />);
    fireEvent.click(ui.getByRole("button", { name: "生成副本并排审阅" }));
    await waitFor(() => expect(ui.getByRole("button", { name: "确认写回原表" })).toBeTruthy());
    fireEvent.click(ui.getByRole("checkbox", { name: "我已核对副本，确认按此计划修改原表。" }));
    fireEvent.click(ui.getByRole("button", { name: "确认写回原表" }));
    await waitFor(() => expect(ui.getByRole("button", { name: "撤销测试执行" })).toBeTruthy());
    // 业务结论始终在场，并且**排在**副本清理结论之前。
    const section = ui.getByText("已写回原表").closest("section")!;
    const description = section.querySelector(".excel-review-description")!.textContent ?? "";
    const copyStatus = section.querySelector(".excel-review-copy-status")!.textContent ?? "";
    expect(description).toContain("原表已按本卡审阅的计划写入");
    expect(section.textContent!.indexOf(description)).toBeLessThan(section.textContent!.indexOf(copyStatus));
    expect(section.textContent).not.toContain("写回失败");
    expect(section.textContent).not.toContain("未写回原表");
    // 副本状态独立成一块，不许混进业务结论里。
    expect(description).not.toContain("副本");
    expect(host.cellStyle(SHEET, "A1")?.fontBold).toBe(true);
    expect(closeLocalExcelWindows).toHaveBeenCalledOnce();
    expect(closeLocalExcelWindows).toHaveBeenCalledWith('url:file:///e:/format-review.xlsx', SHEET, expect.stringMatching(/^EB_[a-f0-9]{20}$/));
    if (status === 'closed') expect(copyStatus).toContain('已关闭');
    else if (status === 'already_closed') expect(copyStatus).toContain('此前已由你关闭');
    else expect(copyStatus).not.toContain('已关闭');
    expect(section.querySelector('.excel-review-recovery')!.textContent).toContain('E:/recovery/review-copy.xlsx');
    expect(section.textContent).not.toContain("正在询问是否保存");
    // 清理副本不是撤销：恢复点仍在，而且真的能撤销。
    fireEvent.click(ui.getByRole("button", { name: "撤销测试执行" }));
    await waitFor(() => expect(ui.getByText("测试恢复完成")).toBeTruthy());
    expect(host.cellStyle(SHEET, "A1")?.fontBold).toBeUndefined();
    expect(host.cellStyle(SHEET, "A1")?.fontSize).toBe(12);
  });
});
