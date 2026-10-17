// @vitest-environment jsdom
/**
 * Review-entry regression coverage after removal of the execution-mode selector.
 * This replaces selector-specific assertions, preserving authorization, ownership,
 * single-submit, failed-write retention and copy-only cleanup checks.
 * Real component/controller; host and authorization transport are doubles.
 */
import { act, cleanup, fireEvent, render } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { useExcelNativeComparison } from "../hooks/useExcelNativeComparison";
import { ExcelNativeReviewActions } from "./ExcelNativeReviewActions";
import type { AnalysisPlan } from "../contracts";
import type { ExcelNativeComparison } from "../excelNativeComparison";

const mocks = vi.hoisted(() => ({ stage: vi.fn(), guard: vi.fn(), close: vi.fn(), mode: vi.fn() }));
vi.mock("../excelNativeComparison", () => ({
  stageExcelNativeComparison: mocks.stage, assertExcelNativeComparisonUnchanged: mocks.guard,
  closeExcelNativeComparison: mocks.close, excelNativeReviewMode: mocks.mode
}));
const plan: AnalysisPlan = { id: "same-plan", title: "Preview", summary: "", assumptions: [], warnings: [],
  actions: [{ type: "writeValues", sheet: "Ledger", range: "C2", values: [[7]] }] };
const session: ExcelNativeComparison = {
  plan, authorization: ["Ledger"], changedCells: 1,
  workbookIdentity: "url:file:///E:/Original.xlsx", originalWindowName: "Original.xlsx",
  copyWindowName: "Book1", copySheetName: "EB_Test", copySheetId: "copy", copyEvidence: "proof", actionEvidence: "proof",
  original: { id: "original", name: "Ledger", address: "A1:C2", startRow: 1, startColumn: 1,
    values: [["x"]], formulas: [["x"]], numberFormat: [["General"]], properties: [[{}]],
    formulaCells: [], rowHeights: [16], columnWidths: [60] }
};
type Props = {
  commit: (s: ExcelNativeComparison | null) => Promise<boolean>;
  beforeOperation?: () => Promise<boolean>;
  targetPlan?: AnalysisPlan;
  authorization?: string[];
};
function Harness({ commit, beforeOperation, targetPlan = plan, authorization = ["Ledger"] }: Props) {
  const controller = useExcelNativeComparison();
  return <>
    <ExcelNativeReviewActions messageId="m1" plan={targetPlan} authorization={authorization}
      controller={controller} disabled={controller.pending} commit={commit} beforeOperation={beforeOperation} />
    <button onClick={() => void controller.save("m1", targetPlan, authorization, commit, "EB_Test")}>旧写回回调</button>
  </>;
}
const click = async (element: HTMLElement) => { await act(async () => { fireEvent.click(element); }); };
async function staged(props: Props) {
  const ui = render(<Harness {...props} />);
  await click(ui.getByText("生成副本并排审阅"));
  return ui;
}
beforeEach(() => {
  vi.clearAllMocks(); mocks.mode.mockReturnValue("native"); mocks.stage.mockResolvedValue(session);
  mocks.guard.mockResolvedValue(undefined); mocks.close.mockResolvedValue("closed");
  vi.stubGlobal("Excel", { run: async (job: (context: unknown) => Promise<void>) => job({}) });
});
afterEach(() => { cleanup(); vi.unstubAllGlobals(); });

describe("入口与授权", () => {
  it("取消执行方式选择及重新排列入口，不重新加入直接写入按钮", () => {
    const ui = render(<Harness commit={vi.fn(async () => true)} />);
    expect(ui.queryByRole("radio")).toBeNull();
    expect(ui.queryByText("确认并直接执行到原表")).toBeNull();
    expect(ui.queryByText("重新排列")).toBeNull();
  });
  it("确认原表变更不冒充保存磁盘文件", async () => {
    const ui = await staged({ commit: vi.fn(async () => true) });
    expect(ui.getByText("确认写回原表").textContent).not.toMatch(/保存到文件|另存|保存磁盘/);
  });
  it("授权检查拒绝时不生成副本、不写原表", async () => {
    const commit = vi.fn(async () => true);
    await staged({ commit, beforeOperation: vi.fn(async () => false) });
    expect(mocks.stage).not.toHaveBeenCalled(); expect(commit).not.toHaveBeenCalled();
  });
});
describe("副本与写回", () => {
  it("生成副本自身不调用原表执行器", async () => {
    const commit = vi.fn(async () => true);
    await staged({ commit });
    expect(mocks.stage).toHaveBeenCalledOnce(); expect(commit).not.toHaveBeenCalled();
  });
  it("确认后先复核、写入一次，再关闭本次副本", async () => {
    const commit = vi.fn(async () => true), ui = await staged({ commit });
    await click(ui.getByText("确认写回原表"));
    expect(mocks.guard).toHaveBeenCalledOnce(); expect(commit).toHaveBeenCalledExactlyOnceWith(session);
    expect(mocks.close).toHaveBeenCalledExactlyOnceWith(session);
    expect(mocks.guard.mock.invocationCallOrder[0]).toBeLessThan(commit.mock.invocationCallOrder[0]);
    expect(commit.mock.invocationCallOrder[0]).toBeLessThan(mocks.close.mock.invocationCallOrder[0]);
  });
  it("写入抛错保留副本，不伪称未写入或验证通过", async () => {
    const ui = await staged({ commit: vi.fn(async () => { throw new Error("回读不符"); }) });
    await click(ui.getByText("确认写回原表"));
    expect(mocks.close).not.toHaveBeenCalled();
    expect(ui.getByText("写回结果不确定")).toBeTruthy();
    expect(ui.queryByText("已写回原表")).toBeNull();
  });
  it("关闭未确认不覆盖已经验证的业务结论", async () => {
    mocks.close.mockResolvedValue("close_unconfirmed");
    const ui = await staged({ commit: vi.fn(async () => true) });
    await click(ui.getByText("确认写回原表"));
    expect(ui.getByText("已写回原表")).toBeTruthy();
    expect(ui.getByText(/关闭尚未确认/)).toBeTruthy();
  });
});
describe("放弃与旧回调", () => {
  it("放弃不写原表，清理仅收到本次副本会话", async () => {
    const commit = vi.fn(async () => true), ui = await staged({ commit });
    await click(ui.getByText("放弃变更"));
    expect(commit).not.toHaveBeenCalled(); expect(mocks.close).toHaveBeenCalledExactlyOnceWith(session);
  });
  it("放弃之后排队的旧写回调用不能再写", async () => {
    const commit = vi.fn(async () => true), ui = await staged({ commit });
    await click(ui.getByText("放弃变更")); await click(ui.getByText("旧写回回调"));
    expect(commit).not.toHaveBeenCalled(); expect(mocks.guard).not.toHaveBeenCalled();
  });
});
describe("版本、重入与异步边界", () => {
  it("授权检查在途时换同 ID 的计划内容，旧确认拒绝写入", async () => {
    let resolve!: (ok: boolean) => void;
    const beforeOperation = vi.fn(async () => true), commit = vi.fn(async () => true);
    const ui = await staged({ commit, beforeOperation });
    beforeOperation.mockImplementation(() => new Promise<boolean>(r => { resolve = r; }));
    await click(ui.getByText("确认写回原表"));
    const changed: AnalysisPlan = { ...plan, actions: [{ type: "writeValues", sheet: "Ledger", range: "C2", values: [[99]] }] };
    ui.rerender(<Harness commit={commit} beforeOperation={beforeOperation} targetPlan={changed} />);
    await act(async () => resolve(true));
    expect(commit).not.toHaveBeenCalled(); expect(mocks.guard).not.toHaveBeenCalled();
  });
  it("授权检查在途时改变授权范围，旧确认拒绝写入", async () => {
    let resolve!: (ok: boolean) => void;
    const beforeOperation = vi.fn(async () => true), commit = vi.fn(async () => true);
    const ui = await staged({ commit, beforeOperation });
    beforeOperation.mockImplementation(() => new Promise<boolean>(r => { resolve = r; }));
    await click(ui.getByText("确认写回原表"));
    ui.rerender(<Harness commit={commit} beforeOperation={beforeOperation} authorization={["Other"]} />);
    await act(async () => resolve(true));
    expect(commit).not.toHaveBeenCalled(); expect(mocks.guard).not.toHaveBeenCalled();
  });
  it("同一轮重复点击只写一次", async () => {
    let release!: () => void;
    const gate = new Promise<void>(r => { release = r; });
    const commit = vi.fn(async () => { await gate; return true; });
    const ui = await staged({ commit }), button = ui.getByText("确认写回原表");
    await act(async () => { fireEvent.click(button); fireEvent.click(button); fireEvent.click(button); });
    expect(commit).toHaveBeenCalledOnce();
    await act(async () => release());
  });
  it("写前复核在途期间换计划，复核回来也不能写", async () => {
    let release!: () => void;
    mocks.guard.mockImplementation(() => new Promise<void>(r => { release = r; }));
    const commit = vi.fn(async () => true), ui = await staged({ commit });
    await click(ui.getByText("确认写回原表"));
    expect((ui.getByText("放弃变更") as HTMLButtonElement).disabled).toBe(true);
    ui.rerender(<Harness commit={commit} targetPlan={{ ...plan, summary: "edited" }} />);
    await act(async () => release());
    expect(commit).not.toHaveBeenCalled(); expect(mocks.close).not.toHaveBeenCalled();
  });
  it("成功后的旧回调不能重放第二次执行", async () => {
    const commit = vi.fn(async () => true), ui = await staged({ commit });
    await click(ui.getByText("确认写回原表")); await click(ui.getByText("旧写回回调"));
    expect(commit).toHaveBeenCalledOnce(); expect(mocks.close).toHaveBeenCalledOnce();
  });
});
describe("清理边界", () => {
  it("关闭目标是副本而不是原工作簿", async () => {
    const ui = await staged({ commit: vi.fn(async () => true) });
    await click(ui.getByText("确认写回原表"));
    expect(mocks.close).toHaveBeenCalledOnce();
    expect(mocks.close.mock.calls[0][0].copySheetName).toBe("EB_Test");
    expect(mocks.close.mock.calls[0][0].originalWindowName).toBe("Original.xlsx");
  });
  it("无法确认副本内容时不冒充已关闭", async () => {
    mocks.close.mockResolvedValue("edits_unverified");
    const ui = await staged({ commit: vi.fn(async () => true) });
    await click(ui.getByText("确认写回原表"));
    expect(ui.getByText("已写回原表")).toBeTruthy();
    expect(ui.queryByText("审阅副本已关闭。")).toBeNull();
    expect(ui.getByText(/无法确认副本是否被额外修改/)).toBeTruthy();
  });
});
