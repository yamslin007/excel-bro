// @vitest-environment jsdom
import { act, cleanup, fireEvent, render } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { useExcelNativeComparison } from "../hooks/useExcelNativeComparison";
import { ExcelNativeReviewActions } from "./ExcelNativeReviewActions";
import type { AnalysisPlan } from "../contracts";
import type { ExcelNativeComparison } from "../excelNativeComparison";

const mocks = vi.hoisted(() => ({ stage: vi.fn(), guard: vi.fn(), close: vi.fn(), mode: vi.fn() }));
vi.mock("../excelNativeComparison", () => ({ stageExcelNativeComparison: mocks.stage, assertExcelNativeComparisonUnchanged: mocks.guard, closeExcelNativeComparison: mocks.close, excelNativeReviewMode: mocks.mode }));
const p: AnalysisPlan = { id: "plan", title: "Preview", summary: "Update", assumptions: [], warnings: [], actions: [{ type: "writeValues", sheet: "Ledger", range: "C2", values: [["LOCAL_PRIVATE_SENTINEL"]] }] };
const session: ExcelNativeComparison = { plan: p, authorization: ["Ledger"], changedCells: 1,
  workbookIdentity: "url:file:///E:/Original.xlsx", originalWindowName: "Original.xlsx",
  copyWindowName: "Book1", copySheetName: "EB_Test", copySheetId: "copy", copyEvidence: "proof", actionEvidence: "proof",
  original: { id: "original", name: "Ledger", address: "A1:C2", startRow: 1, startColumn: 1,
    values: [["x"]], formulas: [["x"]], numberFormat: [["General"]], properties: [[{}]], formulaCells: [], rowHeights: [16], columnWidths: [60] }
};
function Harness({ commit, authorization = ["Ledger"], paneNeedsResize = false, beforeOperation }: { commit: (s: ExcelNativeComparison) => Promise<boolean>; authorization?: string[]; paneNeedsResize?: boolean; beforeOperation?: () => Promise<boolean> }) {
  const controller = useExcelNativeComparison();
  return <>
    <div data-testid="original-message"><ExcelNativeReviewActions messageId="m1" plan={p} authorization={authorization} controller={controller} disabled={controller.pending} commit={commit} paneNeedsResize={paneNeedsResize} beforeOperation={beforeOperation} /></div>
    <div data-testid="later-message">后续普通对话</div>
    <div data-testid="second-plan"><ExcelNativeReviewActions messageId="m2" plan={p} authorization={authorization} controller={controller} disabled={controller.pending} commit={commit} /></div>
    <button onClick={() => void controller.save('m1', p, authorization, commit)}>test direct save without review acknowledgement</button>
  </>;
}
beforeEach(() => {
  vi.clearAllMocks(); mocks.mode.mockReturnValue('native'); mocks.stage.mockResolvedValue(session); mocks.guard.mockResolvedValue(undefined); mocks.close.mockResolvedValue(true);
  vi.stubGlobal("Excel", { run: async (job: (context: unknown) => Promise<void>) => job({}) });
});
it("checks the authorized target before creating a copy AND before saving; confirmation never replays a blocked write", async () => {
  const beforeOperation = vi.fn(async () => false);
  const commit = vi.fn(async () => true);
  const ui = render(<Harness commit={commit} beforeOperation={beforeOperation} />);
  await click(ui.getAllByText("生成副本并排审阅")[0]);
  expect(mocks.stage).not.toHaveBeenCalled(); expect(commit).not.toHaveBeenCalled();
  beforeOperation.mockResolvedValue(true);
  await click(ui.getAllByText("生成副本并排审阅")[0]);
  expect(mocks.stage).toHaveBeenCalledOnce();
  beforeOperation.mockResolvedValue(false);
  await click(ui.getByText("保存变更到原表"));
  expect(commit).not.toHaveBeenCalled(); expect(mocks.guard).not.toHaveBeenCalled();
  expect(beforeOperation).toHaveBeenCalledTimes(3);
});
describe('manual compatibility review', () => {
  const manualSession: ExcelNativeComparison = { ...session, mode:'manual', actionEvidence:'', copyEvidence:'' };
  async function stage(ui: ReturnType<typeof render>) {
    await click(ui.getAllByText('生成副本并排审阅')[0]);
  }
  beforeEach(() => { mocks.mode.mockReturnValue('manual'); mocks.stage.mockResolvedValue(manualSession); });
  it.each(['enabled', 'different_instance', 'unavailable'] as const)('reports scrolling independently from arrangement: %s', async synchronousScrolling => {
    mocks.stage.mockResolvedValue({...manualSession,windowLayout:'arranged',synchronousScrolling});
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await stage(ui);
    if (synchronousScrolling === 'enabled') expect(ui.getByText(/已启用 Excel 原生并排比较与同步滚动/)).toBeTruthy();
    else {
      expect(ui.queryByText(/已启用 Excel 原生并排比较与同步滚动/)).toBeNull();
      expect(ui.getByText(synchronousScrolling === 'different_instance' ? /属于不同 Excel 实例/ : /同步滚动未能确认/)).toBeTruthy();
    }
    expect((ui.getByText('保存变更到原表') as HTMLButtonElement).disabled).toBe(true);
    expect(commit).not.toHaveBeenCalled();
  });
  it('explains the actual pane-width problem without claiming automatic success or granting save', async () => {
    mocks.stage.mockResolvedValue({...manualSession,windowLayout:'arranged'});
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} paneNeedsResize />);
    await stage(ui);
    expect(ui.getByRole('status').textContent).toContain('左侧分隔线');
    expect(ui.getByRole('status').textContent).toContain('右侧 Excel 窗口约 40%');
    expect((ui.getByText('保存变更到原表') as HTMLButtonElement).disabled).toBe(true);
    expect(commit).not.toHaveBeenCalled();
  });
  it('shows a restart instruction for an old backend, never new-order success', async () => {
    mocks.stage.mockResolvedValue({...manualSession,windowLayout:'service_outdated'});
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await stage(ui);
    expect(ui.getByRole('status').textContent).toContain('重启 Python 服务');
    expect(ui.queryByText(/已自动排列/)).toBeNull();
    expect((ui.getByText('保存变更到原表') as HTMLButtonElement).disabled).toBe(true);
    expect(commit).not.toHaveBeenCalled();
  });
  it('automatic layout is not copy verification or permission to write', async () => {
    mocks.stage.mockResolvedValue({...manualSession, windowLayout:'arranged'});
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await stage(ui);
    expect(ui.getByText(/已自动排列：左侧审阅副本，右侧原表与 AI 对话/)).toBeTruthy();
    expect(ui.queryByText(/自动排列未完成/)).toBeNull();
    expect((ui.getByText('保存变更到原表') as HTMLButtonElement).disabled).toBe(true);
    await click(ui.getByText('test direct save without review acknowledgement'));
    expect(commit).not.toHaveBeenCalled();
    await click(ui.getByRole('checkbox')); await click(ui.getByText('保存变更到原表'));
    expect(mocks.guard).toHaveBeenCalledOnce(); expect(commit).toHaveBeenCalledOnce();
    expect(mocks.close).not.toHaveBeenCalled();
  });
  it('requires real checkbox acknowledgement and then runs the original guard + commit; never claims automated copy readback or close', async () => {
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await stage(ui); expect(commit).not.toHaveBeenCalled();
    expect(ui.getByText(/视图 → 并排查看/)).toBeTruthy();
    const save = ui.getByText('保存变更到原表') as HTMLButtonElement;
    expect(save.disabled).toBe(true);
    await click(ui.getByText('test direct save without review acknowledgement'));
    expect(commit).not.toHaveBeenCalled(); expect(mocks.guard).not.toHaveBeenCalled();
    await click(ui.getByRole('checkbox', {name:'我已在 Excel 中核对副本，确认按此计划写回原表'}));
    expect(save.disabled).toBe(false); await click(save);
    expect(mocks.guard).toHaveBeenCalledOnce(); expect(commit).toHaveBeenCalledWith(manualSession);
    expect(mocks.close).not.toHaveBeenCalled(); expect(ui.queryByText('保存变更到原表')).toBeNull();
    expect(ui.getByText(/兼容副本未自动回读或关闭/)).toBeTruthy();
  });
  it('discard releases the review without pretending to have closed the copy, and a fresh copy needs a fresh acknowledgement', async () => {
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await stage(ui); await click(ui.getByRole('checkbox'));
    await click(ui.getByText('放弃变更'));
    expect(commit).not.toHaveBeenCalled(); expect(mocks.close).not.toHaveBeenCalled();
    expect(ui.getByText(/加载项不能自动关闭此窗口/)).toBeTruthy();
    mocks.stage.mockResolvedValue({...manualSession,copySheetName:'EB_New'});
    await stage(ui);
    expect((ui.getByText('保存变更到原表') as HTMLButtonElement).disabled).toBe(true);
  });
  it('guard conflicts block the clicked manual save; partial execution discard never says the original was unwritten', async () => {
    const commit = vi.fn(async () => false), ui = render(<Harness commit={commit} />);
    await stage(ui); await click(ui.getByRole('checkbox'));
    mocks.guard.mockRejectedValueOnce(new Error('原表已变化'));
    await click(ui.getByText('保存变更到原表'));
    expect(commit).not.toHaveBeenCalled(); expect(ui.queryByText('保存变更到原表')).toBeNull();
    await click(ui.getByText('关闭副本'));
    expect(ui.getByText(/关闭副本不撤销既有写入/)).toBeTruthy();
    mocks.stage.mockResolvedValue({...manualSession,copySheetName:'EB_AfterConflict'});
    await stage(ui); await click(ui.getByRole('checkbox'));
    await click(ui.getByText('保存变更到原表')); expect(commit).toHaveBeenCalledOnce();
    await click(ui.getByText('关闭副本'));
    expect(ui.getByText(/关闭副本不撤销既有写入/)).toBeTruthy();
  });
  it('does not publish a ready review when workbook creation fails', async () => {
    mocks.stage.mockRejectedValueOnce(new Error('副本打开被拒绝'));
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await stage(ui); expect(ui.queryByRole('checkbox')).toBeNull();
    expect(ui.queryByText('保存变更到原表')).toBeNull(); expect(commit).not.toHaveBeenCalled();
  });
});
afterEach(() => { cleanup(); vi.unstubAllGlobals(); });
async function click(button: HTMLElement) { await act(async () => { fireEvent.click(button); }); }
describe("native preview decision buttons", () => {
  it("stages without committing, keeps decisions inside the owning message, and only commits on save", async () => {
    const commit = vi.fn(async () => true), store = vi.spyOn(Storage.prototype, "setItem");
    const ui = render(<Harness commit={commit} />);
    await click(ui.getAllByText("生成副本并排审阅")[0]);
    expect(mocks.stage).toHaveBeenCalledWith(p, ["Ledger"]);
    expect(JSON.stringify(mocks.stage.mock.calls[0])).toContain("LOCAL_PRIVATE_SENTINEL");
    expect(commit).not.toHaveBeenCalled(); expect(store).not.toHaveBeenCalled();
    expect(ui.getByTestId("original-message").contains(ui.getByText("保存变更到原表"))).toBe(true);
    expect(ui.getByTestId("later-message").textContent).toBe("后续普通对话");
    expect(ui.getByTestId("second-plan").querySelector("button")?.disabled).toBe(true);
    await click(ui.getByText("保存变更到原表"));
    expect(mocks.guard).toHaveBeenCalledOnce(); expect(commit).toHaveBeenCalledWith(session); expect(mocks.close).toHaveBeenCalledWith(session);
    expect(ui.queryByText("保存变更到原表")).toBeNull();
    store.mockRestore();
  });
  it("discard only closes the reviewed copy, and cancelled close leaves decisions available", async () => {
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await click(ui.getAllByText("生成副本并排审阅")[0]); mocks.close.mockResolvedValueOnce(false);
    await click(ui.getByText("放弃变更"));
    expect(ui.getByText("保存变更到原表")).toBeTruthy(); expect(commit).not.toHaveBeenCalled();
    await click(ui.getByText("放弃变更"));
    expect(ui.queryByText("保存变更到原表")).toBeNull(); expect(commit).not.toHaveBeenCalled();
  });
  it("guard failure prevents commit and invalidates the save button", async () => {
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await click(ui.getAllByText("生成副本并排审阅")[0]); mocks.guard.mockRejectedValue(new Error("现场已变化"));
    await click(ui.getByText("保存变更到原表"));
    expect(commit).not.toHaveBeenCalled(); expect(ui.queryByText("保存变更到原表")).toBeNull();
    expect(ui.getByText("现场已变化")).toBeTruthy(); expect(ui.getByText("关闭副本")).toBeTruthy();
  });
  it("does not replay a successful write when close is cancelled", async () => {
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await click(ui.getAllByText("生成副本并排审阅")[0]); mocks.close.mockResolvedValueOnce(false);
    await click(ui.getByText("保存变更到原表"));
    expect(ui.queryByText("保存变更到原表")).toBeNull(); expect(ui.getByText("关闭副本")).toBeTruthy();
    await click(ui.getByText("关闭副本")); expect(commit).toHaveBeenCalledOnce();
  });
  it("a partial/failed execution cannot be blindly retried, and original execution evidence remains authoritative", async () => {
    const commit = vi.fn(async () => false), ui = render(<Harness commit={commit} />);
    await click(ui.getAllByText("生成副本并排审阅")[0]); await click(ui.getByText("保存变更到原表"));
    expect(ui.queryByText("保存变更到原表")).toBeNull(); expect(commit).toHaveBeenCalledOnce();
    expect(mocks.close).not.toHaveBeenCalled(); expect(ui.getByText(/写回未完成/)).toBeTruthy();
    await click(ui.getByText("关闭副本"));
    expect(ui.getByText(/关闭副本不撤销既有写入/)).toBeTruthy();
    expect(ui.queryByText("已放弃审阅副本，未写回原表。")).toBeNull();
  });
});
