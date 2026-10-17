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
// `staleAcknowledgedCopy`：模拟"旧请求带着**上一份**副本的确认"直接调写回入口（已排队的点击 / 快捷键 / 程序直调）。
// `saveMessageId`：旧请求打向哪张卡。放弃后 m1 保留终态卡片，新的副本由 m2 生成 ⇒ 旧请求要打向"现在在场的那张卡"。
function Harness({ commit, authorization = ["Ledger"], paneNeedsResize = false, beforeOperation, staleAcknowledgedCopy = null, saveMessageId = "m1" }: { commit: (s: ExcelNativeComparison | null) => Promise<boolean>; authorization?: string[]; paneNeedsResize?: boolean; beforeOperation?: () => Promise<boolean>; staleAcknowledgedCopy?: string | null; saveMessageId?: string }) {
  const controller = useExcelNativeComparison();
  return <>
    <div data-testid="original-message"><ExcelNativeReviewActions messageId="m1" plan={p} authorization={authorization} controller={controller} disabled={controller.pending} commit={commit} paneNeedsResize={paneNeedsResize} beforeOperation={beforeOperation} /></div>
    <div data-testid="later-message">后续普通对话</div>
    <div data-testid="second-plan"><ExcelNativeReviewActions messageId="m2" plan={p} authorization={authorization} controller={controller} disabled={controller.pending} commit={commit} /></div>
    <button onClick={() => void controller.save(saveMessageId, p, authorization, commit)}>test direct save without review acknowledgement</button>
    <button onClick={() => void controller.save(saveMessageId, p, authorization, commit, staleAcknowledgedCopy)}>test direct save with a previous copy acknowledgement</button>
  </>;
}
beforeEach(() => {
  vi.clearAllMocks(); mocks.mode.mockReturnValue('native'); mocks.stage.mockResolvedValue(session); mocks.guard.mockResolvedValue(undefined); mocks.close.mockResolvedValue('closed');
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
  await click(ui.getByText("确认写回原表"));
  expect(commit).not.toHaveBeenCalled(); expect(mocks.guard).not.toHaveBeenCalled();
  expect(beforeOperation).toHaveBeenCalledTimes(3);
});
/*
 * 写回入口唯一那个"在 `save()` 之外"的 await 是 `beforeOperation`（授权/目标检查）。
 *
 * 它回来之后，现场可能已经换人：用户在这段等待里点了「放弃变更」，并由**另一条消息**
 * 生成了新副本。旧请求仍打向原来那张卡 —— 它必须被拒，不能拿新副本的现场去写原表。
 */
it("授权检查在途期间现场换掉了（放弃并生成新副本）⇒ 旧请求回来后不写回", async () => {
  const commit = vi.fn(async () => true);
  let calls = 0;
  let release!: (value: boolean) => void;
  const pending = new Promise<boolean>(resolve => { release = resolve; });
  // 第 1 次（生成副本）照常放行；第 2 次（点写回）挂起，把"在途"这个瞬间固定下来。
  const beforeOperation = vi.fn(async () => (++calls === 2 ? pending : true));
  const ui = render(<Harness commit={commit} beforeOperation={beforeOperation} />);
  await click(ui.getAllByText("生成副本并排审阅")[0]);   // m1 生成 EB_Test
  await click(ui.getByText("确认写回原表"));             // 卡在授权检查里
  expect(commit).not.toHaveBeenCalled();

  // 在途期间：放弃（终态）→ 由另一条消息生成新副本 EB_New。
  await click(ui.getByText("放弃变更"));
  mocks.stage.mockResolvedValue({ ...session, copySheetName: "EB_New" });
  await click(ui.getAllByText("生成副本并排审阅")[0]);
  expect(commit).not.toHaveBeenCalled();

  // 放行 ⇒ 旧请求打向 m1，而现在的现场属于另一张卡 ⇒ 一个格都不写。
  await act(async () => { release(true); });
  expect(commit).not.toHaveBeenCalled();
  expect(mocks.guard).not.toHaveBeenCalled();
  expect(ui.getByRole("status").textContent).toContain("没有可保存的已审阅副本");

  // 正向锚点：新的那张卡自己点写回，同一条链路必须放行（不是"永远拒绝写回"）。
  await click(ui.getByText("确认写回原表"));
  expect(mocks.guard).toHaveBeenCalledOnce();
  expect(commit).toHaveBeenCalledOnce();
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
    expect((ui.getByText('确认写回原表') as HTMLButtonElement).disabled).toBe(true);
    expect(commit).not.toHaveBeenCalled();
  });
  it('explains the actual pane-width problem without claiming automatic success or granting save', async () => {
    mocks.stage.mockResolvedValue({...manualSession,windowLayout:'arranged'});
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} paneNeedsResize />);
    await stage(ui);
    // 首屏那一行是用户能读懂的一句话，且是**可收起**的（summary 属于 details）。
    const hint = ui.getByText(/如原表显示过窄/);
    expect(hint.textContent).toContain('可拖动 AI 窗格左边界');
    expect(hint.closest('details')).not.toBeNull();
    // 真实区分两件事：窗口排列已完成 vs AI 比例尚未确认。目标比例收在展开层里。
    expect(ui.getByText(/整个可用宽度约 30%/)).toBeTruthy();
    // 旧口径"把 AI 调到右侧 Excel 窗口约 40%"正是把原表挤没的算法，必须不再出现。
    expect(document.body.textContent).not.toContain('右侧 Excel 窗口约');
    expect(document.body.textContent).not.toContain('已按 35/35/30');
    expect((ui.getByText('确认写回原表') as HTMLButtonElement).disabled).toBe(true);
    expect(commit).not.toHaveBeenCalled();
  });
  it('AI 宽度确实已确认时，一句提示都不出现（不打扰）', async () => {
    mocks.stage.mockResolvedValue({...manualSession,windowLayout:'arranged'});
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await stage(ui);
    expect(ui.queryByText(/如原表显示过窄/)).toBeNull();
    expect(ui.queryByText(/整个可用宽度约/)).toBeNull();
  });
  it('keeps initial staging but removes the rearrange button', async () => {
    mocks.stage.mockResolvedValue({...manualSession,windowLayout:'arranged'});
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await stage(ui);
    expect(ui.queryByRole('button', {name:'重新排列'})).toBeNull();
    expect(commit).not.toHaveBeenCalled();
    expect((ui.getByText('确认写回原表') as HTMLButtonElement).disabled).toBe(true);
  });
  it('shows a restart instruction for an old backend, never new-order success', async () => {
    mocks.stage.mockResolvedValue({...manualSession,windowLayout:'service_outdated'});
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await stage(ui);
    // 开发环境重启提示**只在诊断入口**出现，不进首屏。
    const diagnostics = ui.getByText('诊断详情').closest('details')!;
    expect(diagnostics.textContent).toContain('重启 Python 服务');
    expect(ui.getByText(/诊断详情/).closest('section')!.querySelector('.excel-review-title')!.textContent)
      .toBe('请核对审阅副本');
    expect(ui.queryByText(/已自动排列/)).toBeNull();
    expect((ui.getByText('确认写回原表') as HTMLButtonElement).disabled).toBe(true);
    expect(commit).not.toHaveBeenCalled();
  });
  it('automatic layout is not copy verification or permission to write', async () => {
    mocks.stage.mockResolvedValue({...manualSession, windowLayout:'arranged'});
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await stage(ui);
    // 首屏只说明一次左右布局（卡片说明），不再重复念第二遍"已自动排列：左侧…"。
    expect(ui.getByText('左侧是审阅副本，右侧是原表。原表尚未修改。')).toBeTruthy();
    expect(ui.queryByText(/已自动排列/)).toBeNull();
    expect((ui.getByText('确认写回原表') as HTMLButtonElement).disabled).toBe(true);
    await click(ui.getByText('test direct save without review acknowledgement'));
    expect(commit).not.toHaveBeenCalled();
    await click(ui.getByRole('checkbox')); await click(ui.getByText('确认写回原表'));
    expect(mocks.guard).toHaveBeenCalledOnce(); expect(commit).toHaveBeenCalledOnce();
    // 排列成功既不等于核对通过、也不等于已关闭副本；勾选确认写回后才去关闭。
    expect(mocks.close).toHaveBeenCalledOnce();
  });
  it('requires real checkbox acknowledgement and then runs the original guard + commit; ending the review reports the copy cleanup separately', async () => {
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await stage(ui); expect(commit).not.toHaveBeenCalled();
    expect(ui.getByText(/视图 → 并排查看/)).toBeTruthy();
    const save = ui.getByText('确认写回原表') as HTMLButtonElement;
    expect(save.disabled).toBe(true);
    await click(ui.getByText('test direct save without review acknowledgement'));
    expect(commit).not.toHaveBeenCalled(); expect(mocks.guard).not.toHaveBeenCalled();
    await click(ui.getByRole('checkbox', {name:'我已核对副本，确认按此计划修改原表。'}));
    expect(save.disabled).toBe(false); await click(save);
    expect(mocks.guard).toHaveBeenCalledOnce(); expect(commit).toHaveBeenCalledWith(manualSession);
    // 已经实现：兼容副本也能被自动关闭（本机桥按 nonce 身份关闭），不再宣称"不能自动关闭"。
    expect(mocks.close).toHaveBeenCalledWith(manualSession);
    expect(ui.queryByText('确认写回原表')).toBeNull();
    // 业务结果与副本清理结果**分两处**独立陈述。
    expect(ui.getByText('已写回原表')).toBeTruthy();
    expect(ui.getByText(/原表已按本卡审阅的计划写入，并通过执行侧逐格回读核对。/)).toBeTruthy();
    expect(ui.getByText('审阅副本已关闭。')).toBeTruthy();
    // 磁盘清理结果属于诊断详情，不进主卡结论。
    expect(ui.getByText(/尚无已确认的恢复文件/)).toBeTruthy();
    expect(ui.getByText('诊断详情').closest('details')!.textContent).toContain('尚无已确认的恢复文件');
    expect(ui.getByText(/原表已按本卡审阅的计划写入/).closest('.excel-review-description')!.textContent)
      .not.toContain('磁盘文件');
  });
  it('放弃后立刻撤下确认框与原审阅按钮，旧回调也拒绝写回；新副本要重新勾选', async () => {
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await stage(ui); await click(ui.getByRole('checkbox'));
    await click(ui.getByText('放弃变更'));
    expect(commit).not.toHaveBeenCalled(); expect(mocks.close).toHaveBeenCalledWith(manualSession);
    // 主卡只讲"计划已取消"，副本状态独立且次要。
    expect(ui.getByText('已放弃变更')).toBeTruthy();
    expect(ui.getByText('本次计划未写入原表。')).toBeTruthy();
    expect(ui.getByText('审阅副本已关闭。')).toBeTruthy();
    // 确认框与三个审阅按钮全部撤下。
    expect(ui.queryByRole('checkbox')).toBeNull();
    expect(ui.queryByText('确认写回原表')).toBeNull();
    expect(ui.queryByText('重新排列')).toBeNull();
    expect(ui.queryByText('放弃变更')).toBeNull();
    // **旧回调被拒绝**：即便拿到放弃之前的旧闭包，也不可能把计划写进原表。
    await click(ui.getByText('test direct save without review acknowledgement'));
    expect(commit).not.toHaveBeenCalled(); expect(mocks.guard).not.toHaveBeenCalled();
    expect(ui.getByText(/没有可保存的已审阅副本/)).toBeTruthy();
    // 已放弃是终态：允许对同一计划重新生成副本，而且新副本要重新勾选。
    mocks.stage.mockResolvedValue({...manualSession,copySheetName:'EB_New'});
    await stage(ui);
    expect((ui.getByText('确认写回原表') as HTMLButtonElement).disabled).toBe(true);
  });
  it('确认勾选项绑在**这一份**副本上：带着上一份副本的确认也不能写回新副本', async () => {
    const commit = vi.fn(async () => true);
    /*
     * 放弃后 m1 保留终态卡片，新的副本由 **m2** 那张卡生成（见下：副本表名 EB_New）。
     * 所以"现在在场的那张卡"是 m2 —— 旧请求必须打向它，才能单独检验"确认绑在这一份副本上"
     * 这一条判据，而不是被"根本没有可保存的副本"这条更早的判据先挡掉。
     */
    const ui = render(<Harness commit={commit} staleAcknowledgedCopy="EB_Test" saveMessageId="m2" />);
    await stage(ui);   // 副本 EB_Test（m1）
    await click(ui.getByRole('checkbox'));

    await click(ui.getByText('放弃变更'));
    mocks.stage.mockResolvedValue({...manualSession, copySheetName:'EB_New'});
    await stage(ui);   // 新副本 EB_New（m2）：确认框已复位，写回入口仍禁用
    expect((ui.getByText('确认写回原表') as HTMLButtonElement).disabled).toBe(true);

    /*
     * 一个"旧请求"（已排队的点击 / 快捷键 / 程序直调）带着**上一份副本**的那个
     * "已核对"来点写回。它绝不能把新副本的计划写进原表 —— 用户核对过的那一份
     * （EB_Test）从来没被写，没核对过的那一份（EB_New）反而被写。
     */
    await click(ui.getByText('test direct save with a previous copy acknowledgement'));
    expect(commit).not.toHaveBeenCalled();
    expect(mocks.guard).not.toHaveBeenCalled();
    expect(ui.getByRole('status').textContent).toContain('请先在 Excel 中审阅兼容副本');

    // 正向锚点：给**新的**这份副本重新勾选，同一条链路必须放行。
    await click(ui.getByRole('checkbox'));
    await click(ui.getByText('确认写回原表'));
    expect(commit).toHaveBeenCalledOnce();
  });
  it('guard conflicts block the clicked manual save; a failed/uncertain writeback RETAINS the copy and never says the original was unwritten', async () => {
    const commit = vi.fn(async () => false), ui = render(<Harness commit={commit} />);
    await stage(ui); await click(ui.getByRole('checkbox'));
    mocks.guard.mockRejectedValueOnce(new Error('原表已变化'));
    await click(ui.getByText('确认写回原表'));
    expect(commit).not.toHaveBeenCalled(); expect(ui.queryByText('确认写回原表')).toBeNull();
    await click(ui.getByText('保留副本并结束审阅'));
    // 写回失败/不确定 ⇒ 默认保留副本供核对；不自动关闭、不删除排查材料。
    expect(mocks.close).not.toHaveBeenCalled();
    expect(ui.getByText(/已保留审阅副本供你核对/)).toBeTruthy();
    expect(ui.queryByText(/未写回原表/)).toBeNull();
    mocks.stage.mockResolvedValue({...manualSession,copySheetName:'EB_AfterConflict'});
    await stage(ui); await click(ui.getByRole('checkbox'));
    await click(ui.getByText('确认写回原表')); expect(commit).toHaveBeenCalledOnce();
    await click(ui.getByText('保留副本并结束审阅'));
    expect(mocks.close).not.toHaveBeenCalled();
    expect(ui.getByText(/已保留审阅副本供你核对/)).toBeTruthy();
  });
  it('does not publish a ready review when workbook creation fails', async () => {
    mocks.stage.mockRejectedValueOnce(new Error('副本打开被拒绝'));
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await stage(ui); expect(ui.queryByRole('checkbox')).toBeNull();
    expect(ui.queryByText('确认写回原表')).toBeNull(); expect(commit).not.toHaveBeenCalled();
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
    expect(ui.getByTestId("original-message").contains(ui.getByText("确认写回原表"))).toBe(true);
    expect(ui.getByTestId("later-message").textContent).toBe("后续普通对话");
    expect(ui.getByTestId("second-plan").querySelector("button")?.disabled).toBe(true);
    await click(ui.getByText("确认写回原表"));
    expect(mocks.guard).toHaveBeenCalledOnce(); expect(commit).toHaveBeenCalledWith(session); expect(mocks.close).toHaveBeenCalledWith(session);
    expect(ui.queryByText("确认写回原表")).toBeNull();
    store.mockRestore();
  });
  it("放弃后副本没收尾：清理状态是次要的，绝不恢复审阅状态、也不许推断保存提示", async () => {
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await click(ui.getAllByText("生成副本并排审阅")[0]); mocks.close.mockResolvedValueOnce('close_unconfirmed');
    await click(ui.getByText("放弃变更"));
    // 「关闭尚未确认」≠「宿主在问是否保存」：没有对话框证据时不许这么写。
    expect(ui.getByText(/关闭尚未确认/)).toBeTruthy();
    expect(ui.queryByText(/正在询问是否保存/)).toBeNull();
    // **清理失败不恢复审阅状态**：确认框与写回入口都不许回来。
    expect(ui.queryByText("确认写回原表")).toBeNull();
    expect(ui.queryByRole('checkbox')).toBeNull();
    expect(ui.queryByText("放弃变更")).toBeNull();
    expect(commit).not.toHaveBeenCalled();
    // 主结论仍是"计划已取消"，副本状态另起一处且不冲淡它。
    expect(ui.getByText("已放弃变更")).toBeTruthy();
    expect(ui.getByText("本次计划未写入原表。")).toBeTruthy();
    expect(ui.getByText(/审阅副本仍打开，请核对窗口名称后手动关闭。若副本中有需要保留的修改，请先另存。/)).toBeTruthy();
    // 旧回调同样被拒绝。
    await click(ui.getByText('test direct save without review acknowledgement'));
    expect(commit).not.toHaveBeenCalled();
  });
  it("guard failure prevents commit and invalidates the save button", async () => {
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await click(ui.getAllByText("生成副本并排审阅")[0]); mocks.guard.mockRejectedValue(new Error("现场已变化"));
    await click(ui.getByText("确认写回原表"));
    expect(commit).not.toHaveBeenCalled(); expect(ui.queryByText("确认写回原表")).toBeNull();
    expect(ui.getByText("现场已变化")).toBeTruthy(); expect(ui.getByText("保留副本并结束审阅")).toBeTruthy();
  });
  it("写回成功但副本里还有未经确认的人工修改：业务结论不变，副本状态独立显示", async () => {
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await click(ui.getAllByText("生成副本并排审阅")[0]); mocks.close.mockResolvedValueOnce('extra_edits');
    await click(ui.getByText("确认写回原表"));
    // 副本里有未经确认的修改 ⇒ 不关副本；但**关闭失败不改写业务结论**。
    expect(ui.queryByText("确认写回原表")).toBeNull(); expect(ui.getByText("关闭审阅副本")).toBeTruthy();
    expect(ui.getByText("已写回原表")).toBeTruthy();
    expect(ui.getByText("副本里还有未经确认的人工修改，未自动关闭。")).toBeTruthy();
    // 业务结论一个字都不许被改写成失败。
    expect(ui.queryByText(/写回失败/)).toBeNull(); expect(ui.queryByText(/未写回原表/)).toBeNull();
    // 副本状态不许混用另外两档的措辞。
    expect(ui.queryByText(/无法确认副本是否被额外修改/)).toBeNull();
    expect(ui.queryByText(/关闭尚未确认/)).toBeNull();
    await click(ui.getByText("关闭审阅副本")); expect(commit).toHaveBeenCalledOnce();
  });
  it("a partial/failed execution cannot be blindly retried, and original execution evidence remains authoritative", async () => {
    const commit = vi.fn(async () => false), ui = render(<Harness commit={commit} />);
    await click(ui.getAllByText("生成副本并排审阅")[0]); await click(ui.getByText("确认写回原表"));
    expect(ui.queryByText("确认写回原表")).toBeNull(); expect(commit).toHaveBeenCalledOnce();
    expect(mocks.close).not.toHaveBeenCalled();
    // 执行器跑过但没确认 ⇒ 归"结果不确定"，**不套**成功模板，也不冒充"没写成"。
    expect(ui.getByText("写回结果不确定")).toBeTruthy();
    expect(ui.queryByText("已写回原表")).toBeNull();
    expect(ui.queryByText("写回未完成")).toBeNull();
    expect(ui.getByText(/部分动作已经执行/)).toBeTruthy();
    // 失败阶段没有写回/放弃入口，只有一个结束会话的动作。
    expect(ui.queryByText("放弃变更")).toBeNull();
    expect(ui.queryByRole('checkbox')).toBeNull();
    await click(ui.getByText("保留副本并结束审阅"));
    expect(mocks.close).not.toHaveBeenCalled();
    expect(ui.getByText(/已保留审阅副本供你核对/)).toBeTruthy();
    expect(ui.queryByText("已放弃审阅副本，未写回原表。")).toBeNull();
  });
  it("守卫在写入前拦下时归「写回未完成」，不冒充「结果不确定」", async () => {
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await click(ui.getAllByText("生成副本并排审阅")[0]);
    mocks.guard.mockRejectedValueOnce(new Error("现场已变化"));
    await click(ui.getByText("确认写回原表"));
    // 原表一格未写：这一档与"跑到一半没确认"必须说成两句不同的话。
    expect(ui.getByText("写回未完成")).toBeTruthy();
    expect(ui.queryByText("写回结果不确定")).toBeNull();
    expect(ui.getByText(/也可能一格未改/)).toBeTruthy();
    // 实际原因照实显示。
    expect(ui.getByText("现场已变化")).toBeTruthy();
  });
});
/**
 * 审阅卡的**四个阶段**逐阶段验收：审阅中 / 已放弃 / 写回成功 / 写回失败。
 *
 * 与 WPS 端同一组口径（`packages/core/src/reviewCardPresentation.ts`）：哪个阶段说什么话、
 * 确认框与按钮在不在、副本状态是不是独立显示、不同的结果有没有互相冒充。
 * 这里跑的是**兼容（人工核对）模式** —— 只有它才有"我已核对副本"这句必要说明。
 */
describe("审阅卡四阶段界面（审阅中 / 已放弃 / 写回成功 / 写回失败）", () => {
  const manualSession: ExcelNativeComparison = { ...session, mode: "manual", actionEvidence: "", copyEvidence: "" };
  async function stage(ui: ReturnType<typeof render>) {
    await click(ui.getAllByText("生成副本并排审阅")[0]);
  }
  beforeEach(() => {
    mocks.mode.mockReturnValue("manual");
    mocks.stage.mockResolvedValue(manualSession);
  });

  it("审阅中：首屏只留五项，确认框是真闸，排查材料全部下沉", async () => {
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await stage(ui);

    const card = ui.getByText("请核对审阅副本").closest("section")!;
    // 五项：标题 / 说明 / 范围 / 兼容模式必要说明 / 确认框。
    expect(card.querySelector(".excel-review-title")!.textContent).toBe("请核对审阅副本");
    expect(card.querySelector(".excel-review-description")!.textContent)
      .toBe("左侧是审阅副本，右侧是原表。原表尚未修改。");
    expect(card.querySelector(".excel-review-scope")!.textContent)
      .toBe("范围：Ledger!C2 · 写入数值 · 涉及 1 个单元格");
    expect(card.querySelector(".excel-review-compat")!.textContent)
      .toBe("请在 Excel 中核对副本。确认后将按本计划修改原表，并验证执行结果；副本中的人工修改不会写回。");
    expect(card.querySelector(".excel-manual-review-confirm")!.textContent)
      .toContain("我已核对副本，确认按此计划修改原表。");

    // 三个按钮；写回入口在勾选前**不可点**。
    const buttons = card.querySelector(".excel-native-review-buttons")!;
    expect(Array.from(buttons.querySelectorAll("button")).map((item) => (item.textContent ?? "").trim()))
      .toEqual(["放弃变更", "确认写回原表"]);
    expect((ui.getByText("确认写回原表") as HTMLButtonElement).disabled).toBe(true);

    // 排查材料全部下沉：首屏那五项里不得出现副本标识 / 关闭诊断 / 磁盘清理。
    const front = [".excel-review-title", ".excel-review-description", ".excel-review-scope",
      ".excel-review-compat", ".excel-manual-review-confirm"]
      .map((selector) => card.querySelector(selector)?.textContent ?? "")
      .join(" ");
    expect(front).toContain("请核对审阅副本");
    expect(front).not.toContain("EB_Test");
    expect(front).not.toContain("恢复文件");
    expect(front).not.toContain("关闭结果：");
    // 但**可查到**：诊断详情折叠里逐组都在（收起 ≠ 删除）。
    const diagnostics = ui.getByText("诊断详情").closest("details")!;
    expect(diagnostics.textContent).toContain("本次副本工作表名：EB_Test");
    expect(diagnostics.textContent).toContain("原表工作簿标识：url:file:///E:/Original.xlsx");
    expect(diagnostics.textContent).toContain("恢复文件");
    // 卡面不许出现 Markdown 记号（星号会原样显示给用户）。
    expect(card.textContent).not.toMatch(/\*\*|__/);

    // 真闸的正向锚点：勾选后入口立刻可用，并如实走到 guard + commit。
    await click(ui.getByRole("checkbox"));
    expect((ui.getByText("确认写回原表") as HTMLButtonElement).disabled).toBe(false);
    await click(ui.getByText("确认写回原表"));
    expect(mocks.guard).toHaveBeenCalledOnce();
    expect(commit).toHaveBeenCalledOnce();
  });

  it("已放弃：终态撤下确认框与原审阅按钮，旧回调也拒绝继续写回", async () => {
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await stage(ui);
    // 正面证据：闸门确实接在这条链路上 —— 没勾选就直调保存，一步都不走。
    await click(ui.getByText("test direct save without review acknowledgement"));
    expect(commit).not.toHaveBeenCalled();
    expect(mocks.guard).not.toHaveBeenCalled();

    await click(ui.getByRole("checkbox"));
    await click(ui.getByText("放弃变更"));

    // 主结论：计划已取消、原表未写入；副本状态独立且次要。
    expect(ui.getByText("已放弃变更")).toBeTruthy();
    expect(ui.getByText("本次计划未写入原表。")).toBeTruthy();
    expect(ui.getByText("审阅副本已关闭。")).toBeTruthy();
    // 确认框与三个审阅按钮全部撤下。
    expect(ui.queryByRole("checkbox")).toBeNull();
    expect(ui.queryByText("确认写回原表")).toBeNull();
    expect(ui.queryByText("重新排列")).toBeNull();
    expect(ui.queryByText("放弃变更")).toBeNull();

    // **旧回调被拒绝**：拿到放弃之前的旧入口再调一次，原表一格都不许动。
    await click(ui.getByText("test direct save without review acknowledgement"));
    expect(commit).not.toHaveBeenCalled();
    expect(mocks.guard).not.toHaveBeenCalled();
    expect(ui.getByText(/没有可保存的已审阅副本/)).toBeTruthy();
    // 结论没有被改写。
    expect(ui.queryByText("已写回原表")).toBeNull();
  });

  it("写回成功：业务结论在前，副本状态独立显示，不再保留任何审阅入口", async () => {
    const commit = vi.fn(async () => true), ui = render(<Harness commit={commit} />);
    await stage(ui);
    await click(ui.getByRole("checkbox"));
    await click(ui.getByText("确认写回原表"));

    // 审阅入口全部撤下。
    expect(ui.queryByRole("checkbox")).toBeNull();
    expect(ui.queryByText("确认写回原表")).toBeNull();
    expect(ui.queryByText("放弃变更")).toBeNull();
    expect(ui.queryByText("重新排列")).toBeNull();
    // 业务结论来自 outcome，副本状态**独立**成句、不带业务前缀。
    expect(ui.getByText("已写回原表")).toBeTruthy();
    expect(ui.getByText("原表已按本卡审阅的计划写入，并通过执行侧逐格回读核对。")).toBeTruthy();
    const copyStatus = document.querySelector(".excel-review-copy-status")!;
    expect(copyStatus.textContent).toContain("审阅副本已关闭。");
    expect(copyStatus.textContent).not.toContain("已写回原表");
    // 业务结论排在副本状态**前面**（主次之分）。
    const headline = document.querySelector(".excel-review-description")!;
    expect((headline.compareDocumentPosition(copyStatus) & Node.DOCUMENT_POSITION_FOLLOWING) !== 0)
      .toBe(true);
    // 不同结果不混用文案。
    expect(ui.queryByText("已放弃变更")).toBeNull();
    expect(ui.queryByText(/本次计划未写入原表/)).toBeNull();
    // 副本已收尾 ⇒ 不再提供"再关一次"。
    expect(ui.queryByText("关闭审阅副本")).toBeNull();
  });

  it("写回失败：按实际状态展示，绝不套成功模板、两档互不冒充", async () => {
    // ① 守卫在写入前拦下 ⇒ 原表一格未写，归「写回未完成」。
    const blockedCommit = vi.fn(async () => true);
    const blocked = render(<Harness commit={blockedCommit} />);
    await stage(blocked);
    mocks.guard.mockRejectedValueOnce(new Error("现场已变化"));
    await click(blocked.getByRole("checkbox"));
    await click(blocked.getByText("确认写回原表"));
    expect(blocked.getByText("写回未完成")).toBeTruthy();
    expect(blocked.getByText(/也可能一格未改/)).toBeTruthy();
    expect(blocked.getByText("现场已变化")).toBeTruthy();
    // 不冒充另一档，更不套成功模板 / 不混用已放弃口径。
    expect(blocked.queryByText("写回结果不确定")).toBeNull();
    expect(blocked.queryByText("已写回原表")).toBeNull();
    expect(blocked.queryByText(/验证通过/)).toBeNull();
    expect(blocked.queryByText("已放弃变更")).toBeNull();
    expect(blocked.queryByText(/本次计划未写入原表/)).toBeNull();
    // 失败阶段没有写回/放弃入口，只有一个结束会话的动作。
    expect(blocked.queryByRole("checkbox")).toBeNull();
    expect(blocked.queryByText("确认写回原表")).toBeNull();
    expect(blocked.queryByText("放弃变更")).toBeNull();
    expect(blocked.getByText("保留副本并结束审阅")).toBeTruthy();
    expect(mocks.close).not.toHaveBeenCalled();

    // ② 执行器跑过但没确认 ⇒ 归「结果不确定」，与①必须说成两句不同的话。
    cleanup();
    const uncertainCommit = vi.fn(async () => false);
    const uncertain = render(<Harness commit={uncertainCommit} />);
    await stage(uncertain);
    await click(uncertain.getByRole("checkbox"));
    await click(uncertain.getByText("确认写回原表"));
    expect(uncertain.getByText("写回结果不确定")).toBeTruthy();
    expect(uncertain.getByText(/部分动作已经执行/)).toBeTruthy();
    expect(uncertain.queryByText("写回未完成")).toBeNull();
    expect(uncertain.queryByText(/也可能一格未改/)).toBeNull();
    expect(uncertain.queryByText("已写回原表")).toBeNull();
    expect(uncertain.queryByText(/验证通过/)).toBeNull();
    expect(uncertain.queryByText("已放弃变更")).toBeNull();
  });
});

describe("格式能力检测前移到预览阶段", () => {
  const formatPlan: AnalysisPlan = {
    ...p,
    actions: [{ type: "setFont", sheet: "Ledger", range: "C2", bold: true }]
  };
  function NoticeHarness({ plan }: { plan: AnalysisPlan }) {
    const controller = useExcelNativeComparison();
    return (
      <ExcelNativeReviewActions
        messageId="m1"
        plan={plan}
        authorization={["Ledger"]}
        controller={controller}
        disabled={controller.pending}
        commit={vi.fn(async () => true)}
      />
    );
  }
  function stubExcelApi(isSetSupported: (version: string) => boolean) {
    vi.stubGlobal("Office", {
      context: { requirements: { isSetSupported: (_set: string, version: string) => isSetSupported(version) } }
    });
  }
  it("ExcelApi ≤1.17 宿主：不阻断，但首屏短说明 + 详情点缺 1.19 与受影响步骤", () => {
    stubExcelApi((version) => version !== "1.19");
    const ui = render(<NoticeHarness plan={formatPlan} />);
    const notice = ui.getByText(/ExcelApi 1\.19/).closest(".excel-format-capability-notice");
    expect(notice).not.toBeNull();
    expect(notice!.getAttribute("role")).toBe("status");
    // 首屏短说明点出受影响步骤；技术原因（缺 1.19、直接格式路径）收进「技术原因」详情。
    const summary = notice!.querySelector(".excel-format-capability-summary");
    expect(summary?.textContent).toContain("第 1 步");
    expect(notice!.textContent).toContain("直接格式");
    // 非阻断：入口照常可用，用户仍能生成副本审阅。
    expect(ui.getAllByText("生成副本并排审阅").length).toBeGreaterThan(0);
  });
  it("ExcelApi 1.19 宿主：不打扰用户（无提示）", () => {
    stubExcelApi(() => true);
    const ui = render(<NoticeHarness plan={formatPlan} />);
    expect(ui.queryByText(/excel-format-capability-notice/)).toBeNull();
    expect(document.querySelector(".excel-format-capability-notice")).toBeNull();
  });
  it("连 1.9 都没有的宿主：首屏警示短说明 + 详情点缺 API", () => {
    stubExcelApi((version) => version !== "1.19" && version !== "1.9");
    const ui = render(<NoticeHarness plan={formatPlan} />);
    const notice = ui.getByText(/ExcelApi 1\.9/).closest(".excel-format-capability-notice");
    expect(notice).not.toBeNull();
    expect(notice!.getAttribute("role")).toBe("status");
    expect(notice!.querySelector(".excel-format-capability-summary")?.textContent).toContain("第 1 步");
    expect(notice!.textContent).toContain("无法在执行前保存可恢复的格式依据");
  });
});
