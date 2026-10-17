// @vitest-environment jsdom
import { act, cleanup, fireEvent, render, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { AuthorizedTarget } from "@excel-bro/core/authorizedTarget";
import { useAuthorizedTarget } from "./useAuthorizedTarget";
import { AuthorizedTargetConfirmation } from "../components/AuthorizedTargetConfirmation";

const original: AuthorizedTarget = { fileKey: "PRIVATE_PATH_A", name: "Original.xlsx",
  sheets: [{ id: "1", name: "Sheet1" }, { id: "2", name: "Other" }], selected: ["1"] };
const other: AuthorizedTarget = { ...original, fileKey: "PRIVATE_PATH_B", name: "Other.xlsx" };
afterEach(cleanup);
function mount(options: Partial<Pick<Parameters<typeof useAuthorizedTarget>[0], "reviewActive" | "isReviewCopy" | "analysisEvidence">> = {}) {
  let current: AuthorizedTarget | null = original;
  const inspect = vi.fn(async () => { if (!current) throw new Error("PRIVATE_FAILURE"); return current; });
  const adopt = vi.fn((target: AuthorizedTarget) => { current = target; });
  const operation = vi.fn();
  let guard!: ReturnType<typeof useAuthorizedTarget>;
  function Harness() {
    guard = useAuthorizedTarget({ enabled: true, inspect, adopt, ...options });
    return <>
      <textarea defaultValue="original request and confirmed answer" />
      <button onClick={() => void (async () => {
        if (await guard.check(() => { void guard.check().then((ready) => { if (ready) operation(); }); })) operation();
      })()}>操作</button>
      <button onClick={() => void (async () => {
        const run = () => { void guard.checkForAnalysis().then((ready) => { if (ready) operation(); }); };
        if (await guard.checkForAnalysis(run)) operation();
      })()}>分析需求</button>
      {guard.notice && <AuthorizedTargetConfirmation notice={guard.notice} busy={guard.confirming} error={guard.error}
        onConfirm={(choice, selected) => void guard.confirm(choice, selected)} onCancel={guard.cancel}
        onRecheck={() => void guard.recheck()} />}
    </>;
  }
  const ui = render(<Harness />);
  return { ui, inspect, adopt, operation, current: (value: AuthorizedTarget | null) => { current = value; }, guard: () => guard };
}
async function initialized(h: ReturnType<typeof mount>) { await waitFor(() => expect(h.guard().binding).not.toBeNull()); }
describe("pre-operation target confirmation shared by both hosts", () => {
  it("the existing selector can confirm and resume exactly one queued analysis", async () => {
    const h = mount({ analysisEvidence: vi.fn(async () => ({ key: "private-key", conclusive: false })) });
    await initialized(h);
    fireEvent.click(h.ui.getByText("分析需求"));
    await waitFor(() => expect(h.guard().notice?.reason).toBe("confirmation_required"));
    await act(async () => { expect(await h.guard().confirmScope(original, { continuePendingAnalysis: true })).toBe(true); });
    await waitFor(() => expect(h.operation).toHaveBeenCalledOnce());
    expect(h.guard().notice).toBeNull();
    await act(async () => { await h.guard().confirmScope(original, { continuePendingAnalysis: true }); });
    expect(h.operation).toHaveBeenCalledOnce();
  });
  it("selector confirmation after a blocked write never replays that write", async () => {
    const h = mount(); await initialized(h); h.current(other);
    await act(async () => { expect(await h.guard().check()).toBe(false); });
    await act(async () => { expect(await h.guard().confirmScope(other, { continuePendingAnalysis: true })).toBe(true); });
    expect(h.operation).not.toHaveBeenCalled();
  });
  it("a first weak-evidence scope is a single confirmation, not an identity failure", async () => {
    const h = mount({ analysisEvidence: vi.fn(async () => ({ key: "private-key", conclusive: false })) });
    await initialized(h); fireEvent.click(h.ui.getByText("分析需求"));
    await waitFor(() => expect(h.guard().notice?.reason).toBe("confirmation_required"));
    expect(h.ui.queryAllByRole("radio")).toHaveLength(0);
    expect(h.operation).not.toHaveBeenCalled();
    fireEvent.click(h.ui.getByText("确认范围并继续"));
    await waitFor(() => expect(h.operation).toHaveBeenCalledOnce());
    fireEvent.click(h.ui.getByText("分析需求"));
    await waitFor(() => expect(h.operation).toHaveBeenCalledTimes(2));
    expect(h.ui.queryByRole("dialog")).toBeNull();
  });
  it("retries a transient metadata read without revoking committed consent or showing a dialog", async () => {
    const h = mount({ analysisEvidence: vi.fn(async () => ({ key: "private-key", conclusive: false })) });
    await initialized(h);
    await act(async () => { expect(await h.guard().confirmScope(original)).toBe(true); });
    h.inspect.mockRejectedValueOnce(new Error("PRIVATE_TRANSIENT"));
    fireEvent.click(h.ui.getByText("分析需求"));
    await waitFor(() => expect(h.operation).toHaveBeenCalledOnce());
    expect(h.ui.queryByRole("dialog")).toBeNull();
    expect(h.adopt).toHaveBeenCalledOnce();
  });
  it("retries missing analysis evidence without turning a confirmed scope into another consent request", async () => {
    const evidence = vi.fn(async (): Promise<{key: string; conclusive: boolean} | null> => ({ key: "private-key", conclusive: false }));
    const h = mount({ analysisEvidence: evidence }); await initialized(h);
    await act(async () => { expect(await h.guard().confirmScope(original)).toBe(true); });
    evidence.mockResolvedValueOnce(null);
    fireEvent.click(h.ui.getByText("分析需求"));
    await waitFor(() => expect(h.operation).toHaveBeenCalledOnce());
    expect(h.ui.queryByRole("dialog")).toBeNull();
  });
  it("persistent missing evidence stays blocked; recheck can resume only a fresh pending request", async () => {
    const evidence = vi.fn(async (): Promise<{key: string; conclusive: boolean} | null> => null);
    const h = mount({ analysisEvidence: evidence }); await initialized(h);
    fireEvent.click(h.ui.getByText("分析需求"));
    await waitFor(() => expect(h.guard().notice?.reason).toBe("unavailable"));
    expect(h.ui.queryAllByRole("radio")).toHaveLength(0);
    expect(h.operation).not.toHaveBeenCalled();
    const attempts = evidence.mock.calls.length;
    await act(async () => { await Promise.resolve(); });
    expect(evidence).toHaveBeenCalledTimes(attempts); // no automatic retry loop
    evidence.mockResolvedValue({ key: "private-key", conclusive: true });
    fireEvent.click(h.ui.getByText("重新检查"));
    await waitFor(() => expect(h.operation).toHaveBeenCalledOnce());
    expect(h.ui.queryByRole("dialog")).toBeNull();
    expect(h.adopt).not.toHaveBeenCalled();
  });
  it("unavailable identity cannot be overridden by calling confirmation directly", async () => {
    const h = mount({ analysisEvidence: vi.fn(async () => null) }); await initialized(h);
    fireEvent.click(h.ui.getByText("分析需求"));
    await waitFor(() => expect(h.guard().notice?.reason).toBe("unavailable"));
    await act(async () => { await h.guard().confirm("original", []); await h.guard().confirm("current", ["1", "2"]); });
    expect(h.adopt).not.toHaveBeenCalled();
    expect(h.operation).not.toHaveBeenCalled();
    expect(h.guard().binding?.selected).toEqual(["1"]);
  });
  it("recheck after a metadata pause never replays an old write without an analysis retry", async () => {
    const h = mount(); await initialized(h); h.current(null);
    await act(async () => { expect(await h.guard().check()).toBe(false); });
    h.current(original);
    fireEvent.click(h.ui.getByText("重新检查"));
    await waitFor(() => expect(h.ui.queryByRole("dialog")).toBeNull());
    expect(h.operation).not.toHaveBeenCalled();
    expect(h.adopt).not.toHaveBeenCalled();
  });
  it("a genuinely changed target discovered on recheck still needs explicit choice", async () => {
    const evidence = vi.fn(async (): Promise<{key: string; conclusive: boolean} | null> => null);
    const h = mount({ analysisEvidence: evidence }); await initialized(h);
    fireEvent.click(h.ui.getByText("分析需求"));
    await waitFor(() => expect(h.guard().notice?.reason).toBe("unavailable"));
    h.current(other); evidence.mockResolvedValue({ key: "other", conclusive: true });
    fireEvent.click(h.ui.getByText("重新检查"));
    await waitFor(() => expect(h.guard().notice?.reason).toBe("file_changed"));
    expect(h.ui.getAllByRole("radio")).toHaveLength(2);
    expect(h.operation).not.toHaveBeenCalled();
    expect(h.adopt).not.toHaveBeenCalled();
  });
  it("positive evidence conflict is blocked immediately, not retried until it looks stable", async () => {
    const evidence = vi.fn(async () => ({ key: "new", conclusive: true }));
    const h = mount({ analysisEvidence: evidence }); await initialized(h);
    evidence.mockResolvedValueOnce({ key: "old", conclusive: true });
    fireEvent.click(h.ui.getByText("分析需求"));
    await waitFor(() => expect(h.guard().notice?.reason).toBe("unavailable"));
    expect(evidence).toHaveBeenCalledTimes(3);
    expect(h.operation).not.toHaveBeenCalled();
    expect(h.guard().hasAnalysisConsent("new")).toBe(false);
  });
  it("partly missing evidence cannot hide a positive conflict between the other proofs", async () => {
    const evidence = vi.fn(async (): Promise<{key: string; conclusive: boolean} | null> => ({ key: "new", conclusive: true }));
    const h = mount({ analysisEvidence: evidence }); await initialized(h);
    evidence.mockResolvedValueOnce({ key: "old", conclusive: true })
      .mockResolvedValueOnce({ key: "new", conclusive: true }).mockResolvedValueOnce(null);
    fireEvent.click(h.ui.getByText("分析需求"));
    await waitFor(() => expect(h.guard().notice?.reason).toBe("unavailable"));
    expect(evidence).toHaveBeenCalledTimes(3);
    expect(h.operation).not.toHaveBeenCalled();
  });
  it("explicit sheet selection commits the same authorization used by subsequent analysis", async () => {
    const evidence = vi.fn(async (target: AuthorizedTarget) => ({
      key: JSON.stringify([target.fileKey, target.selected]), conclusive: false
    }));
    const h = mount({ analysisEvidence: evidence }); await initialized(h);
    const chosen = { ...original, selected: ["1", "2"] };
    await act(async () => { expect(await h.guard().confirmScope(chosen)).toBe(true); });
    expect(h.guard().binding?.selected).toEqual(["1", "2"]);
    expect(h.operation).not.toHaveBeenCalled();
    for (let turn = 0; turn < 2; turn++) {
      fireEvent.click(h.ui.getByText("分析需求"));
      await waitFor(() => expect(h.operation).toHaveBeenCalledTimes(turn + 1));
      expect(h.ui.queryByRole("dialog")).toBeNull();
    }
    h.current({ ...chosen, fileKey: "PRIVATE_PATH_B" });
    await act(async () => { expect(await h.guard().checkForAnalysis()).toBe(false); });
    expect(h.operation).toHaveBeenCalledTimes(2);
  });
  it("scope confirmation re-reads identity and never grants a missing or replaced sheet", async () => {
    const h = mount(); await initialized(h);
    h.current({ ...original, sheets: [original.sheets[0]] });
    await act(async () => { expect(await h.guard().confirmScope({ ...original, selected: ["2"] })).toBe(false); });
    expect(h.adopt).not.toHaveBeenCalled();
    h.current({ ...original, fileKey: "PRIVATE_PATH_B" });
    await act(async () => { expect(await h.guard().confirmScope(original)).toBe(false); });
    expect(h.adopt).not.toHaveBeenCalled();
  });
  it("scope confirmation rejects changing host evidence without changing the committed authorization", async () => {
    const evidence = vi.fn(async () => ({ key: "first", conclusive: false }));
    const h = mount({ analysisEvidence: evidence }); await initialized(h);
    evidence.mockResolvedValueOnce({ key: "first", conclusive: false })
      .mockResolvedValue({ key: "replacement", conclusive: false });
    await act(async () => { expect(await h.guard().confirmScope({ ...original, selected: ["1", "2"] })).toBe(false); });
    expect(h.guard().binding?.selected).toEqual(["1"]);
    expect(h.adopt).not.toHaveBeenCalled();
  });
  it.each([{ selected: [] }, { selected: ["1", "1"] }])("scope confirmation rejects empty or duplicate selections: $selected", async ({ selected }) => {
    const h = mount(); await initialized(h);
    await act(async () => { expect(await h.guard().confirmScope({ ...original, selected })).toBe(false); });
    expect(h.adopt).not.toHaveBeenCalled();
  });
  it("scope confirmation rejects a file change during async evidence checks", async () => {
    const evidence = vi.fn(async () => ({ key: "constant", conclusive: false }));
    const h = mount({ analysisEvidence: evidence }); await initialized(h);
    evidence.mockImplementationOnce(async () => { h.current(other); return { key: "constant", conclusive: false }; });
    await act(async () => { expect(await h.guard().confirmScope(original)).toBe(false); });
    expect(h.adopt).not.toHaveBeenCalled();
    expect(h.guard().binding?.fileKey).toBe(original.fileKey);
  });
  it("confirmed weak evidence is revoked on a detected change, even when the original target returns", async () => {
    const h = mount({ analysisEvidence: vi.fn(async (target: AuthorizedTarget) => ({
      key: JSON.stringify([target.fileKey, target.selected]), conclusive: false
    })) }); await initialized(h);
    await act(async () => { expect(await h.guard().confirmScope(original)).toBe(true); });
    h.current(other);
    await act(async () => { expect(await h.guard().check()).toBe(false); h.guard().cancel(); });
    h.current(original);
    await act(async () => { expect(await h.guard().checkForAnalysis()).toBe(false); });
    expect(h.ui.getByRole("dialog")).toBeTruthy();
    expect(h.operation).not.toHaveBeenCalled();
  });
  it("the target dialog can confirm a newly selected scope while inspection still returns the old selection", async () => {
    const h = mount({ analysisEvidence: vi.fn(async (target: AuthorizedTarget) => ({
      key: JSON.stringify([target.fileKey, target.selected]), conclusive: false
    })) }); await initialized(h); h.current(other);
    fireEvent.click(h.ui.getByText("分析需求")); await waitFor(() => h.ui.getByRole("dialog"));
    fireEvent.click(h.ui.getByLabelText(/改用当前文件/));
    fireEvent.click(h.ui.getByLabelText("Other"));
    fireEvent.click(h.ui.getByText("确认目标并继续"));
    await waitFor(() => expect(h.operation).toHaveBeenCalledTimes(1));
    expect(h.guard().binding?.selected).toEqual(["2"]);
    expect(h.ui.queryByRole("dialog")).toBeNull();
  });
  it("scope confirmation cannot replay a queued request and cannot change targets during review", async () => {
    const h = mount(); await initialized(h); h.current(other);
    const queued = vi.fn();
    await act(async () => { await h.guard().check(queued); });
    await act(async () => { expect(await h.guard().confirmScope(other)).toBe(true); });
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(queued).not.toHaveBeenCalled();
    expect(h.ui.queryByRole("dialog")).toBeNull();
    cleanup();
    const review = mount({ reviewActive: true });
    await act(async () => { expect(await review.guard().confirmScope(original)).toBe(false); });
    expect(review.adopt).not.toHaveBeenCalled();
  });
  it("unknown identity explains verification failure without claiming a different same-named file", async () => {
    const h = mount({ analysisEvidence: vi.fn(async () => null) }); await initialized(h);
    fireEvent.click(h.ui.getByText("分析需求")); await waitFor(() => h.ui.getByRole("dialog"));
    expect(h.ui.getByRole("dialog").textContent).toContain("暂时无法核实");
    expect(h.ui.getByRole("dialog").textContent).not.toContain("名称相同不代表");
    expect(h.operation).not.toHaveBeenCalled();
  });
  it("弱宿主证据先确认目标，再分析；同任务补充不重复确认，新的授权边界仍拒绝", async () => {
    const evidence = vi.fn(async (target: AuthorizedTarget) => ({ key: `private:${target.fileKey}`, conclusive: false }));
    const h = mount({ analysisEvidence: evidence }); await initialized(h);
    fireEvent.click(h.ui.getByText("分析需求"));
    await waitFor(() => expect(h.ui.getByRole("dialog")).toBeTruthy());
    expect(h.operation).not.toHaveBeenCalled();
    expect(h.adopt).not.toHaveBeenCalled();
    expect(h.ui.queryByText(/需求分析过程/)).toBeNull();
    expect(h.ui.queryAllByRole("radio")).toHaveLength(0);
    fireEvent.click(h.ui.getByText("确认范围并继续"));
    await waitFor(() => expect(h.operation).toHaveBeenCalledTimes(1));
    await act(async () => { expect(await h.guard().checkForAnalysis()).toBe(true); });
    expect(h.ui.queryByRole("dialog")).toBeNull();
    expect(h.ui.container.textContent).not.toContain("private:");
    h.current(other);
    await act(async () => { expect(await h.guard().checkForAnalysis()).toBe(false); });
    expect(h.operation).toHaveBeenCalledTimes(1);
  });
  it("显式撤销同意后重新确认；未保存且证据缺失不分析", async () => {
    const evidence = vi.fn(async (): Promise<{ key: string; conclusive: boolean } | null> => ({ key: "private-key", conclusive: false }));
    const h = mount({ analysisEvidence: evidence }); await initialized(h);
    fireEvent.click(h.ui.getByText("分析需求")); await waitFor(() => h.ui.getByRole("dialog"));
    fireEvent.click(h.ui.getByText("确认范围并继续"));
    await waitFor(() => expect(h.operation).toHaveBeenCalledTimes(1));
    await act(async () => { h.guard().clearAnalysisConsent(); expect(await h.guard().checkForAnalysis()).toBe(false); });
    expect(h.ui.getByRole("dialog")).toBeTruthy();
    evidence.mockResolvedValue(null);
    fireEvent.click(h.ui.getByText("确认范围并继续"));
    await waitFor(() => h.ui.getByRole("alert"));
    expect(h.operation).toHaveBeenCalledTimes(1);
  });
  it("强身份证据在分析检查中变化，即使元数据未变也先停下", async () => {
    const evidence = vi.fn(async () => ({ key: "first", conclusive: true }));
    const h = mount({ analysisEvidence: evidence }); await initialized(h);
    evidence.mockResolvedValueOnce({ key: "first", conclusive: true }).mockResolvedValue({ key: "replacement", conclusive: true });
    fireEvent.click(h.ui.getByText("分析需求")); await waitFor(() => h.ui.getByRole("dialog"));
    expect(h.operation).not.toHaveBeenCalled();
  });
  it("强身份与有效授权一致时直接分析；不会因为检查而多问一次", async () => {
    const h = mount({ analysisEvidence: vi.fn(async () => ({ key: "identity", conclusive: true })) });
    await initialized(h); fireEvent.click(h.ui.getByText("分析需求"));
    await waitFor(() => expect(h.operation).toHaveBeenCalledTimes(1));
    expect(h.ui.queryByRole("dialog")).toBeNull();
    expect(h.adopt).not.toHaveBeenCalled();
  });
  it("用户取消前置确认后不分析，保留输入；重新提交仍需确认", async () => {
    const h = mount({ analysisEvidence: vi.fn(async () => ({ key: "identity", conclusive: false })) });
    await initialized(h); fireEvent.click(h.ui.getByText("分析需求"));
    await waitFor(() => h.ui.getByRole("dialog"));
    fireEvent.click(h.ui.getByText("暂不操作"));
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(h.operation).not.toHaveBeenCalled();
    expect((h.ui.getByRole("textbox") as HTMLTextAreaElement).value).toContain("original request");
    fireEvent.click(h.ui.getByText("分析需求"));
    await waitFor(() => h.ui.getByRole("dialog"));
    expect(h.operation).not.toHaveBeenCalled();
  });
  it("确认期间强宿主证据变化，不能只凭相同名称与元数据开始分析", async () => {
    const evidence = vi.fn(async () => ({ key: "first", conclusive: false }));
    const h = mount({ analysisEvidence: evidence }); await initialized(h);
    fireEvent.click(h.ui.getByText("分析需求")); await waitFor(() => h.ui.getByRole("dialog"));
    evidence.mockResolvedValueOnce({ key: "first", conclusive: false }).mockResolvedValue({ key: "replacement", conclusive: false });
    fireEvent.click(h.ui.getByText("确认范围并继续"));
    await waitFor(() => h.ui.getByRole("alert"));
    expect(h.operation).not.toHaveBeenCalled(); expect(h.adopt).not.toHaveBeenCalled();
  });
  it("stronger host guard can request consent before operation even when metadata matches", async () => {
    const h = mount(); await initialized(h);
    await act(async () => { expect(await h.guard().requestConfirmation()).toBe(false); });
    expect(h.ui.getByRole("dialog")).toBeTruthy();
    expect(h.adopt).not.toHaveBeenCalled();
    expect(h.operation).not.toHaveBeenCalled();
    expect(h.ui.queryAllByRole("radio")).toHaveLength(0);
    fireEvent.click(h.ui.getByText("确认范围并继续"));
    await waitFor(() => expect(h.adopt).toHaveBeenCalledTimes(1));
    expect(h.operation).not.toHaveBeenCalled(); // Never replay an old mapping or write.
  });
  it("periodic metadata monitoring detects a change without waiting for submit, and preserves dialog choices", async () => {
    vi.useFakeTimers();
    try {
      const h = mount();
      await act(async () => { await Promise.resolve(); await Promise.resolve(); });
      h.current(other);
      await act(async () => { await vi.advanceTimersByTimeAsync(2000); });
      expect(h.ui.getByRole("dialog")).toBeTruthy();
      fireEvent.click(h.ui.getByLabelText("改用当前文件：Other.xlsx"));
      fireEvent.click(h.ui.getByLabelText("Sheet1"));
      const count = h.inspect.mock.calls.length;
      await act(async () => { await vi.advanceTimersByTimeAsync(4000); });
      expect(h.inspect).toHaveBeenCalledTimes(count);
      expect((h.ui.getByLabelText("Sheet1") as HTMLInputElement).checked).toBe(true);
      expect(h.operation).not.toHaveBeenCalled();
    } finally { cleanup(); vi.useRealTimers(); }
  });
  it("focus detects a change before any operation and has no preselected target/sheets", async () => {
    const h = mount(); await initialized(h); h.current(other);
    fireEvent(window, new Event("focus"));
    await waitFor(() => expect(h.ui.getByRole("dialog")).toBeTruthy());
    expect(h.operation).not.toHaveBeenCalled();
    expect(h.adopt).not.toHaveBeenCalled();
    expect(h.ui.getAllByRole("radio").every((input) => !(input as HTMLInputElement).checked)).toBe(true);
    expect((h.ui.getByText("确认目标并继续") as HTMLButtonElement).disabled).toBe(true);
    expect((h.ui.getByRole("textbox") as HTMLTextAreaElement).value).toContain("confirmed answer");
    expect(h.ui.getByRole("dialog").textContent).not.toContain("PRIVATE_PATH");
  });
  it("submit checks even without a focus event; selection only does not operate", async () => {
    const h = mount(); await initialized(h); h.current(other);
    fireEvent.click(h.ui.getByText("操作"));
    await waitFor(() => expect(h.ui.getByRole("dialog")).toBeTruthy());
    fireEvent.click(h.ui.getByLabelText("改用当前文件：Other.xlsx"));
    expect(h.ui.getAllByRole("checkbox").every((input) => !(input as HTMLInputElement).checked)).toBe(true);
    fireEvent.click(h.ui.getByLabelText("Sheet1"));
    expect(h.operation).not.toHaveBeenCalled();
    fireEvent.click(h.ui.getByText("确认目标并继续"));
    await waitFor(() => expect(h.operation).toHaveBeenCalledOnce());
    expect(h.adopt.mock.calls[0][0].selected).toEqual(["1"]);
    expect(h.ui.queryByRole("dialog")).toBeNull();
  });
  it("confirmation re-reads identity and rejects a third file appearing while dialog is open", async () => {
    const h = mount(); await initialized(h); h.current(other);
    fireEvent.click(h.ui.getByText("操作")); await waitFor(() => h.ui.getByRole("dialog"));
    fireEvent.click(h.ui.getByLabelText("改用当前文件：Other.xlsx")); fireEvent.click(h.ui.getByLabelText("Sheet1"));
    h.current({ ...other, fileKey: "PRIVATE_PATH_C" });
    fireEvent.click(h.ui.getByText("确认目标并继续"));
    await waitFor(() => expect(h.ui.getByRole("alert").textContent).toContain("又发生变化"));
    expect(h.adopt).not.toHaveBeenCalled(); expect(h.operation).not.toHaveBeenCalled();
  });
  it("confirmation cannot reuse metadata inspection started before the click", async () => {
    const h = mount(); await initialized(h); h.current(other);
    await act(async () => { await h.guard().check(); });
    fireEvent.click(h.ui.getByLabelText("改用当前文件：Other.xlsx")); fireEvent.click(h.ui.getByLabelText("Sheet1"));
    let release!: (value: AuthorizedTarget) => void;
    h.inspect.mockImplementationOnce(() => new Promise<AuthorizedTarget>((resolve) => { release = resolve; }));
    let pending!: Promise<boolean>;
    act(() => { pending = h.guard().check(); });
    h.current({ ...other, fileKey: "PRIVATE_PATH_C" });
    fireEvent.click(h.ui.getByText("确认目标并继续"));
    await act(async () => { release(other); await pending; });
    await waitFor(() => expect(h.ui.getByRole("alert").textContent).toContain("又发生变化"));
    expect(h.adopt).not.toHaveBeenCalled(); expect(h.operation).not.toHaveBeenCalled();
  });
  it("confirmation failure keeps the chosen target and sheet selection so the user can retry once", async () => {
    const h = mount(); await initialized(h); h.current(other);
    fireEvent.click(h.ui.getByText("操作")); await waitFor(() => h.ui.getByRole("dialog"));
    fireEvent.click(h.ui.getByLabelText("改用当前文件：Other.xlsx")); fireEvent.click(h.ui.getByLabelText("Sheet1"));
    expect((h.ui.getByLabelText("Sheet1") as HTMLInputElement).checked).toBe(true);
    // Identity changes between the click and the confirmation re-read.
    h.current({ ...other, fileKey: "PRIVATE_PATH_C" });
    fireEvent.click(h.ui.getByText("确认目标并继续"));
    await waitFor(() => expect(h.ui.getByRole("alert").textContent).toContain("又发生变化"));
    // The retry is the same round of confirmation, so the selection must survive it.
    expect(h.adopt).not.toHaveBeenCalled(); expect(h.operation).not.toHaveBeenCalled();
    expect((h.ui.getByLabelText("Sheet1") as HTMLInputElement).checked).toBe(true);
    expect((h.ui.getByLabelText("改用当前文件：Other.xlsx") as HTMLInputElement).checked).toBe(true);
    expect((h.ui.getByText("确认目标并继续") as HTMLButtonElement).disabled).toBe(false);
  });
  it("a settled dialog still resets a stale selection when a genuinely new target appears", async () => {
    const h = mount(); await initialized(h);
    // Open on a different target, then move to a third one while the dialog stays open.
    h.current(other);
    fireEvent.click(h.ui.getByText("操作"));
    await waitFor(() => expect(h.ui.getByLabelText("改用当前文件：Other.xlsx")).toBeTruthy());
    fireEvent.click(h.ui.getByLabelText("改用当前文件：Other.xlsx")); fireEvent.click(h.ui.getByLabelText("Sheet1"));
    expect((h.ui.getByLabelText("Sheet1") as HTMLInputElement).checked).toBe(true);
    h.current({ ...other, fileKey: "PRIVATE_PATH_C", name: "Third.xlsx" });
    await act(async () => { expect(await h.guard().check()).toBe(false); });
    await waitFor(() => expect(h.ui.getByText(/改用当前文件：Third\.xlsx/)).toBeTruthy());
    // New round: the choice made against Other.xlsx must not carry over to Third.xlsx.
    expect((h.ui.getByText("确认目标并继续") as HTMLButtonElement).disabled).toBe(true);
    expect((h.ui.getByLabelText(/改用当前文件：Third\.xlsx/) as HTMLInputElement).checked).toBe(false);
  });
  it("a retry on the same target keeps every still-valid sheet checked", async () => {
    const h = mount(); await initialized(h); h.current(other);
    fireEvent.click(h.ui.getByText("操作")); await waitFor(() => h.ui.getByRole("dialog"));
    fireEvent.click(h.ui.getByLabelText("改用当前文件：Other.xlsx"));
    fireEvent.click(h.ui.getByLabelText("Sheet1")); fireEvent.click(h.ui.getByLabelText("Other"));
    // Re-read reports a different identity, but the same target still offers both sheets.
    h.current({ ...other, fileKey: "PRIVATE_PATH_C" });
    fireEvent.click(h.ui.getByText("确认目标并继续"));
    await waitFor(() => expect(h.ui.getByRole("alert").textContent).toContain("又发生变化"));
    expect((h.ui.getByLabelText("Sheet1") as HTMLInputElement).checked).toBe(true);
    expect((h.ui.getByLabelText("Other") as HTMLInputElement).checked).toBe(true);
  });
  it("a retry drops sheet ids the re-read target no longer offers", async () => {
    const h = mount(); await initialized(h); h.current(other);
    fireEvent.click(h.ui.getByText("操作")); await waitFor(() => h.ui.getByRole("dialog"));
    fireEvent.click(h.ui.getByLabelText("改用当前文件：Other.xlsx"));
    fireEvent.click(h.ui.getByLabelText("Sheet1")); fireEvent.click(h.ui.getByLabelText("Other"));
    // Re-read target lost the second sheet entirely.
    h.current({ ...other, fileKey: "PRIVATE_PATH_C", sheets: [other.sheets[0]] });
    fireEvent.click(h.ui.getByText("确认目标并继续"));
    await waitFor(() => expect(h.ui.getByRole("alert").textContent).toContain("又发生变化"));
    expect((h.ui.getByLabelText("Sheet1") as HTMLInputElement).checked).toBe(true);
    expect(h.ui.queryByLabelText("Other")).toBeNull();
    expect((h.ui.getByText("确认目标并继续") as HTMLButtonElement).disabled).toBe(false);
  });
  it("original choice requires returning to original, not silently activating another file", async () => {
    const h = mount(); await initialized(h); h.current(other);
    fireEvent.click(h.ui.getByText("操作")); await waitFor(() => h.ui.getByRole("dialog"));
    fireEvent.click(h.ui.getByLabelText(/继续使用原授权/));
    fireEvent.click(h.ui.getByText("确认目标并继续"));
    await waitFor(() => h.ui.getByRole("alert")); expect(h.operation).not.toHaveBeenCalled();
    h.current(original); fireEvent.click(h.ui.getByLabelText(/继续使用原授权/));
    fireEvent.click(h.ui.getByText("确认目标并继续"));
    await waitFor(() => expect(h.operation).toHaveBeenCalledOnce());
  });
  it("missing identity pauses without false change claims or executing", async () => {
    const h = mount(); await initialized(h); h.current(null);
    fireEvent.click(h.ui.getByText("操作")); await waitFor(() => h.ui.getByRole("dialog"));
    expect(h.ui.getByRole("dialog").textContent).toContain("暂时无法核实操作目标");
    expect(h.ui.getByRole("dialog").textContent).not.toContain("PRIVATE_FAILURE");
    expect(h.operation).not.toHaveBeenCalled();
  });
  it("cancel removes the queued operation, not the input", async () => {
    const h = mount(); await initialized(h); h.current(other);
    fireEvent.click(h.ui.getByText("操作")); await waitFor(() => h.ui.getByRole("dialog"));
    fireEvent.click(h.ui.getByText("暂不操作"));
    expect(h.ui.queryByRole("dialog")).toBeNull();
    expect((h.ui.getByRole("textbox") as HTMLTextAreaElement).value).toContain("confirmed answer");
    expect(h.operation).not.toHaveBeenCalled();
  });
  it("a known review copy is exempt; arbitrary other files are not", async () => {
    const h = mount({ reviewActive: true, isReviewCopy: (target) => target.fileKey === "owned-copy" });
    // First explicit action creates the source baseline, as review monitoring is suspended.
    await act(async () => { expect(await h.guard().check()).toBe(true); });
    h.current({ ...other, fileKey: "owned-copy" });
    await act(async () => { expect(await h.guard().check()).toBe(true); });
    h.current(other);
    await act(async () => { expect(await h.guard().check()).toBe(false); });
    expect(h.ui.getByRole("dialog")).toBeTruthy();
  });
  it("an old write blocked by check is never queued/replayed by target confirmation", async () => {
    const h = mount(); await initialized(h); h.current(other);
    await act(async () => { expect(await h.guard().check()).toBe(false); });
    fireEvent.click(h.ui.getByLabelText("改用当前文件：Other.xlsx")); fireEvent.click(h.ui.getByLabelText("Sheet1"));
    fireEvent.click(h.ui.getByText("确认目标并继续"));
    await waitFor(() => expect(h.adopt).toHaveBeenCalledOnce());
    expect(h.operation).not.toHaveBeenCalled();
  });
});
