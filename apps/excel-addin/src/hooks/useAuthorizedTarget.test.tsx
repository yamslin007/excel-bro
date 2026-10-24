// @vitest-environment jsdom
import { act, cleanup, fireEvent, render, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { AuthorizedTarget } from "@excel-bro/core/authorizedTarget";
import { useAuthorizedTarget } from "./useAuthorizedTarget";
import { AuthorizedTargetConfirmation } from "../components/AuthorizedTargetConfirmation";
import { inspectOfficeAuthorizedTarget, officeAnalysisTargetEvidence } from "../officeAuthorizedTarget";

const original: AuthorizedTarget = { fileKey: "PRIVATE_PATH_A", name: "Original.xlsx",
  sheets: [{ id: "1", name: "Sheet1" }, { id: "2", name: "Other" }], selected: ["1"] };
const other: AuthorizedTarget = { ...original, fileKey: "PRIVATE_PATH_B", name: "Other.xlsx" };
afterEach(cleanup);
function mount(options: Partial<Pick<Parameters<typeof useAuthorizedTarget>[0], "reviewActive" | "isReviewCopy" | "analysisEvidence" | "selectionPending">> = {}) {
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
  it("a newer checkbox gesture invalidates verification before any authorization is published", async () => {
    const h = mount(); await initialized(h);
    let resolve!: (target: AuthorizedTarget) => void;
    h.inspect.mockImplementationOnce(() => new Promise((done) => { resolve = done; }));
    let current = true;
    let result!: Promise<boolean>;
    act(() => { result = h.guard().confirmScope({ ...original, selected: ["1", "2"] }, { isCurrent: () => current }); });
    current = false;
    await act(async () => { resolve(original); expect(await result).toBe(false); });
    expect(h.adopt).not.toHaveBeenCalled();
    expect(h.guard().binding?.selected).toEqual(["1"]);
  });
  it("clearing all checkboxes revokes consent without reading data or replaying work", async () => {
    const evidence = vi.fn(async () => ({ key: "same-scope", conclusive: false }));
    const h = mount({ analysisEvidence: evidence }); await initialized(h);
    await act(async () => { await h.guard().confirmScope(original); });
    expect(h.guard().hasAnalysisConsent("same-scope")).toBe(true);
    evidence.mockClear();
    await act(async () => { expect(await h.guard().confirmScope({ ...original, selected: [] }, { allowEmpty: true })).toBe(true); });
    expect(h.guard().binding?.selected).toEqual([]);
    expect(h.guard().hasAnalysisConsent("same-scope")).toBe(false);
    expect(evidence).not.toHaveBeenCalled();
    expect(h.operation).not.toHaveBeenCalled();
  });
  it("an in-flight monitor cannot report an old scope while new checkboxes are being verified", async () => {
    let pending = false;
    const h = mount({ selectionPending: () => pending }); await initialized(h);
    let resolve!: (target: AuthorizedTarget) => void;
    h.inspect.mockImplementationOnce(() => new Promise((done) => { resolve = done; }));
    let result!: Promise<boolean>;
    act(() => { result = h.guard().check(); });
    pending = true;
    await act(async () => { resolve({ ...original, selected: ["1", "2"] }); expect(await result).toBe(false); });
    expect(h.guard().notice).toBeNull();
    expect(h.adopt).not.toHaveBeenCalled();
  });
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
  it.each(["metadata", "analysis-evidence"])("keeps explicit scope consent across an exhausted %s read and rechecks without reauthorization", async (failure) => {
    const proof = { key: "private-key", conclusive: false };
    const evidence = vi.fn(async (): Promise<typeof proof | null> => proof);
    const h = mount({ analysisEvidence: evidence }); await initialized(h);
    await act(async () => { expect(await h.guard().confirmScope(original)).toBe(true); });
    if (failure === "metadata") h.current(null);
    else evidence.mockResolvedValue(null);
    fireEvent.click(h.ui.getByText("分析需求"));
    await waitFor(() => expect(h.guard().notice?.reason).toBe("unavailable"));
    expect(h.guard().hasAnalysisConsent(proof.key)).toBe(true);
    expect(h.operation).not.toHaveBeenCalled();
    h.current(original); evidence.mockResolvedValue(proof);
    fireEvent.click(h.ui.getByText("重新检查"));
    await waitFor(() => expect(h.operation).toHaveBeenCalledOnce());
    expect(h.guard().notice).toBeNull();
    expect(h.adopt).toHaveBeenCalledOnce(); // the original user selection only
  });
  it.each(["file", "selection", "conflicting-evidence"])("revokes consent on positive %s change even if the old target returns", async (change) => {
    const proof = { key: "private-key", conclusive: false };
    const evidence = vi.fn(async (): Promise<typeof proof | null> => proof);
    const h = mount({ analysisEvidence: evidence }); await initialized(h);
    await act(async () => { expect(await h.guard().confirmScope(original)).toBe(true); });
    if (change === "file") h.current(other);
    else if (change === "selection") h.current({ ...original, selected: ["1", "2"] });
    else evidence.mockResolvedValueOnce(proof).mockResolvedValueOnce({ ...proof, key: "conflict" });
    await act(async () => { expect(await h.guard().checkForAnalysis()).toBe(false); });
    expect(h.guard().hasAnalysisConsent(proof.key)).toBe(false);
    h.current(original);
    await act(async () => { expect(await h.guard().checkForAnalysis()).toBe(false); });
    expect(h.guard().notice?.reason).toBe("confirmation_required");
    expect(h.operation).not.toHaveBeenCalled();
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
  it("rechecking an unavailable target cannot promote a queued write callback to an analysis retry", async () => {
    const h = mount(); await initialized(h); h.current(null);
    fireEvent.click(h.ui.getByText("操作"));
    await waitFor(() => expect(h.guard().notice?.reason).toBe("unavailable"));
    h.current(original);
    fireEvent.click(h.ui.getByText("重新检查"));
    await waitFor(() => expect(h.ui.queryByRole("dialog")).toBeNull());
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 0)); });
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

/**
 * 原地撤销造成的**目录增加**能不能认下来。
 *
 * 撤销会把表放回工作簿，目录因此比基线多几张。认下来是为了不误报"目标已变化"，
 * 但认下来这件事本身要证明：只会认"这次撤销放回的"那几张，授权集合一字不改。
 * 任何一条被放宽都等于借撤销的名义关掉目标守卫。
 */
describe("worksheet restore target rebinding", () => {
  const sheet = (id: string, name: string) => ({ id, name });
  const target = (
    selected: string[],
    sheets: { id: string; name: string }[],
    fileKey = "PRIVATE_PATH_A"
  ): AuthorizedTarget => ({ fileKey, name: "Original.xlsx", sheets, selected });

  it("accepts only the directory growth this restore explains and never widens the authorized set", async () => {
    const h = mount();
    await initialized(h);
    const sheetsAfterDelete = [sheet("1", "Sheet1"), sheet("3", "汇总")];
    // 删表后的现场：Other 已被删除，授权只剩「Sheet1」。
    await act(async () => { h.guard().bind(target(["1"], sheetsAfterDelete)); });

    // 回执说放回了 Other，目录里却没有 ⇒ 这份回执不能当证据。
    expect(h.guard().bindVerifiedRestore(target(["1"], sheetsAfterDelete), ["Other"])).toBe(false);
    // 目录多了一张谁都没授权的表 ⇒ 不是这次撤销造成的。
    expect(h.guard().bindVerifiedRestore(
      target(["1"], [...sheetsAfterDelete, sheet("4", "外部新增")]), ["Other"])).toBe(false);
    // 回执说放回的表确实在目录里，但同时还多了一张没人认领的表 ⇒ 同样拒绝。
    // （这一条只被"多出来的表是否来自本次撤销"那条判据挡住，别的判据管不到。）
    expect(h.guard().bindVerifiedRestore(
      target(["1"], [...sheetsAfterDelete, sheet("2", "Other"), sheet("4", "外部新增")]), ["Other"])).toBe(false);
    // 借撤销之名把放回来的表纳入授权 ⇒ 拒绝。
    expect(h.guard().bindVerifiedRestore(
      target(["1", "2"], [...sheetsAfterDelete, sheet("2", "Other")]), ["Other"])).toBe(false);
    // 已授权的表凭空消失 ⇒ 拒绝（撤销只加表）。
    expect(h.guard().bindVerifiedRestore(target([], [sheet("2", "Other")]), ["Other"])).toBe(false);
    // 换了工作簿 ⇒ 拒绝。
    expect(h.guard().bindVerifiedRestore(
      target(["1"], [...sheetsAfterDelete, sheet("2", "Other")], "PRIVATE_PATH_B"), ["Other"])).toBe(false);
    // 没有"放回了什么"的回执 ⇒ 拒绝。
    expect(h.guard().bindVerifiedRestore(
      target(["1"], [...sheetsAfterDelete, sheet("2", "Other")]), [])).toBe(false);

    // 正面：Other 回到目录里，授权仍是「Sheet1」。
    expect(h.guard().bindVerifiedRestore(
      target(["1"], [sheet("1", "Sheet1"), sheet("2", "Other"), sheet("3", "汇总")]), ["Other"])).toBe(true);
    expect(h.guard().binding?.selected).toEqual(["1"]);
  });
});

/*
 * 基线被推进时，**还在飞**的那一份读数不能拿去和新基线比。
 *
 * 现场：`bindVerifiedShrink` / `bindVerifiedRestore`（删表收尾、原地撤销）以及
 * 从文件入口采用当前工作簿，都是**不经过** `read()` 的直接 bind —— 它们不会等
 * 监视器那一轮读数回来。于是"读数在旧状态开始、在新基线之后返回"是常态：
 * 拿旧读数和新基线比，得到的 `scope_changed` 是假的，用户刚确认完就被要求再确认。
 *
 * 判据分两半：旧读数作废后必须**重读**（一致则放行）；重读仍是真变化就停下，
 * 绝不放行旧请求。
 */
describe("读数在飞期间基线被推进", () => {
  it("旧读数不得判成范围变化：作废重读后一致就放行", async () => {
    const expanded: AuthorizedTarget = { ...original, selected: ["1", "2"] };
    let current: AuthorizedTarget = original;
    /** 每一次读都在**发起那一刻**取快照（真实宿主就是这么返回的）。 */
    const reads: Array<{ snapshot: AuthorizedTarget; release: () => void }> = [];
    const inspect = vi.fn((() => new Promise<AuthorizedTarget>((resolve) => {
      const snapshot = current;
      reads.push({ snapshot, release: () => resolve(snapshot) });
    })) as never);
    const adopt = vi.fn();
    let guard!: ReturnType<typeof useAuthorizedTarget>;
    function Harness() {
      guard = useAuthorizedTarget({ enabled: true, inspect, adopt });
      return <span>{guard.binding ? "bound" : "idle"}</span>;
    }
    render(<Harness />);
    // 首绑：沿用旧的"第一次读成基线"语义。
    await waitFor(() => expect(reads).toHaveLength(1));
    await act(async () => { reads.shift()!.release(); await Promise.resolve(); });
    await waitFor(() => expect(guard.binding?.selected).toEqual(["1"]));

    // 一轮检查开始（读数在旧状态发起，被挂住）。
    let checked: boolean | null = null;
    let operation = 0;
    act(() => { void guard.check().then((ready) => { checked = ready; if (ready) operation += 1; }); });
    await waitFor(() => expect(reads).toHaveLength(1));

    // 期间基线被推进（= 删表收尾 / 原地撤销那条直接 bind 的路）。
    current = expanded;
    act(() => { guard.bind(expanded); });
    expect(guard.binding?.selected).toEqual(["1", "2"]);

    // 旧读数这时才回来 —— 它描述的是推进**之前**的范围。
    await act(async () => { reads.shift()!.release(); await Promise.resolve(); });
    // 实现必须作废它并重读，而不是拿它和新基线比。
    await waitFor(() => expect(reads.length, "旧读数作废后必须重读一次").toBeGreaterThan(0));
    await act(async () => { reads.shift()!.release(); await Promise.resolve(); });

    await waitFor(() => expect(checked, `检查没有给出结论（notice=${guard.notice?.reason ?? "none"}）`).toBe(true));
    expect(guard.notice).toBeNull();
    expect(operation).toBe(1);
  });

  it("重读仍与基线不一致：停下并要求确认，旧请求不重放", async () => {
    const expanded: AuthorizedTarget = { ...original, selected: ["1", "2"] };
    const mysteriouslyNarrowed: AuthorizedTarget = { ...original, selected: ["1"], sheets: [{ id: "1", name: "Sheet1" }] };
    let current: AuthorizedTarget = original;
    const reads: Array<{ snapshot: AuthorizedTarget; release: () => void }> = [];
    const inspect = vi.fn((() => new Promise<AuthorizedTarget>((resolve) => {
      const snapshot = current;
      reads.push({ snapshot, release: () => resolve(snapshot) });
    })) as never);
    let guard!: ReturnType<typeof useAuthorizedTarget>;
    function Harness() {
      guard = useAuthorizedTarget({ enabled: true, inspect, adopt: vi.fn() });
      return <span>{guard.binding ? "bound" : "idle"}</span>;
    }
    render(<Harness />);
    await waitFor(() => expect(reads).toHaveLength(1));
    await act(async () => { reads.shift()!.release(); await Promise.resolve(); });
    await waitFor(() => expect(guard.binding?.selected).toEqual(["1"]));

    let operation = 0;
    let checked: boolean | null = null;
    act(() => { void guard.check(() => { operation += 1; }).then((ready) => { checked = ready; }); });
    await waitFor(() => expect(reads).toHaveLength(1));
    current = expanded;
    act(() => { guard.bind(expanded); });
    expect(guard.binding?.selected).toEqual(["1", "2"]);

    // 重读开始**之前**目录就已经是一份确实不同的清单：范围真的变了。
    const stale = reads.shift()!;
    current = mysteriouslyNarrowed;
    await act(async () => { stale.release(); await Promise.resolve(); });
    await waitFor(() => expect(reads.length).toBeGreaterThan(0));
    await act(async () => { reads.shift()!.release(); await Promise.resolve(); });

    await waitFor(() => expect(checked).toBe(false));
    expect(guard.notice?.reason).toBe("scope_changed");
    expect(operation).toBe(0);
  });
});

/* ────────────────────────────────────────────────────────────────────────────
 * 完整调用链：adopt → bind → checkForAnalysis
 *
 * 现场：按「Sheet11 → Sheet1 → Sheet10」（不是目录顺序）确认三张表后发送需求，
 * 被拦在「暂时无法核实当前目标」。链路上的三个事实本来就不该互相打架：
 *   · 用户**点选**的顺序 → 守卫基线 `bound.selected`；
 *   · `adopt` 写进已提交的**名字**（`App.tsx` 按工作表目录顺序取出来）；
 *   · 下一次 `inspect` 按那些名字映射回来的 `selected`（于是变成目录顺序）。
 * 三者描述的是同一组表，只有顺序不同。判"是不是同一个目标"必须只看集合。
 *
 * 这条用例刻意用**真实的** `inspectOfficeAuthorizedTarget` 与
 * `officeAnalysisTargetEvidence`（只替身宿主对象），否则"证据键"这条判据
 * 会被测试自己写的那份替身悄悄替掉 —— 那正是仓库里反复踩过的坑。
 * ──────────────────────────────────────────────────────────────────────────── */
function stubThreeSheetWorkbook() {
  const document = { url: "file:///C:/Data/Book.xlsx" };
  const sheets = [
    { id: "1", name: "Sheet1", load: vi.fn() },
    { id: "10", name: "Sheet10", load: vi.fn() },
    { id: "11", name: "Sheet11", load: vi.fn() }
  ];
  const worksheets = { items: sheets, load: vi.fn(), getActiveWorksheet: () => sheets[2] };
  vi.stubGlobal("Office", { context: { document } });
  vi.stubGlobal("Excel", { run: async (job: (context: unknown) => unknown) =>
    job({ workbook: { worksheets }, sync: vi.fn(async () => {}) }) });
  return { document, sheets };
}
describe("已确认范围的证据冲突判据（真实适配器 + 真实守卫）", () => {
  afterEach(() => { vi.unstubAllGlobals(); });
  it("已提交名单按目录顺序写回、基线按点选顺序绑定：发送仍不判证据冲突", async () => {
    stubThreeSheetWorkbook();
    /** `adopt` 之后写进已提交的名字。初始就是目录顺序（App 第一次读到的样子）。 */
    let committed: string[] = ["Sheet1", "Sheet10", "Sheet11"];
    let guard!: ReturnType<typeof useAuthorizedTarget>;
    function Harness() {
      guard = useAuthorizedTarget({
        enabled: true,
        inspect: () => inspectOfficeAuthorizedTarget(committed),
        // 与 `App.tsx` 的 adopt 同形：按**工作表目录**顺序取已授权名字。
        adopt: (target) => {
          committed = target.sheets
            .filter((sheet) => target.selected.includes(sheet.id))
            .map((sheet) => sheet.name);
        },
        analysisEvidence: officeAnalysisTargetEvidence
      });
      return null;
    }
    render(<Harness />);
    await waitFor(() => expect(guard.binding).not.toBeNull());

    // 用户点选的顺序：Sheet11 → Sheet1 → Sheet10。
    const clicked: AuthorizedTarget = { ...guard.binding!, selected: ["11", "1", "10"] };
    await act(async () => { expect(await guard.confirmScope(clicked)).toBe(true); });
    // adopt 之后已提交名单变成目录顺序 —— 这就是现场那两个顺序的差别来源。
    expect(committed).toEqual(["Sheet1", "Sheet10", "Sheet11"]);
    expect(guard.binding?.selected).toEqual(["11", "1", "10"]);

    await act(async () => { expect(await guard.checkForAnalysis()).toBe(true); });
    expect(guard.notice).toBeNull();
  });
  it("成员真的变了仍然拦下：少勾一张不是顺序问题", async () => {
    stubThreeSheetWorkbook();
    let committed = ["Sheet1", "Sheet10", "Sheet11"];
    let guard!: ReturnType<typeof useAuthorizedTarget>;
    function Harness() {
      guard = useAuthorizedTarget({
        enabled: true,
        inspect: () => inspectOfficeAuthorizedTarget(committed),
        adopt: (target) => {
          committed = target.sheets
            .filter((sheet) => target.selected.includes(sheet.id))
            .map((sheet) => sheet.name);
        },
        analysisEvidence: officeAnalysisTargetEvidence
      });
      return null;
    }
    render(<Harness />);
    await waitFor(() => expect(guard.binding).not.toBeNull());
    const clicked: AuthorizedTarget = { ...guard.binding!, selected: ["11", "1", "10"] };
    await act(async () => { expect(await guard.confirmScope(clicked)).toBe(true); });

    // 外部少授权一张（同一顺序、同一文件，只有成员不同）：必须判范围变化。
    committed = ["Sheet1", "Sheet10"];
    await act(async () => { expect(await guard.checkForAnalysis()).toBe(false); });
    expect(guard.notice?.reason).toBe("scope_changed");
  });
});
