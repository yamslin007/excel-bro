// @vitest-environment jsdom
import { afterEach, beforeEach, describe, it, expect, vi } from "vitest";
import { useEffect } from "react";
import { cleanup, renderHook, act, waitFor, render, screen, fireEvent } from "@testing-library/react";
import { useWorksheetCheckpoint } from "./useWorksheetCheckpoint";
import {
  WorksheetCheckpointCard,
  WorksheetCheckpointRecoverySection
} from "../components/WorksheetCheckpointCard";
import {
  confirmWorksheetCheckpointBackup,
  executeWorksheetDeletion,
  listWorksheetCheckpoints,
  openWorksheetCheckpoint,
  restoreWorksheetCheckpoint,
  worksheetCheckpointStatus
} from "../api";
vi.mock("../api", () => ({
  executeWorksheetDeletion: vi.fn(),
  worksheetCheckpointStatus: vi.fn(),
  openWorksheetCheckpoint: vi.fn(),
  restoreWorksheetCheckpoint: vi.fn(),
  confirmWorksheetCheckpointBackup: vi.fn(),
  listWorksheetCheckpoints: vi.fn()
}));

const input = {
  workbookPath: "C:\\test.xlsx",
  authorizedSheets: ["A", "B"],
  plan: { id: "p", title: "retain A", summary: "delete B", assumptions: [], warnings: [], actions: [{ type: "deleteWorksheet" as const, sheet: "B" }] }
};
const receipt = {
  operationId: "a".repeat(32),
  status: "completed" as const,
  deletedSheets: ["B"],
  remainingSheets: ["A"],
  recoveryAvailable: true,
  message: "已删除 1 张工作表，并核对剩余工作表。"
};
const scope = { conversationId: "conv-1", workbookIdentity: "url:c:/test.xlsx" };
const context = { ...scope, requestMessageId: "msg-1" };
const storageKey = (host: string) => `excel-bro.worksheet-checkpoints.${host}.v1`;
/** 原地撤销成功的回执：把 `receipt.deletedSheets` 全部放回。 */
const restoredReceipt = (deleted: string[] = ["B"]) => ({
  operationId: receipt.operationId,
  status: "completed" as const,
  restoredSheets: [...deleted],
  failedSheets: [],
  unverifiedSheets: [],
  repairedFormulas: 0,
  repairedNames: 0,
  limitations: [],
  message: `已把 ${deleted.length} 张工作表放回原工作簿。`
});

afterEach(cleanup);
beforeEach(() => {
  localStorage.clear();
  vi.resetAllMocks();
  vi.mocked(worksheetCheckpointStatus).mockResolvedValue(receipt);
  vi.mocked(openWorksheetCheckpoint).mockResolvedValue({ opened: true, recoveryPath: "local" });
  vi.mocked(restoreWorksheetCheckpoint).mockResolvedValue(restoredReceipt() as never);
  vi.mocked(listWorksheetCheckpoints).mockResolvedValue([]);
});

describe.each(["excel", "wps-et"] as const)("%s checkpoint attribution", host => {
  it("records conversation, workbook and request message before dispatch; a lost response never deletes twice", async () => {
    vi.mocked(executeWorksheetDeletion).mockImplementation(async req => {
      expect(req.conversationId).toBe("conv-1");
      expect(req.requestMessageId).toBe("msg-1");
      const stored = JSON.parse(localStorage.getItem(storageKey(host)) ?? "{}");
      const saved = stored.bindings.find((item: { operationId: string }) => item.operationId === req.operationId);
      expect(saved).toMatchObject({ conversationId: "conv-1", workbookIdentity: "url:c:/test.xlsx", requestMessageId: "msg-1" });
      throw new Error("lost response");
    });
    const { result } = renderHook(() => useWorksheetCheckpoint(host));
    await act(async () => { expect(await result.current.execute("task", input, context)).toBe(true); });
    expect(result.current.entries[0].result?.status).toBe("completed");
    expect(result.current.entryForMessage("msg-1", "conv-1")?.operationId).toBe(result.current.entries[0].operationId);
    await act(async () => { expect(await result.current.execute("task", input, context)).toBe(true); });
    expect(executeWorksheetDeletion).toHaveBeenCalledOnce();
  });

  it("reload queries status only for this conversation and workbook, and an in-flight switch keeps the result on its own request", async () => {
    const other = "b".repeat(32);
    localStorage.setItem(storageKey(host), JSON.stringify({ version: 2, bindings: [
      { operationId: receipt.operationId, conversationId: "conv-1", workbookIdentity: "url:c:/test.xlsx", requestMessageId: "msg-1" },
      { operationId: other, conversationId: "conv-2", workbookIdentity: "url:c:/test.xlsx", requestMessageId: "msg-2" },
      { operationId: "c".repeat(32), conversationId: "conv-1", workbookIdentity: "url:c:/other.xlsx", requestMessageId: "msg-3" },
      "d".repeat(32)
    ] }));
    const { result } = renderHook(() => useWorksheetCheckpoint(host));
    await act(async () => { await result.current.hydrate(scope); });
    await waitFor(() => expect(result.current.entries).toHaveLength(1));
    expect(worksheetCheckpointStatus).toHaveBeenCalledOnce();
    expect(worksheetCheckpointStatus).toHaveBeenCalledWith(receipt.operationId);
    expect(result.current.entryForMessage("msg-2", "conv-2")).toBeNull();
    expect(executeWorksheetDeletion).not.toHaveBeenCalled();
    // 第二轮水合不重复查询（刷新后记录不重复）。
    await act(async () => { await result.current.hydrate(scope); });
    expect(worksheetCheckpointStatus).toHaveBeenCalledOnce();
  });

  it("does not guess a conversation or workbook for legacy ids", async () => {
    localStorage.setItem(storageKey(host), JSON.stringify(["d".repeat(32)]));
    const { result } = renderHook(() => useWorksheetCheckpoint(host));
    await act(async () => { await result.current.hydrate(scope); });
    expect(worksheetCheckpointStatus).not.toHaveBeenCalled();
    expect(result.current.entries).toHaveLength(0);
    expect(result.current.entryForMessage("msg-1", "conv-1")).toBeNull();
  });

  it("binds the result card to the message that asked for it and never appends elsewhere", async () => {
    let release!: (value: typeof receipt) => void;
    vi.mocked(executeWorksheetDeletion).mockImplementation(req =>
      new Promise(resolve => { release = value => resolve({ ...value, operationId: req.operationId }); }));
    const { result } = renderHook(() => useWorksheetCheckpoint(host));
    let running!: Promise<boolean>;
    act(() => { running = result.current.execute("task", input, context); });
    await waitFor(() => expect(result.current.entries).toHaveLength(1));
    // 请求在途时切走会话：新会话里没有这张卡。
    expect(result.current.entryForMessage("msg-1", "conv-2")).toBeNull();
    await act(async () => { release(receipt); await running; });
    // 结果仍写回原会话的那条请求。
    expect(result.current.entryForMessage("msg-1", "conv-1")?.result?.status).toBe("completed");
    expect(result.current.entries).toHaveLength(1);
  });

  /*
   * 卡上那颗「↶ 撤销本次执行」现在**真的**在原工作簿里撤销：它调的是 restore，
   * 不是"打开一份副本"。只有服务端回执核验通过（全部放回、无失败、无未核对）时
   * 才写"已撤销"，并且此时才提起副本 —— 副本是「备份与恢复」里的备用入口。
   */
  it("undo restores in place and only then says 已撤销，不打开副本", async () => {
    localStorage.setItem(storageKey(host), JSON.stringify({ version: 2, bindings: [
      { operationId: receipt.operationId, conversationId: "conv-1", workbookIdentity: "url:c:/test.xlsx", requestMessageId: "msg-1" }
    ] }));
    function Host() {
      const controller = useWorksheetCheckpoint(host);
      useEffect(() => { void controller.hydrate(scope); }, []);
      const entry = controller.entryForMessage("msg-1", "conv-1");
      return entry ? <WorksheetCheckpointCard entry={entry} controller={controller} /> : null;
    }
    const { container } = render(<Host />);
    await waitFor(() => expect(screen.getByRole("button", { name: "↶ 撤销本次执行" })).toBeTruthy());
    fireEvent.click(screen.getByRole("button", { name: "↶ 撤销本次执行" }));
    // 对象是**这次执行**，不是"某份文件"：operationId + 本次删掉的表名。
    await waitFor(() => expect(restoreWorksheetCheckpoint)
      .toHaveBeenCalledWith(receipt.operationId, receipt.deletedSheets));
    expect(await screen.findByText("已撤销本次删除。")).toBeTruthy();
    expect(container.textContent).toContain("已恢复 1 张工作表");
    expect(container.textContent).toContain("已放回：B");
    expect(openWorksheetCheckpoint).not.toHaveBeenCalled();
    expect(container.textContent).not.toContain("已打开恢复副本");
    expect(executeWorksheetDeletion).not.toHaveBeenCalled();
    // 终态：入口收起，不给第二次点击的机会。
    expect(screen.queryByRole("button", { name: "↶ 撤销本次执行" })).toBeNull();
  });

  it("a retry after a partial restore never claims 已撤销 and keeps the exit for a second try", async () => {
    vi.mocked(restoreWorksheetCheckpoint).mockResolvedValueOnce({
      ...restoredReceipt(),
      status: "partial",
      restoredSheets: ["B"],
      failedSheets: ["C"],
      message: "已放回 B，C 放回失败。"
    } as never);
    localStorage.setItem(storageKey(host), JSON.stringify({ version: 2, bindings: [
      { operationId: receipt.operationId, conversationId: "conv-1", workbookIdentity: "url:c:/test.xlsx", requestMessageId: "msg-1" }
    ] }));
    const { result } = renderHook(() => useWorksheetCheckpoint(host));
    await act(async () => { await result.current.hydrate(scope); });
    await act(async () => {
      expect(await result.current.restore(receipt.operationId)).toMatchObject({ ok: false });
    });
    expect(result.current.entryForMessage("msg-1", "conv-1")?.result?.restore?.status).toBe("partial");
    // 失败原文留在卡上，并且**不是**删除本身出错。
    const partialError = result.current.entryForMessage("msg-1", "conv-1")?.error;
    expect(partialError).toBeUndefined();
    // 再点一次：不重复造表（服务端幂等）+ 结果如实覆盖。
    await act(async () => {
      expect(await result.current.restore(receipt.operationId)).toMatchObject({ ok: true });
    });
    expect(vi.mocked(restoreWorksheetCheckpoint)).toHaveBeenCalledTimes(2);
    expect(vi.mocked(restoreWorksheetCheckpoint).mock.calls[1]).toEqual([receipt.operationId, receipt.deletedSheets]);
    expect(result.current.entryForMessage("msg-1", "conv-1")?.result?.restore?.status).toBe("completed");
  });

  it("a restore that never answered writes the failure on the card and keeps the exit", async () => {
    vi.mocked(restoreWorksheetCheckpoint).mockRejectedValueOnce(new Error("检查点文件校验失败。"));
    localStorage.setItem(storageKey(host), JSON.stringify({ version: 2, bindings: [
      { operationId: receipt.operationId, conversationId: "conv-1", workbookIdentity: "url:c:/test.xlsx", requestMessageId: "msg-1" }
    ] }));
    function Host() {
      const controller = useWorksheetCheckpoint(host);
      useEffect(() => { void controller.hydrate(scope); }, []);
      const entry = controller.entryForMessage("msg-1", "conv-1");
      return entry ? <WorksheetCheckpointCard entry={entry} controller={controller} /> : null;
    }
    const { container } = render(<Host />);
    await waitFor(() => expect(screen.getByRole("button", { name: "↶ 撤销本次执行" })).toBeTruthy());
    fireEvent.click(screen.getByRole("button", { name: "↶ 撤销本次执行" }));
    const alert = await waitFor(() => {
      const node = container.querySelector('[role="alert"]');
      expect(node).not.toBeNull();
      return node as HTMLElement;
    });
    expect(alert.textContent).toContain("检查点文件校验失败。");
    // 不谎报成功，也不把恢复记录清掉：入口还在，可以重试。
    expect(container.textContent).not.toContain("已撤销本次删除");
    expect(screen.getByRole("button", { name: "↶ 撤销本次执行" })).toBeTruthy();
    expect(openWorksheetCheckpoint).not.toHaveBeenCalled();
  });

  it("a double click cannot dispatch a second restore", async () => {
    let release!: () => void;
    vi.mocked(restoreWorksheetCheckpoint).mockImplementationOnce(() =>
      new Promise(resolve => { release = () => resolve(restoredReceipt() as never); }));
    localStorage.setItem(storageKey(host), JSON.stringify({ version: 2, bindings: [
      { operationId: receipt.operationId, conversationId: "conv-1", workbookIdentity: "url:c:/test.xlsx", requestMessageId: "msg-1" }
    ] }));
    const { result } = renderHook(() => useWorksheetCheckpoint(host));
    await act(async () => { await result.current.hydrate(scope); });
    let first!: Promise<unknown>;
    act(() => { first = result.current.restore(receipt.operationId); });
    // 在途期间的第二次点击必须被挡下（不排队、不重复造表）。
    await act(async () => {
      expect(await result.current.restore(receipt.operationId)).toMatchObject({
        ok: false, error: "上一项检查点操作尚未结束。"
      });
    });
    await act(async () => { release(); await first; });
    expect(restoreWorksheetCheckpoint).toHaveBeenCalledOnce();
  });

  it("an incomplete deletion has no in-place undo and points at the copy instead", async () => {
    vi.mocked(worksheetCheckpointStatus).mockResolvedValue({ ...receipt, status: "partial_or_unknown", message: "删除结果待确认。" });
    localStorage.setItem(storageKey(host), JSON.stringify({ version: 2, bindings: [
      { operationId: receipt.operationId, conversationId: "conv-1", workbookIdentity: "url:c:/test.xlsx", requestMessageId: "msg-1" }
    ] }));
    function Host() {
      const controller = useWorksheetCheckpoint(host);
      useEffect(() => { void controller.hydrate(scope); }, []);
      const entry = controller.entryForMessage("msg-1", "conv-1");
      return entry ? <WorksheetCheckpointCard entry={entry} controller={controller} /> : null;
    }
    const { container } = render(<Host />);
    const blocked = await waitFor(() => {
      const node = container.querySelector(".worksheet-checkpoint-undo-blocked");
      expect(node).not.toBeNull();
      return node as HTMLElement;
    });
    expect(blocked.textContent).toContain("不能原地撤销");
    expect(blocked.textContent).toContain("可打开执行前副本");
    expect(screen.queryByRole("button", { name: "↶ 撤销本次执行" })).toBeNull();
  });

  it("keeps a failed copy open retryable and never hides the record", async () => {
    vi.mocked(openWorksheetCheckpoint).mockRejectedValueOnce(new Error("恢复副本校验失败，未打开。"));
    localStorage.setItem(storageKey(host), JSON.stringify({ version: 2, bindings: [
      { operationId: receipt.operationId, conversationId: "conv-1", workbookIdentity: "url:c:/test.xlsx", requestMessageId: "msg-1" }
    ] }));
    const { result } = renderHook(() => useWorksheetCheckpoint(host));
    await act(async () => { await result.current.hydrate(scope); });
    await act(async () => { expect(await result.current.open(receipt.operationId)).toEqual({ ok: false, error: "恢复副本校验失败，未打开。" }); });
    expect(result.current.entryForMessage("msg-1", "conv-1")?.error).toContain("未打开");
    vi.mocked(openWorksheetCheckpoint).mockResolvedValue({ opened: true, recoveryPath: "local" });
    await act(async () => { expect(await result.current.open(receipt.operationId)).toEqual({ ok: true }); });
    expect(result.current.entryForMessage("msg-1", "conv-1")?.opened).toBe(true);
  });

  it("asks for backup-only consent at the same request and rechecks live authority", async () => {
    const pending = { ...receipt, status: "needs_backup_consent" as const, deletedSheets: [], remainingSheets: [], recoveryAvailable: false, backupSheets: ["#EB_RULES", "#EB_RULES_BACKUP"], message: "需要本地备份许可" };
    vi.mocked(executeWorksheetDeletion).mockImplementation(async req => ({ ...pending, operationId: req.operationId }));
    vi.mocked(confirmWorksheetCheckpointBackup).mockImplementation(async req => ({ ...receipt, operationId: req.operationId }));
    let authorized = true;
    const beforeResume = vi.fn(async () => { if (!authorized) throw new Error("工作簿或授权变化"); });
    const onCompleted = vi.fn();
    const { result } = renderHook(() => useWorksheetCheckpoint(host));
    await act(async () => { expect(await result.current.execute("task", input, { ...context, beforeResume, onCompleted })).toBe(false); });
    const id = result.current.entries[0].operationId;
    expect(result.current.waitingForConsent("task")).toBe(true);
    expect(confirmWorksheetCheckpointBackup).not.toHaveBeenCalled();
    authorized = false;
    await act(async () => { await result.current.confirmBackup(id); });
    expect(result.current.entries[0].error).toContain("授权变化");
    expect(confirmWorksheetCheckpointBackup).not.toHaveBeenCalled();
    authorized = true;
    await act(async () => { await result.current.confirmBackup(id); });
    expect(confirmWorksheetCheckpointBackup).toHaveBeenCalledWith({ operationId: id, backupSheets: pending.backupSheets });
    expect(executeWorksheetDeletion).toHaveBeenCalledOnce();
    expect(vi.mocked(executeWorksheetDeletion).mock.calls[0][0].authorizedSheets).toEqual(["A", "B"]);
    expect(onCompleted).toHaveBeenCalledOnce();
  });

  it("a double click cannot dispatch a second native batch", async () => {
    let release!: (value: typeof receipt) => void;
    vi.mocked(executeWorksheetDeletion).mockImplementation(() => new Promise(resolve => { release = resolve; }));
    const { result } = renderHook(() => useWorksheetCheckpoint(host));
    let first!: Promise<boolean>;
    act(() => { first = result.current.execute("one", input, context); });
    await expect(result.current.execute("two", input, context)).rejects.toThrow("尚未结束");
    await act(async () => { release(receipt); await first; });
    expect(executeWorksheetDeletion).toHaveBeenCalledOnce();
  });

  it("writes a pre-dispatch failure onto the request card instead of a new bottom message", async () => {
    vi.mocked(executeWorksheetDeletion).mockImplementation(async req => ({ ...receipt, operationId: req.operationId }));
    const { result } = renderHook(() => useWorksheetCheckpoint(host));
    await act(async () => { await result.current.execute("task", input, context); });
    act(() => { expect(result.current.reportIssue("msg-1", "conv-1", "授权在检查期间变化，未删除。")).toBe(true); });
    expect(result.current.entryForMessage("msg-1", "conv-1")?.error).toContain("授权");
    expect(result.current.entries).toHaveLength(1);
    expect(result.current.reportIssue("missing", "conv-1", "x")).toBe(false);
  });

  it("lists recovery points by workbook identity for 备份与恢复", async () => {
    const { result } = renderHook(() => useWorksheetCheckpoint(host));
    await act(async () => { await result.current.listRecoverable("C:\\test.xlsx"); });
    expect(listWorksheetCheckpoints).toHaveBeenCalledWith({ workbookPath: "C:\\test.xlsx", host });
  });

  it("shows a legacy recovery point in 备份与恢复 and opens a copy without executing anything", async () => {
    const legacy = "d".repeat(32);
    vi.mocked(listWorksheetCheckpoints).mockResolvedValue([{
      operationId: legacy, status: "completed", deletedSheets: ["B"], remainingSheets: ["A"],
      recoveryAvailable: true, message: "已删除 1 张工作表。", createdAt: null, planTitle: "旧记录"
    }]);
    function Panel() {
      const controller = useWorksheetCheckpoint(host);
      return <WorksheetCheckpointRecoverySection controller={controller} workbookPath={"C:\\test.xlsx"} />;
    }
    render(<Panel />);
    await waitFor(() => expect(listWorksheetCheckpoints).toHaveBeenCalledWith({ workbookPath: "C:\\test.xlsx", host }));
    expect(await screen.findByText("旧记录")).toBeTruthy();
    expect(screen.getByText("时间未知")).toBeTruthy();
    fireEvent.click(screen.getByRole("button", { name: "打开执行前恢复副本" }));
    await waitFor(() => expect(openWorksheetCheckpoint).toHaveBeenCalledWith(legacy));
    expect(await screen.findByText("已打开恢复副本。")).toBeTruthy();
    expect(executeWorksheetDeletion).not.toHaveBeenCalled();
  });
});
