import { useCallback, useEffect, useRef, useState } from "react";
import {
  confirmWorksheetCheckpointBackup,
  executeWorksheetDeletion,
  listWorksheetCheckpoints,
  openWorksheetCheckpoint,
  restoreWorksheetCheckpoint,
  worksheetCheckpointStatus
} from "../api";
import type {
  WorksheetCheckpointSummary,
  WorksheetDeleteCheckpointRequest,
  WorksheetDeleteCheckpointResult,
  WorksheetRestoreResult
} from "../contracts";
import {
  checkpointBindingInScope,
  parseWorksheetCheckpointBindings,
  serializeWorksheetCheckpointBindings,
  withWorksheetCheckpointBinding,
  type WorksheetCheckpointBinding,
  type WorksheetCheckpointScope
} from "@excel-bro/core/worksheetCheckpoint";

/**
 * 一条工作表删除检查点在本机的记录。
 *
 * 归属（会话 / 工作簿 / 请求消息）在**发起那一刻**就固定，后续异步回执只更新
 * 这一条，不会追加到用户后来切到的会话里，也不会被画到别的请求下面。
 */
export type WorksheetCheckpointEntry = {
  operationId: string;
  conversationId: string | null;
  workbookIdentity: string | null;
  requestMessageId: string | null;
  result?: WorksheetDeleteCheckpointResult;
  error?: string;
  deferred?: boolean;
  opening?: boolean;
  opened?: boolean;
  /** 原地撤销在途：按钮据此禁用，防重复点击。 */
  restoring?: boolean;
  /** 原地撤销的失败/冲突原文（与删除本身的 `error` 分开，别混成一次失败）。 */
  restoreError?: string;
};

/** 发起一次检查点操作时必须提供的归属与现场复核。 */
export type WorksheetCheckpointIssueContext = WorksheetCheckpointScope & {
  requestMessageId?: string | null;
  /** 备份授权后再次核对授权/工作簿现场；不通过就不继续删除。 */
  beforeResume?: () => Promise<void>;
  onCompleted?: () => Promise<void> | void;
};

export type WorksheetCheckpointOpenOutcome = { ok: boolean; error?: string };
/** 原地撤销的结果：`ok` 只在**核验过的全部放回**时为真。 */
export type WorksheetRestoreOutcome = { ok: boolean; error?: string; result?: WorksheetRestoreResult };

export function useWorksheetCheckpoint(host: "excel" | "wps-et") {
  const [entries, setEntries] = useState<WorksheetCheckpointEntry[]>([]);
  const [pending, setPending] = useState(false);
  const inFlight = useRef(false);
  const issued = useRef(new Map<string, string>());
  const contexts = useRef(new Map<string, WorksheetCheckpointIssueContext>());
  const receipts = useRef(new Map<string, WorksheetDeleteCheckpointResult>());
  const bindings = useRef<WorksheetCheckpointBinding[]>([]);
  /** 本次挂载已经查过状态的 operationId；刷新后不重复查询，更不会重新执行。 */
  const hydrated = useRef(new Set<string>());
  const storageKey = `excel-bro.worksheet-checkpoints.${host}.v1`;
  const entriesRef = useRef<WorksheetCheckpointEntry[]>(entries);
  entriesRef.current = entries;

  const findBinding = (operationId: string) =>
    bindings.current.find(binding => binding.operationId === operationId) ?? null;

  const patchEntry = (operationId: string, patch: Partial<WorksheetCheckpointEntry>) => {
    if (patch.result) receipts.current.set(operationId, patch.result);
    setEntries(current => {
      const index = current.findIndex(entry => entry.operationId === operationId);
      if (index >= 0) {
        return current.map((entry, position) =>
          position === index ? { ...entry, ...patch } : entry);
      }
      const binding = findBinding(operationId);
      return [...current, {
        operationId,
        conversationId: binding?.conversationId ?? null,
        workbookIdentity: binding?.workbookIdentity ?? null,
        requestMessageId: binding?.requestMessageId ?? null,
        ...patch
      }];
    });
  };

  /**
   * 归属先落盘，再派发请求。
   *
   * 顺序不能反：响应丢失时靠这份记录才能"只查状态、不重复删除"；刷新后也靠它
   * 把结果卡放回**原来的那条请求消息**。
   */
  const rememberBinding = (binding: WorksheetCheckpointBinding) => {
    bindings.current = withWorksheetCheckpointBinding(bindings.current, binding);
    try {
      localStorage.setItem(storageKey, serializeWorksheetCheckpointBindings(bindings.current));
    } catch {
      // 存储不可用时本次会话照常运行；丢失的只是"刷新后回挂到原消息"的能力。
    }
  };

  useEffect(() => {
    try {
      const stored: unknown = JSON.parse(localStorage.getItem(storageKey) ?? "[]");
      // 旧版只存 operationId：解析成"归属未知"，不猜会话、不猜工作簿。
      bindings.current = parseWorksheetCheckpointBindings(stored);
    } catch {
      // 损坏的发现数据不会授权任何写入，也不会被猜成某条请求的归属。
      bindings.current = [];
    }
  }, [storageKey]);

  /**
   * 只查询**当前会话 + 当前工作簿**里已有操作的状态。
   *
   * 别的会话/工作簿的记录不在这里出现，也就不会变成聊天底部的常驻卡；它们仍可
   * 从「备份与恢复」按工作簿身份找回。
   */
  async function hydrate(scope: WorksheetCheckpointScope) {
    const targets = bindings.current.filter(binding =>
      checkpointBindingInScope(binding, scope) && !hydrated.current.has(binding.operationId));
    for (const binding of targets) {
      if (hydrated.current.has(binding.operationId)) continue;
      hydrated.current.add(binding.operationId);
      if (entriesRef.current.some(entry => entry.operationId === binding.operationId && entry.result)) continue;
      await refresh(binding.operationId);
    }
  }

  async function refresh(operationId: string) {
    if (!findBinding(operationId)) return null;
    try {
      const result = await worksheetCheckpointStatus(operationId);
      patchEntry(operationId, { result, error: undefined });
      return result;
    } catch {
      patchEntry(operationId, { error: "未取得执行状态；请查询检查点，不要重复删除。" });
      return null;
    }
  }

  async function execute(
    key: string,
    input: Omit<WorksheetDeleteCheckpointRequest, "operationId" | "host">,
    context: WorksheetCheckpointIssueContext
  ) {
    if (inFlight.current) throw new Error("上一项检查点操作尚未结束。");
    if (!context.conversationId || !context.workbookIdentity) {
      throw new Error("无法确认本次操作属于哪个会话或工作簿，未删除。");
    }
    const existing = issued.current.get(key);
    if (existing) {
      // 同一任务只允许一次原生删除：重新进入只查状态。
      if (context.beforeResume || context.onCompleted) contexts.current.set(existing, context);
      return (await refresh(existing))?.status === "completed";
    }
    const operationId = crypto.randomUUID().replace(/-/g, "");
    const binding: WorksheetCheckpointBinding = {
      operationId,
      conversationId: context.conversationId,
      workbookIdentity: context.workbookIdentity,
      requestMessageId: context.requestMessageId ?? null
    };
    // Persist discovery BEFORE dispatch. A lost response must not cause a second deletion.
    rememberBinding(binding);
    issued.current.set(key, operationId);
    contexts.current.set(operationId, context);
    inFlight.current = true; setPending(true);
    patchEntry(operationId, { result: undefined, error: undefined });
    try {
      const result = await executeWorksheetDeletion({
        ...input,
        host,
        operationId,
        conversationId: binding.conversationId ?? undefined,
        requestMessageId: binding.requestMessageId ?? undefined
      });
      patchEntry(operationId, { result, error: undefined });
      return result.status === "completed";
    } catch {
      return (await refresh(operationId))?.status === "completed";
    } finally { inFlight.current = false; setPending(false); }
  }

  async function confirmBackup(operationId: string) {
    if (inFlight.current) return;
    const context = contexts.current.get(operationId);
    const receipt = receipts.current.get(operationId);
    if (!context || receipt?.status !== "needs_backup_consent" || !receipt.backupSheets?.length) return;
    inFlight.current = true; setPending(true);
    try {
      // Consent is for this operation only; refresh does not restore write authority.
      await context.beforeResume?.();
      let result: WorksheetDeleteCheckpointResult | null;
      try {
        result = await confirmWorksheetCheckpointBackup({ operationId, backupSheets: receipt.backupSheets });
        patchEntry(operationId, { result, error: undefined });
      } catch { result = await refresh(operationId); }
      if (result?.status === "completed") {
        contexts.current.delete(operationId);
        await context.onCompleted?.();
      }
    } catch (reason) {
      patchEntry(operationId, {
        error: reason instanceof Error ? reason.message : "授权检查未通过，未继续删除。"
      });
    } finally { inFlight.current = false; setPending(false); }
  }

  function defer(operationId: string) {
    if (inFlight.current) return;
    contexts.current.delete(operationId);
    patchEntry(operationId, { deferred: true });
  }

  const canResume = (operationId: string) => contexts.current.has(operationId);
  const waitingForConsent = (key: string) =>
    receipts.current.get(issued.current.get(key) ?? "")?.status === "needs_backup_consent";

  /**
   * 打开**执行前副本**（不是撤销整本工作簿）。
   *
   * 只报告"副本已打开"，从不宣称原表已还原；也不保存、不关闭原表。
   */
  async function open(operationId: string): Promise<WorksheetCheckpointOpenOutcome> {
    if (inFlight.current) return { ok: false, error: "上一项检查点操作尚未结束。" };
    inFlight.current = true; setPending(true);
    if (findBinding(operationId)) patchEntry(operationId, { opening: true });
    try {
      await openWorksheetCheckpoint(operationId);
      if (findBinding(operationId)) patchEntry(operationId, { opening: false, opened: true, error: undefined });
      return { ok: true };
    } catch (reason) {
      const error = reason instanceof Error ? reason.message : "恢复副本未打开；当前原表保留。";
      if (findBinding(operationId)) patchEntry(operationId, { opening: false, error });
      return { ok: false, error };
    } finally { inFlight.current = false; setPending(false); }
  }

  /**
   * **原地撤销**本次删除：把表放回当前打开的原工作簿，文件身份不变。
   *
   * 与 `open` 严格分开：
   * - 这里不打开任何新文件，也不把"打开副本"当成功；
   * - 回执自证覆盖关系（`assertWorksheetRestoreResult`）后才写进结果卡；
   * - 在途防重复点击；失败/冲突都保留重试入口与归档。
   */
  async function restore(operationId: string): Promise<WorksheetRestoreOutcome> {
    if (inFlight.current) return { ok: false, error: "上一项检查点操作尚未结束。" };
    const receipt = receipts.current.get(operationId) ?? null;
    if (!receipt || receipt.status !== "completed") {
      return { ok: false, error: "本次删除没有完整完成，不能原地撤销；可打开执行前副本核对。" };
    }
    if (receipt.restore?.status === "completed") {
      // 已经撤销过：不再发请求，避免重复造表。
      return { ok: true, result: receipt.restore };
    }
    inFlight.current = true; setPending(true);
    if (findBinding(operationId)) patchEntry(operationId, { restoring: true, restoreError: undefined });
    try {
      const result = await restoreWorksheetCheckpoint(operationId, receipt.deletedSheets);
      const next: WorksheetDeleteCheckpointResult = { ...receipt, restore: result };
      receipts.current.set(operationId, next);
      if (findBinding(operationId)) patchEntry(operationId, { restoring: false, result: next, restoreError: undefined });
      return { ok: result.status === "completed", result, error: result.status === "completed" ? undefined : result.message };
    } catch (reason) {
      const error = reason instanceof Error ? reason.message : "撤销尚未确认；请核对工作表列表后重试。";
      if (findBinding(operationId)) patchEntry(operationId, { restoring: false, restoreError: error });
      return { ok: false, error };
    } finally { inFlight.current = false; setPending(false); }
  }

  /**
   * 把一条"还没派发就停下"的失败写到对应请求的卡上。
   *
   * 只有已经建立归属的请求才有卡：返回 false 时调用方应写回**发起时的会话**，
   * 而不是当前的会话。
   */
  function reportIssue(requestMessageId: string | null, conversationId: string, message: string): boolean {
    if (!requestMessageId) return false;
    const entry = entriesRef.current.find(item =>
      item.requestMessageId === requestMessageId && item.conversationId === conversationId);
    if (!entry) return false;
    patchEntry(entry.operationId, { error: message });
    return true;
  }

  /** 只读：按工作簿身份列出恢复点，供「备份与恢复」使用（含旧记录）。 */
  async function listRecoverable(workbookPath: string): Promise<WorksheetCheckpointSummary[]> {
    return listWorksheetCheckpoints({ workbookPath, host });
  }

  /**
   * 某条请求的**已取得回执**（没有就返回 null）。
   *
   * 只读本机记录，不发请求、不改状态：删表完成后要用它的 `deletedSheets` /
   * `remainingSheets` 去核对"范围收缩是不是本次操作的预期结果"。
   *
   * 必须先查**同步记账**：`patchEntry` 把回执写进 `receipts` 是同步的，而 `entries`
   * 要等下一轮渲染。`confirmBackup` 恰好是"写完回执 → 立刻 `await onCompleted()`"，
   * 两个动作之间界面还没重画。若这里只读 `entriesRef`，收尾那一步就会拿到 null，
   * 把"有证据的预期收窄"误判成"没有证据"，下一轮守卫照旧弹"尚未写入"。
   */
  function receiptForMessage(
    requestMessageId: string | null,
    conversationId: string
  ): WorksheetDeleteCheckpointResult | null {
    if (!requestMessageId) return null;
    const binding = bindings.current.find(item =>
      item.requestMessageId === requestMessageId && item.conversationId === conversationId);
    const settled = binding ? receipts.current.get(binding.operationId) : undefined;
    if (settled) return settled;
    return entriesRef.current.find(entry =>
      entry.requestMessageId === requestMessageId && entry.conversationId === conversationId
    )?.result ?? null;
  }

  /**
   * 某条请求消息对应的检查点。
   *
   * 会话必须匹配：同一份消息列表只会画属于它自己的结果卡。
   */
  const entryForMessage = useCallback((messageId: string, conversationId: string) =>
    entries.find(entry =>
      entry.requestMessageId === messageId && entry.conversationId === conversationId) ?? null,
  [entries]);

  return {
    entries, pending, execute, refresh, open, restore, confirmBackup, canResume,
    waitingForConsent, defer, hydrate, listRecoverable, entryForMessage, reportIssue,
    receiptForMessage
  };
}
