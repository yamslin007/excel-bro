import { describe, it, expect } from "vitest";
import {
  WORKSHEET_CHECKPOINT_OPENED_NOTICE,
  WORKSHEET_CHECKPOINT_RECOVERY_HINT,
  WORKSHEET_CHECKPOINT_UNDO_LABEL,
  checkpointBindingInScope,
  describeWorksheetCheckpoint,
  isUnattributedCheckpoint,
  parseWorksheetCheckpointBindings,
  serializeWorksheetCheckpointBindings,
  withWorksheetCheckpointBinding,
  type WorksheetCheckpointBinding
} from "./worksheetCheckpoint";

const id = (seed: string) => seed.repeat(32).slice(0, 32);
const binding: WorksheetCheckpointBinding = {
  operationId: id("a"),
  conversationId: "conv-1",
  workbookIdentity: "c:\\book.xlsx",
  requestMessageId: "msg-1"
};
const receipt = {
  operationId: binding.operationId,
  status: "completed" as const,
  deletedSheets: ["B", "C"],
  remainingSheets: ["A"],
  recoveryAvailable: true,
  message: "已删除 2 张工作表，并核对剩余工作表。恢复会打开执行前副本，不覆盖当前原表。"
};

describe("worksheet checkpoint attribution", () => {
  it("keeps legacy bare operationIds unattributed instead of guessing a conversation", () => {
    const parsed = parseWorksheetCheckpointBindings(JSON.parse(JSON.stringify([binding.operationId])));
    expect(parsed).toHaveLength(1);
    expect(parsed[0]).toEqual({ operationId: binding.operationId, conversationId: null, workbookIdentity: null, requestMessageId: null });
    expect(isUnattributedCheckpoint(parsed[0])).toBe(true);
    expect(checkpointBindingInScope(parsed[0], { conversationId: "conv-1", workbookIdentity: "c:\\book.xlsx" })).toBe(false);
  });

  it("round-trips scoped bindings and refuses to overwrite known scope with null", () => {
    const stored = parseWorksheetCheckpointBindings(JSON.parse(serializeWorksheetCheckpointBindings([binding])));
    expect(stored).toEqual([binding]);
    const merged = withWorksheetCheckpointBinding(stored, { ...binding, conversationId: null, requestMessageId: null });
    expect(merged[0].conversationId).toBe("conv-1");
    expect(merged[0].requestMessageId).toBe("msg-1");
  });

  it("requires both an explicit conversation and an explicit workbook identity", () => {
    const scope = { conversationId: "conv-1", workbookIdentity: "c:\\book.xlsx" };
    expect(checkpointBindingInScope(binding, scope)).toBe(true);
    expect(checkpointBindingInScope(binding, { ...scope, conversationId: "conv-2" })).toBe(false);
    expect(checkpointBindingInScope(binding, { ...scope, workbookIdentity: "c:\\other.xlsx" })).toBe(false);
    expect(checkpointBindingInScope(binding, { ...scope, workbookIdentity: null })).toBe(false);
  });

  it("drops malformed ids, duplicates and non-object entries", () => {
    const parsed = parseWorksheetCheckpointBindings([
      "not-an-id",
      binding.operationId,
      binding.operationId,
      { operationId: "b".repeat(32) },
      42
    ]);
    expect(parsed.map(item => item.operationId)).toEqual([binding.operationId, "b".repeat(32)]);
    expect(parseWorksheetCheckpointBindings("{")).toEqual([]);
  });
});

describe("worksheet checkpoint card copy", () => {
  it("uses the same undo entry name as every other result card", () => {
    // 两端、两张卡只有一种叫法；改名必须是一次显式决定，不能一端悄悄漂走。
    expect(WORKSHEET_CHECKPOINT_UNDO_LABEL).toBe("↶ 撤销本次执行");
    // 按钮背后是"打开副本"，说明句必须同时在场，且不得暗示原地回滚。
    expect(WORKSHEET_CHECKPOINT_RECOVERY_HINT).toBe("打开执行前副本，当前原表保留");
    expect(WORKSHEET_CHECKPOINT_RECOVERY_HINT).not.toContain("还原");
  });

  it("states only the short result on success and does not repeat recovery wording", () => {
    const view = describeWorksheetCheckpoint({ receipt });
    expect(view.title).toBe("工作表删除已完成");
    expect(view.summary).toBe("已删除 2 张工作表");
    expect(view.summary).not.toContain("恢复");
    expect(view.deletedSheets).toEqual(["B", "C"]);
    expect(view.recoveryAvailable).toBe(true);
    expect(view.canQueryStatus).toBe(false);
  });

  it("adds an opened-copy notice without claiming the original was restored", () => {
    const view = describeWorksheetCheckpoint({ receipt, opened: true });
    expect(view.summary).toBe("已删除 2 张工作表");
    expect(view.openedNotice).toBe(WORKSHEET_CHECKPOINT_OPENED_NOTICE);
    expect(view.openedNotice).not.toContain("撤销");
    expect(view.openedNotice).not.toContain("还原");
  });

  it("keeps failures and partial outcomes visible, moving only the diagnosis into details", () => {
    const failed = describeWorksheetCheckpoint({
      receipt: { ...receipt, status: "failed", deletedSheets: [], remainingSheets: [], recoveryAvailable: false,
        message: "检查点准备阶段中断，未开始删除。请重新发起任务。" }
    });
    expect(failed.tone).toBe("failed");
    expect(failed.summary).toBe("检查点准备阶段中断，未开始删除。");
    expect(failed.diagnostic).toContain("请重新发起任务");
    expect(failed.canQueryStatus).toBe(true);

    const partial = describeWorksheetCheckpoint({
      receipt: { ...receipt, status: "partial_or_unknown", recoveryAvailable: true,
        message: "删除未全部完成或结果未确认；不会自动重试。" }
    });
    expect(partial.tone).toBe("partial");
    expect(partial.recoveryAvailable).toBe(true);
  });

  it("never offers backup consent after deferral and never silently widens the scope", () => {
    const consent = describeWorksheetCheckpoint({
      receipt: { ...receipt, status: "needs_backup_consent", deletedSheets: [], remainingSheets: [], recoveryAvailable: false,
        backupSheets: ["#EB_RULES", "#EB_RULES_BACKUP"], message: "完整恢复备份还需包含以下未勾选工作表。" }
    });
    expect(consent.consentSheets).toEqual(["#EB_RULES", "#EB_RULES_BACKUP"]);
    expect(consent.summary).toContain("尚未备份或删除");

    const deferred = describeWorksheetCheckpoint({
      receipt: { ...receipt, status: "needs_backup_consent", deletedSheets: [], remainingSheets: [], recoveryAvailable: false,
        backupSheets: ["#EB_RULES"] },
      deferred: true
    });
    expect(deferred.consentSheets).toBeNull();
    expect(deferred.summary).toContain("没有删除任何工作表");
  });

  it("reports an unknown status without pretending deletion finished", () => {
    const unknown = describeWorksheetCheckpoint({ receipt: null, error: "未取得执行状态；请查询检查点，不要重复删除。" });
    expect(unknown.tone).toBe("failed");
    expect(unknown.summary).toContain("不要重复删除");
    expect(unknown.canQueryStatus).toBe(false);
    const pending = describeWorksheetCheckpoint({ receipt: null });
    expect(pending.title).toBe("正在删除工作表");
    expect(pending.canQueryStatus).toBe(true);
  });
});
