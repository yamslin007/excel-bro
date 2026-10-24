/**
 * 工作表删除检查点的**归属**与**展示**判据（Excel / WPS 读同一份）。
 *
 * 这里刻意只放纯判据，不碰 React、不碰宿主：
 * - 归属：一条检查点记录属于哪个会话、哪本工作簿、哪条请求消息；
 * - 展示：这张卡该说哪一句真话、哪些明细默认折叠、恢复入口是否可用。
 *
 * 归属必须显式记录，**不许**用"当前会话 / 当前工作簿 / 列表顺序"反推：
 * 用户切走会话之后异步完成的结果，仍然只属于发起它的那条请求。
 */

export type WorksheetCheckpointStatus =
  | "prepared"
  | "needs_backup_consent"
  | "executing"
  | "completed"
  | "partial_or_unknown"
  | "failed";

/** 与 `WorksheetDeleteCheckpointResult` 结构一致；只读，避免两端各写一份字段解释。 */
export interface WorksheetCheckpointReceipt {
  operationId: string;
  status: WorksheetCheckpointStatus;
  backupSheets?: readonly string[];
  deletedSheets: readonly string[];
  remainingSheets: readonly string[];
  recoveryAvailable: boolean;
  message: string;
  /** 最近一次**原地撤销**的结果；缺失只表示"还没撤销过"。 */
  restore?: WorksheetCheckpointRestore | null;
}

/**
 * 原地撤销（把表放回当前打开的原表）的结果。
 *
 * 与"打开执行前副本"是两件不同的事，字段因此也分开：
 * `restoredSheets` 说的是**真的回到原工作簿**的表，不是"副本里有这些表"。
 */
export interface WorksheetCheckpointRestore {
  operationId: string;
  status: "completed" | "partial" | "failed" | "conflict";
  restoredSheets: readonly string[];
  failedSheets: readonly string[];
  unverifiedSheets: readonly string[];
  repairedFormulas: number;
  repairedNames: number;
  limitations: readonly string[];
  message: string;
}

/**
 * 一条检查点的归属。
 *
 * `null` 表示**归属未知**（旧版只存了 operationId）。这种记录不是"属于当前会话"，
 * 也不能被归到当前工作簿：它只保留恢复能力，从「备份与恢复」按工作簿身份另行查找。
 */
export interface WorksheetCheckpointBinding {
  operationId: string;
  conversationId: string | null;
  workbookIdentity: string | null;
  requestMessageId: string | null;
}

export interface WorksheetCheckpointScope {
  conversationId: string;
  workbookIdentity: string | null;
}

/** 服务端按工作簿列出的恢复点摘要（不含任何单元格内容）。 */
export interface WorksheetCheckpointSummary {
  operationId: string;
  status: WorksheetCheckpointStatus;
  deletedSheets: readonly string[];
  remainingSheets: readonly string[];
  recoveryAvailable: boolean;
  message: string;
  createdAt: string | null;
  planTitle: string | null;
}

export const WORKSHEET_CHECKPOINT_STORAGE_VERSION = 2;
/** 「备份与恢复」列表里的入口文案：一行一条记录，"本次"在那里没有唯一所指。 */
export const WORKSHEET_CHECKPOINT_RECOVERY_LABEL = "打开执行前恢复副本";
/**
 * 结果卡右下角的入口文案。
 *
 * 它现在**就是**原地撤销：把本次删除的工作表放回当前打开的原工作簿，
 * 文件身份不变、后续编辑保留。所以这句文案必须名副其实 —— 只有在真的会
 * 原地恢复时才出现（见 `describeWorksheetCheckpoint().undoAvailable`）。
 */
export const WORKSHEET_CHECKPOINT_UNDO_LABEL = "↶ 撤销本次执行";
/** 按钮旁边那句说明：说清"在原表里撤回"，不再提副本。 */
export const WORKSHEET_CHECKPOINT_UNDO_HINT = "在原工作簿撤回本次删除，保留后续编辑";
/** 原地撤销不可用时，指路到备用入口（副本）。 */
export const WORKSHEET_CHECKPOINT_BACKUP_FALLBACK_HINT =
  "删除前的整份文件仍保留在「备份与恢复」里，可打开执行前副本";
/**
 * 「备份与恢复」里那个备用入口的说明：它只打开副本，不动原表。
 *
 * 原地撤销上线后这句仍然成立，但只属于**副本身份**的入口 —— 不要拿它去解释
 * 结果卡上的撤销按钮，那里的语义已经换成"在原工作簿撤回"。
 */
export const WORKSHEET_CHECKPOINT_RECOVERY_HINT = "打开执行前副本，当前原表保留";
/** 撤销**成功后**的结论句。只说撤销，不说"原表已还原成某个副本"。 */
export const WORKSHEET_CHECKPOINT_RESTORED_NOTICE = "已撤销本次删除。";
/** 打开副本之后的独立一句（保留旧语义，不并入撤销结论）。 */
export const WORKSHEET_CHECKPOINT_OPENED_NOTICE = "已打开恢复副本。";

const OPERATION_ID = /^[a-f0-9]{32}$/;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function scopeField(value: unknown): string | null {
  return typeof value === "string" && value.trim().length > 0 ? value : null;
}

/**
 * 解析本机保存的绑定表。
 *
 * 兼容三种形态：`string[]`（旧版只存 operationId）、`{bindings: [...]}`、以及
 * 元素为对象的新版数组。旧版的裸 id 一律解析成**归属未知**，不猜会话。
 */
export function parseWorksheetCheckpointBindings(raw: unknown): WorksheetCheckpointBinding[] {
  let list: unknown[] = [];
  if (Array.isArray(raw)) list = raw;
  else if (isRecord(raw) && Array.isArray(raw.bindings)) list = raw.bindings;
  const seen = new Set<string>();
  const bindings: WorksheetCheckpointBinding[] = [];
  for (const item of list) {
    if (typeof item === "string") {
      if (!OPERATION_ID.test(item) || seen.has(item)) continue;
      seen.add(item);
      bindings.push({ operationId: item, conversationId: null, workbookIdentity: null, requestMessageId: null });
      continue;
    }
    if (!isRecord(item) || typeof item.operationId !== "string" || !OPERATION_ID.test(item.operationId)) continue;
    if (seen.has(item.operationId)) continue;
    seen.add(item.operationId);
    bindings.push({
      operationId: item.operationId,
      conversationId: scopeField(item.conversationId),
      workbookIdentity: scopeField(item.workbookIdentity),
      requestMessageId: scopeField(item.requestMessageId)
    });
  }
  return bindings;
}

export function serializeWorksheetCheckpointBindings(bindings: readonly WorksheetCheckpointBinding[]): string {
  return JSON.stringify({
    version: WORKSHEET_CHECKPOINT_STORAGE_VERSION,
    bindings: bindings.map(binding => ({
      operationId: binding.operationId,
      conversationId: binding.conversationId,
      workbookIdentity: binding.workbookIdentity,
      requestMessageId: binding.requestMessageId
    }))
  });
}

/**
 * 登记一条归属。
 *
 * 同一个 operationId 只登记一次：新信息（例如后来补上的 requestMessageId）采用
 * **后者优先但不清空已有值**，避免一次不完整的写入把已有归属抹成 null。
 */
export function withWorksheetCheckpointBinding(
  bindings: readonly WorksheetCheckpointBinding[],
  binding: WorksheetCheckpointBinding
): WorksheetCheckpointBinding[] {
  const index = bindings.findIndex(item => item.operationId === binding.operationId);
  if (index < 0) return [...bindings, binding];
  const merged: WorksheetCheckpointBinding = {
    operationId: binding.operationId,
    conversationId: binding.conversationId ?? bindings[index].conversationId,
    workbookIdentity: binding.workbookIdentity ?? bindings[index].workbookIdentity,
    requestMessageId: binding.requestMessageId ?? bindings[index].requestMessageId
  };
  return bindings.map((item, position) => (position === index ? merged : item));
}

/** 归属未知：旧版记录。既不属于当前会话，也不能被当前工作簿"认领"。 */
export function isUnattributedCheckpoint(binding: WorksheetCheckpointBinding): boolean {
  return binding.conversationId === null && binding.workbookIdentity === null && binding.requestMessageId === null;
}

/**
 * 这条记录是否属于当前作用域。
 *
 * 会话必须**完全相同**；工作簿身份必须明确且相同。任一为空即不匹配 ——
 * "当前工作簿没有身份（未保存）"不等于"所有未保存记录的归属"。
 */
export function checkpointBindingInScope(
  binding: WorksheetCheckpointBinding,
  scope: WorksheetCheckpointScope
): boolean {
  if (!binding.conversationId || binding.conversationId !== scope.conversationId) return false;
  if (!binding.workbookIdentity || !scope.workbookIdentity) return false;
  return binding.workbookIdentity === scope.workbookIdentity;
}

/** 卡片的语气映射到共享结果卡的 tone（复用现有样式，不另造一套）。 */
export type WorksheetCheckpointCardTone = "verified" | "failed" | "partial" | "unverified";

export interface WorksheetCheckpointCardView {
  tone: WorksheetCheckpointCardTone;
  title: string;
  /** 首屏那句短结论，例如「已删除 6 张工作表」。 */
  summary: string;
  deletedSheets: string[];
  remainingSheets: string[];
  /** 诊断原文：只在折叠区里出现。 */
  diagnostic: string | null;
  /** 是否还有等待用户决定的备份授权。 */
  consentSheets: string[] | null;
  /** 执行前副本是否还在（「备份与恢复」里的备用入口）。 */
  recoveryAvailable: boolean;
  /** 是否允许再查一次状态（终态不再重复查）。 */
  canQueryStatus: boolean;
  /**
   * 恢复副本**已经打开**时的独立一句。
   *
   * 它不并入 `summary`：打开副本不是执行结果，把两者合成一句就会让人以为
   * "原表已经还原、删除已撤销"。状态句与打开句必须各自成立、各自可读。
   */
  openedNotice: string | null;
  /**
   * 右下的「↶ 撤销本次执行」是否可以真的**原地**撤销。
   *
   * 只有"本次删除完整完成 + 有可信的归档 + 之前没撤销成功"才为 true。
   * 为 false 时必须给 `undoReason`，并保留备用入口的指路话术 ——
   * 一个点了只会打开副本的按钮，不能叫"撤销"。
   */
  undoAvailable: boolean;
  /** 不可撤销 / 上次未成功时的具体原因（首屏一句，可操作）。 */
  undoReason: string | null;
  /** 撤销按钮旁边那句说明。 */
  undoHint: string;
  /** 撤销成功后的结论句；未成功时为 null。 */
  restoredNotice: string | null;
  /** 已回到原工作簿的表名（折叠区明细）。 */
  restoredSheets: string[];
}

function shortList(names: readonly string[], limit = 3): string {
  if (names.length <= limit) return names.join("、");
  return `${names.slice(0, limit).join("、")} 等 ${names.length} 张`;
}

/**
 * 失败/待确认时首屏只放**第一句**结论，完整诊断进折叠区。
 *
 * 不整段丢弃：诊断里的"请重新发起任务""不会自动重试"是用户唯一的下一步，
 * 只是不该在首屏铺满。
 */
function firstSentence(text: string): string {
  const match = /^[^。；;]+[。；;]?/.exec(text.trim());
  return (match ? match[0] : text).trim();
}

/**
 * 一张检查点卡该说什么。
 *
 * 只描述**真实状态**：
 * - 完成只说"已删除 N 张工作表"，撤销方式在按钮旁说一次；
 * - 撤销成功才说"已撤销本次删除"，并只说"已恢复 N 张工作表"；
 * - 失败/部分完成保留原文，但放进折叠区，首屏不铺诊断；
 * - 打开副本之后只加一句"已打开恢复副本"，绝不说原表已还原、删除已撤销。
 */
export function describeWorksheetCheckpoint(input: {
  receipt?: WorksheetCheckpointReceipt | null;
  error?: string | null;
  deferred?: boolean;
  opened?: boolean;
  /** 本次撤销请求是否在途（按钮据此禁用，不用来改结论）。 */
  restoring?: boolean;
}): WorksheetCheckpointCardView {
  const { receipt, error, deferred, opened } = input;
  const restore = receipt?.restore ?? null;
  const restored = restore?.status === "completed";
  const shared = {
    deletedSheets: receipt ? [...receipt.deletedSheets] : [],
    remainingSheets: receipt ? [...receipt.remainingSheets] : [],
    openedNotice: opened ? WORKSHEET_CHECKPOINT_OPENED_NOTICE : null,
    restoredSheets: restore ? [...restore.restoredSheets] : [],
    undoHint: WORKSHEET_CHECKPOINT_UNDO_HINT,
    restoredNotice: restored ? WORKSHEET_CHECKPOINT_RESTORED_NOTICE : null,
    undoAvailable: false,
    undoReason: null as string | null
  };
  const view = (
    rest: Omit<WorksheetCheckpointCardView,
      "deletedSheets" | "remainingSheets" | "openedNotice" | "restoredSheets" |
      "undoHint" | "restoredNotice" | "undoAvailable" | "undoReason"> &
      Partial<Pick<WorksheetCheckpointCardView, "undoAvailable" | "undoReason">>
  ): WorksheetCheckpointCardView => ({ ...shared, ...rest });

  if (!receipt) {
    if (error) {
      return view({
        tone: "failed",
        title: "工作表删除未完成",
        summary: error,
        diagnostic: null,
        consentSheets: null,
        recoveryAvailable: false,
        canQueryStatus: false
      });
    }
    return view({
      tone: "unverified",
      title: "正在删除工作表",
      summary: "正在保存完整检查点并执行，请稍候…",
      diagnostic: null,
      consentSheets: null,
      recoveryAvailable: false,
      canQueryStatus: true
    });
  }

  if (deferred && receipt.status === "needs_backup_consent") {
    return view({
      tone: "unverified",
      title: "已暂缓本次删除",
      summary: "未授权额外备份，也没有删除任何工作表。",
      diagnostic: null,
      consentSheets: null,
      recoveryAvailable: false,
      canQueryStatus: false
    });
  }

  switch (receipt.status) {
    case "completed": {
      /*
       * 撤销的四种现场，措辞各不相同：
       * - 已撤销：终态，不再给入口；
       * - 上次冲突/部分/失败：给原因，允许重试（真撤销），不说"已撤销"；
       * - 从没撤销过：给入口；
       * - 没有归档：不给入口，指路「备份与恢复」。
       */
      if (restored) {
        return view({
          tone: "verified",
          title: "已撤销本次删除",
          summary: `已恢复 ${restore!.restoredSheets.length} 张工作表`,
          diagnostic: (restore!.limitations.length ? restore!.limitations.join("；") : null),
          consentSheets: null,
          recoveryAvailable: receipt.recoveryAvailable,
          canQueryStatus: false
        });
      }
      if (restore && (restore.status === "conflict" || restore.status === "failed" || restore.status === "partial")) {
        const retryable = restore.status !== "conflict" && receipt.recoveryAvailable;
        return view({
          tone: restore.status === "partial" ? "partial" : "failed",
          title: restore.status === "conflict" ? "暂不能撤销本次删除" : "本次删除尚未撤销",
          summary: firstSentence(restore.message),
          diagnostic: restore.message === firstSentence(restore.message) ? null : restore.message,
          consentSheets: null,
          recoveryAvailable: receipt.recoveryAvailable,
          canQueryStatus: true,
          undoAvailable: retryable,
          undoReason: retryable ? null : restore.message
        });
      }
      return view({
        tone: "verified",
        title: "工作表删除已完成",
        summary: `已删除 ${receipt.deletedSheets.length} 张工作表`,
        diagnostic: null,
        consentSheets: null,
        recoveryAvailable: receipt.recoveryAvailable,
        canQueryStatus: false,
        undoAvailable: receipt.recoveryAvailable,
        undoReason: receipt.recoveryAvailable
          ? null
          : "本次没有可用的执行前归档，无法在原工作簿撤销。"
      });
    }
    case "needs_backup_consent":
      return view({
        tone: "unverified",
        title: "请确认本地备份范围",
        summary: `完整恢复备份还需包含 ${shortList(receipt.backupSheets ?? [])}。尚未备份或删除。`,
        diagnostic: null,
        consentSheets: [...(receipt.backupSheets ?? [])],
        recoveryAvailable: false,
        canQueryStatus: false
      });
    case "partial_or_unknown":
      return view({
        tone: "partial",
        title: "删除结果待确认",
        summary: firstSentence(receipt.message),
        diagnostic: receipt.message === firstSentence(receipt.message) ? null : receipt.message,
        consentSheets: null,
        recoveryAvailable: receipt.recoveryAvailable,
        canQueryStatus: true,
        undoReason: receipt.recoveryAvailable
          ? "本次删除没有完整完成，不能原地撤销；可打开执行前副本核对。"
          : "本次没有可用的执行前归档，无法撤销。"
      });
    case "failed":
      return view({
        tone: "failed",
        title: "工作表删除未完成",
        summary: firstSentence(receipt.message),
        diagnostic: receipt.message === firstSentence(receipt.message) ? null : receipt.message,
        consentSheets: null,
        recoveryAvailable: receipt.recoveryAvailable,
        canQueryStatus: true,
        undoReason: receipt.recoveryAvailable
          ? "本次删除没有完整完成，不能原地撤销；可打开执行前副本核对。"
          : "本次没有可用的执行前归档，无法撤销。"
      });
    default:
      return view({
        tone: "unverified",
        title: receipt.status === "executing" ? "正在删除工作表" : "正在准备完整检查点",
        summary: "已在执行中；请勿重复发起，完成前会再次核对工作表列表。",
        diagnostic: null,
        consentSheets: null,
        recoveryAvailable: receipt.recoveryAvailable,
        canQueryStatus: true,
        undoReason: "删除尚未完成，暂时不能撤销。"
      });
  }
}

/** 「备份与恢复」里那一条恢复点的状态措辞。 */
export function worksheetCheckpointStatusLabel(
  status: WorksheetCheckpointStatus,
  restoreStatus?: WorksheetCheckpointRestore["status"] | null
): string {
  // 撤销状态优先：一条已经原地撤销过的记录，再说"已删除工作表"会误导。
  if (restoreStatus === "completed") return "已撤销删除";
  if (restoreStatus === "partial") return "撤销未完成";
  if (restoreStatus === "conflict") return "暂不能撤销";
  switch (status) {
    case "completed": return "已删除工作表";
    case "needs_backup_consent": return "等待备份授权";
    case "partial_or_unknown": return "结果待确认";
    case "failed": return "未完成";
    default: return "执行中";
  }
}

/** 恢复点是否需要用户先处理（用于排序/提示，不改变恢复能力）。 */
export function worksheetCheckpointNeedsAttention(status: WorksheetCheckpointStatus): boolean {
  return status === "partial_or_unknown" || status === "needs_backup_consent";
}
