/**
 * 操作反馈的归属与去重（Excel / WPS **共用判据**，两端只按引用转发）。
 *
 * 这里只回答三件确定性的事，宿主的数据模型与写入语义一律留在适配层：
 *
 *  1. 一条短反馈属于**哪一次执行**：只有 `executionId` 精确相等才算同一件事。
 *     找不到对应执行记录时**保留原消息行**，绝不"就近"塞给最近一张结果卡 ——
 *     那样会把 A 执行的「已撤销」挂到 B 执行的卡下面。
 *  2. 同一次执行的多条短反馈折成**一条紧凑摘要**（按 kind 去重、按副本身份合并原因）。
 *  3. 展示副本清理失败按**真实副本身份**去重：身份是本次会话随机生成的副本标识
 *     （sessionId + copySheetName / copyLocator），**不是**错误文本，也不是工作表序号。
 *
 * 业务结论（工作簿已恢复 / 已确认变更）与收尾结论（展示副本未清理）永不互相冒充：
 * 前者由 `kind` 表达，后者只由 `copy_cleanup_failed` + issues 表达。
 */

export type OperationFeedbackKind =
  | "confirmed"
  | "discarded"
  | "undone"
  | "copy_cleanup_failed";

/** 一份没清理掉的展示副本。identity 是唯一判据，label/reason 都只用于展示。 */
export type ReviewCopyCleanupIssue = {
  identity: string;
  label?: string;
  reason: string;
};

export type OperationFeedbackRecord = {
  executionId: string;
  kind: OperationFeedbackKind;
  text: string;
  issues?: readonly ReviewCopyCleanupIssue[];
};

/** 传给判据的最小消息形状：两端各自的 ChatMessage 都能直接喂进来。 */
export type OperationFeedbackMessage = {
  id: string;
  /** 这条消息本身携带的执行身份（结果卡、恢复记录、计划卡…）。 */
  executedExecutionId?: string | null;
  officeExecutionId?: string | null;
  changeReviewExecutionId?: string | null;
  folderOperationId?: string | null;
  /**
   * Excel 的 `VerificationReport` 没有执行身份，WPS 的同名对象有。
   * 用 `unknown` + 只读探测，避免为了"看起来统一"给两端各造一个假字段。
   */
  verification?: unknown;
  officeResume?: { operationId?: string | null } | null;
  folderResume?: { operationId?: string | null } | null;
  /** 非空 = 这是一条短反馈行，需要归属到某条执行记录下面。 */
  feedback?: OperationFeedbackRecord | null;
};

/**
 * 一条消息能证明自己属于哪些执行。**只认执行身份**，不认消息位置、计划标题、
 * "最近一次结果卡"或对话顺序 —— 那几样都会把不同 executionId 的操作误合并。
 */
function verificationExecutionId(value: unknown): string | null {
  if (!value || typeof value !== "object") return null;
  const id = (value as { executionId?: unknown }).executionId;
  return typeof id === "string" && id.length > 0 ? id : null;
}

export function operationFeedbackAnchorIds(message: OperationFeedbackMessage): string[] {
  const ids = [
    message.executedExecutionId,
    message.officeExecutionId,
    message.changeReviewExecutionId,
    message.folderOperationId,
    verificationExecutionId(message.verification),
    message.officeResume?.operationId,
    message.folderResume?.operationId
  ];
  const out: string[] = [];
  for (const id of ids) {
    if (typeof id !== "string" || id.length === 0) continue;
    if (out.includes(id)) continue;
    out.push(id);
  }
  return out;
}

export type OperationFeedbackGrouping = {
  /** 执行记录消息 id → 挂在它下面的短反馈（已按 kind / 副本身份收敛）。 */
  byAnchor: ReadonlyMap<string, readonly OperationFeedbackRecord[]>;
  /** 已经归属成功的短反馈消息 id：渲染时不再单独占一行。 */
  suppressed: ReadonlySet<string>;
  /** 找不到执行记录的短反馈消息 id：照旧单独渲染，绝不丢弃、绝不瞎挂。 */
  unattached: ReadonlySet<string>;
};

/**
 * 把短反馈挂到**对应的执行记录**上。
 *
 * - 从反馈行往前找最近一条「有执行身份、且身份包含本次 executionId」的消息；
 * - 同一 anchor 上的多条反馈合并成一条摘要（`compactOperationFeedback`）；
 * - 不同 executionId 的反馈不会因为"挨得近"或文案相同被合并；
 * - 反馈行自身即使也带别的执行身份，也不参与当 anchor（避免两条反馈互相挂）。
 */
export function groupOperationFeedback(
  messages: readonly OperationFeedbackMessage[]
): OperationFeedbackGrouping {
  const byAnchor = new Map<string, OperationFeedbackRecord[]>();
  const suppressed = new Set<string>();
  const unattached = new Set<string>();
  for (let index = 0; index < messages.length; index += 1) {
    const feedback = messages[index].feedback;
    if (!feedback || !feedback.executionId) continue;
    let anchor: string | null = null;
    for (let before = index - 1; before >= 0; before -= 1) {
      const candidate = messages[before];
      if (candidate.feedback) continue;
      if (operationFeedbackAnchorIds(candidate).includes(feedback.executionId)) {
        anchor = candidate.id;
        break;
      }
    }
    if (anchor === null) {
      unattached.add(messages[index].id);
      continue;
    }
    byAnchor.set(anchor, [...(byAnchor.get(anchor) ?? []), feedback]);
    suppressed.add(messages[index].id);
  }
  return { byAnchor, suppressed, unattached };
}

/**
 * 按**真实副本身份**去重。相同的错误文本来自两份不同副本 ⇒ 必须算两份；
 * 同一份副本被多个路径重试 ⇒ 只算一份，最新原因覆盖旧原因。
 */
export function dedupeReviewCopyCleanupIssues(
  issues: readonly ReviewCopyCleanupIssue[]
): ReviewCopyCleanupIssue[] {
  const byIdentity = new Map<string, ReviewCopyCleanupIssue>();
  for (const issue of issues) {
    const identity = typeof issue.identity === "string" ? issue.identity.trim() : "";
    // 拿不到真实身份时**不合并**：宁可多报一份，也不把两份副本认成一份。
    if (!identity) {
      byIdentity.set(`__unidentified__:${byIdentity.size}:${issue.reason}`, { ...issue, identity: "" });
      continue;
    }
    byIdentity.set(identity, { ...issue, identity, reason: issue.reason || "未说明原因" });
  }
  return [...byIdentity.values()];
}

/**
 * 副本清理的紧凑结论。只说"展示副本"这一件事，不提工作簿恢复成败，
 * 也不把多条相同原因拼成一长串重复文案。
 */
export function copyCleanupSummary(input: {
  executionId: string;
  issues: readonly ReviewCopyCleanupIssue[];
}): { count: number; issues: ReviewCopyCleanupIssue[]; headline: string; text: string } {
  const issues = dedupeReviewCopyCleanupIssues(input.issues);
  const count = issues.length;
  const headline = count > 0 ? `展示副本未清理 · ${count} 份` : "展示副本已清理";
  const text = count > 0
    ? `${headline}；工作簿的恢复结果另见本次执行记录。展开可看每份副本的原因。`
    : headline;
  return { count, issues, headline, text };
}

/**
 * 固定区域用的一句话：**不把多条相同原因原样串起来**。
 * 两份副本都因"窗口仍在"没关掉时写「窗口仍在（2 份）」，而不是「窗口仍在；窗口仍在」。
 */
export function copyCleanupNotice(summary: {
  count: number;
  headline: string;
  issues: ReviewCopyCleanupIssue[];
}): string {
  if (summary.count === 0) return "";
  const byReason = new Map<string, number>();
  for (const issue of summary.issues) {
    const reason = issue.reason || "未说明原因";
    byReason.set(reason, (byReason.get(reason) ?? 0) + 1);
  }
  const reasons = [...byReason.entries()]
    .map(([reason, count]) => (count > 1 ? `${reason}（${count} 份）` : reason))
    .join("；");
  return `${summary.headline}：${reasons}。工作簿的恢复结果不受影响，展开可看每份副本。`;
}

const KIND_LABEL: Record<OperationFeedbackKind, string> = {
  confirmed: "已确认变更",
  discarded: "已放弃并恢复",
  undone: "已撤销本次执行",
  copy_cleanup_failed: "展示副本未清理"
};

/**
 * 同一次执行的多条短反馈 → 一条紧凑摘要。
 * 同 kind 只留一条；`copy_cleanup_failed` 的副本原因按身份合并并给出数量。
 */
export function compactOperationFeedback(
  records: readonly OperationFeedbackRecord[]
): { records: OperationFeedbackRecord[]; text: string; issueCount: number } {
  const seen = new Set<OperationFeedbackKind>();
  const out: OperationFeedbackRecord[] = [];
  const cleanupIssues: ReviewCopyCleanupIssue[] = [];
  for (const record of records) {
    if (seen.has(record.kind)) {
      if (record.kind === "copy_cleanup_failed") cleanupIssues.push(...(record.issues ?? []));
      continue;
    }
    seen.add(record.kind);
    if (record.kind === "copy_cleanup_failed") cleanupIssues.push(...(record.issues ?? []));
    out.push({ ...record, issues: undefined });
  }
  if (seen.has("copy_cleanup_failed")) {
    const summary = copyCleanupSummary({ executionId: "", issues: cleanupIssues });
    const at = out.findIndex((record) => record.kind === "copy_cleanup_failed");
    out[at] = { ...out[at], text: summary.headline, issues: summary.issues };
  }
  const text = out.map((record) => record.text || KIND_LABEL[record.kind]).join(" · ");
  return { records: out, text, issueCount: cleanupIssues.length };
}

/**
 * 这条 `notice` 是不是已经由结构化 `copyCleanup` 面板表达过的同一件事。
 * 界面据此避免把同一句提示画两遍（一条纯文本 + 一份可展开清单）。
 */
export function isCopyCleanupNotice(
  notice: string | null | undefined,
  headline: string | null | undefined
): boolean {
  return Boolean(notice && headline && notice.startsWith(headline));
}

/** 固定区域用的紧凑摘要：不重复具体原因，只说"还有几件、最要紧的是哪件"。 */
export function operationFeedbackShortLabel(kind: OperationFeedbackKind): string {
  return KIND_LABEL[kind];
}
