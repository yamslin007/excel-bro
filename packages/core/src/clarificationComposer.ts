/**
 * 「澄清回答」的提交判据（Excel / WPS 共用，纯函数）。
 *
 * 这一轮把快捷回答从**消息里**搬到**输入框上方**，并改成"点选只选中、按发送才提交"。
 * 随之而来的是三个必须只有一份实现的问题：
 *
 * 1. **这次点击到底在回答哪道题？** 问题身份 = `conversationId + 消息 id + 澄清 id`。
 *    只带澄清 id 是不够的：换对话、换任务之后，同一个 id 可能还躺在列表里，
 *    按下去就会把答案提交给一条已经作废的问题。
 * 2. **什么时候能提交？** 选项一个没选、自由输入为空、多字段有必填没填 —— 都不许提交。
 *    这条必须是纯函数：两端各写一遍就一定会有一端把"没填完"当成"填完了"。
 * 3. **提交的是什么？** 选项走结构化 resolution（保住 `resolutionMode`），
 *    自由输入走原文，多字段走"标签：值"。三者不许互相冒充 ——
 *    把自由输入当选项提交，等于替用户选了一个他没点的选项。
 *
 * 三条不变量：
 * A. **答案不等于授权。** 本模块只产出"要提交什么"，不碰授权范围；把答案写进
 *    `commitAuthorizedScope` 是另一件事（`clarificationAnswers` 里已划清）。
 * B. **旧问题不许重复提交。** 提交前必须用 `clarificationSubmission` 的绑定校验
 *    对照**当前**消息状态；结论不是 `ok` 就一个字节都不发。
 * C. **不猜。** 选项答案只来自用户点过的选项 id；没有匹配就是没有，绝不按文案相似度匹配。
 */

/** 一道澄清问题的稳定身份。三段都要，缺一段就可能串到别的任务上。 */
export interface ClarificationBinding {
  conversationId: string;
  /** 承载这道题的那条消息 id（不是任务 id —— 一条任务可以有多个回复段）。 */
  messageId: string;
  clarificationId: string;
}

export function clarificationBindingKey(binding: ClarificationBinding): string {
  return `${binding.conversationId}\u0000${binding.messageId}\u0000${binding.clarificationId}`;
}

/** 本地草稿：绑定 + 用户已选/已填的内容。 */
export interface ClarificationDraft {
  binding: ClarificationBinding;
  /** 已选中的选项 id（顺序 = 点击顺序）。 */
  selected: string[];
  /** 自由输入。 */
  freeText: string;
  /** 多字段草稿（键 = 字段 id）。 */
  fields: Record<string, string>;
}

export function emptyClarificationDraft(binding: ClarificationBinding): ClarificationDraft {
  return { binding, selected: [], freeText: "", fields: {} };
}

/**
 * 草稿还属于当前这道题吗。
 *
 * 换对话、换任务、这道题已经作废 —— 三种情况都必须返回 `false`，
 * 让调用方把草稿丢掉。**不刷新成新绑定继续用**：那等于把用户写给旧问题的答案
 * 悄悄送给新问题。
 */
export function draftBelongsTo(
  draft: ClarificationDraft | null | undefined,
  binding: ClarificationBinding | null | undefined
): boolean {
  if (!draft || !binding) return false;
  return clarificationBindingKey(draft.binding) === clarificationBindingKey(binding);
}

/** 最小选项形状（宿主契约里的 `IntentOption` 比这更宽）。 */
export interface ClarificationOptionLike {
  id: string;
  label: string;
  /** 展示用说明；不参与提交内容。 */
  description?: string;
  resolution: string;
  resolutionMode?: "append" | "replace";
}

/** 最小字段形状。 */
export interface ClarificationFieldLike {
  id: string;
  label: string;
}

/**
 * 点一下选项之后新的选中集合。
 *
 * - 单选（默认）：选中它，取消其它；再点同一个 ⇒ 取消（允许反悔）。
 * - 多选：切换该项，顺序保留点击顺序。
 *
 * 为什么不排序：顺序是用户的选择过程，提交时按这个顺序拼答案；
 * 排序会让"我先选 A 再选 B"和"先 B 后 A"变成同一件事，而它们不是。
 */
export function toggleClarificationOption(
  selected: readonly string[],
  optionId: string,
  multiple = false
): string[] {
  const id = optionId.trim();
  if (!id) return [...selected];
  const index = selected.indexOf(id);
  if (multiple) {
    return index >= 0 ? selected.filter((item) => item !== id) : [...selected, id];
  }
  if (index >= 0) return [];
  return [id];
}

/**
 * 这道题**允不允许**多选。
 *
 * 目前协议里的 `IntentClarification` 没有多选标记，所以恒为 `false`：
 * 多选机制（`toggleClarificationOption(…, true)`）已经实现并有用例，
 * 但在协议补上标记之前**不可达**。这里不猜"选项多就是多选" ——
 * 用户点两个本来互斥的选项，程序替他拼成一个答案，是比少一个功能更糟的事。
 */
export function clarificationAllowsMultiple(
  clarification: { multiSelect?: unknown } | null | undefined
): boolean {
  return clarification?.multiSelect === true;
}

/** 提交内容三选一。`option` 带结构化 resolution，不把标签当答案。 */
export type ClarificationSubmission =
  | { ok: true; kind: "option"; options: ClarificationOptionLike[]; answer: string }
  | { ok: true; kind: "fields"; answer: string }
  | { ok: true; kind: "free"; answer: string }
  | { ok: false; reason: string; missing: string[] };

/**
 * 决定这次"发送"到底提交什么。**唯一判据**。
 *
 * 优先级有意如此：**多字段** > **选项** > **自由输入**。
 *  - 多字段是结构化表单，没填完就不许提交（必填验证保留）；
 *  - 选项是用户点过的，比他在自由输入里打的字更明确；
 *  - 自由输入是兜底，也是唯一允许和选项**同时**存在的补充（拼在答案后面）。
 */
export function clarificationSubmission(input: {
  options?: readonly ClarificationOptionLike[] | null;
  selected?: readonly string[] | null;
  freeText?: string | null;
  fields?: readonly ClarificationFieldLike[] | null;
  fieldValues?: Record<string, string> | null;
}): ClarificationSubmission {
  const fields = input.fields ?? [];
  const values = input.fieldValues ?? {};
  if (fields.length > 0) {
    const missing = fields.filter((field) => !values[field.id]?.trim());
    if (missing.length > 0) {
      return {
        ok: false,
        reason: `还有 ${missing.length} 项没有填：${missing.map((field) => field.label).join("、")}`,
        missing: missing.map((field) => field.id)
      };
    }
    return {
      ok: true,
      kind: "fields",
      answer: fields.map((field) => `${field.label}：${values[field.id].trim()}`).join("\n")
    };
  }

  const freeText = (input.freeText ?? "").trim();
  const options = input.options ?? [];
  const selectedIds = input.selected ?? [];
  /*
   * 只认**用户点过的 id**。找不到对应选项就当没选 ——
   * 按标签做相似度匹配会把一个已经不存在的选项"复活"成一个答案。
   */
  const picked = selectedIds
    .map((id) => options.find((option) => option.id === id))
    .filter((option): option is ClarificationOptionLike => Boolean(option));

  if (picked.length > 0) {
    const labels = picked.map((option) => option.label);
    return {
      ok: true,
      kind: "option",
      options: picked,
      // 选项 + 补充说明：两句都要，否则用户的补充会被静默丢掉。
      answer: freeText ? `${labels.join("、")}\n补充：${freeText}` : labels.join("、")
    };
  }

  if (freeText) return { ok: true, kind: "free", answer: freeText };

  return {
    ok: false,
    reason: fields.length > 0 ? "请填写必填项后再发送。" : "请选择一项，或写下你的想法。",
    missing: []
  };
}

/**
 * 提交前对照**当前**问题状态的最后一道闸。
 *
 * `resolveClarification` 自己也有 `status !== "pending"` 判断，但那是在
 * 拿到 message 对象之后。这里先在"我手上这份草稿还算不算数"这一层拦一次：
 * 定时器、异步回读、用户切换对话都可能让草稿变旧，等真的发出去再发现就晚了。
 */
export function clarificationSubmissionAllowed(input: {
  draft: ClarificationDraft | null | undefined;
  binding: ClarificationBinding | null | undefined;
  /** 当前这道题的状态（`pending` 才允许）。 */
  status: string | null | undefined;
  busy?: boolean;
}): { ok: true } | { ok: false; reason: string } {
  if (!input.binding) return { ok: false, reason: "当前没有待回答的问题。" };
  if (input.status !== "pending") {
    return { ok: false, reason: "这道题已经处理过了，不会再提交一次。" };
  }
  if (!draftBelongsTo(input.draft, input.binding)) {
    return { ok: false, reason: "这个问题已经不属于当前任务，请重新选择。" };
  }
  if (input.busy) return { ok: false, reason: "正在处理上一步，稍后再发送。" };
  return { ok: true };
}
