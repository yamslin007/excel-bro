/**
 * 「澄清的答案」与「授权动作」是两种状态，不得互相冒充（L2，宿主无关）。
 *
 * 历史上它们被塞进同一个 `clarification.answer` 字段：点一下"重新确认数据
 * 范围"，卡片就被写成"已回答：已重新确认数据范围并继续"。后果有两层 ——
 *
 * 1. 用户真正的回答被那句套话**覆盖**了，重开后顺着"已回答"继续，等于
 *    拿着一个不存在答案往下走；
 * 2. 授权被记成"答完题"，于是写入资格看起来已经具备 —— 可重新确认范围
 *    只证明了"你现在授权这几张表"，没有证明"你同意写入那个旧对象"。
 *
 * 本模块把这两件事分开：
 *
 * - **答案**是用户对那道题的回答，只能由用户填写/选择产生；
 * - **授权**是一次独立的动作，记在授权版本上（`authorizationRevision`），
 *   它**不产生答案**，也不把卡片推进到"已回答"。
 */

/** 卡片的答案侧状态。只取必要字段，避免本模块绑定到某一端的完整类型。 */
export interface ClarificationAnswerCard {
  status: string;
  answer?: string | null;
}

/**
 * 重新授权之后，澄清卡应该怎么变。
 *
 * 唯一允许的变化是"把**已有**的真实答案标记为已作答"。
 * 没有真实答案时**原样返回** —— 重新授权不等于回答了那道题，
 * 把卡片推进成"已回答"就是在伪造证据。
 */
export function clarificationAfterReauthorization<
  T extends ClarificationAnswerCard
>(clarification: T | undefined): T | undefined {
  if (!clarification) return clarification;
  if (clarification.answer && clarification.answer.trim()) {
    return { ...clarification, status: "resolved" } as T;
  }
  return clarification;
}

/**
 * 重新授权给用户看的那句话。
 *
 * 它是**系统说明**，不是答案：调用方只能把它 `append` 成一条系统消息，
 * **不得**写进 `clarification.answer`。
 */
export const REAUTHORIZATION_NOTICE =
  "已重新确认当前数据范围（这是一次授权动作，不是对那道题的回答）。" +
  "原需求与你已确认的信息都还在，下面按它们重新读取并生成预览；" +
  "写入前仍会再让你确认一次。";

/** 一条用户确认。带**稳定问题身份**，同一道题改答案时才能定位到旧的。 */
export interface KeyedConfirmation {
  /** 稳定的问题身份。同一道题必须复用它，不能靠文案匹配。 */
  questionId?: string;
  /** 展示用标签。 */
  label: string;
  /** 进模型的确认内容本体。 */
  value: string;
}

/**
 * 合并一条新确认：**同一 `questionId` 的后答替换先答**。
 *
 * 为什么必须替换而不是追加：追加会把改过的旧答案原样留在提示词里，
 * 模型同时看到"按 A 处理"和"按 B 处理"两条互相冲突的确认，只能自己挑一个
 * —— 那正是用户刚刚明确做过、不该再由实现替他做的决定。
 *
 * 没有 `questionId` 时无从判断是不是同一道题，保守**追加**：凭猜测替换
 * 会丢掉一条真实答案，比多带一条旧确认更糟。
 */
export function mergeKeyedConfirmation(
  existing: readonly KeyedConfirmation[],
  next: KeyedConfirmation
): KeyedConfirmation[] {
  const questionId = next.questionId?.trim();
  if (!questionId) return [...existing, next];
  const index = existing.findIndex(
    (item) => item.questionId?.trim() === questionId
  );
  if (index < 0) return [...existing, next];
  const merged = [...existing];
  merged[index] = next;
  return merged;
}

/**
 * 把已确认信息渲染成提示词。
 *
 * 顺序 = 确认发生的顺序；同一道题只出现**最后一次**的答案（见
 * `mergeKeyedConfirmation`），所以这里不会同时出现两条互相冲突的确认。
 */
export function renderConfirmedPrompt(
  originalPrompt: string,
  confirmations: readonly KeyedConfirmation[]
): string {
  if (confirmations.length === 0) return originalPrompt;
  return [
    originalPrompt,
    ...confirmations.map((item) => `补充确认：${item.value}`)
  ].join("\n");
}
