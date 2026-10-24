/** Stable provider-limit codes. Never infer a depleted balance from HTTP 429 alone. */
export type ModelUsageErrorCode =
  | "MODEL_QUOTA_EXHAUSTED"
  | "MODEL_CONTEXT_LIMIT"
  | "MODEL_OUTPUT_LIMIT"
  | "MODEL_BUDGET_LIMIT"
  | "MODEL_TURN_LIMIT";

const notices: Record<ModelUsageErrorCode, string> = {
  MODEL_QUOTA_EXHAUSTED: "模型 API 余额不足或调用额度已用尽。请在模型服务商处检查余额、调用额度或套餐，恢复后重试，或在模型设置中切换可用连接。",
  MODEL_CONTEXT_LIMIT: "本次请求超过模型的上下文长度上限。请缩短对话或拆分任务后重试。",
  MODEL_OUTPUT_LIMIT: "本次模型输出达到长度上限，未获得完整结果。请拆分任务或调整模型输出上限后重试。",
  MODEL_BUDGET_LIMIT: "本次任务达到配置的模型费用上限。请检查任务预算设置，或拆分任务后重试。",
  MODEL_TURN_LIMIT: "本次任务达到模型调用轮数上限。请拆分任务后重试。"
};

/** Both hosts use the same notice for JSON, SSE, and status-recovery errors.
 * A model failure is not evidence about a local workbook write or its receipt.
 */
export function modelUsageNotice(code: string): string | null {
  if (!Object.prototype.hasOwnProperty.call(notices, code)) return null;
  return `${notices[code as ModelUsageErrorCode]} 表格是否已写入或回读成功，请以执行结果卡为准。`;
}
