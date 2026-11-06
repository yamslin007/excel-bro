/**
 * 结果"为什么不完整"的唯一出处。
 *
 * 之前只有一个布尔：`complete = warnings.length === 0`，而结果卡把 `complete === false`
 * 一律写成「结果不完整（扫描已达上限）」—— 三种完全不同的原因共用一句最吓人的话。
 * 2026-11-01 真机现场就是这样：三张表缺「列2」被说成"扫描已达上限"。
 *
 * 现在事实由**产生它的那段代码**写下来（`reason`），文案只负责把事实说清楚。因此
 * 这里不提供"猜原因"的入口：没有 reason 就说明结果是完整的。
 *
 * 为什么只有两种 kind —— 另外两种根本到不了结果卡：
 * - **读取失败**：任何一张表读不出来，整条请求按错误返回（`TOOL_LIMIT_EXCEEDED` /
 *   越权 / 宿主异常），**不会**产出只含其余表的"部分结果"。所以它没有卡片，
 *   也不该在这里编码。
 * - **扫描上限**：超限在执行前就抛错（`executeQueryTableTool` 的限额检查，
 *   见 `config/capabilities.json`），同样不产出部分结果。
 * 谁要是以后真做出"部分扫描"，必须**同时**在这里加 kind 与其唯一的生产者，
 * 不允许先把文案写上去。
 */

/** 真正让结果内容不完整的原因。加新 kind 必须同时加一个真实的生产者。 */
export type QueryIncompleteReason =
  | {
      /** 哪张表缺字段（该表整表跳过）。 */
      kind: "missing_field";
      sheet: string;
      /** 该表里找不到的字段（请求要用的那些）。 */
      fields: string[];
    }
  | {
      /** 哪张表没有数据行（文件夹模式逐文件读取时才可能出现）。 */
      kind: "empty_sheet";
      sheet: string;
    };

/**
 * 展示截断：数据是完整的，只是按 `limit` 只显示了前若干行。
 * 与 `complete` 分开：截断**不是**不完整，两者的文案和行为都不一样。
 */
export interface QueryTruncation {
  shownRows: number;
  totalRows: number;
  limit: number;
}

export function describeIncompleteReason(
  reason: QueryIncompleteReason
): string {
  if (reason.kind === "empty_sheet") return `「${reason.sheet}」没有数据行`;
  return `「${reason.sheet}」没有字段：${reason.fields.join("、")}`;
}

/** 一句话说清缺什么、少了哪几张表；无原因时返回 null。 */
export function describeIncompleteReasons(
  reasons: readonly QueryIncompleteReason[]
): string | null {
  if (reasons.length === 0) return null;
  const sheets = [...new Set(reasons.map((reason) => reason.sheet))];
  const details = reasons.map(describeIncompleteReason).join("；");
  return `${details}。这些表没有被计入结果（共 ${sheets.length} 张）。`;
}

export function describeTruncation(truncation: QueryTruncation): string {
  return `仅显示前 ${truncation.shownRows} 行，共 ${truncation.totalRows} 行（limit=${truncation.limit}）。`;
}

/** 结果卡与模型回执共用的完整性事实。`complete` 只由 reason 列表决定。 */
export interface QueryCompleteness {
  complete: boolean;
  reasons: QueryIncompleteReason[];
  truncated: boolean;
  /** 有截断时给出计数，否则为 null。 */
  truncation: QueryTruncation | null;
}

export function completenessOf(props: {
  reasons?: readonly QueryIncompleteReason[];
  truncation?: QueryTruncation | null;
}): QueryCompleteness {
  const reasons = [...(props.reasons ?? [])];
  const truncation = props.truncation ?? null;
  return {
    complete: reasons.length === 0,
    reasons,
    truncated: truncation !== null,
    truncation
  };
}
