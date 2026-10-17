/**
 * 计划的**实际来源依赖**：这份计划的值实际是从哪些工作表读出来的。
 *
 * 为什么单独建这一层：授权范围、实际来源依赖、写入目标此前被挤在同一个字段上
 * （`sourceFingerprintSheets` 被填成本次授权范围），于是"改来源表"的恢复禁令
 * 必然命中每一个正常计划 —— 登记断点恢复成了死代码。三者口径必须分开：
 *
 * - 授权范围：`authorizedSheetNames`，由用户确认，是不可扩大的边界；
 * - 实际来源依赖：本模块产出，回答"这份计划读了谁"；
 * - 写入目标：`plan.actions[].sheet`，本来就精确。
 *
 * 信任边界（关键）：
 * - 记账只发生在**确定性执行器读完之后**（`DataToolResult.sourceSheets` /
 *   数据登记实际读取成功的来源）、本地规则读取、已保存工具模板声明这三处；
 *   没有任何入口接受"模型说它读了谁"；
 * - 因此**不**用"授权范围减去写入目标"来反推来源 —— 那是伪造，不是记账；
 * - 计划里带的 `sourceDependencySheets`（可能由服务端或模型给出）在装配时被
 *   `attestSourceDependency` 覆盖或剥离：本端没有记账 = 依赖不确定 = 不登记恢复。
 *
 * **双端同源**：Excel（Office.js）与 WPS（JSAPI）各有一份桥宿主适配器，但
 * "读了谁"的记账口径必须逐字一致 —— 放在 core 里两端共用同一份实现，避免
 * 一端改了记账规则、另一端还按旧规则登记恢复（症状是同一份计划在两端一个
 * 可恢复、一个不可恢复）。
 */

/**
 * 本模块只关心计划上的一个字段，别的字段原样透传。
 *
 * 刻意**不** import `AnalysisPlan`：那是 Excel 加载项 `contracts.ts` 里的完整
 * 类型，core 里没有、也不该有（core 不认识双端的完整计划形状）。用一个结构化
 * 最小接口，两端都能直接用，且不会把"计划必须有 actions/acceptanceCriteria"
 * 这类宿主侧约束拖进 core。
 */
export interface SourceDependencyPlan {
  sourceDependencySheets?: string[];
}

/** 可信来源的出处，用于如实说明"谁证明了这件事"。 */
export type SourceDependencyOrigin =
  | "tool-execution-record"
  | "deterministic-local-read"
  | "saved-tool-template";

export interface SourceDependencyRecord {
  origin: SourceDependencyOrigin;
  sheets: readonly string[];
}

/** 一轮（桥的 taskId）内实际读过的工作表，按并集累积。 */
const turnReads = new Map<string, { sheets: string[]; origin: SourceDependencyOrigin }>();

/** 只保留最近若干轮：这些记账只对"当轮生成、当轮执行"的计划有用。 */
const MAX_TRACKED_TURNS = 32;

function normalizeSheets(sheets: readonly string[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const raw of sheets) {
    const name = typeof raw === "string" ? raw.trim() : "";
    if (!name) continue;
    const key = name.toLocaleLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(name);
  }
  return result;
}

/**
 * 记账：本机确定性地读过这些工作表。
 *
 * **唯一合法的调用方是"刚真的读过表"的代码**（桥上的确定性执行器）。
 * 同一个 taskId 多次读取按**并集**累积 —— 计划可能依赖其中任意一次的结果，
 * 取并集才是诚实的答案（取交集会漏掉真实依赖）。
 */
export function recordTrustedToolReads(
  turnId: string,
  sheets: readonly string[]
): void {
  const key = typeof turnId === "string" ? turnId.trim() : "";
  if (!key) return;
  const incoming = normalizeSheets(sheets);
  if (incoming.length === 0) return;
  const existing = turnReads.get(key);
  const merged = normalizeSheets([...(existing?.sheets ?? []), ...incoming]);
  // 重插到队尾：Map 的插入顺序就是 LRU 序，淘汰最久未更新的那一轮。
  turnReads.delete(key);
  turnReads.set(key, { sheets: merged, origin: "tool-execution-record" });
  while (turnReads.size > MAX_TRACKED_TURNS) {
    const oldest = turnReads.keys().next().value;
    if (oldest === undefined) break;
    turnReads.delete(oldest);
  }
}

/** 取该轮的可信读取记录；没读过任何表就是 null（不是空数组）。 */
export function trustedToolReads(turnId: string): SourceDependencyRecord | null {
  const entry = turnReads.get(typeof turnId === "string" ? turnId.trim() : "");
  return entry ? { origin: entry.origin, sheets: [...entry.sheets] } : null;
}

/** 清空全部记账（换簿、测试复位）。 */
export function resetTrustedSourceReads(): void {
  turnReads.clear();
}

/**
 * 把可信记录写成计划的 `sourceDependencySheets`。
 *
 * 没有记录时**删除**该字段（而不是留空数组）：空数组会被读成"依赖为空"，
 * 而事实是"不知道依赖谁"。两者在恢复登记里的结论必须不同。
 */
export function attestSourceDependency<T extends SourceDependencyPlan>(
  plan: T,
  record: SourceDependencyRecord | null
): T {
  const sheets = record ? normalizeSheets(record.sheets) : [];
  const next = { ...plan };
  if (sheets.length === 0) {
    delete next.sourceDependencySheets;
    return next;
  }
  next.sourceDependencySheets = sheets;
  return next;
}

/**
 * **结果计划**的来源依赖：服务端结果来源 ∩ 本端可信记账。
 *
 * 为什么不能直接用 `attestSourceDependency`：那一条只认本端的**读取并集**，
 * 会把"本轮读过、但这份结果没用到的表"也算成依赖。反过来说，只认服务端
 * 声明的来源，又变成"采信远端自报"，本端没有任何读过的凭据。
 *
 * 所以两侧都要，且取**交集**：
 *
 * - 服务端 `sourceDependencySheets`（结果编译时写入的 `item.source_sheets`）
 *   回答"这份结果**据以算出**的那些表"—— 由可信协调器按**实际交给执行器的
 *   输入句柄**认定，不是 prepare 阶段的来源并集；
 * - 本端记账（`recordTrustedToolReads`）回答"这台机器**真的读过**哪些表"。
 *
 * 只有同时成立的表才是可信的依赖。交集为空 ⇒ 删字段（依赖不确定），
 * 不给出一个空的"依赖"—— 那会被读成"这份计划不依赖任何来源"。
 *
 * 服务端没给来源（历史登记项）时退回**单侧**记账：本端确实读过，就照实记；
 * 这比"删字段"保留了更多真实信息，且不引入任何未经证明的表。
 */
export function attestResultSourceDependency<T extends SourceDependencyPlan>(
  plan: T,
  records: readonly SourceDependencyRecord[]
): T {
  const declared = normalizeSheets(plan.sourceDependencySheets ?? []);
  const read = normalizeSheets(records.flatMap((record) => [...record.sheets]));
  const trusted =
    declared.length === 0
      ? read
      : read.filter((name) =>
          declared.some((item) => item.toLocaleLowerCase() === name.toLocaleLowerCase())
        );
  const next = { ...plan };
  if (trusted.length === 0) {
    delete next.sourceDependencySheets;
    return next;
  }
  next.sourceDependencySheets = trusted;
  return next;
}
