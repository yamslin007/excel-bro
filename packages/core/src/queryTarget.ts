/**
 * "这次查哪几张表"的唯一判据。
 *
 * 为什么要单独有这一层：**授权集合只是允许访问的上限，不是本次查询范围**。
 * 2026-11-01 真机现场：用户只说「来源!A1:B5」，模型没在参数里绑表，执行器就把
 * 四张已授权表全读了一遍 —— 三张没有「列2」的表各推一条缺字段告警，结果卡还把
 * 那个不完整说成「扫描已达上限」。用户看到的是一份"覆盖了四张表"的结果。
 *
 * 所以这里把"目标表"变成**必须先算出来**的东西：算不出来就 refuse，而不是退回
 * 全部授权表。三条合法来源（优先级从高到低）：
 *
 * 1. `range` 带工作表名（`来源!A1:B5`）—— 区域自己就说明了表；
 * 2. `sheets`（本次要查的表名列表）—— 跨表查询显式列出来；
 * 3. `scope: "active"` —— 只查当前这张表。
 *
 * 剩下的情形只有"可查范围恰好只有一张表"才不需要绑定（没有歧义）。多张表又没绑定
 * ⇒ 报错并列出可选表名与三种写法。**不要**在这里按用户话里有没有出现表名去猜 ——
 * 那是另一层（模型）该做的判断，本地只认参数。
 */

/** 目标判定的入参。字段名与 `QueryTableArguments` 一一对应，不另造口径。 */
export interface QueryTargetRequest {
  /** 本次要查的工作表名（`arguments.sheets`）。 */
  sheets?: readonly string[] | null;
  /** `arguments.range` 里的工作表名（`来源!A1:B5` ⇒ `来源`），没有表名时为 null。 */
  rangeSheet?: string | null;
  /** `arguments.scope`。只区分 `active`，其余一律按"已选集合"处理。 */
  scope?: string | null;
  /** 本次允许访问的工作表（授权范围，或本回合实际读到的表）。 */
  candidates: readonly string[];
  /** 当前工作表名（`scope: "active"` 时的目标）。 */
  activeWorksheet?: string | null;
}

export type QueryTargetResolution =
  | { sheets: string[]; issue: null; issueKind: null }
  | {
      sheets: [];
      issue: string;
      /**
       * 让两端各自映射成自己的错误码（core 不认识宿主错误码）：
       * - `unbound`：没说查哪几张表 —— 参数问题，模型改一下就能过；
       * - `outside` / `conflict` / `empty`：越权或自相矛盾 —— 必须当成授权问题拒绝。
       */
      issueKind: "unbound" | "outside" | "conflict" | "empty";
    };

function uniqueNames(names: readonly string[] | null | undefined): string[] {
  const result: string[] = [];
  for (const name of names ?? []) {
    const text = typeof name === "string" ? name.trim() : "";
    if (text && !result.includes(text)) result.push(text);
  }
  return result;
}

function quoted(names: readonly string[]): string {
  return names.map((name) => `「${name}」`).join("、");
}

/** 未绑定目标表时的说明。把"可查哪些表"和"三种写法"都写出来，模型才改得对。 */
function unboundIssue(candidates: readonly string[]): string {
  return (
    `本次查询没有指明要查哪些工作表。授权范围只是可访问的上限，` +
    `不是本次查询范围（当前可查：${quoted(candidates)}）。` +
    `请用 sheets 列出本次要查的表（如 {"sheets": ["${candidates[0] ?? "来源"}"]}），` +
    `或把 range 写成「${candidates[0] ?? "来源"}!A1:B5」，` +
    `或用 scope: "active" 只查当前工作表。`
  );
}

/**
 * 算出本次查询的目标表。**只有这一个实现**：Excel 的执行器与 WPS 的取表入口都按引用
 * 转发，不各写一份"差不多的"收窄逻辑。
 */
export function resolveQueryTarget(
  request: QueryTargetRequest
): QueryTargetResolution {
  const candidates = uniqueNames(request.candidates);
  const explicit = uniqueNames(request.sheets);
  const rangeSheet =
    typeof request.rangeSheet === "string" ? request.rangeSheet.trim() : "";
  const active =
    typeof request.activeWorksheet === "string"
      ? request.activeWorksheet.trim()
      : "";

  // ① 区域带表名：区域自己就是绑定，`sheets` 只能与它一致，不能比它更宽。
  if (rangeSheet) {
    // `scope: active` 说的是"当前这张表"。区域指到另一张表时两者矛盾，
    // 按越权处理，不替调用方猜他想读哪张。
    if (request.scope === "active" && active && rangeSheet !== active) {
      return {
        sheets: [],
        issue: `区域指定的是「${rangeSheet}」，但 scope 是 active（当前表「${active}」）。两者矛盾：请把 scope 改为 selected，或让区域指向当前表。`,
        issueKind: "conflict"
      };
    }
    if (explicit.length > 0 && !explicit.includes(rangeSheet)) {
      return {
        sheets: [],
        issue:
          `区域指定的工作表是「${rangeSheet}」，与本次的 sheets（${quoted(explicit)}）不一致。` +
          `两者只保留一个：要么去掉 sheets，要么把区域写成 ${explicit[0]}!A1:B5。`,
        issueKind: "conflict"
      };
    }
    if (!candidates.includes(rangeSheet)) {
      return {
        sheets: [],
        issue: `区域指定的工作表「${rangeSheet}」不在本次可查询的范围内（可查：${quoted(candidates) || "无"}）。`,
        issueKind: "outside"
      };
    }
    return { sheets: [rangeSheet], issue: null, issueKind: null };
  }

  // ② 显式列出的表：逐个都必须在可查范围内，不静默丢掉越界的名字。
  if (explicit.length > 0) {
    const outside = explicit.filter((name) => !candidates.includes(name));
    if (outside.length > 0) {
      return {
        sheets: [],
        issue: `本次要查的工作表${quoted(outside)}不在可查询的范围内（可查：${quoted(candidates) || "无"}）。`,
        issueKind: "outside"
      };
    }
    return { sheets: explicit, issue: null, issueKind: null };
  }

  // ③ 当前表。
  if (request.scope === "active") {
    if (!active) {
      return {
        sheets: [],
        issue: "scope 为 active 但没有可用的当前工作表。",
        issueKind: "empty"
      };
    }
    if (candidates.length > 0 && !candidates.includes(active)) {
      return {
        sheets: [],
        issue: `当前工作表「${active}」不在可查询的范围内（可查：${quoted(candidates)}）。`,
        issueKind: "outside"
      };
    }
    return { sheets: [active], issue: null, issueKind: null };
  }

  // ④ 只有一张可查表时无歧义，不必强迫调用方重复一遍表名。
  if (candidates.length === 1) {
    return { sheets: [candidates[0]], issue: null, issueKind: null };
  }
  if (candidates.length === 0) {
    return { sheets: [], issue: "没有可查询的工作表。", issueKind: "empty" };
  }

  // ⑤ 多张表 + 没有绑定 ⇒ refuse。退回"全部已授权表"正是这次要根除的行为。
  return { sheets: [], issue: unboundIssue(candidates), issueKind: "unbound" };
}
