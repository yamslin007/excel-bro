/** This predicate routes a whitelist plan; it never interprets user text. */
export function isWorksheetDeletion(plan: { actions: readonly { type: string }[] }): boolean {
  return plan.actions.length > 0 && plan.actions.every(action => action.type === "deleteWorksheet");
}
export function worksheetDeletionIssue(plan: { actions: readonly { type: string; sheet?: string }[] }, authorized: readonly string[], all: readonly string[]): string | null {
  if (!isWorksheetDeletion(plan)) return "删表检查点只支持纯删表计划，请将其他修改拆成下一步。";
  if (new Set(authorized).size !== authorized.length || authorized.some(name => !all.includes(name))) {
    return "已授权工作表不在当前工作簿中，请重新核对数据范围。";
  }
  const targets = plan.actions.map(action => action.sheet);
  if (targets.some(name => !name || !authorized.includes(name)) || new Set(targets).size !== targets.length) return "删除目标重复或超出授权范围。";
  if (all.every(name => targets.includes(name))) return "至少需要保留一张工作表，请明确要保留的表。";
  return null;
}

/**
 * 删表会让**已授权工作表变少**，这是本次操作的预期结果，不该被目标守卫当成
 * "用户改了授权范围"再提示一次"尚未写入"。
 *
 * 但"预期"必须被证明，不能靠"刚删完就放过"。这里给出的唯一判据是：
 *
 * 1. 删除前后是**同一个工作簿身份**（身份缺失一律不认）；
 * 2. 回执里的剩余表与宿主**现在真读到的**表集合逐名相同；
 * 3. 回执里的删除清单恰好等于"已授权表 − 现在还在的表"（没有多删、没有少删、
 *    没有别的已授权表悄悄消失）；
 * 4. 收窄后仍有已授权表可用（全部删光时不给"自动收缩"，仍走正常重新确认）。
 *
 * 任何一条不成立都返回 `ok: false` —— 调用方必须放弃这次静默同步，让原有的
 * 范围变更提示照常出现。**这个函数只做减法**：返回的 scope 永远是授权表的子集。
 */
export type WorksheetDeletionScopeReconciliation =
  | { ok: true; scope: string[] }
  | {
    ok: false;
    reason:
      | "identity_missing"
      | "workbook_changed"
      | "receipt_mismatch"
      | "nothing_authorized_left";
  };

function uniqueNames(names: readonly string[]): string[] | null {
  const trimmed = names.map(name => name.trim());
  if (trimmed.some(name => !name)) return null;
  return new Set(trimmed).size === trimmed.length ? trimmed : null;
}

function sameNames(a: readonly string[], b: readonly string[]): boolean {
  if (a.length !== b.length) return false;
  const set = new Set(b);
  return a.every(name => set.has(name));
}

/**
 * 工作簿目录里可能混着应用自己的系统表（`#EB_RULES` 等）。
 *
 * 两端的目录读数口径不同：Excel 的目录快照会把它们过滤掉，WPS 的读取器会把它们
 * 列出来。回执里的剩余表是**全簿**口径。比较之前统一按系统表名过滤，否则
 * "有 `#EB_RULES` 的工作簿永远对不上"，这条同步就成了只在没有系统表的簿里生效。
 */
function withoutSystemSheets(names: readonly string[], isSystemSheet: (name: string) => boolean): string[] {
  return names.filter(name => !isSystemSheet(name));
}

export function reconcileWorksheetDeletionScope(input: {
  beforeIdentity: string | null;
  afterIdentity: string | null;
  authorizedSheets: readonly string[];
  receipt: { deletedSheets: readonly string[]; remainingSheets: readonly string[] } | null;
  actualSheets: readonly string[];
  isSystemSheet?: (name: string) => boolean;
}): WorksheetDeletionScopeReconciliation {
  const { beforeIdentity, afterIdentity, authorizedSheets, receipt, actualSheets } = input;
  const isSystemSheet = input.isSystemSheet ?? (() => false);
  if (!beforeIdentity || !afterIdentity) return { ok: false, reason: "identity_missing" };
  if (beforeIdentity !== afterIdentity) return { ok: false, reason: "workbook_changed" };
  if (!receipt) return { ok: false, reason: "receipt_mismatch" };

  const authorized = uniqueNames(authorizedSheets);
  const deleted = uniqueNames(receipt.deletedSheets);
  const remaining = uniqueNames(receipt.remainingSheets);
  const actualRaw = uniqueNames(actualSheets);
  // 重名/空名说明读数本身不可信，不能拿它做"预期收窄"的证据。
  if (!authorized || !deleted || !remaining || !actualRaw) return { ok: false, reason: "receipt_mismatch" };

  const actual = withoutSystemSheets(actualRaw, isSystemSheet);
  const remainingVisible = withoutSystemSheets(remaining, isSystemSheet);
  if (!deleted.length) return { ok: false, reason: "receipt_mismatch" };
  // 2. 回执与真实剩余集合必须一致；多一张（外部新增）或少一张（外部删除）都不认。
  if (!sameNames(remainingVisible, actual)) return { ok: false, reason: "receipt_mismatch" };

  const actualSet = new Set(actual);
  // 3a. 回执声称删掉的表必须都在本次授权范围内，而且现在真的不在了。
  if (deleted.some(name => !authorized.includes(name) || actualSet.has(name))) {
    return { ok: false, reason: "receipt_mismatch" };
  }
  // 3b. 授权表里除删除清单之外的每一张都必须还在；少一张说明有别的变化。
  const survivors = authorized.filter(name => actualSet.has(name));
  if (survivors.length !== authorized.length - deleted.length) {
    return { ok: false, reason: "receipt_mismatch" };
  }
  // 4. 全部授权表都被删掉时不自动收缩：没有可用授权，只能让用户重新选范围。
  if (!survivors.length) return { ok: false, reason: "nothing_authorized_left" };
  return { ok: true, scope: survivors };
}
