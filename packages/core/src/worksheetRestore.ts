/**
 * 原地撤销之后，授权范围能不能"就地认下来"（Excel / WPS 读同一份）。
 *
 * 撤销会把工作表**放回**工作簿，于是目录比授权时多了几张表。若不处理，下一轮
 * 目标守卫会把这当成"用户改了范围 / 有人动了工作簿"，弹一句"尚未写入"——
 * 与刚刚撤销成功的事实直接冲突。
 *
 * 但"放过"必须被证明，而不是"刚撤销完就放过"。这里的判据和删表时的
 * `reconcileWorksheetDeletionScope` 对称：删表只允许**收缩**，撤销只允许
 * **目录增加、授权不变**。
 *
 * 一条边界：撤销**不会**把表加进授权。返回的 `authorized` 永远是原样的授权集合，
 * 只是把守卫基线与目录对齐。
 */

export type WorksheetRestoreScopeReconciliation =
  | { ok: true; authorized: string[] }
  | {
      ok: false;
      reason:
        | "identity_missing"
        | "workbook_changed"
        | "restored_missing"
        | "authorization_changed"
        | "unexpected_sheets";
    };

function uniqueNames(names: readonly string[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const name of names) {
    if (typeof name !== "string") continue;
    const trimmed = name.trim();
    if (!trimmed || seen.has(trimmed)) continue;
    seen.add(trimmed);
    result.push(trimmed);
  }
  return result;
}

export function reconcileWorksheetRestoreScope(input: {
  beforeIdentity: string | null;
  afterIdentity: string | null;
  /** 撤销**前**已提交的授权表名（收缩后的那一份）。 */
  authorizedSheets: readonly string[];
  /** 服务端回执里"真的放回原工作簿"的表名。 */
  restoredSheets: readonly string[];
  /** 撤销后宿主读到的真实目录。 */
  actualSheets: readonly string[];
  isSystemSheet?: (name: string) => boolean;
}): WorksheetRestoreScopeReconciliation {
  const filter = (names: readonly string[]) =>
    input.isSystemSheet ? names.filter((name) => !input.isSystemSheet!(name)) : [...names];

  if (!input.beforeIdentity || !input.afterIdentity) return { ok: false, reason: "identity_missing" };
  if (input.beforeIdentity !== input.afterIdentity) return { ok: false, reason: "workbook_changed" };

  const authorized = uniqueNames(input.authorizedSheets);
  const restored = uniqueNames(input.restoredSheets);
  const actual = uniqueNames(filter(input.actualSheets));
  if (!authorized.length || !restored.length) return { ok: false, reason: "restored_missing" };

  const actualSet = new Set(actual);
  // 授权过的表必须都还在：撤销只加表，不该让任何已授权表消失。
  if (authorized.some((name) => !actualSet.has(name))) return { ok: false, reason: "authorization_changed" };
  // 回执说放回去了，目录里就必须真有 —— 否则这份回执不能当证据。
  if (restored.some((name) => !actualSet.has(name))) return { ok: false, reason: "restored_missing" };

  const authorizedSet = new Set(authorized);
  const restoredSet = new Set(restored);
  // 目录里多的表只能来自这次撤销：其他新增（用户自己建表、外部程序加表）一律不合预期。
  if (actual.some((name) => !authorizedSet.has(name) && !restoredSet.has(name))) {
    return { ok: false, reason: "unexpected_sheets" };
  }
  // 授权集合原样返回：**不**把恢复出来的表自动纳入授权。
  return { ok: true, authorized };
}
