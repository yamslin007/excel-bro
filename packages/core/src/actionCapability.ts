/**
 * L2 语义层：宿主 × 版本 × 动作 的可信能力上界与交集（方案 §4.2/§4.3）。
 *
 * 这一层刻意**不认识任何具体宿主**，只做纯函数：给定一份能力矩阵与一个
 * 已解析的宿主身份，算出"这个宿主上的动作上界"；再与客户端自报集合求交。
 *
 * 为什么服务端必须自己算上界，而不是采信客户端：
 * 客户端自报集合是**请求体里的一段字符串数组**。若把它当权威，一个构造过的
 * 请求就能让规划器生成本不该开放的动作。所以方向被固定为
 * **客户端只能收窄，不能扩大** —— 交集运算天然满足这一点
 * （`A ∩ B ⊆ A` 恒成立），不需要额外的越权检查。
 *
 * 反向的错误同样要挡住：客户端漏报（自报为空）时必须得到空集，而不是
 * "客户端没说，那就放开"。空集意味着没有动作可规划，规划器会明确拒绝，
 * 不会生成一份之后必然被执行的计划。这是保守失败，不是静默放开。
 */

import type { ActionCapability, HostId } from "./ports";

/**
 * 一个宿主在某个版本前缀下的动作表。
 *
 * `versions` 的键是**版本前缀**，匹配规则见 `matchesVersionPrefix`。
 * 未知版本前缀与未知宿主都解析为空表 —— 空表 = 无动作可规划。
 */
export interface HostCapabilityMatrix {
  /** 宿主级总开关：为 `false` 时整个宿主的所有动作都不进计划。 */
  enabled?: boolean;
  versions: Record<string, HostActionTable>;
}

/** 动作表。键是 `ExcelAction` 的动作类型，`clearRange:contents` 这类变体也用同一形态。 */
export type HostActionTable = Record<string, ActionCapability>;

/** 可信能力矩阵：宿主标识 → 该宿主的能力。未列出的宿主 = 未知宿主 = 空上界。 */
export type CapabilityMatrix = Record<string, HostCapabilityMatrix>;

/**
 * 版本前缀匹配。
 *
 * 与 `wpsWriteVersionGate.ts` 的 `matchesVersionPrefix` 同一规则：
 * 相等或后接 `.`。**不用 `startsWith` 裸匹配** —— 那会让 `"12.0"` 命中
 * `"12.01"`，把两个不同版本当成同一个。
 */
export function matchesVersionPrefix(version: string, prefix: string): boolean {
  return version === prefix || version.startsWith(`${prefix}.`);
}

/** 解析宿主身份。三者都不满足时返回 `null`，交由调用方拒绝，不猜成任何一端。 */
export function parseHostId(value: string | null | undefined): HostId | null {
  const host = (value ?? "").trim().toLowerCase();
  if (host === "excel-office" || host === "wps-et" || host === "folder") {
    return host;
  }
  return null;
}

/**
 * 算出一个宿主/版本的**动作上界**。
 *
 * 返回的每个动作都带原因，而不是只给一个名称集合 —— 用户要看到的是
 * "为什么这个动作不能用"，而原因只有在这个函数里才知道
 * （是被宿主开关关掉、版本不认识，还是这个动作本身没验收）。
 *
 * 未知宿主、未知版本返回**空表**，不是"全部放开"。
 */
export function resolveHostActionBound(
  matrix: CapabilityMatrix,
  hostId: string | null | undefined,
  hostVersion: string | null | undefined
): HostActionTable {
  const host = parseHostId(hostId);
  if (!host) return {};
  const entry = matrix[host];
  if (!entry || entry.enabled === false) return {};
  const version = (hostVersion ?? "").trim();
  if (!version) return {};
  // 版本前缀可能互相重叠（例如同时列了 "12" 与 "12.0"）。取**最长的那个**
  // 匹配，让更具体的条目覆盖更宽的条目；否则结果会依赖对象的键顺序。
  const matched = Object.keys(entry.versions)
    .filter((prefix) => matchesVersionPrefix(version, prefix))
    .sort((left, right) => right.length - left.length)[0];
  if (!matched) return {};
  return entry.versions[matched] ?? {};
}

/**
 * 服务端上界 ∩ 客户端自报集合。
 *
 * 结果里**只保留上界中为 `supported` 的动作**。上界里的 `unverified` /
 * `unsupported` 即便客户端自报为可用也不会进入结果 —— 客户端没有资格把
 * 一个未验收的动作变成可规划的动作。
 */
export function intersectActionCapabilities(
  bound: HostActionTable,
  claimed: readonly string[] | null | undefined
): Set<string> {
  const allowed = new Set<string>();
  if (!claimed || claimed.length === 0) return allowed;
  const claimedSet = new Set(claimed);
  for (const [actionType, capability] of Object.entries(bound)) {
    if (capability.state !== "supported") continue;
    if (!claimedSet.has(actionType)) continue;
    allowed.add(actionType);
  }
  return allowed;
}

/**
 * 一步到位：宿主/版本 → 上界 → 与自报集合求交。
 *
 * 分开暴露两个函数是为了让测试能分别锁住"上界算得对不对"与
 * "交集有没有偷偷放开"，而不必每次都构造完整矩阵。
 */
export function resolveAllowedActionTypes(
  matrix: CapabilityMatrix,
  hostId: string | null | undefined,
  hostVersion: string | null | undefined,
  claimed: readonly string[] | null | undefined
): Set<string> {
  return intersectActionCapabilities(
    resolveHostActionBound(matrix, hostId, hostVersion),
    claimed
  );
}

/**
 * 动作键的规范化：`clearRange` 这类动作按 `applyTo` 分变体，键形如
 * `clearRange:contents`。变体与父动作都必须参与判定 ——
 * 只写 `clearRange` 表示"全部变体"，写 `clearRange:contents` 只表示这一个。
 */
export function actionCapabilityKeys(actionType: string, variant?: string | null): string[] {
  if (actionType === "clearRange" && variant) {
    return [`${actionType}:${variant}`, actionType];
  }
  return [actionType];
}

/**
 * 判定一个具体动作（含变体）是否在交集内。
 *
 * 变体规则：**父动作在白名单里但没有对应变体条目时，视为未验收。**
 * 不能因为 `clearRange` 出现过就放开 `clearRange:formats` ——
 * 那正是"一个动作名笼统宣称全部支持"的失误（对齐指导 §2 末段）。
 */
export function isActionAllowedByKeys(
  allowed: ReadonlySet<string>,
  actionType: string,
  variant?: string | null
): boolean {
  if (actionType === "clearRange" && variant) {
    return allowed.has(`${actionType}:${variant}`);
  }
  return allowed.has(actionType);
}
