import type { WpsIdentityCheck } from "./wpsObjectIdentity";

/**
 * 待答澄清在作答前要回答的唯一问题：**还能不能接着这张卡回答**。
 *
 * 三种结果刻意分开，**不得互相冒充**：
 *
 * - `"same"`：能**正面证明**仍是建卡时的授权对象 → 继续回答，复用该轮授权。
 * - `"changed"`：有**正面证据**说明授权对象或表头已经变了 → 旧卡失效，
 *   提示"数据范围已经变化"。
 * - `"unknown"`：**证明不了** → 既不放行，也**不谎称"数据已变化"**。
 *   此时保留用户的原需求与他刚填的回答，请他在数据范围里重新确认原工作簿
 *   与授权工作表，确认后作为**新一轮**重新读取结构、重新规划与预览。
 *
 * 为什么必须分出 `"unknown"`：宿主可能对**同一张表**返回**新的包装对象**，
 * 也可能在编辑数据格后让旧的包装对象失效。此时对象引用对不上，但授权对象
 * 其实没变。把它判成 `"changed"` 会对着只改了一格又撤回的用户说"数据范围已经
 * 变化" —— 那是假话，而且用户完全无从纠正。
 *
 * 判据优先级：**先看有没有变化证据，再看能不能证明相同，最后才落到证明不了**。
 * 这样"引用相同但表头被改写"这类情形仍会判 `"changed"`，不会因为引用相等就放行。
 */
export type WpsClarificationVerdict = "same" | "changed" | "unknown";

export type WpsClarificationScopeFacts = {
  /** 授权身份指纹（工作簿名 + 已授权表的索引与名称，含顺序）。 */
  identity: { recorded: string | null | undefined; current: string | null };
  /** 表头指纹（真实第 1 行单元格值）。 */
  header: { recorded: string | null | undefined; current: string | null };
  /**
   * 内存里的宿主对象身份复核结果。
   * `null` = 本次会话压根没捕获到（面板刚重开、宿主读不到）。
   */
  objectIdentity: WpsIdentityCheck | null;
};

/**
 * 指纹比对。任一侧读不到（`null`）就**没有可比信息** —— 记"没变化证据"，
 * 不记"变了"。把读失败当成变化是这套判据曾经最大的假话来源。
 */
function fingerprintDiffers(
  recorded: string | null | undefined,
  current: string | null
): boolean {
  if (recorded == null || current == null) return false;
  return recorded !== current;
}

export function wpsClarificationVerdict(
  facts: WpsClarificationScopeFacts
): WpsClarificationVerdict {
  // 1. 正面的变化证据优先 —— 它压过"引用相同"。
  if (facts.objectIdentity === "different") return "changed";
  if (fingerprintDiffers(facts.identity.recorded, facts.identity.current)) return "changed";
  if (fingerprintDiffers(facts.header.recorded, facts.header.current)) return "changed";

  // 2. 唯一能放行的路径：对象身份被**正面证明**相同。
  //    名称与表头一致**不算**证明 —— 一张重建的复制表全都满足。
  if (facts.objectIdentity === "same") return "same";

  // 3. 其余一律"证明不了"：可能真变了，也可能只是换了包装对象。
  return "unknown";
}

/** 给用户看的说明。两种失效原因用**不同**的话，不许都写成"数据已经变化"。 */
export function wpsClarificationScopeNotice(verdict: WpsClarificationVerdict): string {
  return verdict === "changed"
    ? "数据范围已经变化，请检查范围后重新发送需求。"
    : "无法确认当前工作簿和授权工作表仍是发起这次提问时的那一份，为避免把回答用到别的表上，请先在数据范围里重新确认原工作簿与授权工作表；确认后会把这条需求和你的回答作为新一轮重新读取和预览。";
}
