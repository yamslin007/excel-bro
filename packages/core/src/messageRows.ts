/**
 * 消息行的**分组呈现判据** —— Excel / WPS 共用（L2 纯语义层，不碰 React）。
 *
 * ---------------------------------------------------------------------------
 * 这一层回答两件事，且**只回答这两件**
 * ---------------------------------------------------------------------------
 *   1. 这一行还要不要印一次作者名（"Excel Bro" 及其来源徽标）；
 *   2. 这一行是不是"一组回复的首行"（决定它上面留多少空隙）。
 *
 * 为什么必须是共享判据：两端各写一份"行为相似"的副本，就会一边把名字省了、
 * 另一边还印着，而用户来回切换时看到的其实是同一段对话。
 *
 * ---------------------------------------------------------------------------
 * 为什么不是"每条消息都印名字"
 * ---------------------------------------------------------------------------
 * 引擎是一段回复拆成**多条消息**落盘的：过程卡一条、结论一条、结果卡一条。
 * 每条都印名字，界面上就会出现两三个并排的"Excel Bro"，读起来像三件事。
 * 参考实现（对标的 ExcelMaster）是"一段回复一个抬头"，本判据就是这件事。
 *
 * 三条不变量：
 *
 * A. **省的是重复，不是信息。** 一组里只有第一行印名字；但两种**都写明了的**
 *    来源徽标互相不同时仍然各自印一次（过程行是 Agent、接着一行是「需求确认」）——
 *    徽标说的是"这一行是谁给的"，省掉它等于把用户要接手的那道题藏起来。
 *    反过来，"本来没有徽标、后来才标出一个"不算换出处（见 `needsOwnLabel`）。
 *
 * B. **过程行会让位。** 单步的消息理解本身没有内容（`isCompactMessageProcess`），
 *    它的名字要让给紧跟的结论行：否则名字下面是一张空的进度条，正文却无名。
 *    这条是旧行为，本次保留 —— 它的反向锚点是一条真实回归（见用例）。
 *
 * C. **换段必须重新署一次。** 「继续」开的是**新的回复段**（`replySegments.ts`），
 *    两段之间可能连一条用户消息都没有。只按"上一条是不是助手"分组的话，新段的
 *    第一行会被当成上一段的续行而省掉名字 —— 用户看到的是一段凭空冒出来的进度。
 *    因此分组边界同时认段身份。
 */

import { isCompactMessageProcess } from "./processPanel";

/** 这条消息里与本判据有关的部分。两端各自从自己的消息类型映射过来。 */
export interface MessageRowLike {
  /** 必须是原始的 `role`：只有 `assistant` 参与署名。 */
  role: string;
  text?: string;
  /** 本条消息所属回复段身份；老消息没有（`null`/`undefined`），按"同一组"处理。 */
  segmentId?: string | null;
  /** 作者名后面的来源徽标文案；`null`/`undefined` = 这一行不带徽标。 */
  badge?: string | null;
  /** 助手消息的过程日志，只用来判"单步消息理解的空过程行"。 */
  processLog?: Parameters<typeof isCompactMessageProcess>[0];
  /** 这一行这次**不渲染**（例如正在作答的澄清问卷由输入框接管）⇒ 不参与分组。 */
  skipped?: boolean;
}

export interface MessageRowPresentation {
  /** 这一行要印作者名。 */
  author: boolean;
  /** 这一行是"一组回复的首行"（CSS 据此给组间空隙，组内只留小间距）。 */
  groupStart: boolean;
}

/**
 * 逐行给出 `{ author, groupStart }`，顺序与入参一一对应。
 *
 * 纯函数、无副作用、无随机：同一串消息永远得到同一份结果 —— 这是它可以被
 * 两端共用、也可以被反向变异检验的前提。
 */
export function messageRowPresentation(
  rows: readonly MessageRowLike[]
): MessageRowPresentation[] {
  const visible = rows.map((row) => !row.skipped);
  const result: MessageRowPresentation[] = rows.map(() => ({
    author: false,
    groupStart: false
  }));
  /**
   * 当前这一组已经由谁署过名：`undefined` = 组内还没人署名，
   * `null` = 署名的那一行不带徽标。
   */
  let groupBadge: string | null | undefined;

  for (let index = 0; index < rows.length; index += 1) {
    if (!visible[index]) continue;
    const row = rows[index];

    // 用户 / 系统消息永远不署名，但它们一定打断分组：气泡上面那条回复到此为止。
    if (row.role !== "assistant") {
      groupBadge = undefined;
      result[index].groupStart = true;
      continue;
    }

    const previous = previousVisibleIndex(rows, visible, index);
    if (previous < 0 || !sameGroup(rows[previous], row)) {
      groupBadge = undefined;
      result[index].groupStart = true;
    }

    // 不变量 B：空过程行把名字让给结论行（结论行因此可能"不是首行却要署名"）。
    if (defersAuthorToNextRow(rows, visible, index)) continue;

    const badge = row.badge ?? null;
    /*
     * 组内第一行印，之后只有**两种都写明了的**出处互相不同才再印一次
     * （`Agent` 的过程后面紧跟「需求确认」—— 那是另一件事，抬头不能共用）。
     *
     * 「本来没有徽标、后来才标出一个」不算换出处：组首那颗抬头已经写过一次，
     * 再印一遍得到的是两颗一模一样的「Excel Bro」。这不是假设 —— WPS 侧的
     * 徽标口径是"没有过程/计划/验证时才标 Agent"，于是每一对
     * 「过程行（无徽标）→ 结论行（Agent）」都会踩到，正是要消除的那个样子。
     */
    const needsOwnLabel =
      badge !== null && groupBadge !== null && groupBadge !== undefined && badge !== groupBadge;
    if (groupBadge === undefined || needsOwnLabel) {
      result[index].author = true;
      groupBadge = badge;
    }
  }

  return result;
}

/** 上一条**会渲染**的消息下标；没有则 -1。跳过不渲染的行，避免把名字让给看不见的一行。 */
function previousVisibleIndex(
  rows: readonly MessageRowLike[],
  visible: readonly boolean[],
  index: number
): number {
  for (let cursor = index - 1; cursor >= 0; cursor -= 1) {
    if (visible[cursor]) return cursor;
  }
  return -1;
}

/**
 * 两条相邻的助手消息是否仍属**同一组**。
 *
 * 只在双方都带段身份、且段不同时才判"换组"：老消息没有段身份，不能因为缺字段
 * 就把整段历史拆成一段一段。
 */
function sameGroup(previous: MessageRowLike, row: MessageRowLike): boolean {
  if (previous.role !== "assistant") return false;
  const before = previous.segmentId ?? null;
  const current = row.segmentId ?? null;
  if (!before || !current) return true;
  return before === current;
}

/**
 * 这一行是不是"只有一张单步消息理解的过程卡、没有文字、而下一行就是结论"。
 *
 * 判据与旧的渲染判断逐字对应（`!text && isCompactMessageProcess && 下一条是带文字的
 * 助手消息`），只把"下一条"从"数组的下一条"改成"下一条**会渲染的**"：把名字让给
 * 一条被跳过（不渲染）的消息，等于名字彻底消失。
 */
function defersAuthorToNextRow(
  rows: readonly MessageRowLike[],
  visible: readonly boolean[],
  index: number
): boolean {
  const row = rows[index];
  if (row.text) return false;
  if (!isCompactMessageProcess(row.processLog)) return false;
  for (let cursor = index + 1; cursor < rows.length; cursor += 1) {
    if (!visible[cursor]) continue;
    return rows[cursor].role === "assistant" && Boolean(rows[cursor].text?.trim());
  }
  return false;
}
