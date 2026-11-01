/**
 * 输入框右下角那**一颗**主操作的唯一判据（Excel / WPS 共用，纯函数）。
 *
 * 改造前这里是三套各推一套的东西：输入框按钮（↑/■/▶）、过程卡上的「暂停」、
 * 以及"暂停请求在飞"时的临时态。后果是同一个位置在不同阶段长出不同的按钮，
 * 用户只能靠试。
 *
 * 本模块把"现在这个位置该是什么、按下去会发生什么"一次算清。它**不新增任何
 * 暂停协议**：能不能停仍然只看 `controllerAlive`（有没有一个未 abort 的对象），
 * 与 `pauseControl.ts` 同一条判据。
 *
 * 六种真实状态（按优先级，顺序即判据）：
 *
 *  1. `pausing`      暂停请求已发出、还没有回执 ⇒ **等待图标 + 禁用**。
 *                    不禁用的话用户会连点，第二次点出去的是第二个请求。
 *  2. `send-supplement` 有字 + 任务在跑 ⇒ 发送（语义是"打断并补充"，
 *                    不是排队再发一条）。文案必须说清这一点。
 *  3. `send`         有字 ⇒ 发送。**已暂停时也是发送**：那句话要按现有安全流程
 *                    重新判断，不能当成"继续"把旧计划原样重放。
 *  4. `reconcile`    现场不明（上次写入没核对清楚）⇒ **核对**，且**与继续分开**。
 *                    把核对和继续做成同一颗按钮，等于允许跳过核对直接续跑。
 *  5. `continue`     已暂停且到达安全边界、没有新话要说 ⇒ 继续。
 *                    **两种"继续"用同一档、不同文案**：`recoveryResume` 是"接着跑
 *                    上次没跑完的写入动作"，其余是"继续分析"。两者调的是不同处理器，
 *                    但都只由用户这一下触发 —— 核对之后**不会自己续跑**。
 *  6. `pause`        任务在跑、存在可中止对象 ⇒ 暂停。
 *                    写入/回读期间（`uninterruptible`）或没有可中止对象时，
 *                    画的是禁用态 `…` 并**给出原因**，绝不画一颗点了没反应的 ■。
 *
 * 一条铁律（位置冲突的解法）：**输入框里有字时，主按钮永远是"发送"**。
 * 核对/续跑都排在 `hasInput` 之后，所以不会出现"点发送却把文字丢掉去核对"。
 * 代价是"有字时核对点不到"—— 这是有意的：界面要说清"先清空输入框，或先把这句发出去"，
 * 而不是把用户正在写的东西静默吞掉。
 */

/** 主操作的六种形态。 */
export type PrimaryActionKind =
  | "send"
  | "send-supplement"
  | "answer"
  | "pause"
  | "pausing"
  | "continue"
  | "reconcile";

/** 点下去实际走哪条既有链路。`null` = 什么都不发（禁用态）。 */
export type PrimaryActionInvoke =
  | "send"
  | "pause"
  | "resume"
  | "reconcile"
  /** 提交对**当前这道澄清题**的回答（不是新一轮需求）。 */
  | "answer"
  | null;

export interface ComposerPrimaryInput {
  /** 输入框里有内容。 */
  hasInput: boolean;
  /** 已暂停，而且已经停在安全边界（不是在写入/回读中间）。 */
  paused: boolean;
  /** 本端认为有任务在途。 */
  taskRunning: boolean;
  /** 存在一个未 abort 的 controller —— **唯一**的"能暂停"依据。 */
  controllerAlive: boolean;
  /** 写入/回读等明确不允许中断的阶段。 */
  uninterruptible?: boolean;
  /** 暂停请求已发出、还没收到回执。 */
  pausing?: boolean;
  /** 现场不明：必须先核对在途动作，不能直接续跑。 */
  needsReconcile?: boolean;
  /**
   * 核对这一步在**本端**叫什么（缺省「核对在途动作」）。
   *
   * 允许宿主给文案，是因为同一颗按钮承接两种真实场景：一次写了一半的**执行**要核对
   * 现场，和一条**规划任务**还没核对过真实进度。两者的准确说法不同（后者是
   * "核对任务状态"），硬统一成一句会让人以为要去核对写入。判据仍然只有一份：
   * 什么时候出现、点下去做什么，都由这里算。
   */
  reconcileLabel?: string;
  /** 继续这一步在本端叫什么（缺省「继续任务」/「继续剩余动作」）。 */
  resumeLabel?: string;
  /**
   * 核对过了、可以接着跑**上次没跑完的写入动作**（恢复续跑，不是继续分析）。
   *
   * 与 `paused` 分开表达：两者都可能同时成立（规划任务暂停 + 一次写入等续跑），
   * 而它们调的是**不同处理器**。混成一个布尔会在"继续"这颗按钮上撒谎。
   */
  recoveryResume?: boolean;
  /**
   * 输入框此刻是在**回答一道澄清题**（而不是提新需求）。
   *
   * 有它时"发送"发出去的是回答：走澄清链路，不是新一轮需求。
   * 判据在面板侧由"当前 pending 问题 + 用户没有显式选择另提新需求"给出。
   */
  answeringClarification?: boolean;
  /** 回答有没有内容（选项选中 / 字段填了 / 自由输入非空）。 */
  hasAnswerContent?: boolean;
}

export interface ComposerPrimaryAction {
  kind: PrimaryActionKind;
  /** 按钮上的字形。 */
  glyph: string;
  /** 无障碍名（`aria-label`）。 */
  label: string;
  /** 悬停说明（`title`）。必须说清"按下去真正会发生什么"。 */
  title: string;
  /** 不可点击。 */
  disabled: boolean;
  /** 显示等待态（与 `disabled` 分开：等待不是"不可用"，是"正在发生"）。 */
  busy: boolean;
  invoke: PrimaryActionInvoke;
  /** 不可暂停/需要核对的**具体原因**；由调用方如实展示，不许吞掉。 */
  reason: string | null;
}

export const PRIMARY_PAUSE_UNAVAILABLE =
  "本端已经没有可中止的执行；正在等状态核对结果，不会静默。";
export const PRIMARY_PAUSE_UNINTERRUPTIBLE =
  "写入与回读期间不能中断，以免结果状态不明。";
export const PRIMARY_PAUSING_TITLE = "正在停止当前处理，请稍候";
export const PRIMARY_CONTINUE_TITLE = "继续上次暂停的任务";
export const PRIMARY_SEND_TITLE = "发送这条需求";
export const PRIMARY_SEND_SUPPLEMENT_TITLE =
  "打断当前处理，带上这句话重新开始（不是排队发送）";
export const PRIMARY_SEND_AFTER_PAUSE_TITLE =
  "发送这条补充要求：会按现有流程重新判断，不会重放上次的计划";
export const PRIMARY_RECONCILE_TITLE =
  "先只读核对上次写入留下了哪些动作；核对与继续是两件事";
export const PRIMARY_RECOVERY_RESUME_LABEL = "继续剩余动作";
export const PRIMARY_RECOVERY_RESUME_TITLE =
  "接着跑上次没跑完的动作：只执行剩余步骤，已经写过的不会重写";

export function composerPrimaryAction(
  input: ComposerPrimaryInput
): ComposerPrimaryAction {
  /*
   * ① 暂停请求在飞：等待图标 + 禁用。
   *
   * 这里是**唯一**允许"按钮不可点但仍占位"的情形之一。必须占位：
   * 按钮消失会让用户以为按错了，再点一次发送框里的字就成了新需求。
   */
  if (input.pausing) {
    return {
      kind: "pausing",
      glyph: "…",
      label: "正在暂停",
      title: PRIMARY_PAUSING_TITLE,
      disabled: true,
      busy: true,
      invoke: null,
      reason: null
    };
  }

  /* ② 有字 + 在跑：发送（语义是打断并补充）。 */
  if (input.hasInput && input.taskRunning) {
    return {
      kind: "send-supplement",
      glyph: "↑",
      label: "打断并补充",
      title: PRIMARY_SEND_SUPPLEMENT_TITLE,
      disabled: false,
      busy: false,
      invoke: "send",
      reason: null
    };
  }

  /*
   * ③ 正在回答一道澄清题：发送 = 提交回答。
   *
   * 放在"有字就发送"之前：同一份文字，在有 pending 问题时是**回答**，
   * 在没有 pending 问题时才是新需求。判反了就会把回答当成一条新指令重跑一遍。
   * 没内容时禁用并给原因，而不是让按钮消失 —— 消失会让人以为这道题不用回答。
   */
  if (input.answeringClarification && !input.taskRunning) {
    if (!input.hasAnswerContent) {
      return {
        kind: "answer",
        glyph: "↑",
        label: "发送回答",
        title: "先选一项或写下你的想法，再发送",
        disabled: true,
        busy: false,
        invoke: null,
        reason: "请选择一项，或写下你的想法。"
      };
    }
    return {
      kind: "answer",
      glyph: "↑",
      label: "发送回答",
      title: "提交对这道问题的回答（不会当成新需求重跑）",
      disabled: false,
      busy: false,
      invoke: "answer",
      reason: null
    };
  }

  /*
   * ④ 有字就发送 —— **包括已暂停的时候**。
   *
   * 已暂停时把这句话当成"继续"是个真实的错误方向：用户写的是新要求，
   * 旧计划里的动作却会被原样重放。所以这里给的是"发送补充要求"，
   * 由现有的安全流程重新判断；`invoke` 也仍是 `send`，不是 `resume`。
   */
  if (input.hasInput) {
    return {
      kind: "send",
      glyph: "↑",
      label: "发送",
      title: input.paused ? PRIMARY_SEND_AFTER_PAUSE_TITLE : PRIMARY_SEND_TITLE,
      disabled: false,
      busy: false,
      invoke: "send",
      reason: null
    };
  }

  /* ⑤ 现场不明：先核对。与"继续"必须是两颗不同的按钮。 */
  if (input.needsReconcile && !input.taskRunning) {
    const label = input.reconcileLabel?.trim() || "核对在途动作";
    return {
      kind: "reconcile",
      glyph: "◎",
      label,
      title: PRIMARY_RECONCILE_TITLE,
      disabled: false,
      busy: false,
      invoke: "reconcile",
      reason: "上次写入的现场还没核对清楚，先核对再决定继续或撤销。"
    };
  }

  /*
   * ⑤b 核对过了、写入还差几步：继续**剩余动作**。
   *
   * 排在"已暂停继续分析"之前：一次写到一半的执行，现场有明确的未完成动作，
   * 它的代价比"重新分析一遍"高得多。文案必须写"剩余动作"，不能写成"继续任务"，
   * 否则用户会以为这一步会重新规划、甚至重放旧计划。
   */
  if (input.recoveryResume && !input.taskRunning) {
    return {
      kind: "continue",
      glyph: "▶",
      label: input.resumeLabel?.trim() || PRIMARY_RECOVERY_RESUME_LABEL,
      title: PRIMARY_RECOVERY_RESUME_TITLE,
      disabled: false,
      busy: false,
      invoke: "resume",
      reason: null
    };
  }

  /* ⑥ 已暂停、到达安全边界、没有新话要说：继续。 */
  if (input.paused && !input.taskRunning) {
    return {
      kind: "continue",
      glyph: "▶",
      label: input.resumeLabel?.trim() || "继续任务",
      title: PRIMARY_CONTINUE_TITLE,
      disabled: false,
      busy: false,
      invoke: "resume",
      reason: null
    };
  }

  /*
   * ⑦ 在跑：只在**真的能停**时才画暂停。
   *
   * 写入/回读期间画一颗可点的 ■，是在承诺一件做不到的事：已经提交给宿主的
   * 写入不会因为按了按钮就立刻停。这里如实给禁用态与原因。
   */
  if (input.taskRunning && input.uninterruptible) {
    return {
      kind: "pause",
      glyph: "…",
      label: PRIMARY_PAUSE_UNINTERRUPTIBLE,
      title: PRIMARY_PAUSE_UNINTERRUPTIBLE,
      disabled: true,
      busy: false,
      invoke: null,
      reason: PRIMARY_PAUSE_UNINTERRUPTIBLE
    };
  }
  if (input.taskRunning && input.controllerAlive) {
    return {
      kind: "pause",
      glyph: "■",
      label: "暂停任务",
      title: "暂停接收规划进度，保留需求；核对检查点后才可继续",
      disabled: false,
      busy: false,
      invoke: "pause",
      reason: null
    };
  }
  if (input.taskRunning) {
    return {
      kind: "pause",
      glyph: "…",
      label: PRIMARY_PAUSE_UNAVAILABLE,
      title: PRIMARY_PAUSE_UNAVAILABLE,
      disabled: true,
      busy: false,
      invoke: null,
      reason: PRIMARY_PAUSE_UNAVAILABLE
    };
  }

  /* 任务结束（或从未开始）：发送。 */
  return {
    kind: "send",
    glyph: "↑",
    label: "发送",
    title: PRIMARY_SEND_TITLE,
    disabled: false,
    busy: false,
    invoke: "send",
    reason: null
  };
}
