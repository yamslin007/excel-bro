import type { ComposerPrimaryAction } from "@excel-bro/core/composerPrimaryAction";
import "./taskExecutionNotice.css";

/**
 * 输入区上方的那一行**状态说明**（Excel / WPS 共用）。
 *
 * 这里**没有按钮** —— 这是本文件存在的全部理由。改造前它是"运行控制条"，
 * 自己长了一颗 ■/▶，和输入框右下角那颗主按钮抢同一个动作：同一个阶段有两个入口，
 * 其中一个不受输入框内容影响，于是"我只是想补充一句话"会被当成"点继续"。
 * 现在暂停/继续/核对/发送全部收敛到右下角那颗（`ComposerPrimaryButton`），
 * 这里只保留**输入框里那颗按钮表达不了的信息**：
 *
 *  - `等待审阅结果`：结果已生成，等用户确认/放弃。它不是"任务在跑"，也不该出现 ■；
 *  - `写入与回读期间不能中断…`：写下去的动作已经交给宿主，按停止不会让它立刻停。
 *    这是一句**原因**，不给按钮 —— 用户要的是"为什么"，不是"再点一次"；
 *  - `已暂停，可从右下角继续`：只有真的存在可恢复检查点时才这么说
 *    （`kind === "continue"` 的前提就是 `paused`，而 `paused` 只由待续跑消息给出）；
 *  - `正在停止当前处理…`：暂停请求已发出、还没回执。
 *
 * 换句话说：**状态留在这里，动作全在右下角。**
 */
export function TaskExecutionNotice({
  action,
  reviewing
}: {
  action: ComposerPrimaryAction;
  reviewing?: boolean;
}) {
  if (reviewing) {
    return (
      <p className="task-execution-notice" role="status" data-notice="reviewing">
        等待审阅结果
      </p>
    );
  }
  const text =
    action.kind === "pausing"
      ? "正在停止当前处理…"
      : action.kind === "pause" && action.disabled
        ? action.reason
        : action.kind === "continue"
          ? "已暂停，可从右下角继续"
          : null;
  if (!text) return null;
  return (
    <p className="task-execution-notice" role="status" data-notice={action.kind}>
      {text}
    </p>
  );
}
