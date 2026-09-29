/**
 * 对话历史的存储状态提示（唯一渲染点）。
 *
 * **为什么把它从 App.tsx 抽出来**：这块提示描述的是一件有四种真实形态的事，
 * 而早先的实现只有一个布尔值 `legacyCleanupNotice`，四种形态被压成了"显示 / 不显示"。
 * 于是最坏的一种组合 —— **v4 根本没写进去、而且压根没有 v3 旧键** —— 也会渲染出
 * "旧版本记录待清理 / 历史已经以脱敏形式保存"。两句话都不成立：没有旧键可清理，
 * 本次历史也从没保存过。布尔值测不出这种谎报，只有把**实际渲染文案**摆到 DOM 上
 * 才能钉住，所以提示文案必须与判定逻辑分开、可被单独渲染测试。
 *
 * 四种状态（判据见 `useConversation` → `historyStorageStatus`）：
 *
 * | 状态 | 事实 | 必须说 | 绝不能说 |
 * | --- | --- | --- | --- |
 * | `save-failed-legacy-present` | v4 写失败，v3 旧键还在 | 本次保存失败；旧记录仍保留、未清理 | "已保存" |
 * | `save-failed-no-legacy` | v4 写失败，且没有 v3 旧键 | 对话历史未保存 | "旧版本记录待清理" |
 * | `saved-legacy-pending` | v4 已保存，v3 删不掉 | 脱敏后的 v4 已保存；旧键仍待清理 | "保存失败" |
 * | `saved-clean` | v4 已保存，v3 已清理 | （不渲染任何提示） | 任何警告 |
 *
 * 文案里**不得出现异常对象原文**（可能带键名、配额乃至被拒绝的值），也不得出现
 * 任何历史内容 —— 这些值仍然可能含凭据。这里只描述"存没存下、旧键清没清掉"。
 */

/** 对话历史的存储状态。四种形态互斥，且必须来自同一次落盘尝试的结论。 */
export type HistoryStorageStatus =
  | "saved-clean"
  | "saved-legacy-pending"
  | "save-failed-legacy-present"
  | "save-failed-no-legacy";

/** 只有这一种状态不需要提示。 */
export function isHistoryStorageNoticeVisible(
  status: HistoryStorageStatus
): boolean {
  return status !== "saved-clean";
}

export interface HistoryStorageNoticeCopy {
  /** 标题（加粗的那行）。 */
  heading: string;
  /** 正文：说清事实与后续动作。 */
  detail: string;
}

/**
 * 状态 → 文案。抽成纯函数是为了让"这句话是否成立"能被直接断言，
 * 而不必去渲染一个 React 组件再读文本。
 */
export function historyStorageNoticeCopy(
  status: HistoryStorageStatus
): HistoryStorageNoticeCopy | null {
  switch (status) {
    case "saved-clean":
      // 两个警告都该消失。
      return null;
    case "save-failed-legacy-present":
      return {
        heading: "对话历史没能保存到本机",
        detail:
          "本次没能把脱敏后的对话历史写入本机存储（可能被宿主设为只读或空间已满），" +
          "本次会话关掉任务窗格后不会被保留。旧版本的原始记录**仍然完整保留**在存储里，" +
          "没有被清理 —— 需要时可以在恢复存储后重新开始迁移。"
      };
    case "save-failed-no-legacy":
      return {
        heading: "对话历史未能保存到本机",
        detail:
          "本次没能把对话历史写入本机存储（可能被宿主设为只读或空间已满），" +
          "本次会话关掉任务窗格后不会被保留。本机存储里没有旧版本记录，无需清理。"
      };
    case "saved-legacy-pending":
      return {
        heading: "旧版本记录待清理",
        detail:
          "脱敏后的对话历史已经保存到本机，界面和后续请求都不会再带出原文；" +
          "但旧版本对话记录的存储项还没能清除（本机存储可能被宿主设为只读）。" +
          "旧存储项会在下次启动或存储恢复后自动清理。"
      };
  }
}

export default function HistoryStorageNotice({
  status
}: {
  status: HistoryStorageStatus;
}) {
  const copy = historyStorageNoticeCopy(status);
  if (!copy) return null;
  return (
    <div className="storage-write-warning" role="alert">
      <strong>{copy.heading}</strong>
      <span>{copy.detail}</span>
    </div>
  );
}
