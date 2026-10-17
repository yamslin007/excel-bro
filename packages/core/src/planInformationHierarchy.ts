/**
 * 任务窗格的信息层级"口径"。
 *
 * 为什么要有这张表：Excel 与 WPS 渲染的是同一个 `AnalysisPlan`，
 * 但两端的计划卡/结果卡是各自写的 JSX。以往两端各自写死同一句中文，
 * 改一处就会悄悄漂移 —— 例如恢复能力被写成"可撤销"，而真实情况是
 * 恢复点只活在当前加载项会话里，关闭后就用不了了。
 * 现在两端共用这里的字面量，测试可以直接断言口径一致。
 *
 * 边界：这里只放**标题与标签**这类口径。执行器真实给出的证据数量、
 * 目标清单、逐项检查明细由各端按自己拿到的东西组织 —— 证据来源不同，
 * 强行共享就是假装对齐。
 */

/** 计划卡两个笔记区块的标题，分别对应 AnalysisPlan 的 assumptions / warnings。 */
export const PLAN_NOTE_SECTIONS = {
  assumptions: "采用的信息",
  warnings: "补充说明"
} as const;

/**
 * 格式改动预览区块的标题。
 *
 * 两端共用这一串：它是"确认前必须看到的格式影响"，不是某一端的装饰文案，
 * 各写一份就会漂移成"Excel 说了 WPS 没说"。
 */
export const PLAN_FORMAT_SECTION = "格式改动" as const;

/**
 * 行高/列宽溢出目标区域时，预览前面必须挂的警示标签。
 * 用户看到"调整 A1:C10 的行高"时默认理解是"只动这一小块"，
 * 这个标签就是用来打破该默认理解的。
 */
export const PLAN_FORMAT_IMPACT_LABEL = "影响整行/整列" as const;

/**
 * 恢复能力的准确标签。必须带"会话内"：恢复点在会话结束后就取不回来，
 * 笼统写成"可撤销"会让用户以为这是可以事后追溯的承诺。
 */
export const SESSION_RECOVERY_LABEL = "会话内可撤销";

/** 恢复点寿命说明的前缀；各端可以在后面接自己的补充句子。 */
export const SESSION_RECOVERY_NOTICE_PREFIX = "恢复点仅在当前加载项会话内可用";

/**
 * 结果卡表示"检查确实通过"的字样。
 * 只有在执行器真的完成核对之后才允许出现；失败、部分完成或未知时不得渲染。
 */
export const VERIFICATION_PASSED_LABEL = "验证通过";
