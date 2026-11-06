import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

/**
 * 底部结构的**双端契约**（读源码，与 `composerResizeParity.test.ts` 同一层证据）。
 *
 * 真实 DOM 的那一半在 `ComposerPendingBar.test.tsx` / `ExecutionFeedbackSummary.test.tsx`
 * 与两端面板集成测试里。这里只钉住"两端不许各写一套"以及
 * **`.composer-dock` 里只准有工具栏 + 输入框**这条结构约束：
 * 旧实现把通知、恢复摘要、历史反馈逐条插在两者之间，于是每来一条提示就
 * 把工具栏从输入框旁边推开一次。
 */
const appSource = readFileSync("src/App.tsx", "utf8");
const wpsSource = readFileSync("../wps-addin/src/WpsAgentPanel.tsx", "utf8");
const sharedStyles = readFileSync("src/styles.css", "utf8");
const conversationStyles = readFileSync("src/conversation.css", "utf8");

const DOCK_OPEN = '<div className="composer-dock">';
const DOCK_CLOSE = "/.composer-dock";

function dockOf(source: string): string {
  const start = source.indexOf(DOCK_OPEN);
  const end = source.indexOf(DOCK_CLOSE, start);
  expect(start, "必须有固定底部容器").toBeGreaterThan(-1);
  expect(end, "容器必须有结束标记").toBeGreaterThan(start);
  return source.slice(start, end);
}

/** 任何一条出现在工具栏与输入框之间，都会让布局随状态跳动。 */
const FORBIDDEN_BETWEEN = [
  "TargetScopeNotice",
  "WorkDirectoryGuide",
  "HistoryStorageNotice",
  "ComposerActionDock",
  "RecoveryPendingEntry",
  "TaskExecutionNotice",
  "composer-service-status",
  "composer-pending",
  "result-review-bar",
  "RecoveryProbeControl"
];

describe("底部工具栏与输入框：Excel / WPS 同一份结构", () => {
  it.each([
    ["Excel", appSource],
    ["WPS", wpsSource]
  ] as const)("%s：待办栏在 dock 之前，两者都从共享层来", (_host, source) => {
    expect(source).toContain("ComposerPendingBar");
    expect(source).toContain("composerPendingItems");
    expect(source.indexOf("<ComposerPendingBar")).toBeLessThan(source.indexOf(DOCK_OPEN));
    expect(source).toContain("groupOperationFeedback");
    expect(source).toContain("ExecutionFeedbackSummary");
  });

  it.each([
    ["Excel", appSource],
    ["WPS", wpsSource]
  ] as const)("%s：dock 里只有工具栏与输入区", (_host, source) => {
    const dock = dockOf(source);
    expect(dock).toContain("scope-entry-row");
    expect(dock).toContain("composer-input-region");
    expect(dock).toContain("composer-box");
    for (const forbidden of FORBIDDEN_BETWEEN) {
      expect(dock, `${forbidden} 不许回到工具栏与输入框之间`).not.toContain(forbidden);
    }
  });

  it("共享 CSS 用 flex 让两者紧贴，不靠负 margin / 绝对叠盖 / 隐藏溢出", () => {
    const rule = conversationStyles.match(/\.chat-shell \.composer-dock\s*\{[^}]*\}/)?.[0] ?? "";
    expect(rule).toContain("display: flex");
    expect(rule).toContain("flex-direction: column");
    expect(rule).not.toContain("margin-top: -");
    expect(rule).not.toContain("position: absolute");
    expect(rule).not.toContain("overflow: hidden");
    expect(conversationStyles).toContain(".chat-shell .composer-dock > .scope-entry-row");
    expect(conversationStyles).toContain(".chat-shell .composer-dock > .composer-input-region");
  });

  it("让位的是常驻摘要，不是底部容器：`.composer` 不许被压到内容以下", () => {
    // 反向证据（2026-11-05 实测）：给 `.chat-shell .composer` 加 `min-height: 0` 后，
    // `check_recovery_composer_layout.mjs` 在 320x640 报「审阅动作行 bottom=728 > 640」。
    const composerRule = conversationStyles.match(/\.chat-shell \.composer\s*\{[^}]*\}/)?.[0] ?? "";
    expect(composerRule).toContain("display: flex");
    expect(composerRule).not.toContain("min-height: 0");
    const pendingRule = conversationStyles.match(/\.chat-shell \.composer-pending\s*\{[^}]*\}/)?.[0] ?? "";
    expect(pendingRule).toContain("flex: 0 1 auto");
    const dockRule = conversationStyles.match(/\.chat-shell \.composer-dock\s*\{[^}]*\}/)?.[0] ?? "";
    expect(dockRule).toContain("flex: 0 0 auto");
  });

  it("消息区独立滚动并占据剩余高度，底部变高时正常让出空间", () => {
    expect(sharedStyles).toMatch(/\.message-stream\s*\{[^}]*min-height:\s*0;[^}]*overflow-y:\s*auto;/);
    expect(sharedStyles).toMatch(/\.message-stream-shell\s*\{[^}]*min-height:\s*0;/);
    expect(sharedStyles).toMatch(/\.chat-shell\s*\{[^}]*minmax\(0, 1fr\)/);
  });

  it("短系统反馈压缩间距，底色只贴在文字上而不是整块白底", () => {
    expect(conversationStyles).toMatch(/\.conversation-message\.system\s*\{\s*margin-bottom:\s*4px;/);
    expect(conversationStyles).toContain(
      ".conversation-message.system > .conversation-reading > .message-text"
    );
    expect(conversationStyles).toContain("background: var(--appearance-reading-bg");
  });

  it("两端都把副本清理结果交给共享的按身份判据", () => {
    expect(appSource).toContain("copyCleanupLoggedRef");
    expect(wpsSource).toContain("copyCleanupLoggedRef");
    expect(appSource).toContain("resultReviewState.copyCleanup");
    expect(wpsSource).toContain("resultReviewState.copyCleanup");
  });
});
