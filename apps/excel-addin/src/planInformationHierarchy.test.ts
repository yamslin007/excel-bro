import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import {
  PLAN_NOTE_SECTIONS,
  SESSION_RECOVERY_LABEL,
  SESSION_RECOVERY_NOTICE_PREFIX,
  VERIFICATION_PASSED_LABEL
} from "@excel-bro/core";

/**
 * 任务窗格信息层级的"两端口径"账。
 *
 * Excel 侧的计划卡/结果卡写在几千行的 App.tsx 里，没有便宜的整体渲染测试；
 * WPS 侧已经用真实 DOM 断言同一批要点（见 WpsAgentPanel.test.tsx）。
 * 所以这里做的是契约级断言：同一句话必须在两端**来自同一个常量**，
 * 而不是两份各自硬编码、靠 review 维持一致的中文。
 * 界面行为的证据仍以 WPS 侧 DOM 测试和实机验收为准。
 */
const appSource = readFileSync("src/App.tsx", "utf8");
const wpsSource = readFileSync("../wps-addin/src/WpsAgentPanel.tsx", "utf8");
/**
 * WPS 结果卡的标题口径在 `wpsExecutionRecovery.ts`（纯判定 + 文案）里，
 * 面板只负责渲染。只查面板源码会漏掉"常量搬了家"这类漂移，
 * 所以这里同时读模块，并**双向**要求：模块引用共享常量、面板消费该模块。
 */
const wpsRecoverySource = readFileSync("../wps-addin/src/wpsExecutionRecovery.ts", "utf8");
const excelStyles = readFileSync("src/styles.css", "utf8");
const wpsStyles = readFileSync("../wps-addin/src/styles.css", "utf8");

describe("任务窗格信息层级：Excel 与 WPS 的同状态口径", () => {
  it("计划卡的两个笔记区块在两端来自同一份常量", () => {
    expect(PLAN_NOTE_SECTIONS.assumptions).toBe("采用的信息");
    expect(PLAN_NOTE_SECTIONS.warnings).toBe("补充说明");
    expect(appSource).toContain("PLAN_NOTE_SECTIONS.assumptions");
    expect(appSource).toContain("PLAN_NOTE_SECTIONS.warnings");
    expect(wpsSource).toContain("PLAN_NOTE_SECTIONS.assumptions");
    expect(wpsSource).toContain("PLAN_NOTE_SECTIONS.warnings");
    // 两端都不许再把标题写死在 JSX 里 —— 写死就等于放弃了"同一份口径"。
    expect(appSource).not.toContain(">采用的信息<");
    expect(appSource).not.toContain(">补充说明<");
    expect(wpsSource).not.toContain(">采用的信息<");
    expect(wpsSource).not.toContain(">补充说明<");
  });

  it("恢复能力只写成会话内口径，不承诺跨关闭撤销", () => {
    expect(SESSION_RECOVERY_LABEL).toContain("会话内");
    expect(SESSION_RECOVERY_NOTICE_PREFIX).toContain("会话内");
    expect(wpsSource).toContain("SESSION_RECOVERY_LABEL");
    expect(wpsSource).toContain("SESSION_RECOVERY_NOTICE_PREFIX");
    // 裸的"可撤销"会让人以为恢复点是持久的 —— 模块侧同样不许出现字面量。
    expect(wpsSource).not.toContain(">可撤销<");
    expect(wpsSource).not.toContain('"会话内可撤销"');
    expect(wpsRecoverySource).not.toContain('"会话内可撤销"');
  });

  it("「验证通过」只留给真的核对过的结果", () => {
    expect(VERIFICATION_PASSED_LABEL).toBe("验证通过");
    expect(appSource).toContain("VERIFICATION_PASSED_LABEL");
    // WPS 侧：模块引用共享常量，面板必须真的消费那份映射（否则模块是死代码）。
    expect(wpsRecoverySource).toContain("VERIFICATION_PASSED_LABEL");
    expect(wpsSource).toContain("WPS_EXECUTION_OUTCOME_LABEL");
    expect(wpsSource).toContain("./wpsExecutionRecovery");
    // Excel 的三态：verified 之外必须是"未独立验证"/"未通过"，不能借"验证通过"。
    expect(appSource).toContain("已执行，部分操作未独立验证");
    expect(appSource).toContain("执行完成，但验证未通过");
    // WPS 只有"核对过"与"已撤销"两态；写失败时根本不该出现结果卡
    // （这条由 WpsAgentPanel.test.tsx 的渲染断言把关）。
    expect(wpsSource).toContain("已撤销并核对恢复");
  });

  it("两端都有对应的卡片样式，且窄窗下笔记区块不溢出", () => {
    expect(excelStyles).toMatch(/\.inline-notes\s*\{[\s\S]*?\}/);
    expect(excelStyles).toMatch(/\.verification-card\s*\{[\s\S]*?\}/);
    expect(wpsStyles).toMatch(/\.wps-plan-notes\s*\{[\s\S]*?\}/);
    expect(wpsStyles).toMatch(/\.wps-plan-card p\.wps-plan-revised\s*\{[\s\S]*?\}/);
    // 430px 屏上动作行改单列、笔记缩进收窄，否则会被挤到计划卡外面。
    expect(wpsStyles).toMatch(
      /@media \(max-width: 430px\)[\s\S]*?\.wps-plan-notes ul\s*\{[\s\S]*?padding-left:\s*16px;/
    );
  });
});
