import { describe, expect, it } from "vitest";
import type { ReviewCopyCleanupState, ReviewCopyCloseOutcome } from "./reviewCopyLifecycle";
import {
  REVIEW_CARD_ACTION_LABELS,
  REVIEW_CARD_COMPAT_DETAIL,
  REVIEW_CARD_COMPAT_NOTICE,
  REVIEW_CARD_CONFIRM_LABEL,
  REVIEW_CARD_COPY_STILL_OPEN,
  REVIEW_CARD_DESCRIPTION,
  REVIEW_CARD_DIAGNOSTICS_TITLE,
  REVIEW_CARD_PANE_HINT_TEXT,
  REVIEW_CARD_SECONDARY_ACTION_LABELS,
  REVIEW_CARD_TITLE,
  REVIEW_CARD_VERIFICATION,
  planTargetCells,
  rangeCellCount,
  reviewCardBusiness,
  reviewCardClaimsSuccess,
  reviewCardHasDiagnostics,
  reviewCardMarkdownLeaks,
  reviewCardPhase,
  reviewChangeSummary,
  reviewCardView,
  reviewScopeLine,
  reviewScopeLocations,
  type ReviewCardView,
  type ReviewOutcome
} from "./reviewCardPresentation";

const formatPlanActions = [
  { type: "setFont", sheet: "格式验收", range: "B2:E5", bold: true },
  { type: "setAlignment", sheet: "格式验收", range: "B2:E5", horizontal: "center" }
];

function copyStatusText(outcome: ReviewCopyCloseOutcome): string {
  if (outcome === "closed") return "审阅副本已关闭。";
  if (outcome === "close_unconfirmed") return "关闭尚未确认，请到副本窗口前确认。";
  if (outcome === "failed") return "副本未能自动关闭，请手动关闭。";
  if (outcome === "extra_edits") return "副本里还有未经确认的人工修改，未自动关闭。";
  return `副本状态：${outcome}。`;
}

/** 副本收尾状态的构造器：shape 与 `ReviewCopyCleanupState['copy']` 一致。 */
function copyState(outcome: ReviewCopyCloseOutcome, overrides: Partial<ReviewCopyCleanupState["copy"]> = {}) {
  return {
    outcome,
    headline: `已放弃变更，原表未修改；${copyStatusText(outcome)}`,
    statusText: copyStatusText(outcome),
    detail: "关闭请求已经发出，但窗口仍然存在。",
    tone: "attention" as const,
    retryable: true,
    needsUserChoice: true,
    ...overrides
  };
}

function view(outcome: ReviewOutcome, overrides: Partial<Parameters<typeof reviewCardView>[0]> = {}): ReviewCardView {
  return reviewCardView({
    outcome,
    scope: { locations: "格式验收!B2:E5", changes: "加粗、水平居中", totalCells: 16 },
    compatNotice: null,
    paneNeedsResize: false,
    paneWidthSettable: true,
    copy: null,
    ...overrides
  });
}

describe("审阅卡：阶段与首屏内容", () => {
  it("业务结果只映射成四个阶段，界面分支没有第五种可能", () => {
    expect(reviewCardPhase("reviewing")).toBe("reviewing");
    expect(reviewCardPhase("discarded")).toBe("discarded");
    expect(reviewCardPhase("verified")).toBe("committed");
    // 写回失败与"结果不确定"都落在同一阶段，但要说不同的实话。
    expect(reviewCardPhase("write_failed")).toBe("failed");
    expect(reviewCardPhase("partial_or_unknown")).toBe("failed");
  });

  it("审阅中：标题、说明、范围、确认框、三个按钮，一个不少", () => {
    const card = view("reviewing");
    expect(card.title).toBe(REVIEW_CARD_TITLE.reviewing);
    expect(card.title).toBe("请核对审阅副本");
    expect(card.description).toBe("左侧是审阅副本，右侧是原表。原表尚未修改。");
    expect(card.scopeLine).toBe("格式验收!B2:E5 · 加粗、水平居中 · 涉及 16 个单元格");
    expect(card.confirmLabel).toBe(REVIEW_CARD_CONFIRM_LABEL);
    expect(card.confirmLabel).toBe("我已核对副本，确认按此计划修改原表。");
    expect(card.actions.map((action) => action.label)).toEqual(["放弃变更", "确认写回原表"]);
    expect(REVIEW_CARD_ACTION_LABELS.confirm).toBe("确认写回原表");
    // 首屏没有次要动作，也没有清理状态。
    expect(card.secondaryAction).toBeNull();
    expect(card.copyStatus).toBeNull();
  });

  it("范围行按计划算：面积用计划声明的范围，不用「实际变化格数」", () => {
    expect(reviewScopeLocations(formatPlanActions)).toBe("格式验收!B2:E5");
    expect(planTargetCells(formatPlanActions)).toBe(16);
    expect(reviewChangeSummary(formatPlanActions)).toBe("加粗、水平居中");
    // 多区域去重 + 面积累加。
    expect(reviewScopeLocations([...formatPlanActions,
      { type: "setFont", sheet: "另一张", range: "A1:B2", bold: true }])).toBe("格式验收!B2:E5、另一张!A1:B2");
    expect(planTargetCells([{ type: "setFont", sheet: "s", range: "A1:B2" }])).toBe(4);
    expect(rangeCellCount("A1")).toBe(1);
    expect(rangeCellCount("$B$2:$E$5")).toBe(16);
    // 解析不出来宁可不说，也不编一个数。
    expect(rangeCellCount("整行")).toBeNull();
    expect(rangeCellCount("E5:B2")).toBeNull();
    expect(planTargetCells([{ type: "x", sheet: "s", range: "整行" }])).toBeNull();
    expect(reviewScopeLine({ locations: "S!A1", changes: "", totalCells: null })).toBe("S!A1");
  });

  it("非格式动作也有短摘要；未知类型不沉默", () => {
    expect(reviewChangeSummary([{ type: "writeValues" }, { type: "writeFormulas" }])).toBe("写入数值、写入公式");
    expect(reviewChangeSummary([{ type: "somethingNew" }])).toBe("修改内容");
  });

  it("AI 宽度提示只在**确实未确认**时出现，且是可收起的一组文案", () => {
    expect(view("reviewing").paneHint).toBeNull();
    const hinted = view("reviewing", { paneNeedsResize: true });
    expect(hinted.paneHint?.summary).toBe(REVIEW_CARD_PANE_HINT_TEXT);
    expect(hinted.paneHint?.summary)
      .toBe("如原表显示过窄，可拖动 AI 窗格左边界，让两张表的可见宽度接近。");
    // 具体目标比例收在展开层里，不在首屏那一行。
    expect(hinted.paneHint?.detail).toContain("整个可用宽度约 30%");
    // 宿主差异说在展开里，不在首屏那一行。
    expect(view("reviewing", { paneNeedsResize: true, paneWidthSettable: false }).paneHint?.detail)
      .not.toBe(view("reviewing", { paneNeedsResize: true, paneWidthSettable: true }).paneHint?.detail);
  });

  it("同步滚动只有确认启用才有一行简短状态", () => {
    expect(view("reviewing").syncScrolling).toBeNull();
    expect(view("reviewing", { syncScrollingEnabled: false, syncScrollingStatus: "未启用同步滚动" }).syncScrolling).toBeNull();
    expect(view("reviewing", { syncScrollingEnabled: true, syncScrollingStatus: "已启用同步滚动。" }).syncScrolling)
      .toBe("已启用同步滚动。");
  });

  it("兼容模式必要说明只由宿主喂入：原生模式不冒充有兼容说明", () => {
    expect(view("reviewing").compatNotice).toBeNull();
    expect(view("reviewing", { compatNotice: REVIEW_CARD_COMPAT_NOTICE }).compatNotice)
      .toBe("请在 Excel 中核对副本。确认后将按本计划修改原表，并验证执行结果；副本中的人工修改不会写回。");
  });
});

describe("审阅卡：放弃是终态", () => {
  it("撤下确认框与全部审阅按钮，主卡只讲「计划已取消」", () => {
    const card = view("discarded");
    expect(card.title).toBe("已放弃变更");
    expect(card.description).toBe("本次计划未写入原表。");
    expect(card.confirmLabel).toBeNull();
    expect(card.actions).toEqual([]);
    expect(card.business.headline).toBe("已放弃变更");
    expect(card.business.detail).toBe("本次计划未写入原表。");
    expect(reviewCardClaimsSuccess(card)).toBe(false);
  });

  it("删除冗余口径：不再出现「写入与验证结果不受影响」或「没有需要清理的磁盘文件」", () => {
    const card = view("discarded", { copy: copyState("failed", { detail: "关闭动作本身失败了。", statusText: "副本未能自动关闭，请手动关闭。" }) });
    const text = [
      card.title, card.description, card.business.headline, card.business.detail,
      card.copyStatus?.headline, card.copyStatus?.detail
    ].join(" ");
    expect(text).not.toContain("写入与验证结果不受影响");
    expect(text).not.toContain("没有需要清理的磁盘文件");
    // 但"计划取消"这条结论必须在。
    expect(text).toContain("未写入原表");
  });

  it("副本仍打开时给手动处理顺序，不替用户回答宿主的提示", () => {
    const open = view("discarded", { copy: copyState("failed") });
    expect(open.copyOpenNotice).toBe(REVIEW_CARD_COPY_STILL_OPEN);
    expect(open.copyOpenNotice).toContain("请先另存");
    expect(open.copyOpenNotice).toContain("手动关闭");
    // 「不要无条件让用户选择不保存」：这里不许出现这句。
    expect(open.copyOpenNotice).not.toContain("不保存");
    // 副本已确认收尾就没有这一行了。
    const closedCopy = copyState("closed", { statusText: "审阅副本已关闭。", tone: "success", retryable: false, needsUserChoice: false });
    expect(view("discarded", { copy: closedCopy }).copyOpenNotice).toBeNull();
    expect(view("discarded", { copy: null }).copyOpenNotice).toBeNull();
  });

  it("清理失败是**次要**状态：不冲淡「计划已取消」的结论", () => {
    const card = view("discarded", { copy: copyState("close_unconfirmed") });
    // 主结论仍然是放弃，而且排在次要状态之前。
    expect(card.business.headline).toBe("已放弃变更");
    expect(card.copyStatus?.headline).not.toContain("已放弃变更");
    expect(card.copyStatus?.headline).toBe("关闭尚未确认，请到副本窗口前确认。");
    expect(reviewCardClaimsSuccess(card)).toBe(false);
    // 已放弃阶段没有一键再关：副本里可能有用户自己要留的修改。
    expect(card.secondaryAction).toBeNull();
  });
});

describe("审阅卡：写回成功与失败不套用同一模板", () => {
  it("写回成功：只讲业务结论与副本状态，不保留任何审阅入口", () => {
    const card = view("verified", {
      copy: copyState("closed", { statusText: "审阅副本已关闭。", tone: "success", retryable: false, needsUserChoice: false }),
      executionDetail: "共 2 个动作、16 个单元格已按计划核对。"
    });
    expect(card.phase).toBe("committed");
    expect(card.title).toBe("已写回原表");
    expect(card.actions).toEqual([]);
    expect(card.confirmLabel).toBeNull();
    expect(card.business.headline).toBe("已写回原表");
    expect(card.business.detail).toContain("逐格回读核对");
    // 实际写回与验证结果必须真的在场。
    expect(card.actualResult).toContain("16 个单元格");
    expect(card.copyStatus?.headline).toBe("审阅副本已关闭。");
    expect(reviewCardClaimsSuccess(card)).toBe(true);
    // 副本已收尾 ⇒ 连次要动作都不需要。
    expect(card.secondaryAction).toBeNull();
  });

  it("写回成功但副本没收尾：业务结论不变，另给一个不会重放写入的次要动作", () => {
    const card = view("verified", { copy: copyState("close_unconfirmed") });
    expect(card.business.headline).toBe("已写回原表");
    expect(card.copyStatus?.headline).toContain("关闭尚未确认");
    expect(card.copyOpenNotice).toBe(REVIEW_CARD_COPY_STILL_OPEN);
    expect(card.secondaryAction).toEqual({ id: "closeCopy", label: REVIEW_CARD_SECONDARY_ACTION_LABELS.closeCopy });
    // 写回成功是**业务**结论，副本没收尾不许把它说成失败。
    expect(card.phase).toBe("committed");
    expect(reviewCardClaimsSuccess(card)).toBe(true);
  });

  it("验证未通过 / 部分完成：按实际状态展示，一个字都不套成功模板", () => {
    for (const outcome of ["write_failed", "partial_or_unknown"] as const) {
      const card = view(outcome);
      expect(card.phase).toBe("failed");
      expect(reviewCardClaimsSuccess(card)).toBe(false);
      const text = [card.title, card.description, card.business.headline, card.business.detail].join(" ");
      // 成功模板专属字样一律不许出现。
      expect(text).not.toContain("已写回原表");
      expect(text).not.toContain("验证通过");
      expect(text).not.toContain("本次计划未写入原表");
      // 失败阶段也没有"确认写回"或"放弃变更"这类审阅入口。
      expect(card.actions).toEqual([]);
      expect(card.confirmLabel).toBeNull();
      // 写回失败/不确定只给一个"结束会话、把副本留在现场"的次要动作。
      expect(card.secondaryAction).toEqual({ id: "retainAndEnd", label: REVIEW_CARD_SECONDARY_ACTION_LABELS.retainAndEnd });
    }
    // 两档失败自己也不能互相冒充。
    expect(REVIEW_CARD_TITLE.write_failed).toBe("写回未完成");
    expect(REVIEW_CARD_TITLE.partial_or_unknown).toBe("写回结果不确定");
    expect(REVIEW_CARD_DESCRIPTION.write_failed).not.toBe(REVIEW_CARD_DESCRIPTION.partial_or_unknown);
    expect(REVIEW_CARD_DESCRIPTION.write_failed).toContain("也可能一格未改");
    expect(REVIEW_CARD_DESCRIPTION.partial_or_unknown).toContain("无法确认");
  });

  it("四种业务的文案彼此不重合，且验证档位互不冒充", () => {
    const outcomes = ["reviewing", "discarded", "verified", "write_failed", "partial_or_unknown"] as const;
    const headlines = outcomes.map((outcome) => view(outcome).business.headline);
    expect(new Set(headlines).size).toBe(headlines.length);
    const titles = outcomes.map((outcome) => view(outcome).title);
    expect(new Set(titles).size).toBe(titles.length);
    // 只有 verified 是 passed；"读不到/没写成"不许借成功档位。
    expect(REVIEW_CARD_VERIFICATION.verified).toBe("passed");
    expect(REVIEW_CARD_VERIFICATION.write_failed).toBe("not_passed");
    expect(REVIEW_CARD_VERIFICATION.partial_or_unknown).toBe("unknown");
    expect(REVIEW_CARD_VERIFICATION.discarded).toBe("not_applicable");
    expect(reviewCardBusiness("verified").verification).toBe("passed");
  });
});

describe("审阅卡：信息分层与渲染安全", () => {
  it("诊断详情承载重启提示、身份、关闭诊断与磁盘清理", () => {
    const card = view("discarded", {
      diagnostics: {
        restart: ["若刚更新代码，请先重启 Python 服务并刷新加载项。"],
        identity: ["本次副本工作表名：EB_0123456789abcdef0123"],
        close: ["关闭结果：close_unconfirmed"],
        cleanup: ["本次没有需要清理的磁盘文件。"]
      }
    });
    expect(REVIEW_CARD_DIAGNOSTICS_TITLE).toBe("诊断详情");
    expect(reviewCardHasDiagnostics(card.diagnostics)).toBe(true);
    // 这些都是诊断层的事，不进主卡的业务结论。
    expect(card.business.detail).not.toContain("重启 Python 服务");
    expect(card.business.detail).not.toContain("EB_0123");
    expect(view("reviewing").diagnostics.close).toEqual([]);
    expect(reviewCardHasDiagnostics(view("reviewing").diagnostics)).toBe(false);
  });

  it("任何阶段的可见文案都不带 Markdown 记号（星号会原样显示给用户）", () => {
    const outcomes: ReviewOutcome[] = ["reviewing", "discarded", "verified", "write_failed", "partial_or_unknown"];
    for (const outcome of outcomes) {
      const card = view(outcome, {
        compatNotice: REVIEW_CARD_COMPAT_NOTICE,
        compatDetail: REVIEW_CARD_COMPAT_DETAIL["excel-compat"],
        paneNeedsResize: true,
        paneWidthSettable: false,
        copy: copyState("close_unconfirmed"),
        executionDetail: "共 3 个动作已按计划核对。",
        diagnostics: { restart: ["请重启 Python 服务。"], identity: ["EB_x"], close: ["close_unconfirmed"], cleanup: ["无文件"] }
      });
      expect(reviewCardMarkdownLeaks(card)).toEqual([]);
    }
    // 逐条常量也过一遍，防止将来在常量里写回星号。
    const literals = [
      ...Object.values(REVIEW_CARD_TITLE), ...Object.values(REVIEW_CARD_DESCRIPTION),
      ...Object.values(REVIEW_CARD_ACTION_LABELS), REVIEW_CARD_COMPAT_NOTICE, REVIEW_CARD_CONFIRM_LABEL,
      REVIEW_CARD_COPY_STILL_OPEN, REVIEW_CARD_PANE_HINT_TEXT, REVIEW_CARD_DIAGNOSTICS_TITLE,
      ...Object.values(REVIEW_CARD_SECONDARY_ACTION_LABELS),
      ...Object.values(REVIEW_CARD_COMPAT_DETAIL).flat()
    ];
    for (const literal of literals) expect(literal).not.toMatch(/\*\*|__|`/);
  });
});
