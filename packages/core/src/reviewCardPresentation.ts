/**
 * 审阅卡的**信息层级与文案口径**（Excel / WPS **共用**）。
 *
 * 为什么要有这一层：两端的审阅卡是各自写的 JSX，以往各自写死同一句中文，
 * 改一处就悄悄漂移。这里把"哪个阶段说什么、哪条信息放在哪一层"收成一份
 * 可测的纯函数，两端只负责把自己拿到的**事实**喂进来。
 *
 * **边界（不可越过）**：
 *  - 这里只做展示口径。执行算法、撤销算法、关闭/清理判据都不在这里 ——
 *    它们在 `reviewCopyLifecycle` 与各端适配器里。
 *  - 宿主能力差异（能不能由代码设窗格宽度、兼容副本能不能回读）通过**入参**
 *    表达，不在这里假装两端一样。
 *  - 文案一律是**纯文本**：这些串直接进 DOM 文本节点，出现 `**粗体**` 之类的
 *    Markdown 记号会原样显示星号给用户。
 *
 * **四条阶段纪律**：
 *  1. 审阅中首屏只留：标题、说明、范围、兼容模式必要说明、确认框、三个按钮；
 *     其余信息（诊断、身份、清理、重启提示）一律下沉。
 *  2. 已放弃是**终态**：立刻撤下确认框与全部审阅按钮，主卡只讲"计划已取消"。
 *  3. 写回成功后只讲业务结论与副本状态，不再保留任何审阅入口。
 *  4. 验证未通过 / 部分完成 **绝不套成功模板** —— 阶段的唯一来源是
 *     `outcome`，不是界面想说什么。
 */

import {
  FORMAT_PROPERTY_LABELS,
  formatActionEdits,
  formatValueText,
  isFormatActionType,
  asFormatAction,
  type FormatPropertyKey,
  type FormatScalar
} from './formatActions';
import { columnIndexFromLetter } from './range';
import capabilities from '../../../config/capabilities.json';
import type {
  ReviewCopyCleanupState,
  ReviewCopyCloseOutcome,
  ReviewCopyStatusTone
} from './reviewCopyLifecycle';

/** 审阅卡的四个阶段。界面分支**只能**从这里取值。 */
export type ReviewCardPhase = 'reviewing' | 'discarded' | 'committed' | 'failed';

/**
 * 业务结果的原始取值。
 *
 * 它比 `ReviewCardPhase` 细：`write_failed` 与 `partial_or_unknown` 都落在
 * `failed` 阶段，但要说不同的实话 —— 前者"没写成"，后者"写成什么样不知道"。
 */
export type ReviewOutcome = 'reviewing' | 'discarded' | 'verified' | 'write_failed' | 'partial_or_unknown';

export type ReviewCardVerification = 'passed' | 'not_passed' | 'partial' | 'unknown' | 'not_applicable';

/**
 * 阶段标题。
 *
 * 按**业务结果**取值而不是按阶段：`write_failed` 与 `partial_or_unknown` 同属
 * `failed` 阶段，但"没写成"和"写成什么样不知道"必须说成两句不同的话。
 * 标题同时也是主卡里那条**业务结论**——不再另起一行重复一遍。
 */
export const REVIEW_CARD_TITLE: Record<ReviewOutcome, string> = {
  reviewing: '请核对审阅副本',
  discarded: '已放弃变更',
  verified: '已写回原表',
  write_failed: '写回未完成',
  partial_or_unknown: '写回结果不确定'
};

/** 标题下面那一行。审阅中是布局说明，终态是**结论的实际内容**。 */
export const REVIEW_CARD_DESCRIPTION: Record<ReviewOutcome, string> = {
  reviewing: '左侧是审阅副本，右侧是原表。原表尚未修改。',
  discarded: '本次计划未写入原表。',
  verified: '原表已按本卡审阅的计划写入，并通过执行侧逐格回读核对。',
  write_failed: '原表可能只改了一部分，也可能一格未改。请以执行结果与恢复点为准，不要重复提交。',
  partial_or_unknown: '部分动作已经执行，但整体结果无法确认。既不要当成成功，也不要重复提交；先核对原表与恢复点。'
};

/** 业务验证结论。**只有** `passed` 才有资格让界面说"验证通过"。 */
export const REVIEW_CARD_VERIFICATION: Record<ReviewOutcome, ReviewCardVerification> = {
  reviewing: 'not_applicable',
  discarded: 'not_applicable',
  verified: 'passed',
  write_failed: 'not_passed',
  partial_or_unknown: 'unknown'
};

/**
 * 兼容（人工核对）模式下**必须**出现在首屏的说明。
 *
 * 它是决策需要的：兼容副本无法被加载项自动回读，所以"确认"意味着
 * "你自己在宿主里看过了"。少这一句，用户会以为加载项已经替他核对过。
 *
 * 只换宿主名，句子结构两端**逐字同源** —— 各写一句就会漂移成
 * "Excel 说了 WPS 没说"。
 */
export function reviewCardCompatNotice(hostName: string): string {
  return `请在 ${hostName} 中核对副本。确认后将按本计划修改原表，并验证执行结果；副本中的人工修改不会写回。`;
}

export const REVIEW_CARD_COMPAT_NOTICE = reviewCardCompatNotice('Excel');

export const REVIEW_CARD_CONFIRM_LABEL = '我已核对副本，确认按此计划修改原表。';

/** 首屏按钮标签。放弃与写回成功后不得再显示审阅操作入口。 */
export const REVIEW_CARD_ACTION_LABELS = {
  discard: '放弃变更',
  confirm: '确认写回原表'
} as const;

export type ReviewCardActionId = keyof typeof REVIEW_CARD_ACTION_LABELS;

export type ReviewCardAction = { id: ReviewCardActionId; label: string };

/** 业务结论（只讲**原表**；副本的事由 `copyStatus` 负责）。 */
export type ReviewCardBusiness = { headline: string; detail: string; verification: ReviewCardVerification };

/**
 * 业务结论。
 *
 * 标题与说明就是结论本身（`REVIEW_CARD_TITLE` / `REVIEW_CARD_DESCRIPTION`），
 * 这里只是把它连同验证档位打包，供**对话面**（WPS 把结论追加到会话里）使用。
 * 主卡直接渲染标题与说明，不再重复一遍 headline。
 */
export function reviewCardBusiness(outcome: Exclude<ReviewOutcome, 'reviewing'>): ReviewCardBusiness {
  return {
    headline: REVIEW_CARD_TITLE[outcome],
    detail: REVIEW_CARD_DESCRIPTION[outcome],
    verification: REVIEW_CARD_VERIFICATION[outcome]
  };
}

/**
 * 副本仍打开时的手动处理指引。
 *
 * 「不要无条件让用户选择『不保存』」：副本里可能有**用户自己要留**的修改，
 * 一句"选不保存"会把它们一起抹掉。所以这里只给顺序（先另存，再关），
 * 不替用户回答宿主的提示。
 */
export const REVIEW_CARD_COPY_STILL_OPEN =
  '审阅副本仍打开，请核对窗口名称后手动关闭。若副本中有需要保留的修改，请先另存。';

/** AI 宽度确实未确认时才出现的**可收起**提示。 */
export const REVIEW_CARD_PANE_HINT_TEXT =
  '如原表显示过窄，可拖动 AI 窗格左边界，让两张表的可见宽度接近。';

export const REVIEW_CARD_DIAGNOSTICS_TITLE = '诊断详情';
export const REVIEW_CARD_COMPAT_DETAIL_TITLE = '兼容模式说明';

/**
 * 兼容模式折叠里的**技术解释**（格式恢复 + 无法自动回读副本）。
 *
 * 放在折叠里而不是首屏：它是"为什么"，不是"要不要按确认"。
 */
export const REVIEW_CARD_COMPAT_DETAIL: Record<'excel-compat' | 'wps-native', readonly string[]> = {
  'excel-compat': [
    '副本是本机生成 OOXML 后由 Excel 打开的未保存工作簿：格式与公式都写在文件内容里，Excel 打开时自行重算，加载项不参与重算。',
    '兼容宿主没有窗口句柄，也读不回副本单元格内容，所以加载项无法自动核对副本，只能由你在 Excel 里用眼睛核对。',
    '确认后加载项只执行本卡已审阅的计划：副本中的人工修改不会进入原表。'
  ],
  'wps-native': [
    '副本由 WPS 的 Worksheet.Copy() 生成（未保存工作簿），加载项不对副本做额外的列宽贴近。',
    '公式动作在副本上试算并回读计算值；计算值不符时不会拿副本冒充结果，也不会据此写入原表。',
    '确认后加载项只执行本卡已审阅的计划：副本中的人工修改不会进入原表。'
  ]
};

export type ReviewCardScope = {
  /** 目标位置，形如 `格式验收!B2:E5`（多区域用「、」连接）。 */
  locations: string;
  /** 改了什么，形如 `加粗、水平居中`。 */
  changes: string;
  /** 计划涉及的目标单元格总数；范围解析不出来时为 null。 */
  totalCells: number | null;
};

export type ReviewCardPaneHint = {
  /** 首屏那一行（可收起）。 */
  summary: string;
  /** 展开后的宿主差异说明。 */
  detail: string;
};

export type ReviewCardStatusLine = {
  headline: string;
  detail: string;
  tone: ReviewCopyStatusTone;
  /** 允许用户再点一次「关闭审阅副本」。 */
  retryable: boolean;
  /** 需要用户自己决定（如何处理已另存的文件 / 是否丢弃副本）。 */
  needsUserChoice: boolean;
};

/**
 * 副本是否**仍打开**、需要用户手动处理。
 *
 * 只有"已确认收尾"的两种结果才算不用管；其余（含"未尝试"）一律按仍打开对待 ——
 * 宁可多提醒一句，也不让用户以为副本已经被处理掉了。
 */
export function reviewCopyStillOpen(outcome: ReviewCopyCloseOutcome): boolean {
  return outcome !== 'closed' && outcome !== 'already_closed';
}

/** 终态阶段里**唯一**的次要动作；审阅中与已放弃阶段没有次要动作。 */
export type ReviewCardSecondaryAction = { id: 'closeCopy' | 'retainAndEnd'; label: string };

export const REVIEW_CARD_SECONDARY_ACTION_LABELS = {
  /** 写回已成功、副本还没收尾：允许再关一次（错点也不会重放写入）。 */
  closeCopy: '关闭审阅副本',
  /** 写回失败/不确定：结束面板会话，把副本留在现场给用户核对。 */
  retainAndEnd: '保留副本并结束审阅'
} as const;

export type ReviewCardDiagnostics = {
  /** 开发环境 / 服务重启提示：**只在**诊断入口出现。 */
  restart: string[];
  /** 排列与同步滚动的未启用指引。 */
  arrangement: string[];
  /** 随机副本标识、窗口名等（用于人工核对窗口）。 */
  identity: string[];
  /** 关闭诊断。 */
  close: string[];
  /** 磁盘清理结果。 */
  cleanup: string[];
};

export type ReviewCardView = {
  phase: ReviewCardPhase;
  outcome: ReviewOutcome;
  title: string;
  description: string;
  scopeLine: string | null;
  compatNotice: string | null;
  /** 「兼容模式说明」折叠里的技术解释；原生模式只有宿主差异那几条。 */
  compatDetail: readonly string[];
  /** 审阅中才有；其余阶段必须为 null。 */
  confirmLabel: string | null;
  /** 首屏按钮；放弃/写回成功后必须为空数组。 */
  actions: readonly ReviewCardAction[];
  paneHint: ReviewCardPaneHint | null;
  /** 业务结论（只讲原表）。审阅中就是那条布局说明，不另行重复。 */
  business: ReviewCardBusiness;
  /** 验证结论。`passed` 才有资格出现"验证通过"字样。 */
  verification: ReviewCardVerification;
  /** 执行器给出的**实际**结果证据（动作数、核对结论、失败原因）；没有则 null。 */
  actualResult: string | null;
  /** 副本仍打开时的手动处理指引。 */
  copyOpenNotice: string | null;
  /** 次要清理状态：独立显示，**不得**冲淡业务结论。 */
  copyStatus: ReviewCardStatusLine | null;
  /**
   * 终态里唯一的次要动作。
   *
   * 已放弃阶段**没有**：那时副本里可能有用户自己要留的修改，卡片只给手动顺序，
   * 不提供一键再关一次。
   */
  secondaryAction: ReviewCardSecondaryAction | null;
  /** 同步滚动**只有确认启用**才有一行简短状态；未启用一律 null。 */
  syncScrolling: string | null;
  diagnostics: ReviewCardDiagnostics;
};

export function reviewCardPhase(outcome: ReviewOutcome): ReviewCardPhase {
  if (outcome === 'reviewing') return 'reviewing';
  if (outcome === 'discarded') return 'discarded';
  if (outcome === 'verified') return 'committed';
  return 'failed';
}

/** 只有审阅中允许出现确认框与审阅按钮。 */
export function reviewCardAllowsReviewActions(phase: ReviewCardPhase): boolean {
  return phase === 'reviewing';
}

/** 只有业务验证通过才算"成功"；其余一律不许套成功模板。 */
export function reviewCardClaimsSuccess(view: Pick<ReviewCardView, 'outcome' | 'verification'>): boolean {
  return view.outcome === 'verified' && view.verification === 'passed';
}

export function reviewScopeLine(scope: ReviewCardScope | null): string | null {
  if (!scope) return null;
  const parts = [scope.locations, scope.changes].filter((part) => part.trim() !== '');
  if (scope.totalCells !== null) parts.push(`涉及 ${scope.totalCells} 个单元格`);
  return parts.join(' · ');
}

/**
 * 计划"改了什么"的**首屏短摘要**。
 *
 * 口径：布尔属性说属性名（`加粗`），对齐类说值本身（`水平居中`），
 * 其余说「属性 值」。同一属性只出现一次，顺序按计划里第一次出现的先后。
 */
export function reviewChangeSummary(actions: readonly { type: string }[]): string {
  const seen = new Set<string>();
  const labels: string[] = [];
  const push = (label: string) => {
    if (seen.has(label)) return;
    seen.add(label);
    labels.push(label);
  };
  for (const action of actions) {
    if (isFormatActionType(action.type)) {
      const format = asFormatAction(action);
      if (!format) continue;
      for (const edit of formatActionEdits(format)) push(formatEditShortLabel(edit.property, edit.value));
      continue;
    }
    push(ACTION_TYPE_LABELS[action.type] ?? '修改内容');
  }
  return labels.join('、');
}

const ACTION_TYPE_LABELS: Record<string, string> = {
  writeValues: '写入数值',
  writeFormulas: '写入公式',
  clearRange: '清空内容',
  createWorksheet: '新建工作表',
  deleteWorksheet: '删除工作表',
  activateWorksheet: '切换工作表',
  renameWorksheet: '重命名工作表',
  sortRange: '排序',
  applyFilter: '筛选',
  clearFilter: '清除筛选'
};

/** 少数几项直接说值比说属性更清楚的属性。 */
const FORMAT_VALUE_SHORT_LABELS: Record<string, string> = {
  'alignmentHorizontal:center': '水平居中',
  'alignmentHorizontal:left': '左对齐',
  'alignmentHorizontal:right': '右对齐',
  'alignmentVertical:center': '垂直居中',
  'alignmentVertical:top': '顶端对齐',
  'alignmentVertical:bottom': '底端对齐',
  'alignmentVertical:justify': '两端对齐'
};

function formatEditShortLabel(property: FormatPropertyKey, value: FormatScalar): string {
  const label = FORMAT_PROPERTY_LABELS[property];
  if (typeof value === 'boolean') return value ? label : `取消${label}`;
  const keyed = FORMAT_VALUE_SHORT_LABELS[`${property}:${String(value)}`];
  return keyed ?? `${label} ${formatValueText(property, value)}`;
}

/**
 * 目标位置的展示串：`Sheet1!A1`；多区域用「、」连接。
 * 同一个 `sheet!range` 重复出现时只算一次。
 */
export function reviewScopeLocations(actions: readonly { type: string; sheet?: string; range?: string }[]): string {
  const seen = new Set<string>();
  const locations: string[] = [];
  for (const action of actions) {
    if (!action.sheet || !action.range) continue;
    const location = `${action.sheet}!${action.range}`;
    if (seen.has(location)) continue;
    seen.add(location);
    locations.push(location);
  }
  return locations.join('、');
}

/**
 * 计划涉及的目标单元格总数。
 *
 * 用**计划声明的范围面积**，不用"实际变化格数"：后者在格式动作上会得到
 * 不同的数，两端就会各报一个数字。范围解析不出来时返回 null —— 宁可不说，
 * 也不编一个数。
 */
export function planTargetCells(actions: readonly { type: string; sheet?: string; range?: string }[]): number | null {
  const seen = new Set<string>();
  let total = 0;
  let parsed = 0;
  for (const action of actions) {
    if (!action.sheet || !action.range) continue;
    const location = `${action.sheet}!${action.range}`;
    if (seen.has(location)) continue;
    seen.add(location);
    const count = rangeCellCount(action.range);
    if (count === null) return null;
    total += count;
    parsed += 1;
  }
  return parsed === 0 ? null : total;
}

/** A1 记法的面积（含 `$` 与单格写法）；解析不出来返回 null。 */
export function rangeCellCount(address: string): number | null {
  const match = /^\s*\$?([A-Za-z]{1,3})\$?(\d{1,7})\s*(?::\s*\$?([A-Za-z]{1,3})\$?(\d{1,7}))?\s*$/.exec(address);
  if (!match) return null;
  const startColumn = columnIndexFromLetter(match[1].toUpperCase());
  const startRow = Number(match[2]);
  const endColumn = match[3] ? columnIndexFromLetter(match[3].toUpperCase()) : startColumn;
  const endRow = match[4] ? Number(match[4]) : startRow;
  if (endColumn < startColumn || endRow < startRow) return null;
  return (endColumn - startColumn + 1) * (endRow - startRow + 1);
}

export type ReviewCardViewInput = {
  outcome: ReviewOutcome;
  scope: ReviewCardScope | null;
  /** 兼容（人工核对）模式的必要说明；原生模式传 null。 */
  compatNotice: string | null;
  /** **真实回读结论**：AI 窗格宽度确实未确认。false ⇒ 首屏不出现任何调宽提示。 */
  paneNeedsResize: boolean;
  /** 宿主能否由代码设置 AI 窗格宽度（Excel 能；WPS 不能）。 */
  paneWidthSettable: boolean;
  /** 副本收尾状态；未尝试时为 null。 */
  copy: ReviewCopyCleanupState['copy'] | null;
  /** 执行器给出的**实际**结果证据；没有则 null。 */
  executionDetail?: string | null;
  /** 「兼容模式说明」折叠里的技术解释。 */
  compatDetail?: readonly string[];
  /** 同步滚动是否已确认启用。true 才显示简短状态。 */
  syncScrollingEnabled?: boolean;
  syncScrollingStatus?: string | null;
  diagnostics?: Partial<ReviewCardDiagnostics>;
};

const EMPTY_DIAGNOSTICS: ReviewCardDiagnostics = {
  restart: [], arrangement: [], identity: [], close: [], cleanup: []
};

/**
 * 把各端的事实折成审阅卡视图。
 *
 * 唯一的"阶段"来源是 `outcome`：界面想过什么不算数。写回阶段若没有拿到
 * `passed` 的验证结论，`reviewCardClaimsSuccess` 会返回 false，调用方据此
 * 禁止渲染任何成功字样。
 */
export function reviewCardView(input: ReviewCardViewInput): ReviewCardView {
  const phase = reviewCardPhase(input.outcome);
  const reviewing = phase === 'reviewing';
  const verification = REVIEW_CARD_VERIFICATION[input.outcome];
  const diagnostics = { ...EMPTY_DIAGNOSTICS, ...input.diagnostics };
  const copyStatus: ReviewCardStatusLine | null = !reviewing && input.copy
    ? {
      headline: input.copy.statusText,
      detail: input.copy.detail,
      tone: input.copy.tone,
      retryable: input.copy.retryable,
      needsUserChoice: input.copy.needsUserChoice
    }
    : null;
  const copyOpen = !reviewing && input.copy !== null && reviewCopyStillOpen(input.copy.outcome);
  return {
    phase,
    outcome: input.outcome,
    title: REVIEW_CARD_TITLE[input.outcome],
    description: REVIEW_CARD_DESCRIPTION[input.outcome],
    scopeLine: reviewScopeLine(input.scope),
    compatNotice: input.compatNotice,
    compatDetail: input.compatDetail ?? [],
    // 确认框是"审阅中"独有的：放弃之后它必须立刻消失，否则用户会以为还能写回。
    confirmLabel: reviewing ? REVIEW_CARD_CONFIRM_LABEL : null,
    actions: reviewing
      ? [
        { id: 'discard', label: REVIEW_CARD_ACTION_LABELS.discard },
        { id: 'confirm', label: REVIEW_CARD_ACTION_LABELS.confirm }
      ]
      : [],
    paneHint: input.paneNeedsResize
      ? { summary: REVIEW_CARD_PANE_HINT_TEXT, detail: paneHintDetail(input.paneWidthSettable) }
      : null,
    business: input.outcome === 'reviewing'
      ? { headline: REVIEW_CARD_TITLE.reviewing, detail: REVIEW_CARD_DESCRIPTION.reviewing, verification }
      : reviewCardBusiness(input.outcome),
    verification,
    actualResult: input.executionDetail ?? null,
    copyOpenNotice: copyOpen ? REVIEW_CARD_COPY_STILL_OPEN : null,
    copyStatus,
    secondaryAction: secondaryActionFor(phase, copyStatus),
    syncScrolling: reviewing && input.syncScrollingEnabled ? (input.syncScrollingStatus ?? null) : null,
    diagnostics
  };
}

function secondaryActionFor(phase: ReviewCardPhase, copyStatus: ReviewCardStatusLine | null): ReviewCardSecondaryAction | null {
  // 写回成功后的副本是纯粹的遗留物：允许再关一次，且这个动作不会被重放成写入。
  if (phase === 'committed') {
    return copyStatus?.retryable
      ? { id: 'closeCopy', label: REVIEW_CARD_SECONDARY_ACTION_LABELS.closeCopy }
      : null;
  }
  // 写回失败/不确定：结束面板会话，把副本留在现场。
  if (phase === 'failed') return { id: 'retainAndEnd', label: REVIEW_CARD_SECONDARY_ACTION_LABELS.retainAndEnd };
  return null;
}

function paneHintDetail(paneWidthSettable: boolean): string {
  const percent = Math.round(capabilities.reviewPaneSizing.reviewAiWidthFraction * 100);
  return paneWidthSettable
    ? `把 AI 调到整个可用宽度约 ${percent}%，让两张表的可视宽度接近。宿主没有确认窗格宽度时，加载项不会替你声称"已按目标比例排好"。`
    : `当前宿主不能由代码设置 AI 窗格宽度，只能手动拖动到可用宽度约 ${percent}%。加载项不会声称已经自动调到目标比例。`;
}

export function reviewCardHasDiagnostics(diagnostics: ReviewCardDiagnostics): boolean {
  return Object.values(diagnostics).some((group) => group.length > 0);
}

/**
 * 渲染前的最后一道闸：文案里**不许**出现 Markdown 记号。
 *
 * 这些串直接进 DOM 文本节点，星号会原样显示。返回违规项便于测试直接指出是哪一条。
 */
export function reviewCardMarkdownLeaks(view: ReviewCardView): string[] {
  const leaks: string[] = [];
  const check = (label: string, text: string | null) => {
    if (text && /\*\*|__|`/.test(text)) leaks.push(`${label}: ${text}`);
  };
  check('title', view.title);
  check('description', view.description);
  check('scopeLine', view.scopeLine);
  check('compatNotice', view.compatNotice);
  check('confirmLabel', view.confirmLabel);
  check('business.headline', view.business.headline);
  check('business.detail', view.business.detail);
  check('actualResult', view.actualResult);
  check('copyOpenNotice', view.copyOpenNotice);
  check('copyStatus.headline', view.copyStatus?.headline ?? null);
  check('copyStatus.detail', view.copyStatus?.detail ?? null);
  check('paneHint.summary', view.paneHint?.summary ?? null);
  check('paneHint.detail', view.paneHint?.detail ?? null);
  view.compatDetail.forEach((entry, index) => check(`compatDetail[${index}]`, entry));
  for (const action of view.actions) check(`action.${action.id}`, action.label);
  check('secondaryAction', view.secondaryAction?.label ?? null);
  for (const [group, entries] of Object.entries(view.diagnostics)) {
    entries.forEach((entry, index) => check(`diagnostics.${group}[${index}]`, entry));
  }
  return leaks;
}
