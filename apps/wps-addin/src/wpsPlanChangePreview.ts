import { isFormatActionType } from "./wpsPlanExecutor";
import { columnLetterFromIndex } from "@excel-bro/core/range";
import { parseExplicitA1Rect } from "../../excel-addin/src/planCore";
import type { CellValue } from "../../excel-addin/src/contracts";
import type { WpsPreparedPlan } from "./wpsPlanExecutor";
import { describeWpsFormatAction } from "./wpsFormatCapability";

/**
 * 确认卡默认可见的「位置 / 修改前 / 修改后」预览。
 *
 * 为什么需要它：确认卡此前把"预计写入内容"整段收进 `<details>`，首屏只剩标题、
 * 目标格数与一串状态说明。用户要在按下「生成右侧副本并排审阅」之前做决定，
 * 却看不到**到底改了什么**。这里把决策必需的最小事实提到首屏。
 *
 * 三条不能松的口径：
 * 1. 只有**真实拿到的原值**才能写进「修改前」。首选来源是 `beforeValues` ——
 *    它来自 `prepareWpsPlan` 生成预览时对目标区域的那**一次**读取，与
 *    `previewFingerprint` 同源。宿主没交出值（`beforeValuesCaptured === false`）
 *    时如实标 `unknown`，不猜、不拿公式顶替。
 * 2. 只取**前几行几列做代表**，并如实报告"共 N 格、此处展示代表样例"。抽样
 *    绝不能被读成全部，所以 `truncated` 必须能被调用方拿到。
 * 3. 不做**额外**宿主读取：全部数据来自已经在内存里的 `WpsPreparedPlan`。
 */

/** 「修改前」的一格。`unknown` = 没有拿到原值，**不许**编造。 */
export type PlanPreviewBefore =
  | { kind: "value"; text: string }
  | { kind: "unknown" };

export type PlanPreviewRow = {
  /** 稳定 key：`定位地址` 已经是唯一的。 */
  key: string;
  address: string;
  before: PlanPreviewBefore;
  after: PlanPreviewBefore;
  changed: boolean;
};

export type PlanPreviewAction = {
  key: string;
  label: string;
  sheetName: string;
  range: string;
  rows: PlanPreviewRow[];
  /**
   * 格式动作（setFont/setAlignment/resizeRange）的人类可读描述（含整行/整列影响
   * 提示）。存在时，首屏不再显示"修改前 → 修改后"的值矩阵对照（格式动作不改值），
   * 而是显示这条格式描述；`rows` 为 `[]`。值动作没有此字段。
   */
  formatDescription?: string;
  /** 这块区域一共多少格（不代表 rows.length）。 */
  totalCells: number;
  /** 是否只展示了代表性子集。 */
  truncated: boolean;
};

export type PlanChangePreview = {
  actions: PlanPreviewAction[];
  totalCells: number;
};

/** 单元格值 → 展示文本。`null` 与 `""` 是**两个**状态，不得折叠成同一个。 */
export function planPreviewCellText(value: CellValue): string {
  if (value === null) return "（空）";
  if (value === "") return "（空字符串）";
  return String(value);
}

/** 动作类型 → 用户能读的说法。与计划卡其它地方共用同一套词汇。 */
export function planActionLabel(type: string): string {
  if (type === "clearRange") return "清空内容";
  if (type === "writeFormulas") return "写入公式";
  if (type === "setFont") return "设置字体";
  if (type === "setFill") return "设置填充";
  if (type === "setNumberFormat") return "设置数字格式";
  if (type === "setBorders") return "设置边框";
  if (type === "setAlignment") return "设置对齐";
  if (type === "resizeRange") return "调整尺寸";
  if (type === "writeValues") return "写入值";
  return `其他操作（${type}）`;
}

function valueMatrixFor(
  action: WpsPreparedPlan["actions"][number]
): CellValue[][] {
  // 清空内容没有目标值：写入后这块就是空白。仍按矩阵形状给出（空串），
  // 好让用户看到"这几格会被清掉"。
  if (action.type === "clearRange") {
    return Array.from({ length: action.rowCount }, () =>
      Array.from({ length: action.columnCount }, () => "" as CellValue)
    );
  }
  return action.values;
}

/**
 * 生成确认卡首屏的代表性预览。
 *
 * @param plan 已准备好、尚未写入的计划。**不产生任何宿主调用**。
 * @param maxRows 每个动作最多展示多少行（默认 3）。列数最多 3 列。
 */
export function buildPlanChangePreview(
  plan: WpsPreparedPlan,
  maxRows = 3
): PlanChangePreview {
  const actions: PlanPreviewAction[] = plan.actions.map((action) => {
    // 格式动作不改值：首屏不显示"旧值 → 旧值"的值矩阵对照（那是误导），
    // 而是显示格式描述（含整行/整列影响提示）。`formatAction` 字段由
    // `prepareFormatAction` 在预览时写入，永远存在。
    if (isFormatActionType(action.type)) {
      return {
        key: `${action.actionIndex}:${action.sheetName}:${action.range}`,
        label: planActionLabel(action.type),
        sheetName: action.sheetName,
        range: action.range,
        rows: [],
        formatDescription: describeWpsFormatAction(
          action.formatAction ?? (action as unknown as Parameters<typeof describeWpsFormatAction>[0])
        ),
        totalCells: action.cellCount,
        truncated: false
      };
    }
    const rect = parseExplicitA1Rect(
      action.range.includes("!") ? (action.range.split("!").pop() ?? action.range) : action.range
    );
    const before = action.beforeValuesCaptured ? action.beforeValues : null;
    const after = valueMatrixFor(action);
    const rows: PlanPreviewRow[] = [];
    const shownRows = Math.max(0, Math.min(maxRows, action.rowCount));
    const shownColumns = Math.max(0, Math.min(3, action.columnCount));

    for (let rowIndex = 0; rowIndex < shownRows; rowIndex += 1) {
      for (let columnIndex = 0; columnIndex < shownColumns; columnIndex += 1) {
        const address = rect
          ? `${columnLetterFromIndex(rect.startColumn - 1 + columnIndex)}${
              rect.startRow + rowIndex
            }`
          : `${action.range}#${rowIndex + 1}-${columnIndex + 1}`;
        const beforeValue = before?.[rowIndex]?.[columnIndex] ?? null;
        const afterValue = after[rowIndex]?.[columnIndex] ?? null;
        const beforeCell: PlanPreviewBefore = before
          ? { kind: "value", text: planPreviewCellText(beforeValue) }
          : { kind: "unknown" };
        const afterText = planPreviewCellText(afterValue);
        rows.push({
          key: address,
          address,
          before: beforeCell,
          after: { kind: "value", text: afterText },
          changed: beforeCell.kind !== "value" || beforeCell.text !== afterText
        });
      }
    }

    return {
      key: `${action.actionIndex}:${action.sheetName}:${action.range}`,
      label: planActionLabel(action.type),
      sheetName: action.sheetName,
      range: action.range,
      rows,
      totalCells: action.cellCount,
      truncated: action.rowCount > shownRows || action.columnCount > shownColumns
    };
  });

  return { actions, totalCells: plan.totalCells };
}
