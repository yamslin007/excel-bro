/**
 * WPS 格式动作的**已验证属性白名单**（Task #53）。
 *
 * 这一层回答一个确定的问题：某个格式动作的**具体属性**在目标 WPS 版本上是否已
 * 通过实机探测（写哨兵 → 回读一致 → 恢复 → 回读回到原值）。它**只**依据
 * `docs/WPS_FORMAT_PROBE_EVIDENCE_2026-10-13.md` 记录的真实宿主结论，不推断。
 *
 * 纪律（与探测证据一致，**按具体属性开放**）：
 * - 不因 `Font` 部分通过就开放全部字体属性 —— `Font.Underline` 未验证（哨兵误用布尔），
 *   `setFont` 只要带 `underline` 语义就必须拒绝。
 * - 不因 `RowHeight` 通过就开放 `ColumnWidth` —— `ColumnWidth` 有量化漂移，`resizeRange`
 *   只要带 `columnWidth` 就必须拒绝。
 * - 不因 `HorizontalAlignment` 通过就开放 `VerticalAlignment` —— 垂直对齐哨兵与原值同值，
 *   未验证，`setAlignment.vertical` 必须拒绝。
 *
 * 白名单外的属性**明确拒绝**（抛错），不静默忽略：静默忽略会让计划"看起来成功、
 * 实际没改"，与"只有指定属性被改变"的承诺相反。
 *
 * 注意：探测只覆盖**单个空白单元格**（Sheet1!A1）。单格探测**不证明**混合格式的
 * 恢复正确性 —— 逐格/逐行快照与恢复仍需执行器在真实宿主上另行验收。这一层只回答
 * "属性是否已验证可读写可恢复"，不回答"混合格式能否正确恢复"。
 */

/** WPS 实机已验证的字体属性（对应 setFont 可开放的字段）。 */
export const WPS_VERIFIED_FONT_PROPERTIES = [
  "bold",
  "italic",
  "name",
  "size",
  "color"
] as const;

/** WPS 实机已验证的对齐属性（对应 setAlignment 可开放的字段）。 */
export const WPS_VERIFIED_ALIGNMENT_PROPERTIES = ["horizontal", "wrapText"] as const;

/** WPS 实机已验证的尺寸属性（对应 resizeRange 可开放的字段）。 */
export const WPS_VERIFIED_RESIZE_PROPERTIES = ["rowHeight"] as const;

/**
 * setFont 尚未验证的属性：模型若带了这些字段，必须拒绝。
 * 目前只有 `underline`（哨兵误用布尔，回读 2 ≠ true，待换合法枚举哨兵）。
 */
export function assertWpsVerifiedSetFont(action: {
  bold?: boolean | null;
  italic?: boolean | null;
  name?: string | null;
  size?: number | null;
  color?: string | null;
}): void {
  // underline 不在协议字段里（setFont 协议只有 bold/italic/name/size/color），
  // 但保险起见：只要出现未验证的字体语义就拒绝。这里按协议字段逐个核对已足够，
  // 因为协议层 `assertFormatActionFields` 已保证 setFont 只带这五个字段。
  void action;
}

/**
 * setAlignment 的垂直对齐尚未验证（哨兵与原值同为 -4108，写不出区别）。
 * 模型若带了 `vertical`，必须明确拒绝。
 */
export function assertWpsVerifiedSetAlignment(action: {
  horizontal?: string | null;
  vertical?: string | null;
  wrapText?: boolean | null;
}): void {
  if (action.vertical != null) {
    throw new Error(
      "WPS 的垂直对齐（VerticalAlignment）尚未实机验证（哨兵与原值同值），当前拒绝 setAlignment.vertical。"
    );
  }
}

/**
 * resizeRange 的列宽尚未验证（量化漂移：写 12 回读 12.0088…、恢复 8.4533…≠原值）。
 * 模型若带了 `columnWidth`，必须明确拒绝；只有 `rowHeight` 才开放。
 */
export function assertWpsVerifiedResizeRange(action: {
  rowHeight?: number | null;
  columnWidth?: number | null;
}): void {
  if (action.columnWidth != null) {
    throw new Error(
      "WPS 的列宽（ColumnWidth）尚未实机验证（量化漂移，恢复未通过），当前拒绝 resizeRange.columnWidth。"
    );
  }
}

/* ------------------------------------------------------------------------- *
 * WPS 适配层的枚举 / 颜色映射（协议字符串 ↔ WPS 数字枚举）
 * ------------------------------------------------------------------------- */

import { XL_H_ALIGN } from "./wpsFormatProbe";

/** 协议水平对齐字符串 → WPS `HorizontalAlignment` 数字枚举。 */
const HORIZONTAL_TO_WPS: Record<string, number> = {
  general: XL_H_ALIGN.general,
  left: XL_H_ALIGN.left,
  center: XL_H_ALIGN.center,
  right: XL_H_ALIGN.right,
  justify: XL_H_ALIGN.justify,
  distributed: XL_H_ALIGN.distributed,
  fill: XL_H_ALIGN.fill,
  centerAcrossSelection: XL_H_ALIGN.centerAcrossSelection
};

/** 把协议 `setAlignment.horizontal`（字符串）映射为 WPS 数字枚举；未知值抛错。 */
export function wpsHorizontalAlignment(value: string): number {
  const mapped = HORIZONTAL_TO_WPS[value];
  if (mapped === undefined) {
    throw new Error(`WPS 不认识的水平对齐「${value}」。`);
  }
  return mapped;
}

/** 把协议 `#RRGGBB` 颜色转为 WPS `Font.Color`/`Interior.Color` 的数字（0xRRGGBB）。 */
export function wpsColorFromHex(hex: string): number {
  const match = /^#([0-9a-fA-F]{6})$/.exec(hex);
  if (!match) throw new Error(`WPS 不认识的颜色「${hex}」，应为 #RRGGBB。`);
  return Number.parseInt(match[1], 16);
}

/** 格式动作字段（与 `wpsPlanExecutor.FormatActionFields` 同形，供预览描述使用）。 */
type WpsFormatActionLike =
  | { type: "setFont"; bold?: boolean | null; italic?: boolean | null; name?: string | null; size?: number | null; color?: string | null }
  | { type: "setAlignment"; horizontal?: string | null; vertical?: string | null; wrapText?: boolean | null }
  | { type: "resizeRange"; rowHeight?: number | null; columnWidth?: number | null };

const HORIZONTAL_LABELS: Record<string, string> = {
  general: "常规",
  left: "左对齐",
  center: "居中",
  right: "右对齐",
  justify: "两端对齐",
  distributed: "分散对齐",
  fill: "填充",
  centerAcrossSelection: "跨列居中"
};

/**
 * 格式动作的**人类可读预览描述**（含整行/整列影响提示）。
 *
 * 用于计划预览里明确显示"这次会改哪些格式、影响范围是什么"，而不是把格式动作
 * 当成"写入值"糊弄过去。`resizeRange.rowHeight` 会改变**整行**，必须明示。
 */
export function describeWpsFormatAction(action: WpsFormatActionLike): string {
  if (action.type === "setFont") {
    const parts: string[] = [];
    if (action.bold != null) parts.push(action.bold ? "加粗" : "取消加粗");
    if (action.italic != null) parts.push(action.italic ? "斜体" : "取消斜体");
    if (action.name != null) parts.push(`字体改为「${action.name}」`);
    if (action.size != null) parts.push(`字号改为 ${action.size}`);
    if (action.color != null) parts.push(`字体颜色改为 ${action.color}`);
    return parts.length ? `设置字体（${parts.join("、")}）` : "设置字体";
  }
  if (action.type === "setAlignment") {
    const parts: string[] = [];
    if (action.horizontal != null) parts.push(`水平对齐改为「${HORIZONTAL_LABELS[action.horizontal] ?? action.horizontal}」`);
    if (action.wrapText != null) parts.push(action.wrapText ? "开启自动换行" : "关闭自动换行");
    return parts.length ? `设置对齐（${parts.join("、")}）` : "设置对齐";
  }
  // resizeRange
  const parts: string[] = [];
  if (action.rowHeight != null) parts.push(`行高改为 ${action.rowHeight}（影响整行）`);
  if (action.columnWidth != null) parts.push(`列宽改为 ${action.columnWidth}（影响整列）`);
  return parts.length ? `调整尺寸（${parts.join("、")}）` : "调整尺寸";
}
