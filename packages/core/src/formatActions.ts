/**
 * 「常用格式操作」的宿主无关语义层（L1/L2，Excel / WPS / 文件夹三端共用）。
 *
 * 这一层只回答三个问题，**不碰任何宿主 API**：
 *
 * 1. 这个格式动作**要改哪些属性、改成什么**（`formatActionEdits`）；
 * 2. 改动的**真实影响范围**是什么（`formatActionImpact`）——
 *    行高/列宽是行级/列级属性，落在 `B3:D5` 上真实改的是第 3–5 **整行**、
 *    B–D **整列**，同行同列区域之外的单元格一并受影响。预览里必须这么说，
 *    否则用户看到的"只改这一小块"是假的。
 * 3. 给定「修改前的属性快照」与「修改后的属性快照」，出一份前后对照
 *    （`buildFormatChangeReview`）。
 *
 * 三条不能松的口径：
 * - **未采集 ≠ 未设置**。某个属性这一侧没读到就是 `captured:false`，
 *   界面必须标"属性未采集"，不许拿默认值冒充。（与 `changeReview.ts`
 *   的 `VALUE_NOT_CAPTURED_TEXT` 同一条规矩，只是对象从单元格值换成属性。）
 * - **只改指定属性**。`formatActionEdits` 只列出动作显式声明的项；没有声明的
 *   属性**不出现**在对照里，也不允许被"补成默认值"。
 * - **空的格式动作不存在**。`setFont`/`setAlignment`/`resizeRange` 一个属性都
 *   没给时，两端协议层已经拒绝；这一层不再兜底，直接返回空列表让调用方
 *   看得见异常。
 *
 * 输入用**结构类型**而不是 `ExcelAction`：共享层不允许反向依赖 `apps/*`
 * （见 `packages/core/src/index.ts` 的分层说明）。结构类型天然兼容，
 * 调用方直接把动作对象传进来即可，不需要断言。
 */

import { columnLetterFromIndex } from "./range";
import type { ChangeReviewStatus } from "./changeReview";

/** 本轮纳入「常用格式操作」的动作类型。 */
export type FormatActionType =
  | "setFill"
  | "setFont"
  | "setNumberFormat"
  | "setBorders"
  | "setAlignment"
  | "resizeRange"
  | "autofit";

/**
 * 格式属性键。命名跟着**宿主属性**走（`fontName` 对应 `font.name`），
 * 不跟着动作名走 —— 多个动作可能落到同一个宿主属性上，按宿主属性归并
 * 之后前后对照才不会出现"同一个属性被列两次、两次说法还不一样"。
 */
export type FormatPropertyKey =
  | "fontName"
  | "fontSize"
  | "fontBold"
  | "fontItalic"
  | "fontColor"
  | "fillColor"
  | "numberFormat"
  | "alignmentHorizontal"
  | "alignmentVertical"
  | "wrapText"
  | "borderTop"
  | "borderBottom"
  | "borderLeft"
  | "borderRight"
  | "borderInsideHorizontal"
  | "borderInsideVertical"
  | "rowHeight"
  | "columnWidth";

/** 对照表的固定列序：字体 → 填充 → 数字格式 → 对齐 → 边框 → 行高列宽。 */
export const FORMAT_PROPERTY_ORDER: readonly FormatPropertyKey[] = [
  "fontName",
  "fontSize",
  "fontBold",
  "fontItalic",
  "fontColor",
  "fillColor",
  "numberFormat",
  "alignmentHorizontal",
  "alignmentVertical",
  "wrapText",
  "borderTop",
  "borderBottom",
  "borderLeft",
  "borderRight",
  "borderInsideHorizontal",
  "borderInsideVertical",
  "rowHeight",
  "columnWidth"
];

export const FORMAT_PROPERTY_LABELS: Record<FormatPropertyKey, string> = {
  fontName: "字体",
  fontSize: "字号",
  fontBold: "加粗",
  fontItalic: "斜体",
  fontColor: "字色",
  fillColor: "填充色",
  numberFormat: "数字格式",
  alignmentHorizontal: "水平对齐",
  alignmentVertical: "垂直对齐",
  wrapText: "自动换行",
  borderTop: "上边框",
  borderBottom: "下边框",
  borderLeft: "左边框",
  borderRight: "右边框",
  borderInsideHorizontal: "内部横线",
  borderInsideVertical: "内部竖线",
  rowHeight: "行高",
  columnWidth: "列宽"
};

/** 属性值未采集时给用户看的占位。测试锁住这一串，防止被换成默认值。 */
export const FORMAT_VALUE_NOT_CAPTURED_TEXT = "属性未采集";

export type FormatScalar = string | number | boolean | null;

export type FormatEdit = {
  property: FormatPropertyKey;
  value: FormatScalar;
};

/**
 * 协议侧格式动作的最小结构形状。
 *
 * `sides` / 枚举值一律收成 `string`，因为这一层不做取值合法性校验
 * （那是两端 `assertFormatActionFields` 与 Pydantic `Literal` 的职责）。
 * 在这里重复一遍枚举只会造出第二个真相来源。
 */
export type FormatActionLike =
  | { type: "setFill"; sheet: string; range: string; color: string }
  | {
      type: "setFont";
      sheet: string;
      range: string;
      bold?: boolean | null;
      italic?: boolean | null;
      name?: string | null;
      size?: number | null;
      color?: string | null;
    }
  | { type: "setNumberFormat"; sheet: string; range: string; formatCode: string }
  | {
      type: "setBorders";
      sheet: string;
      range: string;
      sides: readonly string[];
      style?: string | null;
      color?: string | null;
      weight?: string | null;
    }
  | {
      type: "setAlignment";
      sheet: string;
      range: string;
      horizontal?: string | null;
      vertical?: string | null;
      wrapText?: boolean | null;
    }
  | {
      type: "resizeRange";
      sheet: string;
      range: string;
      rowHeight?: number | null;
      columnWidth?: number | null;
    }
  | { type: "autofit"; sheet: string; range: string };

const FORMAT_ACTION_TYPES: ReadonlySet<string> = new Set<FormatActionType>([
  "setFill",
  "setFont",
  "setNumberFormat",
  "setBorders",
  "setAlignment",
  "resizeRange",
  "autofit"
]);

export function isFormatActionType(value: string): value is FormatActionType {
  return FORMAT_ACTION_TYPES.has(value);
}

/**
 * 把协议里的动作收窄成这一层认得的结构形状。
 *
 * 这里有一次 `as` —— 但**不是**在替协议做校验：动作能走到这里，说明它已经过了
 * `assertEditablePlan` / Pydantic 的字段级校验（颜色是 `#RRGGBB`、对齐是闭合
 * 集合里的值、行高在 0–409）。这一层再校一遍只会造出第二个真相来源。
 *
 * 之所以需要它：`ExcelAction` 的 `setBorders.style` 是字面量联合，而这一层
 * 故意收成 `string`（枚举的合法性由协议层负责），两条线在类型上互不赋值，
 * 但结构上完全一致。
 */
export function asFormatAction(action: { type: string }): FormatActionLike | null {
  return isFormatActionType(action.type) ? (action as FormatActionLike) : null;
}

/** 从一份动作列表里挑出格式动作。 */
export function selectFormatActions<T extends { type: string }>(
  actions: readonly T[]
): FormatActionLike[] {
  const selected: FormatActionLike[] = [];
  for (const action of actions) {
    const format = asFormatAction(action);
    if (format) selected.push(format);
  }
  return selected;
}

const BORDER_SIDE_PROPERTY: Record<string, FormatPropertyKey> = {
  top: "borderTop",
  bottom: "borderBottom",
  left: "borderLeft",
  right: "borderRight",
  insideHorizontal: "borderInsideHorizontal",
  insideVertical: "borderInsideVertical"
};

/**
 * 边框描述的展示串。三条边属性必须一起出现 —— 只说"有边框"会让用户
 * 无法判断"原来是实线现在改成虚线"到底改了什么。
 *
 * `style` / `color` / `weight` 缺省值与 Python 侧一致，缺省就在这里补，
 * 免得调用方各补一次、各补得不一样。
 */
export function borderDescription(action: {
  style?: string | null;
  color?: string | null;
  weight?: string | null;
}): string {
  const style = action.style ?? "continuous";
  const color = action.color ?? "#000000";
  const weight = action.weight ?? "thin";
  return `${style} ${weight} ${color}`;
}

/**
 * 一个格式动作**显式声明**要改的属性。
 *
 * 顺序与 `FORMAT_PROPERTY_ORDER` 一致，调用方不必再排一次。
 * 动作类型不在本轮范围内时返回空数组（调用方应先过 `isFormatActionType`）。
 */
export function formatActionEdits(action: FormatActionLike): FormatEdit[] {
  const edits: FormatEdit[] = [];
  switch (action.type) {
    case "setFill":
      edits.push({ property: "fillColor", value: action.color });
      break;
    case "setFont":
      // 只推显式给出的项：`bold: false` 是"取消加粗"，`bold` 缺省是"不要动"，
      // 两者必须区分，所以判空一律用 `!= null`，不能用真值判断。
      if (action.bold != null) edits.push({ property: "fontBold", value: action.bold });
      if (action.italic != null) edits.push({ property: "fontItalic", value: action.italic });
      if (action.name != null) edits.push({ property: "fontName", value: action.name });
      if (action.size != null) edits.push({ property: "fontSize", value: action.size });
      if (action.color != null) edits.push({ property: "fontColor", value: action.color });
      break;
    case "setNumberFormat":
      edits.push({ property: "numberFormat", value: action.formatCode });
      break;
    case "setBorders": {
      const description = borderDescription(action);
      for (const side of action.sides) {
        const property = BORDER_SIDE_PROPERTY[side];
        // 未知的边名不在这里兜底成某条边：那会把"模型写错了"变成"画了一条
        // 用户没要求的边"。协议层的闭合集合校验负责先拦掉它。
        if (property) edits.push({ property, value: description });
      }
      break;
    }
    case "setAlignment":
      if (action.horizontal != null) {
        edits.push({ property: "alignmentHorizontal", value: action.horizontal });
      }
      if (action.vertical != null) {
        edits.push({ property: "alignmentVertical", value: action.vertical });
      }
      if (action.wrapText != null) {
        edits.push({ property: "wrapText", value: action.wrapText });
      }
      break;
    case "resizeRange":
      if (action.rowHeight != null) edits.push({ property: "rowHeight", value: action.rowHeight });
      if (action.columnWidth != null) {
        edits.push({ property: "columnWidth", value: action.columnWidth });
      }
      break;
    case "autofit":
      // 自适应不改"值"，它按内容重算行高与列宽 —— 两个属性都要出现在对照里，
      // 否则用户看不到这一动到底会动什么。
      edits.push({ property: "rowHeight", value: null });
      edits.push({ property: "columnWidth", value: null });
      break;
  }
  return sortEdits(edits);
}

function sortEdits(edits: FormatEdit[]): FormatEdit[] {
  return [...edits].sort(
    (left, right) =>
      FORMAT_PROPERTY_ORDER.indexOf(left.property) -
      FORMAT_PROPERTY_ORDER.indexOf(right.property)
  );
}

/**
 * 该动作的**真实影响范围**。
 *
 * `target` 永远只有动作自己声明的那块区域，但行高/列宽是**行级/列级**属性：
 * 宿主在 `B3:D5` 上设行高，改的是第 3–5 整行。预览若只写"B3:D5"，
 * 就是在骗用户。
 *
 * 区域解析失败（不是明确的 A1）时返回 `null`，调用方**不得**退化成
 * "就只影响这块区域"的说法 —— 说不清楚就不说。
 */
export type FormatImpact =
  | { kind: "range"; sheet: string; range: string }
  | { kind: "rows"; sheet: string; range: string; rows: string }
  | { kind: "columns"; sheet: string; range: string; columns: string }
  | { kind: "rowsAndColumns"; sheet: string; range: string; rows: string; columns: string };

export function formatActionImpact(action: {
  type: string;
  sheet: string;
  range: string;
  rowHeight?: number | null;
  columnWidth?: number | null;
}): FormatImpact | null {
  const base = { sheet: action.sheet, range: action.range };
  if (action.type !== "resizeRange" && action.type !== "autofit") {
    return { kind: "range", ...base };
  }
  const rect = parseRect(action.range);
  if (!rect) return null;
  const rows = rect.startRow === rect.endRow ? `${rect.startRow}` : `${rect.startRow}–${rect.endRow}`;
  const columns = `${columnLetterFromIndex(rect.startColumn - 1)}${
    rect.startColumn === rect.endColumn ? "" : `–${columnLetterFromIndex(rect.endColumn - 1)}`
  }`;
  if (action.type === "resizeRange") {
    // 只有真的改了那一维才声称影响那一维。写了 `columnWidth` 却说"影响整行"
    // 会让用户以为行高也变了。
    const changesRows = hasRowHeight(action);
    const changesColumns = hasColumnWidth(action);
    if (changesRows && changesColumns) {
      return { kind: "rowsAndColumns", ...base, rows, columns };
    }
    if (changesRows) return { kind: "rows", ...base, rows };
    if (changesColumns) return { kind: "columns", ...base, columns };
    return { kind: "range", ...base };
  }
  return { kind: "rowsAndColumns", ...base, rows, columns };
}

function hasRowHeight(action: { rowHeight?: number | null }): boolean {
  return action.rowHeight != null;
}

function hasColumnWidth(action: { columnWidth?: number | null }): boolean {
  return action.columnWidth != null;
}

/** 影响范围给用户的一句话。`null`（区域说不清）走拒绝对话，不返回兜底文案。 */
export function formatImpactNotice(impact: FormatImpact): string {
  const where = `「${impact.sheet}!${impact.range}」`;
  switch (impact.kind) {
    case "range":
      return `只影响 ${where} 内的单元格。`;
    case "rows":
      return `行高是整行属性：${where} 覆盖第 ${impact.rows} 行，这些行**整行**的高度都会变，区域之外同行的单元格一并受影响。`;
    case "columns":
      return `列宽是整列属性：${where} 覆盖第 ${impact.columns} 列，这些列**整列**的宽度都会变，区域之外同列的单元格一并受影响。`;
    case "rowsAndColumns":
      return `行高与列宽是整行/整列属性：${where} 覆盖第 ${impact.rows} 行、第 ${impact.columns} 列，这些行与列**整行整列**都会变，区域之外的同行同列单元格一并受影响。`;
  }
}

function parseRect(range: string): { startRow: number; endRow: number; startColumn: number; endColumn: number } | null {
  const sheetLess = range.includes("!") ? (range.split("!").pop() ?? range) : range;
  const match = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(
    sheetLess.trim()
  );
  if (!match) return null;
  const startColumn = columnIndexOf(match[1]);
  const endColumn = columnIndexOf(match[3] ?? match[1]);
  const startRow = Number(match[2]);
  const endRow = Number(match[4] ?? match[2]);
  if (!startColumn || !endColumn || !startRow || !endRow) return null;
  if (endRow < startRow || endColumn < startColumn) return null;
  return { startRow, endRow, startColumn, endColumn };
}

function columnIndexOf(letters: string): number {
  let value = 0;
  for (const letter of letters.toUpperCase()) {
    value = value * 26 + (letter.charCodeAt(0) - 64);
  }
  return value;
}

/**
 * 属性值的展示文本。
 *
 * 布尔值必须显示成"是/否"而不是 true/false；行高列宽字号带单位，
 * 否则"12"到底是磅还是字符宽没人分得清。
 */
export function formatValueText(property: FormatPropertyKey, value: FormatScalar): string {
  if (value === null) {
    // 两个来源都会走到这里，语义不同但展示相同：动作没给这一项（"不要动"）、
    // 或它不是固定值（autofit 按内容重算）。调用方要区分就去看 edits。
    return "（按内容自适应）";
  }
  if (typeof value === "boolean") return value ? "是" : "否";
  if (typeof value === "number") {
    if (property === "rowHeight") return `${value} 磅`;
    if (property === "fontSize") return `${value} 磅`;
    if (property === "columnWidth") return `${value} 字符宽`;
    return String(value);
  }
  return value === "" ? "（空字符串）" : value;
}

export type FormatSnapshot = {
  /** 缺键 = 这一项没采集到，不是"它是默认值"。 */
  properties: Partial<Record<FormatPropertyKey, FormatScalar>>;
  /** 整份快照有没有采到宿主属性。false 时逐项都按未采集展示。 */
  captured: boolean;
};

export type FormatChangeRow = {
  key: string;
  property: FormatPropertyKey;
  label: string;
  before: FormatScalar;
  after: FormatScalar;
  beforeText: string;
  afterText: string;
  beforeCaptured: boolean;
  afterCaptured: boolean;
  changed: boolean;
};

export type FormatChangeReview = {
  sheet: string;
  range: string;
  status: ChangeReviewStatus;
  /** 只包含**实际比对到的属性**；未采集的属性不出现在这里。 */
  rows: FormatChangeRow[];
  changedCount: number;
  /** 有属性因为两侧都没采集而没法比对时为 true。 */
  hasUncaptured: boolean;
};

export type BuildFormatChangeReviewInput = {
  sheet: string;
  range: string;
  status: ChangeReviewStatus;
  /** 计划声明要改哪些属性。只有这些属性会被比对。 */
  properties: readonly FormatPropertyKey[];
  before: FormatSnapshot;
  after: FormatSnapshot;
};

/**
 * 出「属性 × 修改前/修改后」对照。
 *
 * 只比对 `properties` 里列出的属性 —— 其余属性本次根本没打算动，
 * 把它们也列出来会让用户以为计划要改那些。
 *
 * 两侧只要有一侧没采集到，这一行就标未采集并且**不算改动**：
 * "读不到"不能被当成"没变"，更不能被当成"变了"。
 */
export function buildFormatChangeReview(
  input: BuildFormatChangeReviewInput
): FormatChangeReview {
  const wanted = FORMAT_PROPERTY_ORDER.filter((property) =>
    input.properties.includes(property)
  );
  const rows: FormatChangeRow[] = wanted.map((property) => {
    const beforeCaptured = input.before.captured && property in input.before.properties;
    const afterCaptured = input.after.captured && property in input.after.properties;
    const before = beforeCaptured ? input.before.properties[property] ?? null : null;
    const after = afterCaptured ? input.after.properties[property] ?? null : null;
    const comparable = beforeCaptured && afterCaptured;
    return {
      key: `${input.sheet}!${input.range}#${property}`,
      property,
      label: FORMAT_PROPERTY_LABELS[property],
      before,
      after,
      beforeText: beforeCaptured ? formatValueText(property, before) : FORMAT_VALUE_NOT_CAPTURED_TEXT,
      afterText: afterCaptured ? formatValueText(property, after) : FORMAT_VALUE_NOT_CAPTURED_TEXT,
      beforeCaptured,
      afterCaptured,
      changed: comparable && !sameFormatValue(property, before, after)
    };
  });
  return {
    sheet: input.sheet,
    range: input.range,
    status: input.status,
    rows,
    changedCount: rows.filter((row) => row.changed).length,
    hasUncaptured: rows.some((row) => !row.beforeCaptured || !row.afterCaptured)
  };
}

/**
 * 属性值比较。颜色一律忽略大小写 —— 宿主回读常把 `#AABBCC` 写成 `#aabbcc`，
 * 直接 `!==` 会把它报成一次改动，用户会以为写入没生效。
 * 其余字符串（数字格式码、字体名、对齐枚举）按原样比，不做归一化：
 * 数字格式码里的大小写与空格是有意义的。
 */
function sameFormatValue(
  property: FormatPropertyKey,
  left: FormatScalar,
  right: FormatScalar
): boolean {
  if (typeof left === "string" && typeof right === "string") {
    if (property === "fontColor" || property === "fillColor") {
      return left.toLowerCase() === right.toLowerCase();
    }
    return left === right;
  }
  return left === right;
}

/**
 * 出口摘要：不带任何属性值，只留定位与计数。
 * 与 `changeReviewSummary` 同一条规矩 —— 要进日志/聊天/持久化的只能走这里。
 */
export type FormatChangeSummary = {
  sheet: string;
  range: string;
  status: ChangeReviewStatus;
  changedCount: number;
  propertyCount: number;
  hasUncaptured: boolean;
};

export function formatChangeSummary(review: FormatChangeReview): FormatChangeSummary {
  return {
    sheet: review.sheet,
    range: review.range,
    status: review.status,
    changedCount: review.changedCount,
    propertyCount: review.rows.length,
    hasUncaptured: review.hasUncaptured
  };
}

/**
 * 一份计划里所有格式动作合起来要改的属性（按区域归并）。
 *
 * 同区域上多个动作改同一属性时只保留一条 —— 后写的那个决定最终值，
 * 前面的不认领（与 `add_deterministic_acceptance_criteria` 的
 * `superseded` 判据同一条规矩，只是这里只做归并、不判归属）。
 */
export type FormatPlanPropertyGroup = {
  sheet: string;
  range: string;
  properties: FormatPropertyKey[];
  impact: FormatImpact | null;
};

export function formatPlanPropertyGroups(
  actions: readonly FormatActionLike[]
): FormatPlanPropertyGroup[] {
  const groups = new Map<string, FormatPlanPropertyGroup>();
  for (const action of actions) {
    const edits = formatActionEdits(action);
    if (!edits.length) continue;
    const key = `${action.sheet}\u0000${action.range}`;
    const existing = groups.get(key);
    const properties = new Set(existing?.properties ?? []);
    for (const edit of edits) properties.add(edit.property);
    const impact = formatActionImpact(action);
    groups.set(key, {
      sheet: action.sheet,
      range: action.range,
      properties: FORMAT_PROPERTY_ORDER.filter((property) => properties.has(property)),
      // 同区域多个动作的影响范围取"更宽"的那个：只要有一个动作影响整行，
      // 这条区域的预览就必须说影响整行。
      impact: widerImpact(existing?.impact ?? null, impact)
    });
  }
  return [...groups.values()];
}

const IMPACT_WIDTH: Record<FormatImpact["kind"], number> = {
  range: 0,
  rows: 1,
  columns: 1,
  rowsAndColumns: 2
};

function widerImpact(left: FormatImpact | null, right: FormatImpact | null): FormatImpact | null {
  if (!left) return right;
  if (!right) return left;
  return IMPACT_WIDTH[right.kind] > IMPACT_WIDTH[left.kind] ? right : left;
}

/**
 * 计划里格式动作的**确认前预览**：按区域归组，列出将改的属性与值，
 * 以及**真实影响范围**。
 *
 * `impactNotice` 只在改动会溢出目标区域（行高/列宽/自适应）时非空 ——
 * 那正是"用户以为只改这一小块、实际整行整列都变了"的唯一场景。
 * 单纯的区间内格式（字体、填充、对齐、边框）不重复念一遍区域，
 * 那只会把真正的警告淹掉。
 */
export type FormatPlanPreviewEntry = {
  key: string;
  sheet: string;
  range: string;
  /** 形如「字体 → 微软雅黑」的可读改项。 */
  changes: string[];
  /** 会溢出目标区域的改动才有；否则为 null。 */
  impactNotice: string | null;
};

/**
 * 区域解析不出来时的兜底警告。
 *
 * 这种情况**不能**简单地"说不清就不说"：行高列宽确实会溢出区域，
 * 用户已经要按确认了，此时保持沉默等于默许他按一个错了的理解去执行。
 * 说"范围不确定"比说"只影响这一小块"诚实。
 */
export const FORMAT_IMPACT_UNRESOLVED_TEXT =
  "行高与列宽是整行/整列属性，但这次的目标区域无法解析为明确的 A1 范围，无法说明具体会影响哪些行列。";

export function formatPlanPreview(
  actions: readonly FormatActionLike[]
): FormatPlanPreviewEntry[] {
  const groups = formatPlanPropertyGroups(actions);
  const changeByRange = new Map<string, string[]>();
  for (const action of actions) {
    for (const edit of formatActionEdits(action)) {
      const key = `${action.sheet}\u0000${action.range}`;
      const list = changeByRange.get(key) ?? [];
      const label = FORMAT_PROPERTY_LABELS[edit.property];
      const text = `${label} → ${formatValueText(edit.property, edit.value)}`;
      // 同区域同属性被两个动作都声明时，后一个才是最终值；但这里只做预览，
      // 两个都列出来会让用户看到自相矛盾的两行 —— 去重时保留**最后**一条。
      const existingIndex = list.findIndex((item) => item.startsWith(`${label} → `));
      if (existingIndex >= 0) list[existingIndex] = text;
      else list.push(text);
      changeByRange.set(key, list);
    }
  }
  return groups.map((group) => {
    const key = `${group.sheet}\u0000${group.range}`;
    const overflows = group.properties.some(
      (property) => property === "rowHeight" || property === "columnWidth"
    );
    let impactNotice: string | null = null;
    if (group.impact && group.impact.kind !== "range") {
      impactNotice = formatImpactNotice(group.impact);
    } else if (overflows && !group.impact) {
      impactNotice = FORMAT_IMPACT_UNRESOLVED_TEXT;
    }
    return {
      key: `${group.sheet}!${group.range}`,
      sheet: group.sheet,
      range: group.range,
      changes: changeByRange.get(key) ?? [],
      impactNotice
    };
  });
}
