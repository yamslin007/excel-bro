import type { ExcelAction } from "./contracts";

/**
 * Excel 端格式能力的**唯一来源**。
 *
 * 背景（实机故障，2026-10-13）：`ProPlus2024Volume` / `16.0.17932` 这类
 * **批量许可的永久版 Office**，按微软官方要求集表**最高只到 ExcelApi 1.17**，
 * 1.18/1.19/1.20 一律"不支持"。而逐格显示格式的可靠读取依赖 1.19 的
 * `Range.getDisplayedCellProperties` —— 于是旧实现在写前直接拒绝，用户看到的是
 * "无法在执行前可靠保存恢复依据"。
 *
 * 本模块把"这台宿主到底能做什么"抽成**纯函数**，做到三件事：
 *  1. 能力判定**前移到规划/预览阶段**（不等用户确认后才报笼统错误）；
 *  2. 明确说出**缺哪个 API**、**影响哪些动作**；
 *  3. 1.9–1.17 的宿主走**直接格式兼容路径**（`getCellProperties`），而不是一刀切拒绝。
 *
 * 这里刻意不依赖 `Office`/`Excel` 全局：调用方把 `isSetSupported` 传进来，
 * 于是本模块可在任何环境（含没有宿主的单元测试）直接验证。
 */

/**
 * 格式动作**实际会修改的属性**名。
 *
 * 粒度要与"按实际修改属性逐格/逐行采集"对齐：`setFont{bold}` 只碰 `fontBold`，
 * 不该因为 `setFont` 这个动作名就把字体名/字号一起算进"受影响属性"。
 */
export type ExcelFormatField =
  | "fill"
  | "fontName"
  | "fontSize"
  | "fontBold"
  | "fontItalic"
  | "fontColor"
  | "fontUnderline"
  | "horizontalAlignment"
  | "verticalAlignment"
  | "wrapText"
  | "indentLevel"
  | "numberFormat"
  | "borders"
  | "rowHeight"
  | "columnWidth";

/** 逐格格式的恢复路径。 */
export type ExcelFormatPath = "displayed" | "direct" | "unavailable";

/** 逐格显示格式读取所依赖的 API（ExcelApi 1.19）。批量许可永久版拿不到它。 */
export const EXCEL_DISPLAYED_FORMAT_REQUIREMENT =
  "ExcelApi 1.19（Range.getDisplayedCellProperties）";

/** 直接格式兼容路径所依赖的 API（ExcelApi 1.9）。 */
export const EXCEL_DIRECT_FORMAT_REQUIREMENT =
  "ExcelApi 1.9（Range.getCellProperties / setCellProperties / getRowProperties / getColumnProperties）";

/** 六个会改格式的动作（与 `contracts.FORMAT_ACTION_TYPES` 同集合，这里只列类型名）。 */
export const EXCEL_FORMAT_ACTION_TYPES = [
  "setFill",
  "setFont",
  "setBorders",
  "setAlignment",
  "setNumberFormat",
  "resizeRange"
] as const satisfies readonly ExcelAction["type"][];

export interface ExcelFormatCapability {
  /** 逐格格式恢复路径：显示值 / 直接格式 / 不可用。 */
  path: ExcelFormatPath;
  /** 本路径下**缺失**的能力（规划/预览阶段要如实说给用户）。 */
  missingApis: string[];
  /** 会走兼容路径的动作类型（`path === "direct"` 时非空）。 */
  compatActionTypes: ExcelAction["type"][];
  /** 因能力不足会被拒绝的动作类型（`path === "unavailable"` 时非空）。 */
  blockedActionTypes: ExcelAction["type"][];
  /** 面向用户的说明（生成期用，不参与判据）。 */
  note: string;
}

const isFormatActionType = (type: string): boolean =>
  (EXCEL_FORMAT_ACTION_TYPES as readonly string[]).includes(type);

/**
 * 把宿主自报的 ExcelApi 版本翻译成格式能力。
 *
 * 判定用"要求集"而不是"方法是否存在"：`isSetSupported` 是 Office 官方推荐的
 * 运行时判定方式，也是清单与宿主唯一稳定的契约。
 */
export function resolveExcelFormatCapability(
  isSetSupported: (set: string, version: string) => boolean
): ExcelFormatCapability {
  const supportsDisplayed = isSetSupported("ExcelApi", "1.19");
  if (supportsDisplayed) {
    return {
      path: "displayed",
      missingApis: [],
      compatActionTypes: [],
      blockedActionTypes: [],
      note: "本机支持逐格显示格式读取（ExcelApi 1.19），格式动作按执行前的显示状态捕获与恢复。"
    };
  }
  const supportsDirect = isSetSupported("ExcelApi", "1.9");
  if (supportsDirect) {
    return {
      path: "direct",
      missingApis: [EXCEL_DISPLAYED_FORMAT_REQUIREMENT],
      compatActionTypes: [...EXCEL_FORMAT_ACTION_TYPES],
      blockedActionTypes: [],
      note:
        `本机 Excel 不支持 ${EXCEL_DISPLAYED_FORMAT_REQUIREMENT}（批量许可永久版最高到 ExcelApi 1.17）。` +
        "格式动作改走直接格式兼容路径：按实际修改的属性逐格读取直接设置在单元格上的格式，" +
        "撤销时先把目标区域的直接格式清空、再写回捕获到的直接格式集合，因此继承样式与默认格式会自动恢复、" +
        "不会被写成固定的显示值。目标区域若含条件格式或合并单元格，会在写入前明确拒绝。"
    };
  }
  return {
    path: "unavailable",
    missingApis: [EXCEL_DISPLAYED_FORMAT_REQUIREMENT, EXCEL_DIRECT_FORMAT_REQUIREMENT],
    compatActionTypes: [],
    blockedActionTypes: [...EXCEL_FORMAT_ACTION_TYPES],
    note:
      `本机 Excel 既没有 ${EXCEL_DISPLAYED_FORMAT_REQUIREMENT}，也没有 ${EXCEL_DIRECT_FORMAT_REQUIREMENT}，` +
      "无法在执行前保存可恢复的格式依据，因此所有会改格式的动作都会被拒绝（值/公式类动作不受影响）。"
  };
}

/**
 * 一个格式动作**实际会修改**的属性。
 *
 * 只列"这次真的要改"的字段（`!= null` 才是要改，`false` 也是要改），
 * 未给出的属性**不在**结果里 —— 因为"不要动"和"要恢复"是两回事。
 * 非格式动作返回空数组。
 */
export function excelFormatActionFields(action: ExcelAction): readonly ExcelFormatField[] {
  switch (action.type) {
    case "setFill":
      return ["fill"];
    case "setFont": {
      const fields: ExcelFormatField[] = [];
      if (action.bold != null) fields.push("fontBold");
      if (action.italic != null) fields.push("fontItalic");
      if (action.name != null) fields.push("fontName");
      if (action.size != null) fields.push("fontSize");
      if (action.color != null) fields.push("fontColor");
      return fields;
    }
    case "setAlignment": {
      const fields: ExcelFormatField[] = [];
      if (action.horizontal != null) fields.push("horizontalAlignment");
      if (action.vertical != null) fields.push("verticalAlignment");
      if (action.wrapText != null) fields.push("wrapText");
      return fields;
    }
    case "resizeRange": {
      const fields: ExcelFormatField[] = [];
      if (action.rowHeight != null) fields.push("rowHeight");
      if (action.columnWidth != null) fields.push("columnWidth");
      return fields;
    }
    case "setBorders":
      return ["borders"];
    case "setNumberFormat":
      return ["numberFormat"];
    default:
      return [];
  }
}

/** 计划里第一个会被本机能力挡住的格式动作下标（供预检定位到具体步骤）。 */
export function blockedFormatActionIndexes(
  actions: readonly ExcelAction[],
  capability: ExcelFormatCapability
): number[] {
  if (capability.blockedActionTypes.length === 0) return [];
  const blocked = new Set<string>(capability.blockedActionTypes);
  const indexes: number[] = [];
  actions.forEach((action, index) => {
    if (blocked.has(action.type)) indexes.push(index);
  });
  return indexes;
}

/** 计划里会走兼容路径的格式动作下标。 */
export function compatFormatActionIndexes(
  actions: readonly ExcelAction[],
  capability: ExcelFormatCapability
): number[] {
  if (capability.compatActionTypes.length === 0) return [];
  const compat = new Set<string>(capability.compatActionTypes);
  const indexes: number[] = [];
  actions.forEach((action, index) => {
    if (compat.has(action.type)) indexes.push(index);
  });
  return indexes;
}

/**
 * 预览阶段的**非阻断**提示：本机走哪条路径、缺什么 API、哪些动作受影响。
 *
 * 返回 null 表示当前宿主走显示值路径（1.19），无话可说。
 * 这里只生成文字，不改变任何判定 —— 真正的逐区域拒绝在执行前由宿主实读决定。
 *
 * `summary` 是首屏的短说明（一两句），`detail` 是收进「详情」里的技术原因。
 * 两者都是纯文本，不含 Markdown 标记，可直接放进 `<p>`/`<details>`。
 */
export function excelFormatCapabilityNotice(
  actions: readonly ExcelAction[],
  capability: ExcelFormatCapability
): { tone: "info" | "warning"; summary: string; detail: string } | null {
  if (capability.path === "displayed") return null;
  const affected =
    capability.path === "direct"
      ? compatFormatActionIndexes(actions, capability)
      : blockedFormatActionIndexes(actions, capability);
  if (affected.length === 0) return null;
  const steps = affected.map((index) => `第 ${index + 1} 步`).join("、");
  const missing = capability.missingApis.join("、");
  if (capability.path === "direct") {
    return {
      tone: "info",
      summary: `格式动作（${steps}）将走直接格式兼容路径，按实际修改的属性逐格恢复。`,
      detail: `缺失 ${missing}（批量许可永久版最高到 ExcelApi 1.17）。撤销时先把目标区域的直接格式清空、再写回捕获到的直接格式集合，因此继承样式与默认格式会自动恢复；目标区域若含条件格式或合并单元格，会在写入前明确拒绝。`
    };
  }
  return {
    tone: "warning",
    summary: `本机不支持逐格读取格式，以下动作将被拒绝：${steps}。`,
    detail: `缺失 ${missing}，无法在执行前保存可恢复的格式依据，因此所有会改格式的动作都会被拒绝（值/公式类动作不受影响）。`
  };
}

/** 供其他模块复用"这是不是一个会改格式的动作"。 */
export function isExcelFormatActionType(type: string): boolean {
  return isFormatActionType(type);
}

/** 枚举/字符串比较：忽略大小写与 `-`/`_`/空格（`Continuous` vs `continuous`）。 */
function normalizedFormatEnum(value: unknown): string {
  return String(value ?? "")
    .replace(/[-_\s]/g, "")
    .toLocaleLowerCase();
}

function sameFormatColor(left: unknown, right: unknown): boolean {
  return String(left ?? "").trim().toUpperCase() === String(right ?? "").trim().toUpperCase();
}

/**
 * 一个格式动作在**逐格直接格式**上的期望值，字段名与 `excel.ts: directFormatFieldValue` 同构。
 *
 * 只含**逐格**属性：`resizeRange` 是行级/列级、`setNumberFormat` 走 `range.numberFormat`，
 * 都不在逐格矩阵里，因此返回空数组（由调用方另走逐行/逐列与数字格式回读）。
 */
export function expectedDirectFormat(
  action: ExcelAction
): ReadonlyArray<{ field: ExcelFormatField; value: unknown }> {
  switch (action.type) {
    case "setFill":
      // 设了填充色即意味着有填充：pattern 必然是 Solid（与宿主语义一致）。
      return [{ field: "fill", value: { pattern: "Solid", color: action.color } }];
    case "setFont": {
      const expected: Array<{ field: ExcelFormatField; value: unknown }> = [];
      if (action.bold != null) expected.push({ field: "fontBold", value: action.bold });
      if (action.italic != null) expected.push({ field: "fontItalic", value: action.italic });
      if (action.name != null) expected.push({ field: "fontName", value: action.name });
      if (action.size != null) expected.push({ field: "fontSize", value: action.size });
      if (action.color != null) expected.push({ field: "fontColor", value: action.color });
      return expected;
    }
    case "setAlignment": {
      const expected: Array<{ field: ExcelFormatField; value: unknown }> = [];
      if (action.horizontal != null) expected.push({ field: "horizontalAlignment", value: action.horizontal });
      if (action.vertical != null) expected.push({ field: "verticalAlignment", value: action.vertical });
      if (action.wrapText != null) expected.push({ field: "wrapText", value: action.wrapText });
      return expected;
    }
    case "setBorders": {
      const style = action.style ?? "continuous";
      const color = action.color ?? "#000000";
      const weight = action.weight ?? "thin";
      const sides = Object.fromEntries(
        action.sides.map((side) => [side, { style, color, weight }])
      );
      return [{ field: "borders", value: sides }];
    }
    default:
      return [];
  }
}

/**
 * "现场取到的逐格直接格式值"与"动作要求的期望值"是否一致。
 *
 * 沿用执行器回读判据的口径：**枚举忽略大小写/分隔符、颜色忽略大小写、数值允许 0.1 容差**，
 * 其余按原样比较。`actual === null` 表示"该属性未直接设置" —— 对"要求直接设置"的期望
 * 一律判不一致（这正是"未直接设置 ≠ 已设置成默认值"）。
 */
export function formatDirectValueMatches(
  field: ExcelFormatField,
  actual: unknown,
  expected: unknown
): boolean {
  switch (field) {
    case "fill": {
      if (actual == null || expected == null) return actual === expected;
      const seen = actual as { pattern?: unknown; color?: unknown };
      const want = expected as { pattern?: unknown; color?: unknown };
      return (
        normalizedFormatEnum(seen.pattern) === normalizedFormatEnum(want.pattern) &&
        sameFormatColor(seen.color, want.color)
      );
    }
    case "fontBold":
    case "fontItalic":
    case "wrapText":
      return actual === expected;
    case "fontName":
      return actual != null && String(actual) === String(expected);
    case "fontUnderline":
      return normalizedFormatEnum(actual) === normalizedFormatEnum(expected);
    case "fontSize":
      return (
        typeof actual === "number" &&
        typeof expected === "number" &&
        Math.abs(actual - expected) <= 0.1
      );
    case "fontColor":
      return actual != null && sameFormatColor(actual, expected);
    case "horizontalAlignment":
    case "verticalAlignment":
      return normalizedFormatEnum(actual) === normalizedFormatEnum(expected);
    case "indentLevel":
      return actual === expected;
    case "borders": {
      if (actual == null || expected == null) return actual === expected;
      const seen = actual as Record<string, { style?: unknown; color?: unknown; weight?: unknown }>;
      const want = expected as Record<string, { style?: unknown; color?: unknown; weight?: unknown }>;
      return Object.entries(want).every(([side, border]) => {
        const got = seen[side];
        if (!got) return false;
        return (
          normalizedFormatEnum(got.style) === normalizedFormatEnum(border.style) &&
          sameFormatColor(got.color, border.color) &&
          normalizedFormatEnum(got.weight) === normalizedFormatEnum(border.weight)
        );
      });
    }
    default:
      return actual === expected;
  }
}
