import type { WpsApplication, WpsRange, WpsWorksheet } from "./wpsHost";
import { reportDevProbe } from "./devProbe";

/**
 * WPS 格式能力受控探测（Task #47 的「开发诊断入口」）。
 *
 * 目的：`wpsHost.ts` 里 `WpsRange` 的格式成员是**类型补齐**（依据官方文档写成），
 * 但**尚未在目标 WPS 版本上实机逐项验证**。类型里写了 ≠ 宿主真的可用。
 *
 * 这份探测对每个格式成员做**五态判定**，绝不把"成员存在"冒充成"已验证"：
 *
 *   - `member_exists`（成员存在）：读不抛、不是 undefined/null。只说明"有这个字段"。
 *   - `write_verified`（写入验证通过）：写入合法哨兵后，回读**恰好等于哨兵**。
 *     哨兵必须**与执行前原值不同**，且是**该成员在 WPS 接口里的合法值**（见下方
 *     枚举常量表）；否则"回读一致"不能证明写入有效。
 *   - `restore_verified`（恢复验证通过）：把原值写回后，**再次回读并比较**，
 *     必须回到原值。**没有回读的恢复不算恢复** —— 静默失败必须暴露。
 *   - `skipped`（未执行）：如 AutoFit 只有"可调用"可查、没有可回读的哨兵，
 *     就标 `skipped`（未实际执行），**不计入通过数**。
 *   - `failed`（失败）：任何一步失败，并带具体失败阶段与原因。
 *
 * **关联属性与整轮基线**：填充（Interior）与边框（Borders）这类属性是**一组**——
 * 改其中一个字段可能连带改动同组其它字段（例如设 `Pattern` 会改 `PatternColor`）。
 * 因此探测按"关联组"保存**整组基线**，恢复后核对**整组**是否回到基线，而不只是
 * 恢复过的那个字段。**一旦某个成员恢复失败或恢复状态不确定，立即停止后续所有
 * 写入探测**，剩余成员标 `skipped`，避免在已被污染的状态上继续写入。
 * 最终报告区分"单项恢复通过"（`restore_verified`）与"整轮基线恢复通过"
 * （`baselineRestored`）。
 *
 * 探测目标必须是**独立测试工作簿里的单个空白单元格**（无值、无公式），
 * 运行前由面板校验并让用户确认。**这里只探测、只上报，不做任何业务写入。**
 * 结果不构成"已验证" —— 最终验收必须由用户在真实 WPS 里跑完并逐项核对。
 */

/* ------------------------------------------------------------------------- *
 * WPS ET 枚举常量（依据 open.wps.cn / solution.wps.cn / developer.kdocs.cn）
 * ------------------------------------------------------------------------- */

/** XlHAlign —— Range.HorizontalAlignment 的合法值。 */
export const XL_H_ALIGN = {
  general: 1,
  left: -4131,
  center: -4108,
  right: -4152,
  justify: -4130,
  distributed: -4117,
  fill: 5,
  centerAcrossSelection: 7
} as const;

/** XlVAlign —— Range.VerticalAlignment 的合法值。 */
export const XL_V_ALIGN = {
  top: -4160,
  center: -4108,
  bottom: -4107,
  justify: -4130,
  distributed: -4117
} as const;

/** XlPattern —— Interior.Pattern 的合法值（数字，不是字符串）。 */
export const XL_PATTERN = {
  none: -4142,
  solid: 1
} as const;

/** XlLineStyle —— Border.LineStyle 的合法值。 */
export const XL_LINE_STYLE = {
  none: -4142,
  continuous: 1,
  dash: -4115,
  dot: -4118,
  double: -4119
} as const;

/** XlBordersIndex —— Borders.Item(index) 的合法索引。 */
export const XL_BORDERS_INDEX = {
  edgeLeft: 7,
  edgeTop: 8,
  edgeBottom: 9,
  edgeRight: 10
} as const;

/** XlBorderWeight —— Border.Weight 的合法值。 */
export const XL_BORDER_WEIGHT = {
  hairline: 1,
  thin: 2,
  medium: -4138,
  thick: 4
} as const;

/* ------------------------------------------------------------------------- *
 * 探测结果模型
 * ------------------------------------------------------------------------- */

/** 单个成员的五态结论。 */
export type WpsFormatMemberStatus =
  | "member_exists" // 成员存在，但未做写入/恢复验证（哨兵法不适用时）
  | "write_verified" // 写入 + 回读一致（但注意：尚未证明恢复）
  | "restore_verified" // 写入验证 + 恢复后回读一致，全部通过
  | "skipped" // 未执行（如 AutoFit 未实际调用，或前置失败导致后续停止）
  | "failed"; // 某一步失败

/** 单个成员探测结论。 */
export type WpsFormatMemberProbe = {
  member: string;
  status: WpsFormatMemberStatus;
  /** 失败时：具体失败在哪一步。 */
  failedAt: "read" | "write" | "readback" | "restore" | "restore_readback" | "baseline" | null;
  /** 写入使用的哨兵值（只读展示，便于人工核对是否合法）。 */
  sentinel: unknown;
  /** 探测前读到的原值（只读展示）。 */
  original: unknown;
  /** 一句话说明/失败原因。 */
  note: string;
};

/** 一整轮探测的汇总。 */
export type WpsFormatProbeReport = {
  host: "wps-et";
  hostVersion: string | null;
  workbookName: string;
  sheetName: string;
  probeCell: string;
  startedAt: string;
  /** 探测目标身份（面板确认绑定到这里；换簿/换表后失效）。 */
  target: WpsFormatProbeTarget;
  /** 本轮实际探测的能力组（按探测顺序）。留痕，便于把证据按组归档。 */
  groups: WpsFormatProbeGroup[];
  members: WpsFormatMemberProbe[];
  /** 整轮基线是否恢复通过：只有**所有**关联组都回到基线才为 true。 */
  baselineRestored: boolean;
  /** 是否因"恢复失败/不确定"而提前停止后续写入探测。 */
  aborted: boolean;
  /** 汇总计数（供面板一键展示）。 */
  summary: {
    total: number;
    restoreVerified: number;
    writeVerified: number;
    memberExists: number;
    skipped: number;
    failed: number;
  };
};

/** 探测目标身份（确认绑定的锚点）。 */
export type WpsFormatProbeTarget = {
  workbookName: string;
  workbookPath: string | null;
  sheetName: string;
  probeCell: string;
};

/* ------------------------------------------------------------------------- *
 * 能力组（可选探测的入口）
 * ------------------------------------------------------------------------- */

/**
 * 探测的**能力组**。同一轮可以只探测选中的组，使填充/边框这类此前因前置
 * 失败被 abort 而未测到的组，能在**新的独立测试基线**上单独跑，不必先让
 * 前面所有组都通过。
 *
 * 组边界与产品动作对齐（便于把「探测通过」映射到「接入哪个动作」）：
 * - `font`：setFont 依赖的 Font.Bold/Italic/Name/Size/Color/Underline
 * - `alignment`：setAlignment 依赖的 Horizontal/Vertical/WrapText
 * - `fill`：setFill 依赖的 Interior.Color/Pattern
 * - `borders`：setBorders 依赖的 Borders.Item(LineStyle/Weight/ColorIndex)
 * - `size`：resizeRange 依赖的 RowHeight/ColumnWidth
 * - `autofit`：autofit 依赖的 EntireRow/EntireColumn.AutoFit（只探测可调用）
 */
export type WpsFormatProbeGroup =
  | "font"
  | "alignment"
  | "fill"
  | "borders"
  | "size"
  | "autofit";

/** 全部能力组（默认探测范围）。顺序即探测顺序。 */
export const WPS_FORMAT_PROBE_GROUPS: readonly WpsFormatProbeGroup[] = [
  "font",
  "alignment",
  "size",
  "fill",
  "borders",
  "autofit"
];

/** 能力组的中文展示名（面板用）。 */
export const WPS_FORMAT_PROBE_GROUP_LABELS: Record<WpsFormatProbeGroup, string> = {
  font: "字体（加粗/斜体/名称/字号/颜色/下划线）",
  alignment: "对齐与换行（水平/垂直/自动换行）",
  fill: "填充（背景色/图案）",
  borders: "边框（线型/粗细/颜色）",
  size: "尺寸（行高/列宽）",
  autofit: "自适应（仅探测可调用，不实际执行）"
};

/* ------------------------------------------------------------------------- *
 * 读取辅助
 * ------------------------------------------------------------------------- */

/** 读一个值；宿主没暴露（undefined/null）或抛错都算"不可读"。 */
function safeRead(read: () => unknown): { ok: boolean; value: unknown } {
  try {
    const value = read();
    if (value === undefined || value === null) return { ok: false, value };
    return { ok: true, value };
  } catch {
    return { ok: false, value: undefined };
  }
}

/** 判断两个值是否"探测等价"（数字与布尔严格相等；其余走 Object.is，NaN 归一）。 */
function sameValue(left: unknown, right: unknown): boolean {
  if (typeof left === "number" && typeof right === "number") {
    return Number.isNaN(left) ? Number.isNaN(right) : left === right;
  }
  return Object.is(left, right);
}

/* ------------------------------------------------------------------------- *
 * 关联属性组
 * ------------------------------------------------------------------------- */

/**
 * 关联属性组：一组可能互相连带改变的格式字段。
 *
 * 探测前读整组基线；组内每个成员单独做写入验证；**恢复后核对整组基线**，
 * 而不只是恢复过的那个字段。任何一组成员恢复失败/不确定 → 整轮 abort。
 */
type ProbeGroup = {
  /** 组名（用于报告与定位）。 */
  name: string;
  /** 组内成员定义：名字 + 读 + 写 + 哨兵。 */
  members: Array<{
    name: string;
    read: () => unknown;
    write: (value: unknown) => void;
    sentinel: unknown;
  }>;
  /** 整组基线读取：返回一个可比较的"组指纹"（数组，逐字段）。 */
  baseline: () => unknown[];
};

/* ------------------------------------------------------------------------- *
 * 目标资格校验（运行前）
 * ------------------------------------------------------------------------- */

/**
 * 读取目标的当前身份（工作簿路径 + 工作表名 + 目标地址）。
 * 面板用它把"确认"绑定到真实目标身份。
 */
export function readWpsFormatProbeTargetIdentity(
  application: WpsApplication,
  probeCell: string
): WpsFormatProbeTarget | null {
  const workbook = application.ActiveWorkbook;
  if (!workbook) return null;
  const sheet = application.ActiveSheet ?? workbook.ActiveSheet;
  if (!sheet) return null;
  const fullName = String(workbook.FullName ?? workbook.Path ?? "").trim();
  return {
    workbookName: String(workbook.Name || "未命名工作簿"),
    workbookPath: fullName || null,
    sheetName: String(sheet.Name),
    probeCell
  };
}

/**
 * 校验探测目标：必须是**独立测试工作簿**里的**单个空白单元格**（无值、无公式）。
 *
 * 注意：
 * - ".xlsx 后缀 + 空单元格"**不构成**"独立测试工作簿"的证明 —— 那只是**必要条件**。
 *   真正的"独立可丢弃"由用户在面板上**明确确认**（绑定目标身份），本函数只做技术资格校验。
 * - 无法读取目标的值/公式时（`Value2` 读取抛错、`Formula` 读取抛错）**一律拒绝**，
 *   不能把"读取能力缺失"一概解释为"空白"。
 */
export function assertWpsFormatProbeTarget(
  application: WpsApplication,
  probeCell: string
): { ok: true; target: WpsFormatProbeTarget } | { ok: false; reason: string } {
  const workbook = application.ActiveWorkbook;
  if (!workbook) return { ok: false, reason: "WPS 表格中当前没有打开的工作簿。" };
  const sheet = application.ActiveSheet ?? workbook.ActiveSheet;
  if (!sheet) return { ok: false, reason: "无法定位当前工作表。" };

  const fullName = String(workbook.FullName ?? workbook.Path ?? "").trim();
  if (!/\.xlsx$/i.test(fullName)) {
    return {
      ok: false,
      reason: "探测只对已保存为 .xlsx 的工作簿开放（当前工作簿未保存或非 .xlsx）。"
    };
  }

  // 必须是单格地址：A1 形式，且不能是 "A1:B2" 这种区域。
  const singleCell = /^\$?[A-Za-z]{1,3}\$?\d+$/.test(probeCell);
  if (!singleCell) {
    return { ok: false, reason: `探测目标「${probeCell}」不是单个单元格地址，请用 A1 形式（如 A1、$B$3）。` };
  }

  const range = sheet.Range?.(probeCell);
  if (!range) {
    return { ok: false, reason: `无法读取探测单元格「${sheet.Name}!${probeCell}」。` };
  }

  // 无值：Value2 必须可读且为空。读取抛错 → 拒绝（不能把读取缺失当空白）。
  let value: unknown;
  try {
    value = range.Value2;
  } catch (error) {
    return {
      ok: false,
      reason: `无法读取探测单元格「${sheet.Name}!${probeCell}」的值（${error instanceof Error ? error.message : "读取抛错"}），无法确认其为空。`
    };
  }
  if (value !== undefined && value !== null && value !== "") {
    return {
      ok: false,
      reason: `探测单元格「${sheet.Name}!${probeCell}」有值（${String(value).slice(0, 20)}），请改用空白单元格。`
    };
  }

  // 无公式：Formula 必须可读且不是 "=" 开头。读取抛错 → 拒绝。
  let formula: unknown;
  try {
    formula = range.Formula;
  } catch (error) {
    return {
      ok: false,
      reason: `无法读取探测单元格「${sheet.Name}!${probeCell}」的公式（${error instanceof Error ? error.message : "读取抛错"}），无法确认其无公式。`
    };
  }
  if (typeof formula === "string" && formula.startsWith("=")) {
    return {
      ok: false,
      reason: `探测单元格「${sheet.Name}!${probeCell}」含公式（${formula.slice(0, 20)}），请改用空白单元格。`
    };
  }

  const target: WpsFormatProbeTarget = {
    workbookName: String(workbook.Name || "未命名工作簿"),
    workbookPath: fullName || null,
    sheetName: String(sheet.Name),
    probeCell
  };

  return { ok: true, target };
}

/** 判断两份目标身份是否"同一目标"（确认绑定失效的依据）。 */
export function sameProbeTarget(left: WpsFormatProbeTarget, right: WpsFormatProbeTarget): boolean {
  return (
    left.workbookPath === right.workbookPath &&
    left.sheetName === right.sheetName &&
    left.probeCell === right.probeCell
  );
}

/* ------------------------------------------------------------------------- *
 * 单个成员探测（组内）
 * ------------------------------------------------------------------------- */

/**
 * 探测组内单个成员："读原值 → 写哨兵 → 回读哨兵 → 写回原值 → 回读原值"。
 *
 * 返回 `restoreState`：`restored`（确认回到原值）/ `uncertain`（写回抛错或回读不确定）。
 * 调用方（组）据 `uncertain` 决定是否 abort 整轮。
 */
type MemberProbeResult = {
  probe: WpsFormatMemberProbe;
  restoreState: "restored" | "uncertain" | "failed";
};

function probeGroupMember(
  name: string,
  read: () => unknown,
  write: (value: unknown) => void,
  sentinel: unknown
): MemberProbeResult {
  const original = safeRead(read);
  if (!original.ok) {
    return {
      restoreState: "failed",
      probe: {
        member: name,
        status: "failed",
        failedAt: "read",
        sentinel,
        original: undefined,
        note: "读不到（undefined/null 或读抛错）。"
      }
    };
  }

  // 哨兵必须与原值不同；相同则无法证明写入有效，标 skipped。
  if (sameValue(original.value, sentinel)) {
    return {
      restoreState: "restored",
      probe: {
        member: name,
        status: "skipped",
        failedAt: null,
        sentinel,
        original: original.value,
        note: `哨兵 ${String(sentinel)} 与原值相同，写不出区别，跳过写入验证。`
      }
    };
  }

  // 写哨兵。即便抛错，也要尝试恢复（setter 可能先改后抛）。
  let writeError: string | null = null;
  try {
    write(sentinel);
  } catch (error) {
    writeError = error instanceof Error ? error.message : "写哨兵抛错。";
  }

  // 写哨兵后**立即回读哨兵**（此时还没恢复，读到的才是写哨兵的效果）。
  const readback = safeRead(read);
  const writeVerified = !writeError && readback.ok && sameValue(readback.value, sentinel);

  // 无论写是否抛错、回读是否一致，都进入恢复流程：先把原值写回。
  let restoreError: string | null = null;
  try {
    write(original.value);
  } catch (error) {
    restoreError = error instanceof Error ? error.message : "恢复写回抛错。";
  }

  // 恢复后回读，比较是否回到原值。
  const restoredRead = safeRead(read);
  const restoredOk = restoredRead.ok && sameValue(restoredRead.value, original.value);

  // 写哨兵本身抛错 → 判写失败，恢复状态取决于是否回到原值。
  if (writeError) {
    if (restoredOk) {
      return {
        restoreState: "restored",
        probe: {
          member: name,
          status: "failed",
          failedAt: "write",
          sentinel,
          original: original.value,
          note: `写哨兵抛错（${writeError}）；已确认恢复到原值。`
        }
      };
    }
    return {
      restoreState: "uncertain",
      probe: {
        member: name,
        status: "failed",
        failedAt: "write",
        sentinel,
        original: original.value,
        note: `写哨兵抛错（${writeError}）；恢复后回读 ${String(restoredRead.value)} ≠ 原值 ${String(original.value)}，恢复状态不确定。`
      }
    };
  }

  if (!writeVerified) {
    return {
      restoreState: restoredOk ? "restored" : "uncertain",
      probe: {
        member: name,
        status: "failed",
        failedAt: "readback",
        sentinel,
        original: original.value,
        note: `写后回读 ${String(readback.value)} ≠ 哨兵 ${String(sentinel)}。恢复：${restoreError ? `失败（${restoreError}）` : restoredOk ? "已验证回到原值" : `回读 ${String(restoredRead.value)} ≠ 原值`}`
      }
    };
  }

  // 写入验证通过；检查恢复是否成功。
  if (restoreError || !restoredOk) {
    return {
      restoreState: "uncertain",
      probe: {
        member: name,
        status: "failed",
        failedAt: restoreError ? "restore" : "restore_readback",
        sentinel,
        original: original.value,
        note: restoreError
          ? `恢复写回抛错（${restoreError}）。`
          : `恢复后回读 ${String(restoredRead.value)} ≠ 原值 ${String(original.value)}（静默恢复失败）。`
      }
    };
  }

  return {
    restoreState: "restored",
    probe: {
      member: name,
      status: "restore_verified",
      failedAt: null,
      sentinel,
      original: original.value,
      note: "写入验证通过，恢复后回读回到原值。"
    }
  };
}

/**
 * 只探测"成员是否作为函数存在"（AutoFit 无可回读的哨兵值）。
 * **标 `skipped`，不计入通过数** —— "可调用"不等于"执行成功"。
 */
function probeCallableMember(name: string, call: () => (() => void) | undefined): WpsFormatMemberProbe {
  try {
    const fn = call();
    if (typeof fn !== "function") {
      return {
        member: name,
        status: "failed",
        failedAt: "read",
        sentinel: undefined,
        original: undefined,
        note: "成员不是函数（undefined/null 或非函数）。"
      };
    }
    return {
      member: name,
      status: "skipped",
      failedAt: null,
      sentinel: undefined,
      original: undefined,
      note: "成员可调用，但未实际执行（避免改动整行/整列尺寸），不计入已验证。"
    };
  } catch (error) {
    return {
      member: name,
      status: "failed",
      failedAt: "read",
      sentinel: undefined,
      original: undefined,
      note: error instanceof Error ? error.message : "读取抛错。"
    };
  }
}

/* ------------------------------------------------------------------------- *
 * 整轮探测
 * ------------------------------------------------------------------------- */

/**
 * 对指定单元格跑一整轮格式成员探测。**只读写探测单元格，不碰任何业务区域。**
 *
 * 规则：
 * 1. **在任何写入前**复用 `assertWpsFormatProbeTarget` 校验目标资格；不通过直接抛错。
 *    （面板禁用按钮不能替代这里的执行入口校验。）
 * 2. 关联属性按组处理：填充（Interior）、边框（Borders）各为一组，探测前存整组基线，
 *    恢复后核对整组。
 * 3. 一旦某成员恢复失败/不确定，**立即 abort**，后续成员全部标 `skipped`，零写入。
 * 4. 最终报告区分"单项恢复通过"与"整轮基线恢复通过"（`baselineRestored`）。
 */
export function runWpsFormatProbe(
  application: WpsApplication,
  options: { probeCell?: string; groups?: readonly WpsFormatProbeGroup[] } = {}
): WpsFormatProbeReport {
  const probeCell = options.probeCell ?? "A1";
  // 未指定时探测全部组；指定时按给定顺序探测（顺序即组内定义的先后）。
  const selectedGroups = options.groups ?? WPS_FORMAT_PROBE_GROUPS;
  const groupSet = new Set<WpsFormatProbeGroup>(selectedGroups);

  // ① 在任何写入前复用目标校验（面板禁用按钮不能替代）。
  const check = assertWpsFormatProbeTarget(application, probeCell);
  if (!check.ok) {
    throw new Error(check.reason);
  }
  const target = check.target;

  const workbook = application.ActiveWorkbook!;
  const sheet = application.ActiveSheet ?? workbook.ActiveSheet!;
  const range = sheet.Range?.(probeCell);
  if (!range) throw new Error(`WPS 无法读取探测单元格「${sheet.Name}!${probeCell}」。`);

  const members: WpsFormatMemberProbe[] = [];
  let aborted = false;

  // 关联组：填充。整组基线 = [Color, Pattern]。
  const interiorGroup: ProbeGroup = {
    name: "Interior",
    baseline: () => [range.Interior?.Color, range.Interior?.Pattern],
    members: [
      { name: "Interior.Color", read: () => range.Interior?.Color, write: (v) => { if (range.Interior) range.Interior.Color = v as number; }, sentinel: 0xff0000 },
      { name: "Interior.Pattern", read: () => range.Interior?.Pattern, write: (v) => { if (range.Interior) range.Interior.Pattern = v; }, sentinel: XL_PATTERN.solid }
    ]
  };

  // 关联组：边框（下边）。整组基线 = [LineStyle, Weight, ColorIndex]。
  const bottomBorder = () => range.Borders?.Item?.(XL_BORDERS_INDEX.edgeBottom);
  const bordersGroup: ProbeGroup = {
    name: "Borders.Item(edgeBottom)",
    baseline: () => [bottomBorder()?.LineStyle, bottomBorder()?.Weight, bottomBorder()?.ColorIndex],
    members: [
      { name: "Borders.Item.LineStyle", read: () => bottomBorder()?.LineStyle, write: (v) => { const b = bottomBorder(); if (b) b.LineStyle = v; }, sentinel: XL_LINE_STYLE.continuous },
      { name: "Borders.Item.Weight", read: () => bottomBorder()?.Weight, write: (v) => { const b = bottomBorder(); if (b) b.Weight = v; }, sentinel: XL_BORDER_WEIGHT.thin },
      { name: "Borders.Item.ColorIndex", read: () => bottomBorder()?.ColorIndex, write: (v) => { const b = bottomBorder(); if (b) b.ColorIndex = v; }, sentinel: 3 }
    ]
  };

  // Font（独立成员，无关联组，但逐个处理，任一失败/不确定也 abort）。
  const fontMembers = [
    { name: "Font.Bold", read: () => range.Font?.Bold, write: (v: unknown) => { if (range.Font) range.Font.Bold = v; }, sentinel: true },
    { name: "Font.Italic", read: () => range.Font?.Italic, write: (v: unknown) => { if (range.Font) range.Font.Italic = v; }, sentinel: true },
    { name: "Font.Name", read: () => range.Font?.Name, write: (v: unknown) => { if (range.Font) range.Font.Name = v; }, sentinel: "黑体" },
    { name: "Font.Size", read: () => range.Font?.Size, write: (v: unknown) => { if (range.Font) range.Font.Size = v; }, sentinel: 14 },
    { name: "Font.Color", read: () => range.Font?.Color, write: (v: unknown) => { if (range.Font) range.Font.Color = v; }, sentinel: 0x0000ff },
    { name: "Font.Underline", read: () => range.Font?.Underline, write: (v: unknown) => { if (range.Font) range.Font.Underline = v; }, sentinel: true }
  ];

  // 对齐/换行（独立成员）。
  const alignmentMembers = [
    { name: "HorizontalAlignment", read: () => range.HorizontalAlignment, write: (v: unknown) => { range.HorizontalAlignment = v; }, sentinel: XL_H_ALIGN.center },
    { name: "VerticalAlignment", read: () => range.VerticalAlignment, write: (v: unknown) => { range.VerticalAlignment = v; }, sentinel: XL_V_ALIGN.center },
    { name: "WrapText", read: () => range.WrapText, write: (v: unknown) => { range.WrapText = v; }, sentinel: true }
  ];

  // 行高/列宽（改变整行/整列）。
  const sizeMembers = [
    { name: "RowHeight", read: () => range.RowHeight, write: (v: unknown) => { range.RowHeight = v; }, sentinel: 20 },
    { name: "ColumnWidth", read: () => range.ColumnWidth, write: (v: unknown) => { range.ColumnWidth = v; }, sentinel: 12 }
  ];

  /**
   * 处理一个"独立成员列表"（非关联组）：逐个探测，任一恢复不确定 → abort。
   */
  const runIndependentMembers = (
    membersList: Array<{ name: string; read: () => unknown; write: (value: unknown) => void; sentinel: unknown }>
  ): void => {
    for (const m of membersList) {
      if (aborted) {
        members.push({ member: m.name, status: "skipped", failedAt: null, sentinel: m.sentinel, original: undefined, note: "因前置恢复失败/不确定，本轮已停止后续写入探测。" });
        continue;
      }
      const result = probeGroupMember(m.name, m.read, m.write, m.sentinel);
      members.push(result.probe);
      if (result.restoreState === "uncertain") aborted = true;
    }
  };

  /**
   * 处理一个关联组：逐个探测组内成员，之后核对整组基线。
   * 任一成员恢复不确定 → abort，后续组/成员标 skipped。
   */
  const runGroup = (group: ProbeGroup): void => {
    if (aborted) {
      for (const m of group.members) {
        members.push({
          member: m.name,
          status: "skipped",
          failedAt: null,
          sentinel: m.sentinel,
          original: undefined,
          note: "因前置恢复失败/不确定，本轮已停止后续写入探测。"
        });
      }
      return;
    }

    const groupBaseline = group.baseline();

    // 逐个探测组内成员。
    let groupUncertain = false;
    for (const m of group.members) {
      if (aborted || groupUncertain) {
        members.push({
          member: m.name,
          status: "skipped",
          failedAt: null,
          sentinel: m.sentinel,
          original: undefined,
          note: "因前置恢复失败/不确定，本轮已停止后续写入探测。"
        });
        continue;
      }
      const result = probeGroupMember(m.name, m.read, m.write, m.sentinel);
      members.push(result.probe);
      if (result.restoreState === "uncertain") {
        groupUncertain = true;
        aborted = true;
      }
    }

    // 核对整组基线（即使组内全部单项通过，也要看整组是否回到基线 —— 连带变化）。
    if (!groupUncertain) {
      const afterBaseline = group.baseline();
      const baselineOk = groupBaseline.length === afterBaseline.length &&
        groupBaseline.every((value, index) => sameValue(value, afterBaseline[index]));
      if (!baselineOk) {
        // 关联属性变化被发现：整组没有回到基线。
        aborted = true;
        members.push({
          member: `${group.name}（整组基线）`,
          status: "failed",
          failedAt: "baseline",
          sentinel: undefined,
          original: undefined,
          note: `整组基线未恢复：${JSON.stringify(groupBaseline)} → ${JSON.stringify(afterBaseline)}。关联属性发生连带变化。`
        });
      }
    }
  };

  // 按 WPS_FORMAT_PROBE_GROUPS 的顺序，只探测选中的组。
  for (const group of WPS_FORMAT_PROBE_GROUPS) {
    if (!groupSet.has(group)) continue;
    switch (group) {
      case "font":
        runIndependentMembers(fontMembers);
        break;
      case "alignment":
        runIndependentMembers(alignmentMembers);
        break;
      case "size":
        runIndependentMembers(sizeMembers);
        break;
      case "fill":
        runGroup(interiorGroup);
        break;
      case "borders":
        runGroup(bordersGroup);
        break;
      case "autofit":
        // AutoFit 只探测"可调用"，标 skipped，不计入通过（不写任何尺寸）。
        members.push(probeCallableMember("EntireRow.AutoFit", () => range.EntireRow?.AutoFit));
        members.push(probeCallableMember("EntireColumn.AutoFit", () => range.EntireColumn?.AutoFit));
        break;
    }
  }

  // 整轮基线恢复：没有任何 failed（含 baseline 失败）且没有 aborted。
  const baselineRestored = !aborted && members.every((m) => m.status !== "failed");

  const report: WpsFormatProbeReport = {
    host: "wps-et",
    hostVersion: application.Version ? String(application.Version) : null,
    workbookName: target.workbookName,
    sheetName: target.sheetName,
    probeCell,
    startedAt: new Date().toISOString(),
    target,
    groups: [...selectedGroups],
    members,
    baselineRestored,
    aborted,
    summary: {
      total: members.length,
      restoreVerified: members.filter((m) => m.status === "restore_verified").length,
      writeVerified: members.filter((m) => m.status === "write_verified").length,
      memberExists: members.filter((m) => m.status === "member_exists").length,
      skipped: members.filter((m) => m.status === "skipped").length,
      failed: members.filter((m) => m.status === "failed").length
    }
  };

  // DEV 下上报汇总计数；逐项结果由面板展示/复制。
  try {
    reportDevProbe("taskpane", {
      probeTotal: report.summary.total,
      probeRestoreVerified: report.summary.restoreVerified,
      probeWriteVerified: report.summary.writeVerified,
      probeSkipped: report.summary.skipped,
      probeFailed: report.summary.failed,
      probeBaselineRestored: baselineRestored ? 1 : 0
    });
  } catch {
    // 上报只是辅助，探测结论以返回值与面板展示为准。
  }

  return report;
}
