/**
 * 参数化列映射的确定性计算（L2，宿主无关）。
 *
 * **与 `server/app/column_mapping.py` 是镜像关系**，行为必须逐格一致。
 * 一致性不靠人工对读源码,靠共享夹具
 * `config/column_mapping_cases.json` —— 两侧各自加载同一批用例并断言同一批
 * 期望值,任何单边修改都会让一侧变红。改这里必须同时改 Python 侧。
 *
 * 为什么两份实现而不是一份:服务端需要它（文件夹通道由服务端直接读写 xlsx),
 * 两端客户端也需要它（WPS 的源值不上传服务端,`wpsReadonlyAgent.ts` 送的
 * `dataRows` 恒为空 —— 这是刻意的隐私/体积取舍,不是缺陷）。
 * 让服务端算 = WPS 永远算不出来;让客户端算 = 文件夹通道拿不到。
 *
 * 四处刻意的"拒绝"而非"兜底",理由同 Python 侧:
 * 1. 来源值不在映射表内 → 抛错,不猜。错误的分类在单元格层面看不出来。
 * 2. 源/目标在同一表上重叠 → 抛错。逐格落地时读写先后决定结果,
 *    而这个顺序在三条通道上不一致。
 * 3. 形状不符 → 抛错。截断或补齐都会写出用户没要求的格子。
 * 4. 整格含**未声明**的分隔形式 → 抛错并要求澄清。猜一个分隔符去拆,
 *    拆错了在单元格层面同样看不出来。
 *
 * 多值单元格的处理**完全由规则数据驱动**（`delimiters` + `valueRank`）,
 * 本模块不内置任何分隔符字面量,也不认识任何等级名。规则怎么给就怎么算;
 * 规则没给的能力,宁可澄清也不脑补 —— 这是与"拒绝而非兜底"同一个立场。
 */

import type { CellValue } from "./contracts";
import {
  blankCellOutcome,
  blankPolicyFromLegacy,
  isBlankValue,
  type BlankCellOutcome,
  type BlankPolicy
} from "./blankPolicy";
import { columnLetterFromIndex } from "./range";

export class ColumnMappingError extends Error {}

/**
 * 源区域**真的有空白格**，而空白策略尚未确定。
 *
 * 这是**待澄清**，不是错误：实现不替用户选一个默认标签（历史上那正是
 * 空白被一律写成「/」或「未分类」的来源）。调用方应当把它翻成一张带
 * 四个选项的澄清卡（保留原值 / 清空 / 写指定值 / 写空串），而不是把
 * `message` 原样当报错文案丢给用户。
 *
 * 反过来：**源区域一格空白都没有时不该抛**。没读到空白就先问"空白怎么办"
 * 是本模块刻意消灭的行为 —— 那时用户只能凭空回答一个不存在的问题。
 */
export class BlankPolicyNeededError extends ColumnMappingError {
  /** 真实读到的空白格数量。界面据此说明规模，不含任何单元格内容。 */
  readonly blankCount: number;

  constructor(blankCount: number) {
    super(
      `源区域有 ${blankCount} 个空白格，但还没有确定空白应当怎么处理，已拒绝替你选一个默认标签。`
    );
    this.blankCount = blankCount;
  }
}

/**
 * 需要跳过的空白位置太多，一次预览无法用有限个连续写入段表达。
 *
 * 宁可说清楚**为什么表达不了**，也不把"保留原值"偷偷降级成"写个空值"
 * —— 那会让用户以为自己选的是保留，目标格的公式与格式却被改掉了。
 */
export class BlankSplitLimitError extends ColumnMappingError {
  readonly requiredActions: number;
  readonly maxActions: number;

  constructor(requiredActions: number, maxActions: number) {
    super(
      `按当前空白分布需要拆成 ${requiredActions} 段写入，超过单次预览的上限 ` +
        `${maxActions} 段。请改用「写入指定值」或缩小来源区域。`
    );
    this.requiredActions = requiredActions;
    this.maxActions = maxActions;
  }
}

/**
 * 来源格未命中映射表,且**能确定**归因是"类别映射缺口"。
 *
 * 典型情形:已声明分隔符把它拆开了（或它本来就是一个整值）,拆出的项里
 * 有映射表没覆盖的类别。补救动作单一 —— **补类别映射**,可补全后重试。
 *
 * 与 `NeedsClarificationError` 刻意分开:那个是"该补映射还是该补分隔符
 * 说不准",补救动作有两个可能,必须让用户选。合成一个错误会让用户
 * 在两件事之间瞎猜,而猜错的那一半会把分隔符问题当成缺失类别补掉。
 */
export class UnmappedCategoryError extends ColumnMappingError {
  /** 未覆盖的**项**（拆开后逐项、或整格）,去重并按首次出现顺序。 */
  readonly values: readonly CellValue[];

  constructor(values: readonly CellValue[]) {
    const rendered = values.map((value) => JSON.stringify(value)).join("、");
    super(`来源值 ${rendered} 不在映射表内，已拒绝猜测。请补全映射表后重试。`);
    this.values = values;
  }
}

/**
 * 来源格无法在不猜测的前提下归类,而**归因本身也不确定**。
 *
 * 情形:整格未命中、且像是用了某种没被声明为可分分隔符的写法。
 * 此时既可能是"用错了分隔符",也可能是"这个整格本身就缺映射"
 * —— 实现**不判断是哪一个**,只把"这格处理不了"报上去。
 *
 * 与 `UnmappedCategoryError` 的区别是**补救动作的数量**:
 * 这个有两种可能（补分隔符 / 补映射）,必须让用户选。
 */
export class NeedsClarificationError extends ColumnMappingError {
  /** 处理不了的**整格原文**,去重并按首次出现顺序。 */
  readonly cells: readonly CellValue[];

  constructor(cells: readonly CellValue[]) {
    // 文案里**不嵌入原文**:这个 message 会经 `ColumnMappingClarification.message`
    // 下发到客户端，而原文属于本端数据。要展示原文的界面有自己的源值，
    // 不需要（也不该）从服务端把它带回来。
    super(
      `有 ${cells.length} 个来源格无法归类：可能是用了未声明的分隔符，` +
        "也可能是映射表缺少这个类别。已拒绝猜测。"
    );
    this.cells = cells;
  }
}

export class ShapeMismatchError extends ColumnMappingError {}

export class OverlappingRangeError extends ColumnMappingError {}

/**
 * 待澄清的**归因**。用户界面据此给出**具体选项**,而不是只弹一句
 * "请补充信息" —— 归因不同,能做的补救动作不同:
 *
 * - `matchedDelimiterMissingCategories`：分隔符已经声明且对上了,
 *   缺的是类别映射。可选动作只有「补充类别映射」这一条是有效的。
 * - `unresolvedSeparatorOrMapping`：整格处理不了,而且**分不清**是
 *   "用了别的分隔符"还是"缺这个类别的映射"。这时必须给全三个选项:
 *   「补充分隔符」「补充类别映射」「把这格整体当作一个类别」。
 * - `blankPolicyNeeded`：**读到源区域确实有空白格**,而空白怎么处理还没定。
 *   这是一次**由真实数据触发**的澄清 —— 没读到空白就不该出现,也不许在
 *   读之前先问。可选项只有空白那四种,不给类别/分隔符那组。
 */
export type ColumnMappingClarificationReason =
  | "matchedDelimiterMissingCategories"
  | "unresolvedSeparatorOrMapping"
  | "blankPolicyNeeded";

/**
 * 待澄清卡片的内容。**不含来源格的原始内容** —— 原始内容留在本端,
 * 由 UI 就地展示;下发/回传的只有归因与统计。
 *
 * 字段名与协议（`contracts.ts` 的 `ColumnMappingClarification`、
 * `models.py` 的同名模型）**逐字一致**，这样同一份数据可以在
 * "服务端下发 → 客户端消费"之间原样传递，不需要一层改名。
 *
 * 注意 `resolvedBy` 是**服务端的声明**，界面**不要直接信它**：
 * 可选项一律由 `reason` 推导（见 `columnMappingClarificationCard`）。
 * 保留这个字段是为了让服务端能表达自己的意图，以及让两端能核对分歧。
 */
export interface ColumnMappingClarification {
  reason: ColumnMappingClarificationReason;
  message: string;
  /** 未覆盖的类别（已去重）,供界面列出"要补哪些"。 */
  missingCategories: string[];
  /** 归因不确定时,无法处理的整格数量（不回传原文）。 */
  unresolvedCellCount: number;
  /**
   * `blankPolicyNeeded` 时**真实读到**的空白格数量。
   *
   * 必须来自一次真实读取。没有读证据就不许声称"源区域有空白" ——
   * 那是凭空提问，用户无从判断该不该回答。
   */
  blankCellCount?: number;
  /** 服务端认为可行的动作。顺序即推荐顺序。 */
  resolvedBy: ColumnMappingResolution[];
}

/**
 * 澄清卡片上可以给用户选的动作。**归因决定哪些动作有意义** ——
 * 分隔符已经对上的情形给出"补充分隔符"只会误导用户去改一个没问题的设置。
 */
export type ColumnMappingResolution =
  | "addCategories"
  | "specifyOtherDelimiter"
  | "treatWholeCellAsCategory"
  /** 空白：保留目标原值（该位置不写入）。 */
  | "keepTarget"
  /** 空白：清空目标内容（clearRange:contents，不是写空串）。 */
  | "clearTarget"
  /** 空白：写入用户指定的值，需要他一并填出写什么。 */
  | "writeBlankValue"
  /** 空白：显式写空字符串常量，与「清空」不是一回事。 */
  | "writeEmptyString";

/**
 * 拆分后取最高级时的一条次序声明。`rank` 大者优先。
 *
 * 等级**由规则数据给出**,本模块不认识任何具体等级名 —— 它在意的只有
 * 数值大小,以及"没声明时按 `mapping` 的声明顺序"这条确定性兜底。
 */
export interface ColumnMappingRankEntry {
  value: string;
  rank: number;
}

export interface ColumnMappingInput {
  sourceSheet: string;
  targetSheet: string;
  sourceRange: string;
  targetRange: string;
  sourceValues: readonly (readonly CellValue[])[];
  /** 来源值 → 目标标签。键顺序不影响结果,但决定无 `valueRank` 时的优先级。 */
  mapping: Readonly<Record<string, string>>;
  /**
   * 空白来源格的处理策略。**省略即"未定"** —— 只有真的读到空白时才需要决定。
   *
   * 与 `blankLabel` 的关系：`blankPolicy` 优先；没给 `blankPolicy` 时才按
   * 旧字段折（`blankPolicyFromLegacy`）。旧字段非空串 = `kind:"value"`，
   * 语义与旧版逐格一致，不改变任何一条既有计划的写法。
   */
  blankPolicy?: BlankPolicy;
  /**
   * 旧协议的"空白标签"。保留只为兼容旧数据，新调用方一律用 `blankPolicy`。
   *
   * 它**只能**表达"写一个什么标签"，表达不了"保留目标原值"，也表达不了
   * "清空"与"写空串"的区别 —— 那三种意图请改用 `blankPolicy`。
   */
  blankLabel?: string;
  /** 单次预览允许拆成的最大写入段数。省略用 `DEFAULT_MAX_MAPPING_ACTIONS`。 */
  maxSplitActions?: number;
  /**
   * 候选分隔符。省略 = 不拆分。
   *
   * 与 `valueRank` 一样是**平铺字段**,不折进 `semantics` 子对象:
   * 它们来自服务端下发的规则,规则是平的;`config/column_mapping_cases.json`
   * 与 `ColumnMappingRule` 也都按平的写。折一层会让"夹具里有、运行时没有"
   * 这种错在类型检查下隐形 —— 这里的字段名与夹具逐字对齐才有意义。
   */
  delimiters?: readonly string[];
  /** 拆分后取最高级用的次序,`rank` 大者优先。省略 = 按 `mapping` 声明顺序。 */
  valueRank?: readonly ColumnMappingRankEntry[];
  /**
   * 规则**明示为分隔符候选、但尚未确认可用来切分**的字符。
   *
   * 整格没命中、却含有这里的字符时,说明真实的分隔方式还没问清楚,
   * 一律提示澄清。它与 `delimiters` 的分工是:
   * - 出现在 `delimiters` 里 = 用户已确认"按它拆";
   * - 只出现在这里 = 用户知道"它像分隔符",但还没确认要不要拆。
   *
   * 之所以要这个字段,而不是去猜"这格看起来像不像多值":
   * 猜的边界无法解释也无法复现,而它直接决定"要不要停下来问用户"。
   */
  undeclaredSeparators?: readonly string[];
  /**
   * 服务端已经判定"这条需求还缺参数／缺类别、不能算"。
   *
   * 客户端拿到它就不该算、更不该写:归因已经确定,界面上要呈现的是
   * 具体的补救选项,而不是一份可能建立在猜测之上的矩阵。
   * 带上时按 `reason` 抛对应的错误,让调用方走澄清流程。
   */
  clarification?: ColumnMappingClarification;
}

export interface ColumnMappingAction {
  type: "writeValues";
  sheet: string;
  range: string;
  values: string[][];
}

/**
 * 清空目标内容的动作。**只用于 `blankPolicy: clear`**，且只清内容。
 *
 * 为什么不是"写一个空字符串"：空串常量在部分宿主上留下的是
 * `ISBLANK() === FALSE` 的常量格，与真空单元格在后续统计里行为不同。
 * 用户说的是"清空"，就该给 `clearRange:contents`。
 */
export interface ColumnMappingClearAction {
  type: "clearRange";
  sheet: string;
  range: string;
  applyTo: "contents";
}

export type ColumnMappingPlanAction =
  | ColumnMappingAction
  | ColumnMappingClearAction;

/**
 * 一次列映射预览最多允许拆成几段写入。
 *
 * "保留原值"靠**把位置排除出写入范围**实现，位置零散时目标区域必须拆成
 * 多段连续区间。段数无上限会让预览卡变成几百个动作，验收与撤销都不可控；
 * 超限就**明说**，不降级成写入空值。
 */
export const DEFAULT_MAX_MAPPING_ACTIONS = 64;

const CELL = String.raw`\$?([A-Za-z]+)\$?(\d+)`;
const RANGE_PATTERN = new RegExp(`^${CELL}(?::${CELL})?$`);

function columnNumber(name: string): number {
  let result = 0;
  for (const character of name.toUpperCase()) {
    result = result * 26 + (character.charCodeAt(0) - 64);
  }
  return result;
}

interface RangeBox {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}

/** 1 基闭区间。与 Python 侧 `_range_box` 同语义。 */
function rangeBox(address: string): RangeBox {
  const match = RANGE_PATTERN.exec(address.trim());
  if (!match) throw new ColumnMappingError(`无法解析区域「${address}」。`);
  const startColumn = columnNumber(match[1]);
  const startRow = Number(match[2]);
  const endColumn = columnNumber(match[3] ?? match[1]);
  const endRow = Number(match[4] ?? match[2]);
  if (endRow < startRow || endColumn < startColumn) {
    throw new ColumnMappingError(`区域「${address}」的起止顺序颠倒。`);
  }
  return { startRow, startColumn, endRow, endColumn };
}

/** 空白 = 没有值,或纯空白字符串。数字 0 与 false **不是**空白。 */
function isBlank(value: CellValue): boolean {
  return isBlankValue(value);
}

/**
 * 空白格的**占位值**。
 *
 * 只出现在"这一格是空白、且策略要求跳过/清空"的矩阵位置里，而那些位置
 * 一定会被 `splitTargetActions` 排除出写入范围 —— 它**永远不会**被写进
 * 工作表。留一个占位是为了让矩阵保持与目标区域同形，便于比对与试算。
 */
const BLANK_PLACEHOLDER = "";

/**
 * 逐格映射。未命中值**全部**收集后一次抛出。
 *
 * 一次只报一个会让调用方反复试错;映射表通常人工维护,一轮拿到完整缺口
 * 才能一次补完。
 *
 * 判断顺序刻意是**显式 → 空白 → 拆分 → 未知**:
 * - 显式项排第一,所以调用方可以把 `""` 或整格原文写进 `mapping`,
 *   精确覆盖"空串"与"看起来该拆但想整体映射"两类意图;
 * - 空白排第二,所以真空白永远走 `blankLabel`,不会被拆分或未知逻辑碰到;
 * - 拆分排第三,只对**没被显式覆盖、也不是空白**的整格发生。
 */
export interface MappingCellGrid {
  /** 与目标区域同形的目标矩阵。被跳过/清空的位置是占位值，不会落地。 */
  values: string[][];
  /** 与 `values` 同形的落地方式。非空白格一律是 `write`。 */
  outcomes: BlankCellOutcome[][];
}

/**
 * 逐格映射,并给出每一格的**落地方式**。
 *
 * 之所以要把 `outcomes` 一并算出来:仅凭值矩阵无法区分"写一个空字符串"
 * 与"这个位置不写入" —— 两者在 `string[][]` 里长得一模一样,而它们的
 * 语义完全相反。让调用方拿着值矩阵去猜,历史上正是"保留原值被实现成
 * 写空值"的根源。
 */
export function mapColumnValuesWithOutcomes(input: {
  sourceValues: readonly (readonly CellValue[])[];
  mapping: Readonly<Record<string, string>>;
  blankPolicy?: BlankPolicy;
  blankLabel?: string;
  delimiters?: readonly string[];
  valueRank?: readonly ColumnMappingRankEntry[];
  undeclaredSeparators?: readonly string[];
}): MappingCellGrid {
  const policy = input.blankPolicy ?? blankPolicyFromLegacy(input.blankLabel);
  const delimiters = input.delimiters ?? [];
  const undeclaredSeparators = input.undeclaredSeparators ?? [];
  const valueRank = input.valueRank ?? [];
  // 无显式 valueRank 时用 mapping 的声明顺序兜底。JS 字符串键保持插入顺序,
  // 所以这里的下标就是"规则声明的先后",不需要额外参数。
  const explicitRanks = new Map<string, number>();
  for (const entry of valueRank) {
    explicitRanks.set(entry.value, entry.rank);
  }
  const mappingOrder = new Map<string, number>();
  let nextOrder = 0;
  for (const key of Object.keys(input.mapping)) {
    mappingOrder.set(key, nextOrder);
    nextOrder += 1;
  }

  const mapped: string[][] = [];
  const outcomeRows: BlankCellOutcome[][] = [];
  const unmatched: CellValue[] = [];
  const unresolved: CellValue[] = [];
  const seenUnmatched = new Set<string>();
  const seenUnresolved = new Set<string>();

  const recordUnmatched = (value: CellValue, key: string) => {
    if (seenUnmatched.has(key)) return;
    seenUnmatched.add(key);
    unmatched.push(value);
  };

  for (const row of input.sourceValues) {
    const mappedRow: string[] = [];
    const outcomeRow: BlankCellOutcome[] = [];
    /**
     * 未命中位置的**占位**。
     *
     * 这些格子必然让本函数在结尾抛错（未覆盖 / 归因不确定），矩阵不会被
     * 返回，所以它的 outcome 无意义 —— 但矩阵必须保持同形，否则调用方
     * 拿到的形状与目标区域不一致。
     */
    const pushUnmatchedPlaceholder = () => {
      mappedRow.push(BLANK_PLACEHOLDER);
      outcomeRow.push({ kind: "write", value: BLANK_PLACEHOLDER });
    };
    const pushMapped = (value: string) => {
      mappedRow.push(value);
      outcomeRow.push({ kind: "write", value });
    };
    for (const cell of row) {
      // 1. 先试**整格精确匹配**。这是优先级最高的判定:调用方把整格原文
      //    写进 mapping,就是在说"这一格整体算一个类别",拆分不该推翻它。
      if (typeof cell === "string" && Object.hasOwn(input.mapping, cell)) {
        pushMapped(input.mapping[cell]);
        continue;
      }
      // 2. 空白:源本来没有值,按**空白策略**落地,不参与拆分。
      //
      //    策略是 `undecided` 时这里先记成 `skip` 占位:本函数不抛错,
      //    由 `buildColumnMappingActions` 在确认"确实有空白"之后再决定
      //    是否升级为一次澄清 —— 没读到空白就不该拦。
      if (isBlank(cell)) {
        const outcome =
          policy.kind === "undecided"
            ? { kind: "skip" as const }
            : blankCellOutcome(policy);
        mappedRow.push(
          outcome.kind === "write" ? outcome.value : BLANK_PLACEHOLDER
        );
        outcomeRow.push(outcome);
        continue;
      }

      const key = typeof cell === "string" ? cell : stringifyKey(cell);

      // 3. 没有分隔符声明 = 一个格子只当一个值,与扩展前逐格一致。
      if (delimiters.length === 0) {
        if (Object.hasOwn(input.mapping, key)) {
          pushMapped(input.mapping[key]);
          continue;
        }
        recordUnmatched(cell, key);
        pushUnmatchedPlaceholder();
        continue;
      }

      // 4. 用**已声明**的分隔符切分:取出现位置最靠前的那个。
      //
      //    位置优先而非声明序优先:声明序是**优先级**（谁更重要）,
      //    不是**切分方式**（按什么切）。用它决定切分会让
      //    `["；", ","]` 这类表达反直觉。
      let delimiter: string | undefined;
      let delimiterIndex = Number.MAX_SAFE_INTEGER;
      for (const candidate of delimiters) {
        if (candidate.length === 0) continue;
        const index = key.indexOf(candidate);
        if (index >= 0 && index < delimiterIndex) {
          delimiter = candidate;
          delimiterIndex = index;
        }
      }

      // 4a. 整格里没有任何**已声明**的分隔符。
      if (delimiter === undefined) {
        if (Object.hasOwn(input.mapping, key)) {
          pushMapped(input.mapping[key]);
          continue;
        }
        // 没命中,但含有规则**明示为分隔符写法**、却未被声明为可切分的字符:
        // 这时归因**不确定** —— 可能分隔符没声明全,也可能这格本身就缺映射。
        // 两种补救动作不同,所以不替用户选,报上去让他决定。
        const looksLikeOtherSeparator = undeclaredSeparators.some(
          (candidate) => candidate.length > 0 && key.includes(candidate)
        );
        if (looksLikeOtherSeparator) {
          recordUnresolved(cell, key);
          pushUnmatchedPlaceholder();
          continue;
        }
        // 看不出任何分隔形式:归因确定 —— 就是映射表缺这个类别。
        recordUnmatched(cell, key);
        pushUnmatchedPlaceholder();
        continue;
      }

      // 4b. 用已声明的分隔符拆开。
      const items = key
        .split(delimiter)
        .map((item) => item.trim())
        .filter((item) => item.length > 0);

      // 整格以分隔符开头/结尾,拆完一项不剩:当整格未覆盖处理。
      if (items.length === 0) {
        recordUnmatched(cell, key);
        pushUnmatchedPlaceholder();
        continue;
      }

      // 5. 拆出的每一项都必须能命中映射表。
      //
      //    拆分**已经按用户声明的分隔符完成**,所以这里能确定归因:
      //    缺的是**类别映射**,不是分隔符。按项收集未覆盖的类别,
      //    不按整格收 —— 用户要的是"把丙补进映射表",
      //    给他整格原文等于让他自己去拆。
      const missing = items.filter(
        (item) => !Object.hasOwn(input.mapping, item)
      );
      if (missing.length > 0) {
        for (const item of missing) recordUnmatched(item, item);
        pushUnmatchedPlaceholder();
        continue;
      }

      // 6. 取规则指定的最高级:显式 valueRank 先看,没有就看 mapping 声明顺序。
      const winner = pickHighestRank(items, explicitRanks, mappingOrder);
      pushMapped(input.mapping[winner]);
    }
    mapped.push(mappedRow);
    outcomeRows.push(outcomeRow);
  }

  // 先说"归因不确定"的那类:它需要用户做选择,信息量更大。
  // 两类都可能有;归因不确定的更紧急,所以排在前面抛。
  if (unresolved.length > 0) throw new NeedsClarificationError(unresolved);
  if (unmatched.length > 0) throw new UnmappedCategoryError(unmatched);
  return { values: mapped, outcomes: outcomeRows };

  function recordUnresolved(value: CellValue, key: string) {
    if (seenUnresolved.has(key)) return;
    seenUnresolved.add(key);
    unresolved.push(value);
  }
}

/**
 * 只取目标矩阵。与 `mapColumnValuesWithOutcomes` 同一套判定,但没有
 * "这一格该不该写"的信息 —— 新调用方一律用带 outcome 的那个版本。
 */
export function mapColumnValues(input: {
  sourceValues: readonly (readonly CellValue[])[];
  mapping: Readonly<Record<string, string>>;
  blankPolicy?: BlankPolicy;
  blankLabel?: string;
  delimiters?: readonly string[];
  valueRank?: readonly ColumnMappingRankEntry[];
  undeclaredSeparators?: readonly string[];
}): string[][] {
  return mapColumnValuesWithOutcomes(input).values;
}

/**
 * 多值取最高级。比较的是"优先级",不是数值大小:
 * `valueRank` 里 rank 越大越优先;没有 rank 的项按 `mapping` 声明顺序决定。
 */
function pickHighestRank(
  items: readonly string[],
  explicitRanks: ReadonlyMap<string, number>,
  mappingOrder: ReadonlyMap<string, number>
): string {
  let winner = items[0];
  for (const item of items) {
    if (compareRank(item, winner, explicitRanks, mappingOrder) > 0) winner = item;
  }
  return winner;
}

function compareRank(
  left: string,
  right: string,
  explicitRanks: ReadonlyMap<string, number>,
  mappingOrder: ReadonlyMap<string, number>
): number {
  const leftRank = explicitRanks.get(left);
  const rightRank = explicitRanks.get(right);
  if (leftRank !== undefined || rightRank !== undefined) {
    // 只给了一边的等级时,另一边的等级按"最低"处理:显式等级是调用方的
    // 明确声明,未声明的项不该反过来压过它。
    const leftValue = leftRank ?? Number.NEGATIVE_INFINITY;
    const rightValue = rightRank ?? Number.NEGATIVE_INFINITY;
    if (leftValue !== rightValue) return leftValue > rightValue ? 1 : -1;
  }
  const leftOrder = mappingOrder.get(left) ?? Number.MAX_SAFE_INTEGER;
  const rightOrder = mappingOrder.get(right) ?? Number.MAX_SAFE_INTEGER;
  if (leftOrder === rightOrder) return 0;
  // 声明越靠前越优先,所以顺序号越小越优先 —— 返回反号。
  return leftOrder < rightOrder ? 1 : -1;
}

/**
 * 非字符串来源值的键化。**必须与 Python 的 `str()` 对齐**,
 * 否则同一张表在两端会得出不同的未命中集合。
 *
 * `true` 在 JS 里 `String(true) === "true"`,Python 是 `"True"`。
 * 布尔按 Python 形态键化;数字两端一致（整数不带 `.0`,JS 本来就不带）。
 */
function stringifyKey(value: CellValue): string {
  if (value === true) return "True";
  if (value === false) return "False";
  return String(value);
}

/**
 * 目标矩阵 → 写入动作序列。
 *
 * **没有任何需要跳过或清空的位置时，只产出一个覆盖整块区域的动作** ——
 * 与扩展前逐格一致，也是绝大多数映射的真实形态。
 *
 * 一旦有位置要跳过（保留原值）或清空，整块连续区域就不再成立：把它整块
 * 写进去等于把"跳过"写成"写了个空值"，目标格的公式与格式当场被替换掉。
 * 这时按**列内连续段**拆开：同一列里连续若干行落地方式相同就合成一段，
 * 被跳过的位置**不出现在任何动作的 range 里**。
 *
 * 段数必须有上限：零散的空白能把一列拆成上百段，预览、验收与撤销都不可控。
 * 超限就抛 `BlankSplitLimitError` 明说，绝不降级成写空值。
 */
export function splitTargetActions(input: {
  targetSheet: string;
  /** 目标区域的 1 基闭区间。 */
  target: { startRow: number; startColumn: number; endRow: number; endColumn: number };
  /** 未拆分时的完整区域地址，用于"不需要拆分"的快路径。 */
  targetRange: string;
  values: string[][];
  outcomes: BlankCellOutcome[][];
  maxActions?: number;
}): ColumnMappingPlanAction[] {
  const maxActions = input.maxActions ?? DEFAULT_MAX_MAPPING_ACTIONS;
  const rowCount = input.target.endRow - input.target.startRow + 1;
  const columnCount = input.target.endColumn - input.target.startColumn + 1;

  let needsSplit = false;
  for (const row of input.outcomes) {
    for (const outcome of row) {
      if (outcome.kind !== "write") {
        needsSplit = true;
        break;
      }
    }
    if (needsSplit) break;
  }
  if (!needsSplit) {
    return [
      {
        type: "writeValues",
        sheet: input.targetSheet,
        range: input.targetRange,
        values: input.values
      }
    ];
  }

  const actions: ColumnMappingPlanAction[] = [];
  for (let columnOffset = 0; columnOffset < columnCount; columnOffset += 1) {
    const letter = columnLetterFromIndex(input.target.startColumn + columnOffset - 1);
    let runStart: number | null = null;
    let runKind: "write" | "clear" | null = null;
    const flush = (endRowOffset: number) => {
      if (runStart === null || runKind === null) return;
      const startRow = input.target.startRow + runStart;
      const endRow = input.target.startRow + endRowOffset;
      const range = `${letter}${startRow}:${letter}${endRow}`;
      if (runKind === "clear") {
        actions.push({
          type: "clearRange",
          sheet: input.targetSheet,
          range,
          applyTo: "contents"
        });
      } else {
        const values: string[][] = [];
        for (let offset = runStart; offset <= endRowOffset; offset += 1) {
          values.push([input.values[offset][columnOffset]]);
        }
        actions.push({
          type: "writeValues",
          sheet: input.targetSheet,
          range,
          values
        });
      }
      runStart = null;
      runKind = null;
    };
    for (let rowOffset = 0; rowOffset < rowCount; rowOffset += 1) {
      const outcome = input.outcomes[rowOffset][columnOffset];
      const kind = outcome.kind === "write" ? "write" : outcome.kind === "clear" ? "clear" : null;
      if (kind === null) {
        flush(rowOffset - 1);
        continue;
      }
      if (runKind !== kind) {
        flush(rowOffset - 1);
        runStart = rowOffset;
        runKind = kind;
        continue;
      }
    }
    flush(rowCount - 1);
  }

  if (actions.length > maxActions) {
    throw new BlankSplitLimitError(actions.length, maxActions);
  }
  return actions;
}

export function buildColumnMappingActions(
  input: ColumnMappingInput
): ColumnMappingPlanAction[] {
  // 服务端已经判定"这条需求缺参数／缺类别",本端就不该算。放在所有校验
  // 之前:归因已经确定,再算出来的矩阵必然是猜的,不能让它有出场的可能。
  if (input.clarification) {
    if (input.clarification.reason === "unresolvedSeparatorOrMapping") {
      throw new NeedsClarificationError(input.clarification.missingCategories);
    }
    throw new UnmappedCategoryError(input.clarification.missingCategories);
  }
  if (!input.sourceSheet.trim() || !input.targetSheet.trim()) {
    throw new ColumnMappingError("源或目标工作表名为空。");
  }
  // **不再**要求先给出空白标签：那是"还没读源数据就先问空白怎么办"的根源。
  // 策略未定不是错误 —— 只有下面真的读到空白格时才升级为一次澄清。

  const source = rangeBox(input.sourceRange);
  const target = rangeBox(input.targetRange);
  const sourceRows = source.endRow - source.startRow + 1;
  const sourceColumns = source.endColumn - source.startColumn + 1;
  const targetRows = target.endRow - target.startRow + 1;
  const targetColumns = target.endColumn - target.startColumn + 1;

  if (
    input.sourceValues.length !== sourceRows ||
    input.sourceValues.some((row) => row.length !== sourceColumns)
  ) {
    throw new ShapeMismatchError(
      `源值矩阵形状与区域「${input.sourceRange}」（${sourceRows}×${sourceColumns}）不一致。`
    );
  }
  if (targetRows !== sourceRows || targetColumns !== sourceColumns) {
    throw new ShapeMismatchError(
      `目标区域「${input.targetRange}」（${targetRows}×${targetColumns}）` +
        `与源区域「${input.sourceRange}」（${sourceRows}×${sourceColumns}）形状不一致。`
    );
  }

  if (input.sourceSheet === input.targetSheet) {
    const rowsOverlap =
      source.startRow <= target.endRow && target.startRow <= source.endRow;
    const columnsOverlap =
      source.startColumn <= target.endColumn &&
      target.startColumn <= source.endColumn;
    if (rowsOverlap && columnsOverlap) {
      throw new OverlappingRangeError(
        `源区域「${input.sourceRange}」与目标区域「${input.targetRange}」` +
          "在同一工作表上重叠，读写先后顺序在各通道不一致，已拒绝。"
      );
    }
  }

  const policy = input.blankPolicy ?? blankPolicyFromLegacy(input.blankLabel);
  const grid = mapColumnValuesWithOutcomes({
    sourceValues: input.sourceValues,
    mapping: input.mapping,
    blankPolicy: policy,
    delimiters: input.delimiters,
    valueRank: input.valueRank,
    undeclaredSeparators: input.undeclaredSeparators
  });

  /*
   * 策略未定 + **真的读到了空白** → 待澄清，不替用户挑默认标签。
   * 一格空白都没有时这里不触发 —— 那时"空白怎么办"是个不存在的问题，
   * 拿它去拦预览就是让用户凭空回答，也是重授权死循环的起点之一。
   */
  if (policy.kind === "undecided") {
    let blankCount = 0;
    for (const row of grid.outcomes) {
      for (const outcome of row) {
        if (outcome.kind === "skip") blankCount += 1;
      }
    }
    if (blankCount > 0) throw new BlankPolicyNeededError(blankCount);
  }

  return splitTargetActions({
    targetSheet: input.targetSheet,
    target,
    targetRange: input.targetRange,
    values: grid.values,
    outcomes: grid.outcomes,
    maxActions: input.maxSplitActions
  });
}
