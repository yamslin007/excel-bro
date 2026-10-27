/**
 * 计划里**固定值**的来源锚点：这些值是对哪一版源数据算出来的。
 *
 * 为什么需要单独一层：一份计划里可能有"执行时求值"的动作（公式、复制、实时计算），
 * 也可能有**烘焙矩阵**（模型/隔离容器算完后原样写进去的值）。后者在源数据变化之后
 * 依然会照原样写进单元格 —— 单元格层面完全看不出它是旧结果。
 *
 * 这里只回答一个问题：**锚点与现场还是同一版吗**。判据是三态，不是布尔：
 *
 * - `unchanged`：现场指纹与锚点相同，可以写；
 * - `changed`：现场变了 —— 这份计划的固定值已经过期，不能写；
 * - `unverifiable`：现场读不到（或锚点缺失）—— **不等于**没变，按失败关闭处理。
 *
 * 锚点从哪来（这一层不猜，调用方给）：
 * - 本端在**模型实际读到数据那一刻**记下的指纹（桥绑定的指纹），
 *   而不是"计划到达面板那一刻"再读一次 —— 后者会把"分析期间用户改了源表"
 *   自比自等，异常值恰好相等，于是旧结果被静默写入（2026-10-25 实测缺陷）；
 * - SDK 计算结果计划读取失败或没有本轮桥时，锚点为 `null`，必须拒绝写入。
 *   可信 plan_range_edit 编译器证明的字面值赋值不依赖计算来源；仅内存封印有效。
 *   本地规则、执行时计算及文件夹模式由各自可信入口选择独立来源复核，
 *   不能由模型响应自报豁免，也不能用另一动作的来源快照替整份 SDK 计划作证。
 *
 * **双端同源**：Excel 与 WPS 的写前复核读这一份判据。各写一份"看起来一样"的
 * 比较，正是"一端拦得住、另一端照写"那类缺陷的来源。
 */

/** 本模块只关心计划上的两个字段，别的字段原样透传。 */
import { isTrustedLiteralPlan } from "./literalPlanOrigin";

export interface SourceAnchorPlan {
  sourceFingerprint?: string | null;
  sourceFingerprintSheets?: readonly string[];
}

/** 锚点：指纹 + 它是按哪些工作表算出来的（两者必须成对，否则不可复算）。 */
export interface PlanSourceAnchor {
  fingerprint: string;
  sheets: readonly string[];
}

/**
 * 现场与锚点不一致时给用户的**固定**文案（两端共用这一份字面量）。
 *
 * 说清三件事：哪一步之后变的（预览/分析）、这份计划写的是什么（固定值）、
 * 现在做了什么（未写入）。不给"已自动重试"这种做不到的暗示。
 */
export const PLAN_SOURCE_STALE_MESSAGE =
  "数据来源在预览后发生了变化，这份计划写入的是预览时算出的固定值，未写入；请按当前数据重新生成。";

/** 读不到现场时的固定文案：与"变了"分开，因为它们要的下一步不一样。 */
export const PLAN_SOURCE_UNVERIFIABLE_MESSAGE =
  "无法核对当前数据来源是否仍与这份计划一致，未写入；请重新读取后生成计划。";

/**
 * **本应有锚点却没有**时的固定文案。
 *
 * 与上面两条并列而不是合并：它们的成因不同，但用户要做的事是同一步（重新生成）。
 * 分开写是为了不把"没记录过"说成"变了" —— 后者会让人以为数据被谁改过。
 */
export const PLAN_SOURCE_MISSING_MESSAGE =
  "这份计划写入的是算好的固定值，但没有记录它是按哪一版数据算出来的，未写入；请重新读取数据后生成计划。";

/** 锚点覆盖的表与当前授权范围不是同一批：不能拿它当"这一版"的证据。 */
export const PLAN_SOURCE_SCOPE_MISMATCH_MESSAGE =
  "这份计划记录的来源范围与当前授权范围不一致，无法确认它是按哪一版数据算出来的，未写入；请重新生成计划。";

/**
 * 动作体里携带**算好的固定值**的写动作。
 *
 * `writeFormulas` 不在其中：公式引用的是单元格，宿主执行时自己求值。
 * `filterRange.values` / `setDataValidation.values` 是用户点名的取值而非算出的
 * 结果，也不算烘焙。
 */
export const BAKED_WRITE_ACTION_TYPES = ["writeValues", "writeTable"] as const;

/**
 * 这份计划里有没有"算好的固定值"—— 也就是**本来需不需要**来源锚点。
 *
 * 这是保守的动作类型分类；可信字面值例外由闸门验证封印。不看指纹是否为空（空代表缺失），
 * 而看计划写了什么。执行时求值的计划源数据改了正是它该重算的场景，不该被拦。
 */
export function planHasBakedValues(
  plan: { actions?: readonly { type?: string | null }[] } | null | undefined
): boolean {
  const actions = plan?.actions ?? [];
  return actions.some(
    (action) =>
      typeof action?.type === "string" &&
      (BAKED_WRITE_ACTION_TYPES as readonly string[]).includes(action.type)
  );
}

/** 锚点为什么不可用。四种成因都要能说出来，而不是合并成一个 `null`。 */
export type PlanSourceAnchorProblem =
  /** 计划上没有指纹（装配时没取到，或本来就没有桥）。 */
  | "no-fingerprint"
  /** 有指纹没有覆盖表：不知道它按哪些表算的，无法复算。 */
  | "no-sheets"
  /** 锚点来自别的轮次（`taskId` 不一致）：那是另一次读取的现场。 */
  | "foreign-task"
  /** 锚点覆盖的表与当前授权范围不是同一批。 */
  | "scope-mismatch";

/**
 * 写前判定的结论。
 *
 * `ok` 带 `basis`，是为了让"不需要锚点"和"比对通过"在证据里分得开 —— 两者都
 * 可以写，但只有后者证明过来源没变。
 */
export type PlanSourceGateVerdict =
  | { verdict: "ok"; basis: "runtime-evaluated" | "independent-review" | "anchor" | "literal" }
  | { verdict: "changed" }
  | { verdict: "unverifiable" }
  | { verdict: "refused"; problem: PlanSourceAnchorProblem };

/** 计划上能推断出的缺失成因（装配方另有更精确的成因时优先采用）。 */
function inferAnchorProblem(plan: SourceAnchorPlan | null | undefined): PlanSourceAnchorProblem {
  const fingerprint =
    typeof plan?.sourceFingerprint === "string" ? plan.sourceFingerprint.trim() : "";
  return fingerprint ? "no-sheets" : "no-fingerprint";
}

/**
 * 这份计划要不要**现在**去读一次现场指纹。
 *
 * 读宿主是有代价的，而不需要锚点的路径（可信字面值、执行时求值、本机已独立复核）
 * 根本用不上它。两端都先问这一句，避免各自写一份"什么时候该读"。
 */
export function planSourceComparisonNeeded(input: {
  plan: { actions?: readonly { type?: string | null }[] } | null | undefined;
  anchor: PlanSourceAnchor | null;
  independentSourceReview?: boolean;
}): boolean {
  if (isTrustedLiteralPlan(input.plan)) return false;
  if (!planHasBakedValues(input.plan)) return false;
  if (input.independentSourceReview) return false;
  return input.anchor !== null;
}

/**
 * 锚点覆盖的表是不是**都还在**当前授权范围内。
 *
 * 用"包含"而不是"相等"：
 * - 锚点 ⊂ 授权（比如值复制计划只按它真读的那张表记账）是**更精确**的锚点，
 *   按它比对没有放水；
 * - 锚点 ⊄ 授权（用户收窄了范围，或这枚锚点本来就是别一轮的范围）意味着写前
 *   复核要读的表已经不在授权内 —— 读不了也就比不了，必须拒绝。
 */
export function anchorSheetsWithin(
  anchor: PlanSourceAnchor,
  authorizedSheets: readonly string[]
): boolean {
  if (anchor.sheets.length === 0) return false;
  const scope = new Set(
    authorizedSheets.map((name) => String(name ?? "").trim().toLocaleLowerCase()).filter(Boolean)
  );
  return anchor.sheets.every((name) => scope.has(name.trim().toLocaleLowerCase()));
}

/**
 * 写前闸门的唯一判据（双端共用）。
 *
 * 顺序与理由：
 * 0. 可信编译器返回且动作未被修改的字面值赋值 → 不需要计算来源锚点；
 * 1. 没有固定值的计划 → 不需要锚点（执行时求值）；
 * 2. 本机已对源区域做过独立复核 → 不需要计划级锚点（本地计算/计算动作）；
 * 3. 有固定值却没锚点 → **拒绝**（不是"放行"，也不是拿现读补一个）；
 * 4. 锚点覆盖的表已不在授权范围内 → 拒绝（写前复核读不到它们，也就比不了）；
 * 5. 现场读不到 → 拒绝（"读不到"不等于"没变"）；
 * 6. 否则比对指纹。
 */
export function judgePlanSourceGate(input: {
  plan: { actions?: readonly { type?: string | null }[] } & SourceAnchorPlan;
  anchor: PlanSourceAnchor | null;
  /** `anchor` 为 null 时由装配方给出的成因；缺省按计划字段推断。 */
  problem?: PlanSourceAnchorProblem;
  authorizedSheets: readonly string[];
  liveFingerprint: string | null;
  independentSourceReview?: boolean;
}): PlanSourceGateVerdict {
  if (isTrustedLiteralPlan(input.plan)) return { verdict: "ok", basis: "literal" };
  if (!planHasBakedValues(input.plan)) return { verdict: "ok", basis: "runtime-evaluated" };
  if (input.independentSourceReview) return { verdict: "ok", basis: "independent-review" };
  if (!input.anchor) {
    return { verdict: "refused", problem: input.problem ?? inferAnchorProblem(input.plan) };
  }
  if (!anchorSheetsWithin(input.anchor, input.authorizedSheets)) {
    return { verdict: "refused", problem: "scope-mismatch" };
  }
  const compared = comparePlanSourceAnchor(input.anchor, input.liveFingerprint);
  if (compared === "unverifiable") return { verdict: "unverifiable" };
  if (compared === "changed") return { verdict: "changed" };
  return { verdict: "ok", basis: "anchor" };
}

/** 结论 → 给用户的固定文案（`ok` 没有话要说）。 */
export function planSourceGateMessage(verdict: PlanSourceGateVerdict): string | null {
  switch (verdict.verdict) {
    case "ok":
      return null;
    case "changed":
      return PLAN_SOURCE_STALE_MESSAGE;
    case "unverifiable":
      return PLAN_SOURCE_UNVERIFIABLE_MESSAGE;
    case "refused":
      return verdict.problem === "scope-mismatch"
        ? PLAN_SOURCE_SCOPE_MISMATCH_MESSAGE
        : PLAN_SOURCE_MISSING_MESSAGE;
  }
}

/** 计划自带的锚点；缺任一半（指纹或覆盖表）都算没有锚点。 */
export function planSourceAnchor(
  plan: SourceAnchorPlan | null | undefined
): PlanSourceAnchor | null {
  const fingerprint =
    typeof plan?.sourceFingerprint === "string" ? plan.sourceFingerprint.trim() : "";
  const sheets = (plan?.sourceFingerprintSheets ?? [])
    .filter((name): name is string => typeof name === "string")
    .map((name) => name.trim())
    .filter(Boolean);
  if (!fingerprint || sheets.length === 0) return null;
  return { fingerprint, sheets };
}

/** 现场读数与锚点的关系。三态：不同结论给不同下一步。 */
export type PlanSourceVerdict = "unchanged" | "changed" | "unverifiable";

export function comparePlanSourceAnchor(
  anchor: PlanSourceAnchor,
  liveFingerprint: string | null
): PlanSourceVerdict {
  const live = typeof liveFingerprint === "string" ? liveFingerprint.trim() : "";
  if (!live) return "unverifiable";
  return live === anchor.fingerprint ? "unchanged" : "changed";
}

/**
 * 从**这一轮的桥绑定**取锚点。
 *
 * 绑定的指纹是"这一轮第一次受控读取时固定的现场"，也就是模型拿到的数据那一版：
 * - `taskId` 必须与计划所属轮次一致，否则拿的是别的轮次的现场；
 * - 指纹为 `null`（桥在推迟读源、或读失败）⇒ 没有锚点，不猜。
 *
 * 两个宿主都把 `binding()` 交给它，避免一端用绑定、一端用"到达时现读"。
 */
export function sourceAnchorFromBinding(
  binding:
    | {
        taskId: string;
        sourceFingerprint: string | null;
        authorizedSheetNames: readonly string[];
      }
    | null
    | undefined,
  turnId: string | null | undefined
): PlanSourceAnchor | null {
  if (!binding) return null;
  const expected = typeof turnId === "string" ? turnId.trim() : "";
  if (!expected || binding.taskId !== expected) return null;
  return planSourceAnchor({
    sourceFingerprint: binding.sourceFingerprint,
    sourceFingerprintSheets: binding.authorizedSheetNames
  });
}

/** 计划带上锚点（不可变返回，便于"计划到达时装配一次"）。 */
export function stampPlanSourceAnchor<T extends SourceAnchorPlan>(
  plan: T,
  anchor: PlanSourceAnchor | null
): T {
  if (!anchor) return { ...plan, sourceFingerprint: null, sourceFingerprintSheets: [] };
  return {
    ...plan,
    sourceFingerprint: anchor.fingerprint,
    sourceFingerprintSheets: [...anchor.sheets]
  };
}

/** 两批表名是否同一集合（顺序无关；用于"能不能用这枚锚点"的判断）。 */
export function sameSheetSet(
  left: readonly string[],
  right: readonly string[]
): boolean {
  const normalize = (names: readonly string[]) =>
    [...new Set(names.map((name) => name.trim().toLocaleLowerCase()))]
      .filter(Boolean)
      .sort();
  return JSON.stringify(normalize(left)) === JSON.stringify(normalize(right));
}
