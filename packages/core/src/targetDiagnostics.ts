/**
 * 「这次为什么要求确认范围」的**最小诊断**（只有布尔值与原因码）。
 *
 * 存在理由（2026-11-01 真机）：需求在真机上仍弹「请确认本次使用的数据范围」，
 * 而提示标题只有一种文案，无法区分是①元数据比对失败 ②分析证据不足
 * ③调用方强制确认 ④**旧的 notice 没被清掉**。WPS 任务窗格开不了 devtools，
 * 所以只能把判定现场的结构化布尔值送出来。
 *
 * 纪律（不可放宽）：
 * - 只收布尔值 / 原因码 / 计数。**不收**文件名、路径、工作表名、身份键、证据键、
 *   单元格内容 —— 这些一律不进任何观测通道。
 * - 这不是授权或判据的一部分：拿掉它，行为不变。
 */

/** 哪一道闸拦下的：元数据比对，还是分析证据闸。 */
export type TargetDiagnosticGate = "metadata" | "analysis";

/** 拦住的原因码（每个都有唯一判别点）。 */
export type TargetDiagnosticBranch =
  /** `check()`：读不出目标身份。 */
  | "metadata_unavailable"
  /** `check()`：文件变了。 */
  | "metadata_file_changed"
  /** `check()`：工作表集合/名字变了。 */
  | "metadata_scope_changed"
  /** `check(onRetry, true)`：调用方自己要求一次确认（与证据强弱无关）。 */
  | "forced_confirmation"
  /** 分析证据读不出来（重试用尽）。 */
  | "evidence_missing"
  /** 三份证据键不一致（冲突），不是"变了"。 */
  | "evidence_conflict"
  /** 证据一致但只是弱证据：既无对象级证明，也无文档级证明。 */
  | "weak_evidence"
  /** 有文档级证据，但调用方明确拒绝以它放行。 */
  | "document_bound_rejected";

/** 放行时走的是哪一条（用来排除"提示来自别处"）。 */
export type TargetDiagnosticAllowedVia = "conclusive" | "consent" | "documentBound";

/** 工作表稳定 ID 的可读程度（只看得到几成，不记录具体值）。 */
export type TargetStableIdGrade = "all" | "partial" | "none" | "unknown";

/**
 * 宿主适配层能提供的、**只用于诊断**的观测事实。
 *
 * 它们不是判据的一部分：`documentBound` 已经由适配层算好，这里只是把"为什么"
 * 摊开成布尔值/等级，好在真机上区分"没有路径身份"与"有路径但读不到稳定 ID"。
 */
export type TargetDiagnosticHost = {
  /** 每张授权表是否都报出了稳定 ID。 */
  stableId?: TargetStableIdGrade;
  /** 本次读数是否拿到了路径身份（未保存工作簿为 false）。 */
  pathIdentity?: boolean;
};

export type TargetDiagnostic = {
  v: 1;
  stage: "target";
  outcome: "blocked" | "allowed";
  /** 拦住时的闸；放行时固定 `analysis`。 */
  gate: TargetDiagnosticGate;
  branch?: TargetDiagnosticBranch;
  via?: TargetDiagnosticAllowedVia;
  /** 对象级证据是否成立（不是"身份相同"，是"这一瞬间两次读拿到同一对象"）。 */
  conclusive: boolean;
  /** 文档级证据是否成立（路径身份 + 授权集合与活体读数一致）。 */
  documentBound: boolean;
  /** 该证据键是否已有内存同意（刷新后必然为 false）。 */
  hasConsent: boolean;
  /** 调用方是否要求"不许靠文档级证据放行"。 */
  explicitRequired: boolean;
  stableId: TargetStableIdGrade;
  /** 本次读数是否拿到路径身份（未保存工作簿为 false）。 */
  pathIdentity: boolean;
  /** 当前提示轮次；同一轮 = 还是同一条提示，没有开新的一轮。 */
  round: number;
  /** 本页面内第几次判定（单调递增，用来区分"新请求"与"旧提示"。 */
  seq: number;
};

/** 造一条诊断。纯函数，宿主只需填自己知道的那几个布尔值。 */
export type TargetDiagnosticInput = {
  outcome: "blocked" | "allowed";
  gate: TargetDiagnosticGate;
  branch?: TargetDiagnosticBranch;
  via?: TargetDiagnosticAllowedVia;
  conclusive?: boolean;
  documentBound?: boolean;
  hasConsent?: boolean;
  explicitRequired?: boolean;
  stableId?: TargetStableIdGrade;
  pathIdentity?: boolean;
  round: number;
  seq: number;
};

export function targetDiagnostic(input: TargetDiagnosticInput): TargetDiagnostic {
  return {
    v: 1,
    stage: "target",
    outcome: input.outcome,
    gate: input.gate,
    ...(input.branch ? { branch: input.branch } : {}),
    ...(input.via ? { via: input.via } : {}),
    conclusive: input.conclusive === true,
    documentBound: input.documentBound === true,
    hasConsent: input.hasConsent === true,
    explicitRequired: input.explicitRequired === true,
    stableId: input.stableId ?? "unknown",
    pathIdentity: input.pathIdentity === true,
    round: input.round,
    seq: input.seq
  };
}
