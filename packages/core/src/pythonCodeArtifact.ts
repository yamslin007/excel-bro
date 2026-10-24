/**
 * Python 代码工件 —— 代码块要显示的**真实**状态（Excel / WPS 共用，L2 纯语义层）。
 *
 * ---------------------------------------------------------------------------
 * 这一层要防的四件事
 * ---------------------------------------------------------------------------
 * 1. **没有证据就不要说"正在写 Python"。** 只有当上游真的宣告进入代码生成阶段
 *    （`stage: "generating"`）才允许这么说。当前 SDK 通道把整段代码随工具调用
 *    一次性送达，因此它的第一条状态是 `generated`（代码已生成），**不是**
 *    `generating` —— 界面不许把"完整代码已到"演成"正在逐字写"。
 * 2. **增量必须可去重。** 供应商给的片段有重放、乱序、重复结束三种情况。同一次
 *    生成尝试（`attemptId`）内按 `version` 单调推进；低版本一律忽略，绝不把
 *    同一段代码粘两遍。
 * 3. **未完成的草稿不许执行。** `complete === false` 的工件被显式标成草稿，
 *    界面上的"运行"能力根本不存在（本仓库也没有绕过校验的执行入口）。
 * 4. **"已运行"必须指向确切版本。** `executedVersion` 只有在真的进入执行阶段后
 *    才被接受；与现实版本不一致时，界面只能说"运行的不是屏幕上这份代码"。
 *
 * 本模块不碰宿主、不碰 React、不发事件，所有时间/上限由调用方给。
 */

export type PythonCodeStage =
  /** 上游确有"正在生成"的证据（增量流）。 */
  | "generating"
  /** 代码已完整送达。 */
  | "generated"
  /** 正在校验（形状/长度/授权等确定性检查）。 */
  | "validating"
  /** 正在隔离环境执行（**只有执行器真的启动后**才允许出现）。 */
  | "running"
  /** 隔离执行完成，结果已登记。 */
  | "analysis_completed"
  /** 正在整理结果（模型总结阶段）。 */
  | "summarizing"
  /** 用户暂停导致这次执行中断。 */
  | "paused"
  /**
   * 未收到执行确认，包括明确未启动以及旧镜像无法确认是否进入执行的情形。
   * 具体原因来自可信服务端，不以创建进程代替代码执行证据。
   */
  | "not_started"
  | "generation_failed"
  | "validation_failed"
  | "execution_failed";

export const PYTHON_CODE_STAGE_LABEL: Record<PythonCodeStage, string> = {
  generating: "正在生成 Python 分析代码",
  generated: "代码已生成",
  validating: "正在校验代码",
  running: "正在隔离环境运行",
  analysis_completed: "分析完成",
  summarizing: "正在整理结果",
  paused: "已暂停",
  not_started: "未确认执行",
  generation_failed: "代码生成失败",
  validation_failed: "代码校验未通过",
  execution_failed: "执行失败"
};

/**
 * 正文正在按 `codeId` 取回时显示的话。
 *
 * 代码不进账本 ⇒ 事件里只有状态。界面在取回途中必须说"正在取回"，而不是
 * 直接说"没有保留" —— 后者是把没做完的事说成结论。
 */
export const PYTHON_CODE_BODY_LOADING_NOTE = "正在取回代码…";

/**
 * 取不回来时的唯一说法（登记表过期 / 任务对不上 / 本机服务未连通都一样）。
 *
 * 三种原因对用户是同一件事：手上没有这段代码。界面不得展示内部路径、
 * 请求 ID 或原始错误，也不得用状态标签假装代码还在。
 */
export const PYTHON_CODE_BODY_UNAVAILABLE_NOTE =
  "这次只保留了状态，代码正文没有保留；可以让我重新生成一份。";

/** 终态分支：同一次生成尝试里不会被"更晚的好消息"改写。 */
const FAILED_STAGES: ReadonlySet<PythonCodeStage> = new Set([
  "generation_failed",
  "validation_failed",
  "execution_failed",
  "paused"
]);

/** 已经进入过执行阶段的那些状态：只有它们才可能携带 `executedVersion`。 */
const EXECUTION_STAGES: ReadonlySet<PythonCodeStage> = new Set([
  "running",
  "analysis_completed",
  "summarizing",
  "paused",
  "execution_failed"
]);

/** 正常路径上的先后次序；失败分支不参与比较（直接终态）。 */
const STAGE_RANK: Record<PythonCodeStage, number> = {
  generating: 0,
  generated: 1,
  validating: 2,
  running: 3,
  analysis_completed: 4,
  summarizing: 5,
  paused: 3,
  not_started: 3,
  generation_failed: 1,
  validation_failed: 2,
  execution_failed: 3
};

export interface PythonCodeLimits {
  /** 单次展示/接收的代码字符上限。 */
  maxChars: number;
  /** 折叠时预览的行数。 */
  previewLines: number;
  /** 超过这个行数的代码默认折叠（含行号/滚动）。 */
  collapseAfterLines: number;
}

export const DEFAULT_PYTHON_CODE_LIMITS: PythonCodeLimits = {
  maxChars: 20000,
  previewLines: 6,
  collapseAfterLines: 24
};

export interface PythonCodeArtifact {
  /** 一次"代码工件"的稳定身份（同一次分析里多次生成 = 多个工件）。 */
  codeId: string;
  /** 一次生成**尝试**的身份。重新生成 = 新尝试。 */
  attemptId: string;
  stage: PythonCodeStage;
  /** 已收到的代码文本（可能是前缀，也可能是空——见 `codeRetained`）。 */
  code: string;
  /** 上游是否宣告本次生成结束。 */
  complete: boolean;
  /** 因上限被截断：此时"复制全部"不可用。 */
  truncated: boolean;
  /** 真正被校验并执行的版本号；从未进入执行为 `null`。 */
  executedVersion: number | null;
  /** 工件自身的版本号（每次内容推进 +1）。 */
  version: number;
  /** 最近一次推进该工件的服务端事件序号（用于乱序/重放判别）。 */
  sequence?: number;
  /**
   * 增量内容**缺片**：收到的片段版本跳过了中间某一版。
   *
   * 缺一片就不能把拼出来的东西当完整代码：`complete` 会被压成 `false`，
   * "复制全部"也不提供。补齐（收到中间片或收到完整快照）后自动恢复。
   */
  gap: boolean;
  /** 服务端报告的正文字符数（即使正文没有被前端拿到）。 */
  chars?: number;
  /**
   * 正文是否随事件保留到了本端。
   *
   * 代码**不进持久账本**（仓库既有边界：模型生成的代码只在任务内存短期存在），
   * 因此断线重连 / 恢复轮次拿到的是"只有状态、没有正文"的工件。这时界面必须
   * 说清楚"正文没有保留"，不能凭标题假装代码还在。
   */
  codeRetained: boolean;
  /**
   * 没能开始执行时的**具体原因**（服务端给的固定文案）。
   *
   * 只带这一句，不带路径、请求 ID、原始供应商错误。原因必须能落地到界面上，
   * 否则"未开始执行"就成了一句没有信息量的话。
   */
  reason?: string;
}

export interface PythonCodeEvent {
  codeId: string;
  attemptId: string;
  stage: PythonCodeStage;
  /** 内容版本；缺省视为 `1`。 */
  version?: number;
  /** 完整代码（"只有完整代码"的供应商走这条）。 */
  code?: string;
  /** 增量片段（真有增量的供应商走这条）。 */
  delta?: string;
  /** 正文字符数（正文未随事件下发时也能如实告知规模）。 */
  chars?: number;
  truncated?: boolean;
  complete?: boolean;
  /** 实际被执行的版本；非执行阶段会被忽略。 */
  executedVersion?: number | null;
  sequence?: number;
  /** 这一帧是否携带了正文；`false` 表示本端只知道状态。 */
  codeRetained?: boolean;
  /** 未开始执行的具体原因（服务端固定文案）。 */
  reason?: string;
}

export type PythonCodeIgnoreReason =
  | "replay"
  | "stale"
  | "unexecuted-version";

export interface PythonCodeApplyResult {
  artifact: PythonCodeArtifact;
  ignored: PythonCodeIgnoreReason | null;
}

function lineCount(text: string): number {
  if (!text) return 0;
  let lines = 1;
  for (let index = 0; index < text.length; index += 1) {
    if (text[index] === "\n") lines += 1;
  }
  return lines;
}

function clip(text: string, maxChars: number): { code: string; truncated: boolean } {
  if (text.length <= maxChars) return { code: text, truncated: false };
  return { code: text.slice(0, maxChars), truncated: true };
}

/**
 * 把一条代码事件归并进工件。
 *
 * ---------------------------------------------------------------------------
 * 判别顺序（判重用**事件身份**，绝不用文本相似）
 * ---------------------------------------------------------------------------
 * 1. **迟到帧**：`sequence <= 已归并序号` → 整帧丢弃。这一条同时挡住"旧尝试的
 *    迟到事件覆盖当前尝试"（旧尝试的序号一定更小），所以它必须在换尝试之前判。
 * 2. **换尝试**：`attemptId` 变了 → 工件重建（这是"重新生成"，不是"续写"）。
 * 3. **内容推进**：只有 `version` 严格变大的帧才改内容。
 *    - 完整快照（`code`）→ 整体替换，并**清掉缺片标记**（快照能补齐）；
 *    - 增量片段（`delta`）→ **无条件追加**，不按文本判重。
 *      "这一段和末尾一样所以是重复"是错的：`f(g(1` + `)` + `)` 里第二个 `)`
 *      是真实内容。真正该判重的是身份，不是文本。
 * 4. **重放**：`version` 没变大的一帧，内容一律不动，只在阶段更靠后时推进阶段。
 * 5. **缺片**：增量帧的 `version` 跳过了中间版本（`version > 已收版本 + 1`）→
 *    记 `gap`，此后 `complete` 一律压成 `false`（缺片就不能算完整）。
 * 6. `executedVersion` 只在执行阶段被接受；否则记为"忽略"，避免界面拿一个
 *    从不存在的版本号声称"已运行"。
 */
export function applyPythonCodeEvent(
  current: PythonCodeArtifact | null,
  event: PythonCodeEvent,
  limits: PythonCodeLimits = DEFAULT_PYTHON_CODE_LIMITS
): PythonCodeApplyResult {
  const version = Math.max(1, Math.floor(event.version ?? 1));
  const hasSnapshot = typeof event.code === "string";
  const hasDelta = typeof event.delta === "string";
  const incomingText = hasSnapshot ? event.code! : hasDelta ? event.delta! : null;

  // 1) 迟到 / 同序号重放：整帧丢弃（含旧尝试的迟到帧）。
  if (
    current &&
    event.sequence != null &&
    current.sequence != null &&
    event.sequence <= current.sequence
  ) {
    return {
      artifact: current,
      ignored: event.sequence < current.sequence ? "stale" : "replay"
    };
  }

  const isNewAttempt = !current || current.attemptId !== event.attemptId;

  if (isNewAttempt) {
    const base = incomingText ?? "";
    const clipped = clip(base, limits.maxChars);
    // 新尝试的第一帧若是**增量**且版本 > 1，说明前面几片没收到 —— 一开始就是
    // 残缺的，不能假装完整。
    const gap = hasDelta && version > 1;
    return {
      artifact: {
        codeId: event.codeId,
        attemptId: event.attemptId,
        stage: event.stage,
        code: clipped.code,
        complete:
          !gap &&
          (event.complete ?? (event.stage === "generated" && hasSnapshot)),
        truncated: Boolean(event.truncated) || clipped.truncated,
        executedVersion:
          EXECUTION_STAGES.has(event.stage) && event.executedVersion != null
            ? event.executedVersion
            : null,
        version,
        gap,
        codeRetained: base.length > 0 || event.codeRetained !== false,
        ...(event.chars != null ? { chars: event.chars } : {}),
        ...(event.sequence != null ? { sequence: event.sequence } : {}),
        ...(event.reason ? { reason: event.reason } : {})
      },
      ignored: null
    };
  }

  let ignored: PythonCodeIgnoreReason | null = null;
  let code = current.code;
  let truncated = current.truncated;
  let complete = current.complete;
  let gap = current.gap;
  let nextVersion = current.version;

  if (version > current.version) {
    if (hasSnapshot) {
      // 完整快照：整体替换。它同时是"补齐"——缺片标记消掉。
      const clipped = clip(event.code!, limits.maxChars);
      code = clipped.code;
      truncated = Boolean(event.truncated) || clipped.truncated;
      gap = false;
      if (event.complete != null) complete = event.complete;
      nextVersion = version;
    } else if (hasDelta) {
      // 增量：无条件追加。文本相同不代表事件重复。
      if (version > current.version + 1) gap = true;
      code = code + event.delta!;
      const clipped = clip(code, limits.maxChars);
      code = clipped.code;
      truncated = truncated || clipped.truncated;
      gap = gap || Boolean(event.truncated);
      if (event.complete != null) complete = event.complete;
      nextVersion = version;
    } else {
      // 只推进阶段，不动内容。
      if (event.complete != null) complete = event.complete;
      nextVersion = version;
    }
    if (gap) complete = false;
  } else {
    // 同版本或更低版本 = 这一版内容已经处理过。内容不动，只在阶段更靠后时推进。
    ignored = "replay";
    if (event.complete === true && !gap) complete = true;
  }

  let stage = current.stage;
  if (event.stage !== current.stage) {
    const advancing =
      FAILED_STAGES.has(event.stage) ||
      STAGE_RANK[event.stage] > STAGE_RANK[current.stage];
    const allowed = !FAILED_STAGES.has(current.stage) && advancing;
    if (allowed) stage = event.stage;
  }
  if (stage !== current.stage) ignored = null;

  let executedVersion = current.executedVersion;
  if (event.executedVersion != null) {
    if (EXECUTION_STAGES.has(stage)) {
      executedVersion = event.executedVersion;
    } else if (!EXECUTION_STAGES.has(current.stage)) {
      ignored = "unexecuted-version";
    }
  }

  return {
    artifact: {
      ...current,
      codeId: event.codeId || current.codeId,
      stage,
      code,
      complete,
      truncated,
      executedVersion,
      version: nextVersion,
      gap,
      codeRetained: code.length > 0 || current.codeRetained,
      ...(event.chars != null ? { chars: event.chars } : {}),
      ...(event.sequence != null ? { sequence: event.sequence } : {}),
      ...(event.reason ? { reason: event.reason } : {})
    },
    ignored
  };
}

export interface PythonCodeView {
  label: string;
  /** 本端是否真的拿到了正文。 */
  hasCode: boolean;
  /**
   * 正文没拿到，但**还没试过**按 `codeId` 取回。
   *
   * 代码不进账本，所以"事件里没有正文"是常态而不是终点：真正的结论要么是取回
   * 成功，要么是取不回来（登记表过期 / 任务对不上）。在试过之前不许说
   * "正文没有保留" —— 那是把没做过的事说成事实。
   */
  retrievable: boolean;
  /** 代码是否还在生成中（此时不提供"复制全部"）。 */
  draft: boolean;
  /** 折叠状态下的预览（按行截断，带省略标记）。 */
  preview: string;
  /** 展开后要展示的代码正文。 */
  body: string;
  /** 超过折叠阈值，需要"查看 / 收起"。 */
  collapsible: boolean;
  /** 是否因上限被截断（此时不得提供"复制全部"）。 */
  truncated: boolean;
  /** 增量缺片：拼出来的内容不完整，不能当完整代码用。 */
  gap: boolean;
  /** 服务端报告的字符数（正文没保留时也能说清规模）。 */
  chars: number | null;
  /** 能否复制全部：必须真的拿到了全文。 */
  canCopyAll: boolean;
  /** 界面上"已运行"的代码是否就是屏幕上这一版。 */
  executedCodeMatches: boolean;
  /** 失败/暂停/正文缺失时给用户的说明（不含内部路径、请求 ID、原始供应商错误）。 */
  note?: string;
}

/**
 * 由工件推导界面要展示的东西。
 *
 * **没有任何"执行"入口**：这个仓库只有"生成计划 → 预览 → 用户确认 → 白名单执行"
 * 一条写入路径，代码块是只读展示。
 *
 * 正文缺失（断线重连后的恢复轮次、账本里的历史帧）时仍然返回视图：状态是真实的，
 * 只是要说清"代码正文没有保留"，而不是把整块藏起来假装没发生过分析。
 */
export function pythonCodeView(
  artifact: PythonCodeArtifact | null | undefined,
  limits: PythonCodeLimits = DEFAULT_PYTHON_CODE_LIMITS
): PythonCodeView | null {
  if (!artifact) return null;
  const hasCode = artifact.code.length > 0;
  const lines = lineCount(artifact.code);
  const draft = !artifact.complete || artifact.stage === "generating";
  const previewLines = artifact.code.split("\n", limits.previewLines);
  const clippedPreview = previewLines.join("\n");
  const collapsible = hasCode && lines > limits.collapseAfterLines;
  const executedCodeMatches =
    hasCode &&
    artifact.executedVersion !== null &&
    artifact.executedVersion === artifact.version &&
    !draft;
  const retrievable = !hasCode;
  const note = noteForStage(artifact.stage, draft, hasCode, artifact.gap, artifact.reason);
  return {
    label: PYTHON_CODE_STAGE_LABEL[artifact.stage],
    hasCode,
    retrievable,
    draft,
    preview: collapsible ? clippedPreview : artifact.code,
    body: artifact.code,
    collapsible,
    truncated: artifact.truncated,
    gap: artifact.gap,
    chars: artifact.chars ?? (hasCode ? artifact.code.length : null),
    canCopyAll: hasCode && artifact.complete && !artifact.truncated && !artifact.gap,
    executedCodeMatches,
    ...(note ? { note } : {})
  };
}

function noteForStage(
  stage: PythonCodeStage,
  draft: boolean,
  hasCode: boolean,
  gap: boolean,
  reason?: string
): string | undefined {
  // 没有执行确认时不猜是否运行过；明确的拒绝原因由服务端提供。
  if (stage === "not_started") {
    return reason ?? "未收到代码开始执行的确认，表格没有被改动。";
  }
  // 缺片优先于"草稿"：草稿说的是"还没写完"，缺片说的是"写完但少了中间一段"，
  // 后者更严重 —— 拼出来的代码不能拿去用，也不许复制。
  if (gap) return "代码片段没有收齐（缺少中间一段），不能把这段当作完整代码。";
  if (stage === "generation_failed") return "这次没有拿到可运行的代码，可以让我重新生成。";
  if (stage === "validation_failed") return "代码没有通过本地校验，不会被执行。";
  if (stage === "execution_failed") return "隔离环境里的执行没有成功，表格没有被改动。";
  if (stage === "paused") return "已暂停：这段代码没有继续运行。";
  if (draft && hasCode) return "未完成草稿，不能运行。";
  return undefined;
}

/**
 * 复制用的正文。
 *
 * 只有真的取得全文（未截断、已生成完）时才允许"复制全部"；否则返回 `null`，
 * 调用方不得提供该入口 —— 让用户复制半截代码去做别的事，比不提供更糟。
 */
export function pythonCodeCopyText(
  artifact: PythonCodeArtifact | null | undefined,
  limits: PythonCodeLimits = DEFAULT_PYTHON_CODE_LIMITS
): string | null {
  const view = pythonCodeView(artifact, limits);
  if (!view || !view.canCopyAll) return null;
  return artifact?.code ?? null;
}

/** 从活动步骤里挑出要展示的代码工件：最后一条带工件的步骤胜出。 */
export function latestPythonCode<T extends { pythonCode?: PythonCodeArtifact | null }>(
  steps: readonly T[]
): PythonCodeArtifact | null {
  for (let index = steps.length - 1; index >= 0; index -= 1) {
    const candidate = steps[index].pythonCode;
    if (candidate) return candidate;
  }
  return null;
}
