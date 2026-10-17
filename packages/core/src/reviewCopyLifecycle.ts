import capabilities from '../../../config/capabilities.json';

export const REVIEW_COPY_LIFECYCLE = capabilities.reviewCopyLifecycle;

/**
 * 审阅副本生命周期（Excel / WPS **共用状态语义**，宿主操作留在各自适配器）。
 *
 * 本模块只做三件事，都是确定性的：
 *  1. 副本身份/所有权的**形状**（谁创建的、拿什么认它、有没有磁盘文件）；
 *  2. 结束时把「业务结果 × 关闭结果 × 临时文件结果」映射成**一组文案与可用动作**；
 *  3. 临时文件路径的**归一化与越界校验**（只允许专用临时目录内的我方登记文件）。
 *
 * 它不碰宿主：关闭窗口、删文件都由适配器做。这样两端不会各写一套"关闭失败却报成功"的口径。
 *
 * **铁律**：
 *  - 副本身份不得靠"工作簿N"、`EB_` 前缀、活动窗口或文件名相似。真正的判据是本次会话
 *    随机生成的 `copySheetName`（20 位十六进制 nonce）+ 宿主侧的未保存/身份元数据 + 唯一性。
 *  - 未保存的工作簿**不是**磁盘文件。`tempFilePath === null` 时绝不允许出现"已删除临时文件"。
 *  - 写回失败/部分完成/结果不确定 ⇒ **保留**副本供核对，不自动关闭、不删除。
 *  - 重试/重复清理必须幂等：已由用户关闭过 ⇒ `already_closed` 属于成功收尾，不是失败。
 *
 * **无提示丢弃副本前必须过三项判断，且不可互相顶替**：
 *  ① 副本身份可信（nonce + 未保存标记 + 唯一性）；
 *  ② 没被另存/替换（否则那是用户自己的文件）；
 *  ③ **副本里没有额外的人工修改** —— ①和②都成立时第三人仍可能往副本里敲了东西。
 *  ① 与 ② 成立**不能**推出 ③；③ 判不出来就不许静默丢弃。判据入口是
 *  `describeCopyDiscardSafety(verdict)`：只有逐格读回并确认 `unchanged` 才放行，
 *  `changed` ⇒ `extra_edits`，`unknown` ⇒ `edits_unverified`。
 *
 * **关闭请求发出但窗口仍在，不等于"宿主在问是否保存"**：
 *  没有真实对话框证据时只能报 `close_unconfirmed`（关闭尚未确认）并提示用户查看窗口，
 *  不许推断原因 —— 关得慢、被别的模态框挡住，都会让窗口多停一会儿。
 *
 * `user_cancelled` 已被删除：没有任何适配器能观察到"用户点了取消"。`Window.close()`
 * 返回 `false` 是**歧义**信号（被取消 / 被推迟 / 被拒），一律按 `close_unconfirmed` 报。
 */

export type ReviewCopyKind = 'excel-native' | 'excel-compat' | 'wps-native';

export type ReviewCopyOwnership = {
  kind: ReviewCopyKind;
  /** 本次审阅会话标识（同一轮预览/审阅内不变，用于防止跨会话误关）。 */
  sessionId: string;
  /** 本次随机生成的副本工作表名 —— 真正的清理判据，不是前缀匹配。 */
  copySheetName: string;
  /** 宿主侧定位串（Excel 窗口名 / WPS 副本名与全名）；只用于展示与二次核对，不单独构成判据。 */
  copyLocator: string;
  /** 原表定位串；原表**永远**不在可关闭/可删除集合里。 */
  originalLocator: string;
  /** 本次登记的临时文件（规范化绝对路径）；没有磁盘文件时为 null。 */
  tempFilePath: string | null;
};

/** 关闭副本的结果。`already_closed` 是**成功**收尾，`close_unconfirmed` 不代表失败。 */
export type ReviewCopyCloseOutcome =
  | 'closed'
  | 'already_closed'
  | 'close_unconfirmed'
  | 'user_saved'
  | 'identity_lost'
  | 'ambiguous'
  | 'unsupported'
  | 'failed'
  | 'not_attempted'
  | 'extra_edits'
  | 'edits_unverified';

/**
 * 第三项判断：副本相对"我们写入后的最终快照"有没有变化。
 *
 * `unknown` 是**独立于** `changed` 的一档：读不到副本不等于副本没被改过，
 * 两者都不许静默丢弃，但要说清是哪一种，避免把"没读到"说成"确认被改"。
 */
export type ReviewCopyContentVerdict = 'unchanged' | 'changed' | 'unknown';

/** 业务结果。决定"能不能自动丢弃副本"的第一道闸。 */
export type ReviewCopyBusinessOutcome = 'verified' | 'write_failed' | 'partial_or_unknown' | 'discarded';

/**
 * 状态的**强调级别**，纯展示用。
 *
 * 它只决定"这条次要信息该多显眼"，**不**表达成败：`attention` 不等于失败，
 * `success` 也不等于业务成功（副本关掉了 ≠ 原表写对了）。
 */
export type ReviewCopyStatusTone = 'success' | 'attention' | 'neutral';

export type TempFileCleanupOutcome = 'none' | 'deleted' | 'in_use' | 'pending' | 'out_of_scope';

export type ReviewCopyCleanupState = {
  copy: {
    outcome: ReviewCopyCloseOutcome;
    /** 业务前缀 + 副本状态连成一句（对话/一条龙陈述面用）。 */
    headline: string;
    /**
     * **只讲副本自身**的状态，不含业务结论。
     *
     * 审阅卡要「业务结果」与「副本关闭状态」**独立显示**，所以这里单独留一份；
     * 两个面同源，不会出现"卡片说关了、对话说没关"。
     */
    statusText: string;
    detail: string;
    /** 供界面决定次要状态的强调级别；不改写任何结论。 */
    tone: ReviewCopyStatusTone;
    /** 允许用户再点一次「重试关闭」。 */
    retryable: boolean;
    /** 需要用户自己决定（是否丢弃/如何处理已另存的文件）。 */
    needsUserChoice: boolean;
  };
  file: {
    outcome: TempFileCleanupOutcome;
    headline: string;
    detail: string;
    retryable: boolean;
    /** 相对根目录的展示名（不暴露完整路径）。 */
    label: string | null;
  };
};

/**
 * 副本收尾的**状态表**。
 *
 * 每一条都只有 `statusText`（副本自身的状态），业务结论由 `BUSINESS_PREFIX` 前缀拼接。
 * 这样拆开是为了让审阅卡能把"原表怎么样了"和"副本怎么样了"**分两处独立显示**，
 * 而对话面仍能拿到合成后的一整句 —— 两个面同源，不会各写一套。
 *
 * 文案里**不许出现 Markdown 记号**（`**粗体**` 之类）：这些串直接进 DOM 文本节点，
 * 星号会原样显示给用户。
 */
type ReviewCopyStatusEntry = {
  /** 只讲副本自身。 */
  statusText: string;
  detail: string;
  tone: ReviewCopyStatusTone;
  retryable: boolean;
  needsUserChoice: boolean;
};

const COPY_HEADLINES: Record<ReviewCopyCloseOutcome, ReviewCopyStatusEntry> = {
  closed: {
    statusText: '审阅副本已关闭。',
    detail: '关闭的是本次流程创建的那一份副本；原表与原表所在窗口未被触碰。',
    tone: 'success',
    retryable: false,
    needsUserChoice: false
  },
  already_closed: {
    statusText: '审阅副本此前已由你关闭，本次无需再处理。',
    detail: '重复清理是幂等的：找不到本次副本即视为已收尾，不报错误。',
    tone: 'success',
    retryable: false,
    needsUserChoice: false
  },
  close_unconfirmed: {
    statusText: '关闭尚未确认，请到副本窗口前确认。',
    detail: '关闭请求已经发出，但窗口仍然存在。加载项没有对话框证据，所以不推断原因：可能是宿主在询问是否保存，也可能是别的东西挡住了关闭、或关闭本身较慢。请切到该窗口确认后再决定；副本内容不会写回原表。',
    tone: 'attention',
    retryable: true,
    needsUserChoice: true
  },
  extra_edits: {
    statusText: '副本里还有未经确认的人工修改，未自动关闭。',
    detail: '副本中留有生成对照之后新增的修改（与你无关的改动也可能是别的程序写进去的）。为避免静默丢弃，加载项没有关闭副本；这些修改也不会写回原表。请核对后手动处理。',
    tone: 'attention',
    retryable: true,
    needsUserChoice: true
  },
  edits_unverified: {
    statusText: '无法确认副本是否被额外修改，未自动关闭。',
    detail: '宿主没能给出副本的当前读数，因此无法证明"没有额外人工修改"。宁可留给你处理：加载项不关闭副本、不丢弃任何内容。请自行确认后关闭。',
    tone: 'attention',
    retryable: true,
    needsUserChoice: true
  },
  user_saved: {
    statusText: '审阅副本已被另存为文件，未自动关闭。',
    detail: '你自己保存/另存的文件不在本流程的所有权内：不会关闭、不会删除，请自行决定如何处理。',
    tone: 'neutral',
    retryable: false,
    needsUserChoice: true
  },
  identity_lost: {
    statusText: '无法确认副本身份，未自动关闭。',
    detail: '副本可能已被替换、改名或由别的窗口占用。为避免关掉你的其它文档，这里只提示，不动手。',
    tone: 'attention',
    retryable: true,
    needsUserChoice: true
  },
  ambiguous: {
    statusText: '匹配到多份候选副本，未自动关闭。',
    detail: '有不止一个窗口带着本次会话的副本标记。请手动确认后再关闭，避免误关无关工作簿。',
    tone: 'neutral',
    retryable: false,
    needsUserChoice: true
  },
  unsupported: {
    statusText: '本宿主不支持自动关闭副本，请手动关闭。',
    detail: '当前宿主没有可用的关闭接口（或该接口被宿主拒绝）。请手动关闭副本，并自行决定是否保留其中的内容。',
    tone: 'neutral',
    retryable: false,
    needsUserChoice: true
  },
  failed: {
    statusText: '副本未能自动关闭，请手动关闭。',
    detail: '关闭动作本身失败了。这不是业务失败：原表的写入与验证结果由执行结果决定。',
    tone: 'attention',
    retryable: true,
    needsUserChoice: true
  },
  not_attempted: {
    statusText: '本次未尝试关闭副本。',
    detail: '未发起关闭请求。请手动关闭副本，并自行决定是否保留其中的内容。',
    tone: 'neutral',
    retryable: true,
    needsUserChoice: true
  }
};

/**
 * 业务结论前缀。与 `COPY_HEADLINES` 拼接成完整的收尾陈述。
 *
 * 只说**原表**发生了什么，一个字都不提副本 —— 副本的事由 `statusText` 负责。
 */
const BUSINESS_PREFIX: Record<'verified' | 'discarded', string> = {
  verified: '已写回并验证；',
  discarded: '已放弃变更，原表未修改；'
};

const RETAIN_MESSAGES: Record<'write_failed' | 'partial_or_unknown', { headline: string; statusText: string; detail: string }> = {
  write_failed: {
    headline: '写回未完成，已保留审阅副本供你核对。',
    statusText: '已保留审阅副本供你核对，未自动关闭。',
    detail: '写回失败时不自动关闭、不删除任何东西：排查材料必须留在现场。按执行结果与恢复点核对原表后，再决定是否手动关闭副本。'
  },
  partial_or_unknown: {
    headline: '写回结果不确定，已保留审阅副本供你核对。',
    statusText: '已保留审阅副本供你核对，未自动关闭。',
    detail: '部分完成或结果无法确认时不会自动丢弃副本，也不会把原表结果说成已确认。请先核对原表与实际结果，再处理副本。'
  }
};

/** 关闭结果里"允许继续清理磁盘文件"的集合。 */
const CLOSE_ALLOWS_FILE_CLEANUP: ReadonlySet<ReviewCopyCloseOutcome> = new Set(['closed', 'already_closed', 'close_unconfirmed']);

/**
 * **第三项判断**的判定入口：副本相对"我们写入后的最终快照"有没有变化。
 *
 * 为什么必须单独一项：副本身份可信（①）、没被另存/替换（②）都成立，**也推不出**
 * ③"副本里没有额外人工修改" —— 用户在并排审阅期间完全可以往副本里敲字。
 * 前两项是"这是不是我们的副本"，第三项是"这份副本还是不是我们看到的样子"。
 *
 * 只有适配器**逐格读回并核对过**才允许传 `unchanged`；读不到一律 `unknown`，
 * 不许把"没读到"当成"应该没改"。
 *
 * @returns `null` 表示可以安全丢弃（可以调用关闭）；否则返回必须交回用户决定的关闭结果。
 */
export function describeCopyDiscardSafety(verdict: ReviewCopyContentVerdict): ReviewCopyCloseOutcome | null {
  if (verdict === 'unchanged') return null;
  return verdict === 'changed' ? 'extra_edits' : 'edits_unverified';
}

export function reviewCopyCleanupState(input: {
  business: ReviewCopyBusinessOutcome;
  copy: ReviewCopyCloseOutcome;
  file?: TempFileCleanupOutcome;
  /** 本次登记的临时文件（规范绝对路径）；没有磁盘文件时为 null/undefined。 */
  tempFilePath?: string | null;
  tempRoot?: string;
}): ReviewCopyCleanupState {
  const tempFilePath = input.tempFilePath ?? null;
  const fileOutcome: TempFileCleanupOutcome = input.file ?? (tempFilePath === null ? 'none' : 'pending');
  const file = tempFileState(fileOutcome, tempFilePath, input.tempRoot);
  // 只有副本已确认收尾（含"用户自己关掉了"）才允许声称磁盘文件已清理；否则如实说"仍保留"。
  const fileFor = (outcome: ReviewCopyCloseOutcome) => CLOSE_ALLOWS_FILE_CLEANUP.has(outcome)
    ? file
    : { ...file, outcome: fileOutcome === 'deleted' ? 'out_of_scope' as const : fileOutcome };
  if (input.business === 'write_failed' || input.business === 'partial_or_unknown') {
    const message = RETAIN_MESSAGES[input.business];
    return {
      copy: { outcome: 'not_attempted', headline: message.headline, statusText: message.statusText,
        detail: message.detail, tone: 'attention', retryable: false, needsUserChoice: true },
      file: fileFor('not_attempted')
    };
  }
  // 未知/未取到的关闭结果一律按 `failed` 处理：宁可让用户手动关闭，也绝不因为
  // 适配器回了个不认识的值（或什么都没回）就报"已关闭"。
  const template: ReviewCopyStatusEntry | undefined = COPY_HEADLINES[input.copy];
  const outcome: ReviewCopyCloseOutcome = template ? input.copy : 'failed';
  const resolved = template ?? COPY_HEADLINES.failed;
  const headline = `${BUSINESS_PREFIX[input.business === 'verified' ? 'verified' : 'discarded']}${resolved.statusText}`;
  return {
    copy: { outcome, headline, statusText: resolved.statusText, detail: resolved.detail,
      tone: resolved.tone, retryable: resolved.retryable, needsUserChoice: resolved.needsUserChoice },
    file: fileFor(outcome)
  };
}

function tempFileState(outcome: TempFileCleanupOutcome, path: string | null, root?: string): ReviewCopyCleanupState['file'] {
  const label = path ? path.split(/[\\/]/).filter(Boolean).slice(-1)[0] ?? null : null;
  switch (outcome) {
    case 'none':
      // 未保存工作簿**不是**磁盘文件：这里必须说清"没有文件"，不能说"已清理"。
      return { outcome, headline: '本次没有需要清理的磁盘文件。', label: null, retryable: false,
        detail: '审阅副本是宿主内存中的未保存工作簿（由宿主从内存内容创建），不是磁盘文件；只有加载项登记过的专用临时文件才会被删除。' };
    case 'deleted':
      return { outcome, headline: `已清理本次登记的临时文件${label ? `（${label}）` : ''}。`, label, retryable: false,
        detail: '只删除本次会话登记、且规范化路径位于专用临时目录内的文件。' };
    case 'in_use':
      return { outcome, headline: `临时文件正被占用，已保留待清理${label ? `（${label}）` : ''}。`, label, retryable: true,
        detail: '文件被占用时不会强杀宿主进程、也不会强行删除；释放后可重试清理。' };
    case 'pending':
      return { outcome, headline: '临时文件尚未清理，已保留待处理。', label, retryable: true,
        detail: '副本尚未可靠结束，磁盘清理推迟；不清算未确认归属的文件。' };
    case 'out_of_scope':
      return { outcome, headline: '临时文件路径不在专用临时目录内，已拒绝清理。', label, retryable: false,
        detail: root ? `只允许清理 ${root} 内的本次登记文件；越界路径一律拒绝。` : '只允许清理专用临时目录内的本次登记文件；越界路径一律拒绝。' };
  }
}

/** 各宿主"实际创建了什么"的权威说法（交付与 UI 共用，避免两边各写一套）。 */
export function reviewCopyCreationSummary(kind: ReviewCopyKind): { created: string; canClose: string; hasDiskFile: boolean } {
  if (kind === 'excel-native') return {
    created: '通过 Excel `createWorkbook` 新建的未保存工作簿窗口（只含本次授权表与计划结果）。',
    canClose: 'Office.js 的 `Window.close()`；调用前先逐格回读副本，确认它仍是"本次写入后的样子"。对不上或读不到就只提示、不关闭（宿主自行决定是否询问保存）。',
    hasDiskFile: false
  };
  if (kind === 'excel-compat') return {
    created: '通过 Excel `createWorkbook(base64)` 新建的未保存工作簿（本机生成的 OOXML，含公式与变化标记）。',
    canClose: 'Office.js 在兼容宿主没有窗口句柄；改用本地桥对身份已核实的该副本窗口发送 WM_CLOSE（宿主自行决定是否询问保存）。兼容副本无法回读，所以这里从不宣称"已确认没有额外修改"——关不掉或宿主提问一律交回用户。',
    hasDiskFile: false
  };
  return {
    created: 'WPS 里由 `Worksheet.Copy()` 生成的未保存单表副本工作簿。',
    canClose: 'WPS 工作簿的 `Close(false)`，只在副本逐格回读仍与本次审阅的最终快照一致时调用；发现额外修改（`extra_edits`）或读不到读数（`edits_unverified`）都不关闭，交回用户决定。',
    hasDiskFile: false
  };
}

/**
 * 临时文件路径校验：只接受**本机盘符绝对路径**且严格位于 `root` 之内。
 *
 * 拒绝：空值、相对路径、UNC（`\\server\share`）、带 `..` 的路径、根之外、
 * 以及任何不可归一化的输入（控制字符、把根本身当文件）。
 * 不解析符号链接：宿主只登记自己创建的文件，绝不跟随用户提供的路径。
 *
 * 这是"磁盘清理拒绝越出专用临时目录"的**唯一判据**：任何一条不满足都不许删。
 */
export function validateTempFilePath(root: string, candidate: string): { ok: boolean; normalized: string | null; reason: string } {
  const normalize = (value: string) => value.replace(/\\/g, '/').replace(/\/+$/, '');
  if (typeof candidate !== 'string' || candidate.trim() === '') return { ok: false, normalized: null, reason: 'empty' };
  if (/[\u0000-\u001f]/.test(candidate)) return { ok: false, normalized: null, reason: 'control-characters' };
  const path = normalize(candidate.trim());
  const base = normalize(root);
  // 只认本机盘符绝对路径：UNC 与相对路径都不是"专用临时目录里的登记文件"。
  const localAbsolute = /^[a-z]:\//i;
  if (!localAbsolute.test(path)) return { ok: false, normalized: null, reason: 'not-absolute-local-path' };
  if (!localAbsolute.test(base)) return { ok: false, normalized: null, reason: 'invalid-root' };
  if (path.split('/').some(segment => segment === '..')) return { ok: false, normalized: null, reason: 'traversal' };
  if (path.length <= base.length) return { ok: false, normalized: null, reason: 'outside-root' };
  const within = path.toLowerCase().startsWith(`${base.toLowerCase()}/`);
  if (!within) return { ok: false, normalized: null, reason: 'outside-root' };
  return { ok: true, normalized: path, reason: 'ok' };
}

/**
 * 把"这次登记了哪个临时文件"折成清理结果。**不碰磁盘**。
 *
 * 全产品当前**没有**任何磁盘临时文件（副本一律是宿主内存里的未保存工作簿），
 * 所以 `tempFilePath === null` 时必须给出 `none`，而不是编造"已删除临时文件"。
 * 一旦将来真的登记了文件，路径必须先过 `validateTempFilePath`；越界一律
 * `out_of_scope`，被占用一律 `in_use`（保留待清理，不强杀宿主）。
 */
export function planTempFileCleanup(input: {
  tempFilePath: string | null;
  tempRoot: string;
  registered?: boolean;
  locked?: boolean;
}): TempFileCleanupOutcome {
  if (input.tempFilePath === null) return 'none';
  if (!input.registered) return 'out_of_scope';
  if (!validateTempFilePath(input.tempRoot, input.tempFilePath).ok) return 'out_of_scope';
  return input.locked ? 'in_use' : 'deleted';
}
