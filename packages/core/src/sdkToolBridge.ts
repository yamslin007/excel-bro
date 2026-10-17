/**
 * Host-independent task tool bridge state machine.
 *
 * The SDK runs in the Python backend; the local host adapter (Office.js / WPS)
 * executes ``query_table`` on the *open* workbook and posts a receipt back. This
 * module is the client-side half that both hosts share: it turns the opaque
 * progress pointer into a poll, tracks which calls are still awaiting a receipt
 * and decides — before hitting the network — whether a receipt would be
 * accepted, a duplicate, a conflict or simply late.
 *
 * It is deliberately pure: no React, no host API, no fetch. The two add-in
 * panels own the transport and the deterministic query execution; this module
 * owns the semantics so they cannot drift apart.
 */

export type BridgeHostId = "excel-office" | "wps-et" | "folder";
export type BridgeToolName = "query_table" | "prepare_python_data";

export interface SdkBridgeBinding {
  taskId: string;
  hostId: BridgeHostId;
  workbookName: string;
  /**
   * 稳定工作簿身份（``url:`` / ``wps-path:`` / ``name:`` 命名空间）。
   * 只比名字无法隔离两个同名工作簿；服务端按此判定归属，回执也要原样回带。
   */
  workbookIdentity: string;
  sourceFingerprint: string | null;
  /**
   * **绑定时的授权范围快照**（本轮用户确认过的工作表）。
   *
   * 与 ``sourceFingerprint`` 同一个道理：它记录的是"这一轮是在什么授权下开始的"。
   * 执行前复核要拿它跟**当前**授权比 —— 只看当前授权是否非空是不够的，授权
   * **收窄**（Sheet2 被撤）必须拒读，否则会拿少一张表的范围继续算，静默给出
   * 与用户授权不符的答案。
   */
  authorizedSheetNames: string[];
}

export interface BridgeToolPointer {
  operationId: string;
  tool: BridgeToolName;
}

export interface PendingBridgeCall {
  operationId: string;
  tool: BridgeToolName;
  arguments: unknown;
  requestFingerprint: string;
  expiresAt: number;
}

export interface BridgePollResponse {
  calls: PendingBridgeCall[];
  generation: number;
}

/** 回执的受理结果：accepted / duplicate / rejected（被拒时带固定错误码）。 */
export interface BridgeReceiptResponse {
  status: "accepted" | "duplicate" | "rejected";
  code?: string | null;
}

export type BridgeReceiptDecision = "accepted" | "duplicate" | "conflict" | "late";

export interface SdkBridgeState {
  binding: SdkBridgeBinding | null;
  generation: number;
  cancelled: boolean;
  /** A pointer arrived; the host should poll the bridge. */
  pollRequested: boolean;
  /** Calls handed out by the server that have not yet been delivered. */
  pending: PendingBridgeCall[];
  /** operationId → request fingerprint already delivered. */
  delivered: Record<string, string>;
}

const OPERATION_ID = /^op_[0-9a-f]{32}$/;
const FINGERPRINT = /^[0-9a-f]{64}$/;

/**
 * 桥承载的工具名集合（与 ``BridgeToolName`` 同源）。
 *
 * 指针解析与轮询校验必须认**同一集合**：指针放行的工具、轮询却拒收，
 * 服务端的调用就会被判无效，而调用方 ``drain`` 会把异常吞掉 —— 症状是
 * "服务端一个回执都收不到、只能干等超时"，而不是任何看得见的报错。
 */
const BRIDGE_TOOL_NAMES: readonly BridgeToolName[] = ["query_table", "prepare_python_data"];

export function isBridgeToolName(value: unknown): value is BridgeToolName {
  return BRIDGE_TOOL_NAMES.includes(value as BridgeToolName);
}

export function createSdkBridgeState(
  binding: SdkBridgeBinding | null = null
): SdkBridgeState {
  return {
    binding,
    generation: 0,
    cancelled: false,
    pollRequested: false,
    pending: [],
    delivered: {}
  };
}

/** Strict parse of the opaque pointer carried by a progress event. */
export function bridgeToolPointer(value: unknown): BridgeToolPointer | null {
  if (value === null || typeof value !== "object") {
    return null;
  }
  const pointer = value as Partial<BridgeToolPointer>;
  if (
    typeof pointer.operationId !== "string" ||
    !OPERATION_ID.test(pointer.operationId) ||
    !isBridgeToolName(pointer.tool)
  ) {
    return null;
  }
  return { operationId: pointer.operationId, tool: pointer.tool };
}

/**
 * Apply a progress event. A ``bridgeTool`` pointer requests a poll; a
 * ``cancelled`` status on a bridge step marks the session cancelled so the host
 * stops polling and rejects any late receipt locally.
 */
export function applyBridgeProgressEvent(
  state: SdkBridgeState,
  event: { status?: string; stepId?: string; bridgeTool?: unknown } | null | undefined
): SdkBridgeState {
  if (!event) {
    return state;
  }
  const pointer = bridgeToolPointer(event.bridgeTool);
  const isBridgeStep =
    typeof event.stepId === "string" && event.stepId.startsWith("bridge.");
  if (pointer === null && !isBridgeStep) {
    return state;
  }
  const next: SdkBridgeState = {
    ...state,
    pending: [...state.pending],
    delivered: { ...state.delivered }
  };
  if (event.status === "cancelled" || event.status === "failed") {
    next.cancelled = true;
    next.pending = [];
  }
  if (pointer !== null) {
    next.pollRequested = true;
  }
  return next;
}

export function assertBridgePollResponse(value: unknown): asserts value is BridgePollResponse {
  if (value === null || typeof value !== "object") {
    throw new Error("桥轮询响应不是对象");
  }
  const response = value as Partial<BridgePollResponse>;
  if (
    typeof response.generation !== "number" ||
    !Number.isInteger(response.generation) ||
    !Array.isArray(response.calls)
  ) {
    throw new Error("桥轮询响应缺少 generation/calls");
  }
  for (const call of response.calls) {
    if (
      !call ||
      typeof call.operationId !== "string" ||
      !OPERATION_ID.test(call.operationId) ||
      !isBridgeToolName(call.tool) ||
      typeof call.requestFingerprint !== "string" ||
      !FINGERPRINT.test(call.requestFingerprint)
    ) {
      throw new Error("桥轮询返回了无效的待执行调用");
    }
  }
}

/** Merge a poll response: keep only not-yet-delivered calls of this session. */
export function applyBridgePollResponse(
  state: SdkBridgeState,
  response: BridgePollResponse
): SdkBridgeState {
  assertBridgePollResponse(response);
  const pending = response.calls.filter(
    (call) => state.delivered[call.operationId] === undefined
  );
  return {
    ...state,
    generation: response.generation,
    pending,
    pollRequested: false,
    delivered: { ...state.delivered }
  };
}

export function pendingBridgeCall(
  state: SdkBridgeState,
  operationId: string
): PendingBridgeCall | null {
  return state.pending.find((call) => call.operationId === operationId) ?? null;
}

/**
 * Decide, before sending, what the server will answer. Mirrors the backend so
 * the host never re-runs a query it already delivered and never posts a receipt
 * for an operation the server has already retired.
 */
export function decideBridgeReceipt(
  state: SdkBridgeState,
  operationId: string,
  requestFingerprint: string
): BridgeReceiptDecision {
  if (state.cancelled) {
    return "late";
  }
  const already = state.delivered[operationId];
  if (already !== undefined) {
    return already === requestFingerprint ? "duplicate" : "conflict";
  }
  const call = pendingBridgeCall(state, operationId);
  if (call === null) {
    return "late";
  }
  return call.requestFingerprint === requestFingerprint ? "accepted" : "conflict";
}

export function markBridgeDelivered(
  state: SdkBridgeState,
  operationId: string,
  requestFingerprint: string
): SdkBridgeState {
  return {
    ...state,
    pending: state.pending.filter((call) => call.operationId !== operationId),
    delivered: { ...state.delivered, [operationId]: requestFingerprint }
  };
}

export function cancelBridge(state: SdkBridgeState): SdkBridgeState {
  return { ...state, cancelled: true, pending: [], pollRequested: false };
}

/** Whether the host should issue a poll right now. */
export function shouldPollBridge(state: SdkBridgeState): boolean {
  return state.binding !== null && !state.cancelled && state.pollRequested;
}

/**
 * 工作簿身份的单一来源。宿主没有稳定身份时退回 ``name:<名字>``，与服务端
 * ``local_results.fallback_workbook_identity`` 保持一致 —— 两端不能各算一套。
 */
export function bridgeWorkbookIdentity(
  name: string,
  identity?: string | null
): string {
  const trimmed = typeof identity === "string" ? identity.trim() : "";
  return trimmed ? trimmed : `name:${name}`;
}

// ---------------------------------------------------------------------------
// 宿主适配器编排：轮询 → 复核授权 → 真实读取+确定性计算 → 回执 → SDK 继续
// ---------------------------------------------------------------------------

export type BridgeCallOutcome =
  | { kind: "result"; result: unknown }
  | { kind: "error"; errorCode: string };

/** 复核当前授权的结果：拒绝时给出固定错误码，不把宿主细节透出去。 */
export type BridgeAuthorizationVerdict =
  | { ok: true }
  | { ok: false; errorCode: string };

/** 执行前重读的真实宿主现场（两个宿主同形）。 */
export interface BridgeLiveAuthorization {
  workbookIdentity: string;
  sourceFingerprint: string | null;
  authorizedSheetNames: string[];
  activeWorksheet: string;
}

/**
 * 本次调用要读哪些工作表。
 *
 * ``scope:"active"`` 只读当前活动表；其余情况读的就是**绑定时的授权范围** ——
 * ``query_table`` 的参数里并不枚举工作表名，所以"会读到哪些表"等于这一轮用户
 * 确认过的集合。把它显式取出来，才谈得上跟当前授权做比较。
 */
export function bridgeRequestedSheetNames(
  binding: SdkBridgeBinding,
  call: PendingBridgeCall,
  live: { activeWorksheet: string }
): string[] {
  if (call.tool === "prepare_python_data") {
    // 数据登记明确枚举来源工作表：要读的正是 ``sources[].sheet``，不是整份授权。
    // 任何一张不在当前授权内都拒绝，否则会拿越权范围去读表。
    const args = (call.arguments ?? {}) as { sources?: unknown };
    if (!Array.isArray(args.sources)) return [];
    const names: string[] = [];
    for (const item of args.sources) {
      if (item && typeof item === "object" && typeof (item as { sheet?: unknown }).sheet === "string") {
        names.push((item as { sheet: string }).sheet);
      }
    }
    return names;
  }
  const args = (call.arguments ?? {}) as { scope?: unknown };
  return args.scope === "active"
    ? [live.activeWorksheet]
    : [...binding.authorizedSheetNames];
}

/**
 * 现场读不到时用的**固定**拒绝码（两端共用这一份字面量）。
 *
 * 语义是"这轮请求的授权**没能核实**，请重新确认数据范围" —— 与
 * `verifyBridgeAuthorization` 里"范围不被覆盖"的结论给的是同一个可执行的下一步。
 */
export const BRIDGE_UNVERIFIABLE_ERROR_CODE = "unauthorized_scope";

/**
 * 执行前的授权复核 —— 两个宿主必须用同一份判据。
 *
 * 三类变化一律拒绝：换簿（身份不同）、源数据变化（指纹不同）、**授权收窄**
 * （绑定时的范围不再被当前授权完全覆盖）。最后一条最容易被漏掉：只看"当前授权
 * 非空"会放行一个少了一张表、静默算出错答案的查询。
 */
export function verifyBridgeAuthorization(
  binding: SdkBridgeBinding,
  call: PendingBridgeCall,
  live: BridgeLiveAuthorization
): BridgeAuthorizationVerdict {
  if (live.workbookIdentity !== binding.workbookIdentity) {
    // 名字相同但身份不同 = 换了一个工作簿，绝不能拿旧授权去读。
    return { ok: false, errorCode: "source_changed" };
  }
  // 范围先判：收窄是"授权"问题（要重新勾选），不是"源变了"（要重新查询）。
  // 两边给的都是可执行的下一步，判反了用户会被指向错误的补救动作。
  const requested = bridgeRequestedSheetNames(binding, call, live);
  const allowed = new Set(live.authorizedSheetNames);
  if (requested.length === 0 || requested.some((name) => !allowed.has(name))) {
    return { ok: false, errorCode: "unauthorized_scope" };
  }
  if (
    binding.sourceFingerprint !== null &&
    live.sourceFingerprint !== null &&
    live.sourceFingerprint !== binding.sourceFingerprint
  ) {
    return { ok: false, errorCode: "source_changed" };
  }
  return { ok: true };
}

/**
 * 读一次**当前**现场再复核；读不到现场就 fail-closed 拒绝。
 *
 * 为什么必须在这里接住异常：复核抛出会被调用方 `drain` 吞掉（"桥失败不能拖垮
 * 整轮规划"），于是服务端**一个回执都收不到**，只能干等到
 * `agentToolBridge.waitTimeoutSeconds` —— 那正是"无人回执的等待"。
 *
 * 拿不到证据 ≠ 没变化：读不到当前授权/身份/来源时必须**明确拒绝**。
 * 判据与去重、错误码都在共享核心里，两个宿主不各写一份。
 */
export async function verifyBridgeAuthorizationOrRefuse(
  binding: SdkBridgeBinding,
  call: PendingBridgeCall,
  readLive: () => BridgeLiveAuthorization | Promise<BridgeLiveAuthorization>
): Promise<BridgeAuthorizationVerdict> {
  let live: BridgeLiveAuthorization;
  try {
    live = await readLive();
  } catch {
    return { ok: false, errorCode: BRIDGE_UNVERIFIABLE_ERROR_CODE };
  }
  return verifyBridgeAuthorization(binding, call, live);
}

export interface BridgeHostPorts {
  /** 拉取本次会话待执行的调用（本机回环）。 */
  poll(binding: SdkBridgeBinding): Promise<BridgePollResponse>;
  /**
   * **执行前的授权复核**：工作簿身份/来源指纹/工作表范围任一变化都必须拒绝，
   * 否则会拿旧授权去读新数据。返回 ``{ok:false}`` 时不再读宿主。
   *
   * 允许返回 Promise：真实宿主的"当前授权/身份/来源"往往要**重新读一次**
   * （Office.js / JSAPI 都是异步 API）。同步快照在这里是不够的 —— 它证明不了
   * 等待期间用户已经收窄或撤销了授权。
   */
  verify(
    binding: SdkBridgeBinding,
    call: PendingBridgeCall
  ): BridgeAuthorizationVerdict | Promise<BridgeAuthorizationVerdict>;
  /** 在真实宿主上读取并跑**现有确定性** ``query_table``。 */
  execute(binding: SdkBridgeBinding, call: PendingBridgeCall): Promise<BridgeCallOutcome>;
  /** 回执（本机回环）。 */
  receipt(
    binding: SdkBridgeBinding,
    call: PendingBridgeCall,
    outcome: BridgeCallOutcome
  ): Promise<BridgeReceiptResponse>;
}

export interface BridgeOutcomeEntry {
  operationId: string;
  decision: BridgeReceiptDecision;
  /** 服务端返回的 accepted / duplicate / rejected；本地短路时为 null。 */
  status: string | null;
  /** 回执被服务端拒收时给出的固定错误码。 */
  code: string | null;
}

export interface BridgeServiceResult {
  state: SdkBridgeState;
  polled: boolean;
  /** 真正执行了本地查询的次数。 */
  executed: number;
  outcomes: BridgeOutcomeEntry[];
}

/**
 * 服务一次桥：需要时轮询，然后逐个处理待执行调用。
 *
 * 单次调用内**不会**重复执行同一 ``operationId`` —— 已经交付过的调用会在本地
 * 短路，服务端仍会给 ``duplicate``，但那条路径不需要我们再读一次宿主。
 */
export async function serviceBridgeOnce(
  state: SdkBridgeState,
  ports: BridgeHostPorts
): Promise<BridgeServiceResult> {
  const binding = state.binding;
  if (binding === null || state.cancelled) {
    return { state, polled: false, executed: 0, outcomes: [] };
  }

  let next = state;
  let polled = false;
  if (shouldPollBridge(next)) {
    const response = await ports.poll(binding);
    next = applyBridgePollResponse(next, response);
    polled = true;
  }

  const outcomes: BridgeOutcomeEntry[] = [];
  let executed = 0;
  for (const call of [...next.pending]) {
    if (next.cancelled) break;
    const decision = decideBridgeReceipt(next, call.operationId, call.requestFingerprint);
    if (decision !== "accepted") {
      // 已交付 / 会话已取消：本地短路，不重复读宿主。同时把它移出待执行列表，
      // 否则一条残留的 pending 会让调用方一直空转。
      next = {
        ...next,
        pending: next.pending.filter((item) => item.operationId !== call.operationId)
      };
      outcomes.push({
        operationId: call.operationId,
        decision,
        status: null,
        code: null
      });
      continue;
    }
    const verdict = await ports.verify(binding, call);
    let outcome: BridgeCallOutcome;
    if (verdict.ok) {
      outcome = await ports.execute(binding, call);
      executed += 1;
    } else {
      outcome = { kind: "error", errorCode: verdict.errorCode };
    }
    const response = await ports.receipt(binding, call, outcome);
    outcomes.push({
      operationId: call.operationId,
      decision,
      status: response.status,
      code: response.code ?? null
    });
    // accepted/duplicate 才认为服务端已收下；rejected 留给下一轮指针重试。
    if (response.status !== "rejected") {
      next = markBridgeDelivered(next, call.operationId, call.requestFingerprint);
    } else {
      next = {
        ...next,
        pending: next.pending.filter((item) => item.operationId !== call.operationId)
      };
    }
  }
  return { state: next, polled, executed, outcomes };
}

/**
 * 反复服务直到没有新的指针（服务端一次只发一个指针，正常情况一轮就够）。
 * ``maxRounds`` 是防呆上限，避免服务端异常时客户端空转。
 */
export async function runBridgeUntilIdle(
  state: SdkBridgeState,
  ports: BridgeHostPorts,
  options: { maxRounds?: number } = {}
): Promise<BridgeServiceResult> {
  const maxRounds = options.maxRounds ?? 8;
  let current = state;
  let polled = false;
  let executed = 0;
  const outcomes: BridgeOutcomeEntry[] = [];
  for (let round = 0; round < maxRounds; round += 1) {
    if (current.cancelled || (current.pending.length === 0 && !shouldPollBridge(current))) {
      break;
    }
    const result = await serviceBridgeOnce(current, ports);
    current = result.state;
    polled = polled || result.polled;
    executed += result.executed;
    outcomes.push(...result.outcomes);
    if (!result.polled && result.outcomes.length === 0) break;
  }
  return { state: current, polled, executed, outcomes };
}
