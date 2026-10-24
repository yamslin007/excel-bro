/**
 * 模型选择的唯一规则。
 *
 * Excel 与 WPS 两端共用同一份实现，避免出现“界面显示 A、请求发送 B”的
 * 不一致。所有函数都是纯函数，便于行为测试。
 */

export interface SelectableModel {
  id: string;
  available: boolean;
  agentAvailable?: boolean;
  agentUnavailableReason?: string | null;
  formulaAvailable?: boolean;
}

/** Older catalogues lack purpose metadata; the server remains the final guard. */
export function agentModelAvailable(model: SelectableModel | undefined): boolean {
  return Boolean(model?.available && model.agentAvailable !== false);
}

export function formulaModelAvailable(model: SelectableModel): boolean {
  return model.available && model.formulaAvailable !== false;
}

export function agentModelUnavailableReason(model: SelectableModel): string {
  return model.agentUnavailableReason || "此连接暂不可用于 SDK Agent，请检查连接配置。";
}

export function canonicalModelId(id: string): string {
  const prefix = "claude-agent-sdk:";
  return id.startsWith(prefix) ? `connection:${id.slice(prefix.length)}` : id;
}

export function keepUnavailableSelection(models: SelectableModel[], id: string): boolean {
  return models.some((model) => model.id === canonicalModelId(id) && model.agentAvailable === false);
}

/** 同一模型 ID 的多条连接按保存顺序编号；编号只用于展示，选择仍使用唯一连接 ID。 */
export function connectionModelSuffix(
  connections: readonly { id: string; modelId: string }[],
  connectionId: string
): string {
  const target = connections.find((connection) => connection.id === connectionId);
  if (!target) return "";
  const key = target.modelId.trim().toLocaleLowerCase();
  const matching = connections.filter((connection) =>
    connection.modelId.trim().toLocaleLowerCase() === key
  );
  return matching.length > 1
    ? ` (${matching.findIndex((connection) => connection.id === connectionId) + 1})`
    : "";
}

/**
 * 旧版本保存在本地的虚拟模型 ID（“基础模式”）。它不是一个真实连接，
 * 目录加载成功后必须被迁移掉；保留常量只用于识别和清理旧记录。
 */
export const LEGACY_LOCAL_MODEL_ID = "local";

/** 没有任何可用模型时使用的显式空选择状态。 */
export const NO_MODEL_SELECTED = "";

/**
 * 已选连接的 Agent 规划统一使用 Claude Agent SDK。
 *
 * SDK 适配器只拿得到表结构（`planningReadPolicy: "none"`），需要数据时必须由
 * 模型经本机工具桥调 `query_table` 去读真实宿主。因此它的判据是**两件事**的
 * 唯一来源，不能各写一份：
 *
 * - 规划读策略是否收窄到 `"none"`；
 * - 本地意图路径是否要**放弃预置计算**（预置的 `dataResults` 会被服务端的
 *   "不上传数据行"守卫拒绝，新工具桥就永远拿不到执行权）。
 */
export function isSdkPlanningModel(modelId: string | null | undefined): boolean {
  // Model IDs identify credentials, not a second Agent implementation. Endpoint
  // compatibility is validated by the server; it cannot fall back to old loops.
  return isRealModelId(modelId);
}

/** The SDK adapter currently supports header-only planning, including free-text continuation. */
export function planningReadPolicyForModel(
  modelId: string | null | undefined, requested: "ifNeeded" | "none"
): "ifNeeded" | "none" {
  return isSdkPlanningModel(modelId) ? "none" : requested;
}

/**
 * 这轮规划是否**必须**有"流式工具投递"能力。
 *
 * SDK 适配器拿不到数据行（`planningReadPolicy: "none"`），模型要取数只能经本机
 * 工具桥下发 `query_table`；而工具请求**只能**经 SSE 进度流投递到面板。一旦传输
 * 降级成非流式（旧后端没有 `/api/turn/stream`、或代理缓冲 SSE），服务端会为一次
 * 永远送不到本机的工具请求干等到 `agentToolBridge.waitTimeoutSeconds` 超时 ——
 * 那正是"进入无人回执的等待"。
 *
 * 所以这类请求在**退回非流式之前**就要被明确拒绝，而不是静默降级。
 * 判据与 `isSdkPlanningModel` 同源：适配器一旦能走别的取数通道，这里跟着一起改。
 */
export function requiresStreamingToolDelivery(
  modelId: string | null | undefined
): boolean {
  return isSdkPlanningModel(modelId);
}

export function isRealModelId(modelId: string | null | undefined): boolean {
  const value = (modelId ?? "").trim();
  return value !== "" && value !== LEGACY_LOCAL_MODEL_ID;
}

/**
 * 按固定优先级确定新任务使用的模型：
 *
 * 1. 用户本次明确指定且在目录中有效的模型
 * 2. 本机保存且仍有效的真实模型 ID
 * 3. 服务端声明且仍有效的默认模型 ID
 * 4. 目录中的首个可用真实模型
 * 5. 无可用模型时返回空选择（未配置状态）
 *
 * 返回值可能为空串，表示“没有可用模型”。调用方必须处理该状态，
 * 不得回退到基础模式。
 */
export function chooseAvailableModel(
  models: SelectableModel[],
  currentModelId: string,
  defaultModelId: string | null | undefined,
  preferredModelId?: string
): string {
  const usable = models.filter((model) => agentModelAvailable(model) && isRealModelId(model.id));
  currentModelId = canonicalModelId(currentModelId);
  preferredModelId = preferredModelId ? canonicalModelId(preferredModelId) : undefined;
  if (usable.length === 0) return NO_MODEL_SELECTED;
  const has = (candidate: string | null | undefined) =>
    typeof candidate === "string" &&
    isRealModelId(candidate) &&
    usable.some((model) => model.id === candidate);
  if (has(preferredModelId)) return preferredModelId as string;
  if (has(currentModelId)) return currentModelId;
  // An existing incompatible choice requires a user decision, never a new provider.
  if (keepUnavailableSelection(models, currentModelId)) return NO_MODEL_SELECTED;
  if (has(defaultModelId)) return defaultModelId as string;
  return usable[0].id;
}

/**
 * 判断本地保存的模型选择是否需要写回。
 *
 * 目录加载成功后发现旧记录是 `local`，或指向已删除的连接时，必须迁移为
 * 当前有效选择；目录加载**失败**时不调用本函数，以免根据空目录删除
 * 仍然有效的选择。
 */
export function shouldPersistSelection(
  storedModelId: string,
  resolvedModelId: string
): boolean {
  return storedModelId !== resolvedModelId;
}

/**
 * 计算“当前应展示的模型 ID”。
 *
 * 展示值必须与请求实际发送的 modelId 完全一致：因此这里只认解析后的
 * 有效选择，解析为空就返回空（未配置），绝不退回目录首项或 `local`。
 */
export function resolveDisplayedModelId(
  models: SelectableModel[],
  selectedModelId: string,
  defaultModelId: string | null | undefined
): string {
  return chooseAvailableModel(models, selectedModelId, defaultModelId);
}
