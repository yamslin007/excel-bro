import {
  ApiRequestError,
  checkIntent,
  understandMessage,
  createAssistantResponse,
  isLocalServiceConnectionError,
  type ModelOption
} from "../../excel-addin/src/api";
import {
  continueWpsPlanning,
  getWpsPlanningTaskStatus,
  streamWpsPlanning,
  wpsPlanningResumeDecision,
  type WpsPlanningTaskStatus
} from "./wpsPlanningStream";
import { newPlanningRequestId } from "@excel-bro/core";
import { waitForPlanning } from "@excel-bro/core/planningReplay";
import type { PlanningContinuationEvidence } from "@excel-bro/core/planningContinuationEvidence";
import { assertKnownResultKind } from "./wpsPlanningTypes";
import capabilities from "../../../config/capabilities.json";
import {
  BlankPolicyNeededError,
  NeedsClarificationError,
  UnmappedCategoryError
} from "@excel-bro/core";
import type {
  AssistantResponse,
  AnalysisPlan,
  CellValue,
  ColumnMappingRule,
  DataToolRequest,
  DataToolResult,
  IntentCheckRequest,
  IntentCheckResponse,
  MessageUnderstandingResponse,
  IntentScopeContext,
  ImageAttachment,
  LocalResultCard,
  PlanRequest,
  PlanningPauseResult,
  ResultContext,
  SdkBridgeToolPointer,
  TaskProgressEvent,
  TaskProgressStatus,
  WorkbookSnapshot
} from "../../excel-addin/src/contracts";
import {
  executeQueryTableData,
  type SheetValues
} from "../../excel-addin/src/dataTools";
import { detectSheetFields } from "../../excel-addin/src/tableSchema";
import { isSdkPlanningModel } from "../../excel-addin/src/modelSelection";
import {
  buildWpsColumnMappingPlan,
  sliceMappingSource,
  type MappingSourceSheet
} from "./wpsColumnMappingPlan";
import type { WpsSourceSnapshot } from "./wpsPlanExecutor";
import { wpsReportedActionTypes } from "./wpsHostAdapter";
import { readWpsHostProbe, type WpsApplication } from "./wpsHost";
import {
  authorizeWpsSheets,
  listWpsWorksheets,
  readAuthorizedWpsSheets,
  readAuthorizedWpsSheetSchemas,
  wpsWorkbookBridgeIdentity,
  type WpsSheetReadResult
} from "./wpsWorkbookReader";

/**
 * 需要在写入前澄清的一轮结果。
 *
 * `columns` 非空 = **列映射澄清**，来自服务端下发的规则 `clarification`。
 * 两者共用同一张澄清卡，但归因与可选项不同 —— 卡片的选项只由 `reason`
 * 推导（`@excel-bro/core` 的 `columnMappingClarificationCard`），本端不猜测。
 *
 * `columns` 为空 = 既有的需求澄清（模型说不清要做什么，走 `response`）。
 */
export type WpsTurnClarification = {
  /** 列映射澄清：服务端已判定规则缺参数/缺类别，**未读源值、未计算**。 */
  columns?: {
    clarification: NonNullable<
      Extract<ColumnMappingRule, { clarification?: unknown }>["clarification"]
    >;
    scopeLabel: string;
    summary: string;
  };
  /** 需求澄清（模型层）。 */
  response?: Extract<IntentCheckResponse, { kind: "clarification" }>;
  /** 本地读源列后发现的未覆盖值；只允许留在面板内存，不得进历史或请求。 */
  localMapping?: {
    rule: ColumnMappingRule;
    context: { summary: string; warnings: string[]; assumptions: string[] };
    values: readonly CellValue[];
    reason: "unmapped" | "separator_or_mapping" | "blank";
    /**
     * `reason === "blank"` 时**真实读到**的空白格数量。
     *
     * 它来自一次真实读取。没有读证据就不许声称"源区域有空白"。
     */
    blankCount?: number;
    /** 已声明拆分后缺的是"项"，不是整格；不能暗用映射声明顺序决定优先级。 */
    allowExactMapping: boolean;
  };
};

export type WpsTurnResult =
  | {
      kind: "answer";
      message: string;
      resultContext?: ResultContext | null;
      provider: "model" | "local" | "system";
      outcome?: "completed" | "failed";
      /**
       * 服务端这次规划的身份。**断线后要先拿它查真实状态**，
       * 有有效检查点才继续 —— 不能靠"重新发一遍需求"冒充续跑。
       */
      turnId?: string | null;
      /**
       * SDK 本地结果卡指针：**只带引用与结构，没有任何值**。
       * 真实值由本机回环端点按 `resultRef` 取回，不进模型/历史/持久账本。
       */
      localResult?: LocalResultCard | null;
    }
  | {
      kind: "clarification";
      clarification: WpsTurnClarification;
      turnId?: string | null;
      /**
       * 结构扫描**真实读到的表头行**（每张已授权表一条，顺序与授权序一致）。
       *
       * 为什么要把它带出来：澄清卡需要一个"这份结构还是不是原来那份"的判据，
       * 而它必须落在**表头单元格的真实内容**上 —— 工作表名只是身份，表头是语义。
       * 之前面板拿 `listWpsWorksheets().sheets.map(([index, name]))` 冒充表头，
       * 于是"表头判据"退化成身份判据的副本，两条判据吃同一份数据、互相遮蔽。
       *
       * 只含表头行（`schemaOnly` 扫描天然只取第 1 行），数据行不进这里，
       * 因此它落到澄清卡上也不会把单元格值写进对话历史。
       */
      headerRows?: string[][];
    }
  | {
      kind: "plan";
      planOrigin?: "literal_range_edit" | null;
      plan: AnalysisPlan;
      turnId?: string | null;
      /**
       * 计划是否在本端算出（`local_plan_rule` 那一支）。
       *
       * 为 `true` 时**不得**送 `/api/plans/validate`：`writeValues` 的验收条件
       * 里带着完整的值矩阵（`models.py` 的 `RangeEqualsCriterion.expected`），
       * 上传等于把刚才刻意留在本地的源值送出去 —— 单射映射可逆。
       *
       * 跳过服务端校验在这条路径上不丢东西：WPS 的执行验证是自洽的（写入前后
       * 指纹比对 + 回读，见 `wpsPlanExecutor.ts`），全程不读 `acceptanceCriteria`；
       * 而授权与存在性由 `prepareWpsPlan` 对着活宿主查，比服务端拿名字列表比
       * 更严。区域也已在服务端下发规则时规范化过。
       */
      computedLocally?: boolean;
      sourceSnapshots?: WpsSourceSnapshot[];
    }
  | {
      /**
       * 规划**暂停**（CAP-02）。这不是失败，是一份结构化的、可继续的现场。
       *
       * 关键纪律：暂停卡上的按钮**不在这里决定**。本分支只把服务端给的
       * `pause` 载荷（含 `revision` / `lastSequence` / `canContinue` /
       * `pauseType`）原样端上来，由 `wpsPlanningResumeDecision` 从结构化状态
       * 推导 —— 绝不用助手文本或业务关键词猜"能不能继续"。
       */
      kind: "planning_pause";
      pause: PlanningPauseResult;
    };

export type WpsReadonlyTurnProgress = {
  stage: "reading" | "intent" | "calculating" | "answering";
  stepId?: string;
  title?: string;
  sheetName?: string;
  rowsRead?: number;
  totalRows?: number;
  toolName?: string;
  message: string;
  /**
   * 这一步**真实**的结构化状态。
   *
   * 以前这里没有这个字段，于是整条链上"完成"永远表达不出来：转换层只能说
   * "还在进行"，下游只能靠"下一步出现了"倒推上一步完成 —— 那是对用户撒谎。
   * 本地阶段（读表结构、算数）由 `runWpsTurn` 在真正做完之后补发 `completed`；
   * 服务端事件则由 `event.status` 原样带过来。
   */
  status?: TaskProgressStatus;
  /**
   * 服务端原始事件（只有服务端事件才有）。
   *
   * **整条带着，不逐字段 copy。** 曾经为了"只取要用的几个字段"手工挑，
   * 结果 `bridgeTool` 被漏掉，续跑时模型发的读表请求永远到不了本端、
   * 服务端一路等到超时（表现是"继续分析卡住"）。字段清单会随协议增长，
   * 挑字段的写法注定再漏一次；交给下游按 `TaskProgressEvent` 整体消费。
   */
  event?: TaskProgressEvent;
  /**
   * 本轮规划在服务端的**任务身份**（`turnId`）。
   *
   * 意图判断一返回就带上它，且**只上报一次**（`intent` 阶段）。
   * 存在的理由：断线恢复要拿它去查真实状态；不在这里透出来的话，
   * 面板手上根本没有身份，只能建议用户重发需求 —— 那正是 CAP-02 第 4 条
   * 要禁止的"重发即重新规划"。
   */
  turnId?: string;
  /** 客户端实际收到的事件位置；不等于服务端最新位置。 */
  sequence?: number;
  /**
   * SDK 工具桥指针：服务端的模型正在等本机执行一次本地查询。
   *
   * **只有操作 ID 与工具名**；请求参数留在本机易失桥里由宿主适配器拉取，
   * 不进模型提示、不进聊天历史、不进持久账本。原样透传给面板，
   * 由面板交给桥控制器去轮询/复核/读取/回执。
   */
  bridgeTool?: SdkBridgeToolPointer | null;
  pythonResult?: LocalResultCard | null;
};

/**
 * 把规划流的 `TaskProgressEvent` 折成面板消费的进度对象。
 *
 * **唯一实现**，首次规划与"继续分析"共用。续跑路径曾经各自拼了一份"看起来一样"
 * 的对象，只 copy 了 stage/turnId/sequence/stepId/title/message，**漏掉了
 * `bridgeTool`** —— 于是续跑时模型发出的读表请求永远到不了本端，服务端一路等到
 * 超时（表现为"继续分析卡住"，不是报错）。
 *
 * 现在改成**整条事件挂上去**（`event`），显示用的几个字段由它派生。这样协议每加
 * 一个字段，转换层不需要改，也不会再漏。
 */
export function wpsProgressFromTaskEvent(
  turnId: string,
  event: TaskProgressEvent,
  fallbackStepId = "planning-stream",
  fallbackMessage = "正在规划"
): WpsReadonlyTurnProgress {
  return {
    stage: "answering",
    turnId,
    sequence: event.sequence,
    stepId: event.stepId || fallbackStepId,
    title: event.title,
    // 状态**照抄事件**，不在这里改写成"进行中"。
    status: event.status,
    message: event.summary || event.title || fallbackMessage,
    bridgeTool: event.bridgeTool ?? null,
    pythonResult: event.pythonResult ?? null,
    event
  };
}

export type WpsReadonlyTurnInput = {
  application: WpsApplication;
  selectedSheetIndexes: readonly number[];
  prompt: string;
  taskContext?: IntentCheckRequest["taskContext"];
  modelId?: string | null;
  images?: ImageAttachment[];
  conversation?: IntentCheckRequest["conversation"];
  intentConfirmed?: boolean;
  clarificationRound?: number;
  /** Already classified by the panel; never contains workbook context. */
  messageUnderstanding?: MessageUnderstandingResponse;
  signal?: AbortSignal;
  /** Metadata-only UI guard, repeated before each source read after model awaits. */
  beforeRead?: () => Promise<void>;
  /** Await the host bridge after intent supplies the task id, before planning starts. */
  preparePlanningBridge?: (turnId: string) => Promise<boolean>;
  onProgress?: (progress: WpsReadonlyTurnProgress) => void;
};

type WpsReadonlyAgentDependencies = {
  understandMessage?: typeof understandMessage;
  checkIntent: typeof checkIntent;
  createAssistantResponse: typeof createAssistantResponse;
  streamPlanning: typeof streamWpsPlanning;
  getPlanningStatus: typeof getWpsPlanningTaskStatus;
  continuePlanning: typeof continueWpsPlanning;
  executeQuery: (request: DataToolRequest, sheets: SheetValues[]) => DataToolResult;
  readSheets: typeof readAuthorizedWpsSheets;
  readSchemas: typeof readAuthorizedWpsSheetSchemas;
};

const DEFAULT_DEPENDENCIES: WpsReadonlyAgentDependencies = {
  understandMessage,
  checkIntent,
  createAssistantResponse,
  streamPlanning: streamWpsPlanning,
  getPlanningStatus: getWpsPlanningTaskStatus,
  continuePlanning: continueWpsPlanning,
  executeQuery: executeQueryTableData,
  readSheets: readAuthorizedWpsSheets,
  readSchemas: readAuthorizedWpsSheetSchemas
};

/**
 * 只有"这条流式通道根本不存在"才允许回退到一次性 `/api/turn`。
 *
 * 三类判据，缺一不可地**窄**：
 *
 *  - `isLocalServiceConnectionError`：本地 AI 服务压根没起来（fetch 层报错）；
 *  - 404 / 405：旧服务端没有 `/api/turn/stream` 这个路由。
 *
 * **业务错误绝不在这里**。模型未配置、能力不匹配、任务不存在（400/409）
 * 都是"服务端收到了请求并且明确拒绝"，回退只会掩盖真实原因，还可能触发
 * 第二次模型调用与重复计费。因此 404 `PLANNING_TASK_NOT_FOUND` 这种
 * "端点存在、但这条任务不在"也必须原样抛给调用方，由它去查状态。
 */
function isStreamEndpointUnavailable(reason: unknown): boolean {
  if (isLocalServiceConnectionError(reason)) return true;
  return (
    reason instanceof ApiRequestError && (reason.status === 404 || reason.status === 405)
  );
}

function columnName(index: number): string {
  let value = Math.max(1, index);
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function usedRangeAddress(result: WpsSheetReadResult): string | null {
  if (result.rowCount <= 0 || result.columnCount <= 0) return null;
  const first = `${columnName(result.startColumn)}${result.startRow}`;
  const last = `${columnName(result.startColumn + result.columnCount - 1)}${result.startRow + result.rowCount - 1}`;
  return first === last ? first : `${first}:${last}`;
}

/**
 * 结构扫描**实际取到的表头行**（只取第 1 行）。
 *
 * `readAuthorizedWpsSheetSchemas` 的取值面已经被收窄到表头行，所以 `values[0]`
 * 就是真实表头；把它取出来交给澄清卡做指纹，才能让"表头判据"真的看表头内容，
 * 而不是去看工作表名（那是身份判据的事）。
 *
 * 之所以只取第 1 行、并且把 `null`/`undefined` 一律折成 `""`：
 * 指纹会被持久化进对话历史，必须稳定、可比较，且不含数据行。
 */
function sheetHeaderRows(results: WpsSheetReadResult[]): string[][] {
  return results.map((result) => {
    const header = result.values[0] ?? [];
    return header.map((cell) => (cell == null ? "" : String(cell)));
  });
}

/** 与 core/Python 的来源值键化一致；用户补的整格映射必须命中同一个键。 */
export function wpsMappingSourceKey(value: CellValue): string {
  if (value === true) return "True";
  if (value === false) return "False";
  return String(value);
}

type LocalMappingContext = { summary: string; warnings: string[]; assumptions: string[] };

function localMappingResult(
  rule: ColumnMappingRule,
  sourceSheet: MappingSourceSheet,
  context: LocalMappingContext,
  headerRows?: string[][]
): WpsTurnResult {
  try {
    const plan = buildWpsColumnMappingPlan(rule, sourceSheet, context);
    return {
      kind: "plan",
      computedLocally: true,
      sourceSnapshots: [{
        sheetName: sourceSheet.name,
        range: rule.sourceRange,
        values: sliceMappingSource(sourceSheet, rule.sourceRange)
      }],
      plan
    };
  } catch (error) {
    /*
     * 空白策略未定，且**这一轮真实读到了空白格**。
     *
     * 停在这是对的：它发生在读源值之后，不是"还没读就先问"。四个选项由
     * 界面给出，本端不替用户挑 —— 静默挑一个默认标签（曾经就是这么做的，
     * 空白一律写成「未分类」）会让用户在单元格层面看不出那是实现替他选的。
     */
    if (error instanceof BlankPolicyNeededError) {
      return {
        kind: "clarification",
        headerRows,
        clarification: {
          localMapping: {
            rule,
            context,
            // **不带任何来源值**：空白格没有内容可带，而"有几格空白"这个
            // 计数足够让用户判断规模。带值会违反"源值只留本端"。
            values: [],
            reason: "blank",
            blankCount: error.blankCount,
            allowExactMapping: false
          }
        }
      };
    }
    if (!(error instanceof UnmappedCategoryError) && !(error instanceof NeedsClarificationError)) throw error;
    return {
      kind: "clarification",
      headerRows,
      clarification: {
        localMapping: {
          rule,
          context,
          values: error instanceof UnmappedCategoryError ? error.values : error.cells,
          reason: error instanceof UnmappedCategoryError ? "unmapped" : "separator_or_mapping",
          allowExactMapping: error instanceof NeedsClarificationError || !rule.delimiters?.length
        }
      }
    };
  }
}

/** 用户在本机补全整格映射后重新读授权源表；不重发意图/规则或单元格值。 */
export async function resumeWpsLocalMapping(input: {
  application: WpsApplication;
  selectedSheetIndexes: readonly number[];
  rule: ColumnMappingRule;
  context: LocalMappingContext;
  additions: readonly { source: string; target: string }[];
  signal?: AbortSignal;
  onProgress?: WpsReadonlyTurnInput["onProgress"];
  beforeRead?: WpsReadonlyTurnInput["beforeRead"];
}, dependencies: WpsReadonlyAgentDependencies = DEFAULT_DEPENDENCIES): Promise<WpsTurnResult> {
  if (input.additions.some((item) => !item.source || !item.target.trim())) {
    throw new Error("请为每个未覆盖的来源值明确填写目标值。");
  }
  const existing = new Set(input.rule.mapping.map((item) => item.source));
  if (input.additions.some((item) => existing.has(item.source))) {
    throw new Error("补充映射与已有规则冲突，请重新生成预览。");
  }
  const authorization = authorizeWpsSheets(input.application, input.selectedSheetIndexes);
  const sourceSheet = await readMappingSource(
    { ...input, prompt: "本地映射复核" }, authorization, input.rule.sourceSheet, dependencies
  );
  // Recheck after async host reads as well; never build a preview for a newly switched target.
  await input.beforeRead?.();
  const rule: ColumnMappingRule = {
    ...input.rule,
    mapping: [...input.rule.mapping, ...input.additions]
  };
  return localMappingResult(rule, sourceSheet, input.context);
}

export function buildWpsReadonlyContexts(
  application: WpsApplication,
  results: WpsSheetReadResult[]
): { scope: IntentScopeContext; workbook: WorkbookSnapshot } {
  const metadata = listWpsWorksheets(application);
  const probe = readWpsHostProbe(application);
  const sheets = results.map((result) => ({
    name: result.sheet.name,
    usedRange: usedRangeAddress(result),
    rowCount: result.rowCount,
    columnCount: result.columnCount,
    headers: detectSheetFields(result.values).slice(0, capabilities.intentContext.maxFieldsPerSheet)
  }));
  const authorizedNames = new Set(sheets.map((sheet) => sheet.name));
  const activeWorksheet = authorizedNames.has(probe.activeSheetName)
    ? probe.activeSheetName
    : sheets[0]?.name ?? "";
  const selectedRange = authorizedNames.has(probe.activeSheetName)
    ? probe.selectionAddress
    : null;
  const workbook: WorkbookSnapshot = {
    name: metadata.name,
    capturedAt: new Date().toISOString(),
    activeWorksheet,
    selectedRange,
    /*
     * 稳定工作簿身份，供 SDK 工具桥判归属。必须与宿主适配器
     * （`wpsWorkbookBridgeIdentity`）算出来的一致，否则桥上的轮询会
     * 被判成"不属于本次任务"。
     */
    workbookIdentity: wpsWorkbookBridgeIdentity(application),
    worksheets: sheets.map((sheet) => ({
      ...sheet,
      dataRows: [],
      truncated: sheet.rowCount > 0
    }))
  };
  const scope: IntentScopeContext = {
    workbookName: metadata.name,
    sourceMode: "workbook",
    selectionMode: "manual",
    activeWorksheet,
    selectedRange,
    totalWorksheetCount: metadata.sheets.length,
    // WPS W2 is read-only; unselected names are deliberately not exposed to planning.
    worksheetNames: sheets.map((sheet) => sheet.name),
    sheets
  };
  return { scope, workbook };
}

/**
 * 把读回来的工作表按请求的 `scope` 收窄成本地计算要吃的矩阵。
 *
 * 导出是为了让 SDK 工具桥的宿主适配器**复用同一份**收窄规则 —— 桥上的
 * `query_table` 与本地分析走的是同一个 `query_table`，两份实现迟早会分叉。
 */
export function sheetsForRequest(
  request: DataToolRequest,
  results: WpsSheetReadResult[],
  activeWorksheet: string
): SheetValues[] {
  const selected =
    request.arguments.scope === "active"
      ? results.filter((result) => result.sheet.name === activeWorksheet)
      : results;
  if (selected.length === 0) {
    throw new Error("查询要求读取当前工作表，但当前表不在已授权范围内。");
  }
  return selected.map((result) => ({
    name: result.sheet.name,
    values: result.values
  }));
}

function finalizeResponse(response: AssistantResponse): WpsTurnResult {
  if (response.kind === "clarification") {
    return { kind: "clarification", clarification: { response }, turnId: response.turnId };
  }
  if (response.kind === "answer") {
    return {
      kind: "answer",
      message: response.message,
      resultContext: response.resultContext,
      provider: response.provider,
      outcome: response.outcome,
      // SDK 本地结果卡指针原样上抛：面板据此挂卡，值由本机回环按引用取回。
      localResult: response.localResult ?? null,
      turnId: response.turnId
    };
  }
  if (response.kind === "local_plan_rule") {
    // 这一支由 runWpsTurn 在拿得到授权工作表值的地方处理，不该走到这里。
    // 明确抛错而不是静默当成计划：真走到了说明调用顺序被改坏了。
    throw new Error("本地计算规则需要先读取授权工作表，不能在此处直接转成计划。");
  }
  return { kind: "plan", plan: response.plan, turnId: response.turnId, planOrigin: response.planOrigin };
}

export async function runWpsTurn(
  input: WpsReadonlyTurnInput,
  dependencies: WpsReadonlyAgentDependencies = DEFAULT_DEPENDENCIES
): Promise<WpsTurnResult> {
  const prompt = input.prompt.trim();
  if (prompt.length < 2) throw new Error("请输入至少两个字符的问题。");
  /*
   * 本地阶段的**真实结束**必须由这里补发。
   *
   * 以前本地阶段只发"开始了"，一次都没发过"做完了"，于是下游只能靠
   * "下一步出现了"倒推上一步完成 —— 那是对用户的猜测：最后一步永远停在半途，
   * 中途某一步却可能被提前宣布完成。这里把"开始 / 完成 / 失败"配对发出去，
   * 下游只需要照抄状态，不需要推理。
   */
  const startedStages = new Map<string, WpsReadonlyTurnProgress>();
  const emitLocalStage = (progress: WpsReadonlyTurnProgress) => {
    if (progress.stepId) startedStages.set(progress.stepId, progress);
    input.onProgress?.(progress);
  };
  const finishLocalStage = (
    stepId: string,
    status: TaskProgressStatus = "completed",
    title?: string
  ) => {
    const started = startedStages.get(stepId);
    // 没开始过就不补：凭空插入一条"已完成"会让时间线多出一步用户没见过的事。
    if (!started) return;
    startedStages.delete(stepId);
    input.onProgress?.({
      ...started,
      title: title ?? started.title,
      status,
      message: status === "completed"
        ? `${title ?? started.title ?? stepId}：已完成`
        : `${title ?? started.title ?? stepId}：未完成`
    });
  };
  /** 失败/中止时把所有开着的本地阶段如实收尾，不留"永远在转"的行。 */
  const abandonLocalStages = (status: TaskProgressStatus) => {
    for (const stepId of [...startedStages.keys()]) finishLocalStage(stepId, status);
  };
  if (!input.messageUnderstanding) emitLocalStage({ stage: "intent", stepId: "message-intent", title: "理解消息", status: "started", message: "先理解当前消息，判断是否需要工作簿上下文；尚未读取字段或数据。" });
  const messageUnderstanding = input.messageUnderstanding ?? await (dependencies.understandMessage ?? understandMessage)({
    prompt, imageCount: input.images?.length ?? 0,
    conversation: input.conversation ?? [], modelId: input.modelId || null
  }, input.signal).catch((error: unknown) => {
    abandonLocalStages("failed");
    throw error;
  });
  finishLocalStage("message-intent");
  if (input.signal?.aborted) throw new DOMException("Aborted", "AbortError");
  if (messageUnderstanding.kind === "conversation") {
    return { kind: "answer", message: messageUnderstanding.message, provider: messageUnderstanding.provider, resultContext: null };
  }
  if (messageUnderstanding.kind === "task_recovery") {
    throw new Error("恢复请求必须通过原任务卡片核对，不得新建规划或读取工作表。");
  }
  emitLocalStage({ stage: "reading", stepId: "read-schemas", title: "读取已授权表结构", status: "started", message: "只读取字段结构，不向意图模型发送原始数据行。" });
  await input.beforeRead?.();
  const authorization = authorizeWpsSheets(input.application, input.selectedSheetIndexes);
  const schemas = await dependencies.readSchemas(input.application, authorization, {
    signal: input.signal
  }).catch((error: unknown) => {
    abandonLocalStages("failed");
    throw error;
  });
  finishLocalStage("read-schemas");
  input.signal?.throwIfAborted();
  const { scope, workbook } = buildWpsReadonlyContexts(input.application, schemas);

  emitLocalStage({ stage: "intent", stepId: "intent", title: "结合字段理解任务", status: "started", message: "结合已授权字段判断是否需要澄清、查询或写入计划。" });
  const intent = await waitForPlanning(dependencies.checkIntent(
    {
      prompt,
      scope,
      taskContext: input.taskContext,
      imageCount: input.images?.length ?? 0,
      intentConfirmed: input.intentConfirmed ?? false,
      clarificationRound: input.clarificationRound ?? 0,
      conversation: input.conversation ?? [],
      modelId: input.modelId || null
    },
    input.signal
  ), input.signal).catch((error: unknown) => {
    abandonLocalStages(input.signal?.aborted ? "cancelled" : "failed");
    throw error;
  });
  input.signal?.throwIfAborted();
  finishLocalStage("intent");
  if (intent.kind === "conversation") {
    return { kind: "answer", message: intent.message, provider: intent.provider, resultContext: null };
  }
  /*
   * 意图判断一旦拿到身份就**立刻上报**，早于任何可能失败的网络调用。
   *
   * 若等到规划请求之后才上报，"服务端已收到并开始规划、但流断在半路"
   * 这一档就没有身份可查 —— 而它恰恰是最需要查状态的那一档。
   */
  if (intent.turnId) {
    input.onProgress?.({
      stage: "intent",
      stepId: "planning-identity",
      title: "已建立规划任务身份",
      // 这是一条**事实**，不是一个还在进行的动作：拿到即完成。
      // 不写 `status` 的话，下游只能把它当成"永远在转"的一步。
      status: "completed",
      message: "已从服务端取得本次规划的任务身份，断线后可用它查询真实进度。",
      turnId: intent.turnId
    });
  }
  if (intent.kind === "clarification") {
    return {
      kind: "clarification",
      clarification: { response: intent },
      headerRows: sheetHeaderRows(schemas)
    };
  }

  let dataResults: DataToolResult[] = [];
  let confirmedPrompt = intent.confirmedPrompt;
  /**
   * SDK 适配器下**不能**预置本地计算结果。
   *
   * SDK 只拿得到表结构，模型要自己调 `query_table`，由本机桥在真实宿主上执行
   * 确定性查询。若这里仍把算好的 `dataResults` 塞进请求，服务端的"不上传数据行"
   * 守卫会直接拒绝整条请求 —— 新工具桥就永远拿不到执行权（这正是本轮要修的
   * "旧路径挡住新工具桥"）。跳过预置，读表推迟到桥指针回来的那一刻。
   */
  const sdkPlanning = isSdkPlanningModel(input.modelId);
  if (intent.kind === "tool_request" && !sdkPlanning) {
    await input.beforeRead?.();
    emitLocalStage({ stage: "reading", stepId: "read-data", title: "读取授权数据", status: "started", message: "只为已确定的本地查询读取完整数据。" });
    const readSheetIds = new Set<string>();
    const results = await dependencies.readSheets(input.application, authorization, {
      signal: input.signal,
      onProgress: ({ sheet, rowsRead, totalRows }) => {
        const stepId = `read-data:${sheet.name}`;
        readSheetIds.add(stepId);
        emitLocalStage({
          stage: "reading",
          stepId,
          title: `读取「${sheet.name}」`,
          sheetName: sheet.name,
          rowsRead,
          totalRows,
          // 逐表的进度是"进行中"，不是完成；完成由下面的循环补发。
          status: rowsRead >= totalRows ? "completed" : "progress",
          message: `正在读取「${sheet.name}」：${rowsRead} / ${totalRows} 行`
        });
      }
    }).catch((error: unknown) => {
      abandonLocalStages("failed");
      throw error;
    });
    for (const stepId of readSheetIds) finishLocalStage(stepId);
    finishLocalStage("read-data");
    emitLocalStage({ stage: "calculating", stepId: "query", title: "本地确定性计算", status: "started", toolName: "query_table", message: "在本机计算完整授权数据，不以模型快照代替全表。" });
    const querySheets = sheetsForRequest(intent.request, results, scope.activeWorksheet);
    dataResults = [dependencies.executeQuery(intent.request, querySheets)];
    finishLocalStage("query");
  }

  emitLocalStage({ stage: "answering", stepId: "answer", title: "生成回答或计划", status: "started", message: "根据意图结果与本地计算摘要组织回答。" });
  const hostVersion = String(input.application.Version ?? "").trim() || null;
  const planRequest: PlanRequest = {
    turnId: intent.turnId,
    planningGoal: intent.planningGoal ?? (intent.kind === "tool_request" ? "read_only" : "write_preview"),
    prompt: confirmedPrompt,
    taskContext: input.taskContext,
    workbook,
    dataResults,
    planningReadPolicy: "none" as const,
    images: input.images?.map(({ name, mediaType, data }) => ({ name, mediaType, data })) ?? [],
    modelId: input.modelId || null,
    // 规划前能力收窄（方案 §4.3）：服务端按可信配置算上界，与本端自报求交。
    // 自报只能收窄，不能扩大 —— 服务端不采信这段内容作为授权凭据。
    hostId: "wps-et",
    hostVersion,
    selfReportedActions: wpsReportedActionTypes(hostVersion)
  };
  /**
   * 规划走**流式**通道（CAP-02）：事件增量展示，且服务端会持久化
   * 请求/事件/检查点，断线后能用同一个 `turnId` 查状态、按检查点继续。
   *
   * 但流式是**服务端能力**，不是本端能假定的：旧服务端没有
   * `/api/turn/stream`（404/405），或者代理不支持流式（无 body）。
   * 这两种情况退回既有的一次性 `/api/turn` —— 不是"降级丢结果"，
   * 两条通道返回的是**同一个结构化联合**，分支处理完全一致。
   *
   * 注意顺序：**先试流式，且只在真正的"连不上/端点不存在"时回退**。
   * 业务错误（模型未配置、能力不匹配）必须原样抛出，不能被回退吞成
   * "再试一次一次性接口"，那会掩盖真实原因并可能产生第二次模型调用。
   */
  const turnIdForStream = intent.turnId ?? "";
  if (input.preparePlanningBridge) {
    let ready = false;
    if (turnIdForStream && !input.signal?.aborted) {
      try {
        ready = await input.preparePlanningBridge(turnIdForStream);
      } catch {
        // Do not expose host exception details or fall back to an unbridged request.
        ready = false;
      }
    }
    if (input.signal?.aborted) throw new DOMException("Aborted", "AbortError");
    if (!ready) throw new Error("本机读表桥未就绪，未发送规划请求。请重试本次需求。");
  }
  let response: AssistantResponse | null = null;
  let pause: PlanningPauseResult | null = null;
  if (turnIdForStream) {
    let streamed: Awaited<ReturnType<typeof streamWpsPlanning>> | null = null;
    let fallBack = false;
    try {
      streamed = await dependencies.streamPlanning(
        planRequest as unknown as Record<string, unknown>,
        turnIdForStream,
        {
          signal: input.signal,
          onProgress: (event) => {
            input.onProgress?.(wpsProgressFromTaskEvent(turnIdForStream, event));
          }
        }
      );
    } catch (reason) {
      // 只有"连不上本地服务"或"端点不存在"才回退；其余原样抛出。
      if (isStreamEndpointUnavailable(reason)) {
        fallBack = true;
      } else {
        throw reason;
      }
    }
    if (!fallBack && streamed) {
      // 先按结构化种类分流：暂停**不是** `AssistantResponse`，不能硬塞进同一支。
      // `assertKnownResultKind` 保证未知种类在此报错，而不是降级成 answer。
      const kind = assertKnownResultKind(streamed);
      if (kind === "planning_pause") {
        pause = streamed as PlanningPauseResult;
      } else {
        response = streamed as AssistantResponse;
      }
    } else {
      response = await dependencies.createAssistantResponse(planRequest, input.signal);
    }
  } else {
    // 没有 turnId 就无法走持久化流式通道（服务端要靠它去重与恢复）。
    response = await dependencies.createAssistantResponse(planRequest, input.signal);
  }
  if (pause) {
    /*
     * 暂停是结构化终态，不是失败：原样交给调用方落"可操作暂停卡"。
     *
     * 但"生成回答或计划"这一步**没有完成**，不能留一行永远在转：调用方会按
     * `awaiting_input` 收尾，这里先把本地阶段如实标成"未完成"，让时间线上
     * 不留悬念。判据只写在事件里，不靠调用方再猜一次。
     */
    finishLocalStage("answer", "warning");
    return { kind: "planning_pause", pause };
  }
  if (!response) {
    abandonLocalStages("failed");
    throw new Error("规划没有返回任何结构化结果。");
  }
  // 规划这一轮已经拿到结构化结果，"生成回答或计划"到此真正结束。
  finishLocalStage("answer");
  if (response.kind === "clarification") {
    return {
      kind: "clarification", clarification: { response },
      headerRows: sheetHeaderRows(schemas)
    };
  }
  if (response.kind === "local_plan_rule") {
    // 服务端已经判定"这条规则还缺参数/缺类别，算不出正确答案"。
    // 这时**不要**读源值、更不要算：与 Excel 端同一判据 —— 算出来的矩阵必然
    // 是猜的，而猜错在单元格层面看不出来。把归因接进澄清卡，让用户补完再跑。
    //
    // 位置很关键：这段必须在 `readMappingSource` **之前**。放后面等于先读了
    // 整片源列才告诉用户"其实缺映射" —— 白读一次源值。
    if (response.rule.clarification) {
      return {
        kind: "clarification",
        clarification: {
          columns: {
            clarification: response.rule.clarification,
            scopeLabel: `${response.rule.sourceSheet}!${response.rule.sourceRange}`,
            summary: response.summary
          }
        },
        headerRows: sheetHeaderRows(schemas)
      };
    }
    // 服务端只给了规则（它手上没有源值）。到这里才读源区域：上面 schema 扫描
    // 只取字段名，不含数据行。读出的值**只在本端用**，算完不回传。
    input.onProgress?.({ stage: "reading", stepId: "read-mapping-source", title: "读取授权源列", status: "started", message: "映射源值只留在本机，不上传服务端计划校验。" });
    const sourceSheet = await readMappingSource(
      input,
      authorization,
      response.rule.sourceSheet,
      dependencies
    );
    input.onProgress?.({ stage: "reading", stepId: "read-mapping-source", title: "读取授权源列", status: "completed", message: "已读取映射源列，值只在本机使用。" });
    input.onProgress?.({ stage: "calculating", stepId: "map-column", title: "计算目标列", status: "started", toolName: "map_column", message: "在本机按映射规则生成目标值。" });
    const mapped = localMappingResult(response.rule, sourceSheet, {
      summary: response.summary,
      warnings: response.warnings,
      assumptions: response.assumptions
    }, sheetHeaderRows(schemas));
    input.onProgress?.({ stage: "calculating", stepId: "map-column", title: "计算目标列", status: "completed", toolName: "map_column", message: "已按映射规则算出目标列。" });
    // 带上这次规划的身份：断线后要先拿它查真实状态，而不是重新发一遍需求。
    return { ...mapped, turnId: response.turnId } as WpsTurnResult;
  }
  return finalizeResponse(response);
}

/** 读出规则指定的那一张授权表。表不在授权范围内时拒绝，不自动扩大范围。 */
async function readMappingSource(
  input: WpsReadonlyTurnInput,
  authorization: ReturnType<typeof authorizeWpsSheets>,
  sheetName: string,
  dependencies: WpsReadonlyAgentDependencies
): Promise<MappingSourceSheet> {
  await input.beforeRead?.();
  const results = await dependencies.readSheets(input.application, authorization, {
    signal: input.signal,
    onProgress: ({ sheet, rowsRead, totalRows }) => {
      input.onProgress?.({
        stage: "reading",
        stepId: `read-mapping-source:${sheet.name}`,
        title: `读取「${sheet.name}」`,
        sheetName: sheet.name,
        rowsRead,
        totalRows,
        message: `正在读取「${sheet.name}」：${rowsRead} / ${totalRows} 行`
      });
    }
  });
  const match = results.find((result) => result.sheet.name === sheetName);
  if (!match) {
    throw new Error(
      `映射规则要读取「${sheetName}」，但它不在本次已授权的工作表里。`
    );
  }
  return {
    name: match.sheet.name,
    startRow: match.startRow,
    startColumn: match.startColumn,
    columnCount: match.columnCount,
    values: match.values
  };
}

/** Compatibility alias for extensions/tests importing the W2 name. */
export const runWpsReadonlyTurn = runWpsTurn;
export type WpsReadonlyTurnResult = WpsTurnResult;

/**
 * 查询一条规划任务的**真实状态**（CAP-02）。
 *
 * 存在的意义：断线之后，界面上那份"上次看到的进度"已经不可信 —— 权威现场
 * 只在服务端账本里。因此在任何"继续"之前，必须先用原 `turnId` 查询一次。
 *
 * 只查询，**不创建任务**：服务端侧只是读账本，不会触发新的规划或工具调用。
 */
export async function inspectWpsPlanningTask(
  turnId: string,
  signal?: AbortSignal,
  dependencies: WpsReadonlyAgentDependencies = DEFAULT_DEPENDENCIES
): Promise<WpsPlanningTaskStatus> {
  if (!turnId.trim()) throw new Error("规划任务身份不能为空。");
  return dependencies.getPlanningStatus(turnId, signal);
}

/**
 * 面板重开后，重新核对一条历史暂停的**真实现场**（CAP-02）。
 *
 * 与 `inspectWpsPlanningTask` 的区别只在**入口用途**：这个函数明确表达
 * "我从历史里翻出一条暂停、现在要拿服务端的真实状态把它的按钮重算一遍"。
 * 因此它顺便把结构化快照补回 `PlanningPauseResult` 形状 —— 界面上的暂停卡
 * 需要一个可以继续走的载荷，而历史里只留了定位与归因（不含连接元数据）。
 *
 * **只查询，不继续**。重开之后绝不能自动续跑：那会让用户在完全没点过任何
 * 东西的情况下触发一次模型调用。这里只负责把"现在能不能续"如实算出来。
 */
export async function reinspectWpsPlanningPause(
  entry: { turnId: string; pausedAt?: PlanningPauseResult | null },
  signal?: AbortSignal,
  dependencies: WpsReadonlyAgentDependencies = DEFAULT_DEPENDENCIES
): Promise<{
  status: WpsPlanningTaskStatus;
  view: WpsPlanningResumeView;
  /**
   * 可以挂回暂停卡的载荷。**优先用服务端刚返回的真实状态**；
   * 服务端没给暂停成因时退回历史快照 —— 但只用于显示，不用于判定按钮
   * （按钮一律由 `view` 决定，那个是从真实状态推的）。
   */
  payload: PlanningPauseResult;
}> {
  const status = await inspectWpsPlanningTask(entry.turnId, signal, dependencies);
  const view = wpsPlanningResumeView(status);
  return {
    status,
    view,
    payload: planningPauseFromStatus(status, entry.pausedAt ?? null)
  };
}

/** 续跑方案的对外视图（按钮文案与能不能点，只由服务端结构化状态推导）。 */
export type WpsPlanningResumeView = {
  action: "continue_checkpoint" | "regenerate" | "none";
  actionLabel: string;
  detail: string;
  canContinue: boolean;
  reason: string;
};

/**
 * 把一条真实状态响应还原成可挂回暂停卡的 `PlanningPauseResult`。
 *
 * 断线恢复（第 4 条）拿到的是 `GET /status` 的**状态视图**，而暂停卡需要一个
 * `planning_pause` 载荷。这里做**补齐**而不是"造数"：
 *
 *  - 服务端给了的字段（`revision` / `lastSequence` / `canContinue` /
 *    `pauseType` / `reason`）原样用；
 *  - 服务端没给的执行器计量字段，只有历史快照里有才复用，否则用**保守默认**——
 *    一律不编造"还剩多少预算"这类数字，界面据此显示的也会是"未知"。
 */
export function planningPauseFromStatus(
  status: WpsPlanningTaskStatus,
  pausedAt?: PlanningPauseResult | null
): PlanningPauseResult {
  const base = pausedAt ?? ({} as PlanningPauseResult);
  return {
    ...base,
    kind: "planning_pause",
    turnId: status.turnId,
    status: status.status === "interrupted" ? "interrupted" : "waiting_user",
    revision: status.revision,
    lastSequence: status.lastSequence,
    canContinue: status.canContinue === true,
    pauseType: status.pauseType ?? base.pauseType ?? "interrupted",
    reason: status.reason ?? base.reason ?? "本轮处理暂停，当前进度已保存。"
  } as PlanningPauseResult;
}

/**
 * 由一条真实状态响应推导"下一步能做什么"。
 *
 * 这是**唯一**允许产出继续按钮的地方。它直接吃
 * `getWpsPlanningTaskStatus` 的返回：`canContinue` / `pauseType` /
 * `error.code` 都是服务端给的判据，本端不额外"乐观放行"。
 */
export function wpsPlanningResumeView(status: WpsPlanningTaskStatus): WpsPlanningResumeView {
  return wpsPlanningResumeDecision(status);
}

export type WpsPlanningContinueResult = {
  status: WpsPlanningTaskStatus;
  /** 续跑后的结构化结果：可能是答案/计划/澄清，也可能再次暂停。 */
  outcome: Awaited<ReturnType<typeof continueWpsPlanning>>;
};

/**
 * 按服务端检查点继续一次暂停的规划（CAP-02）。
 *
 * 三条不可放宽的纪律：
 *
   *  1. **先查状态再继续。** `expectedRevision` 用刚查到的真实 revision 做 CAS；
   *     `afterSequence` 用客户端实际已消费的位置（未知从零补收），不能冒用服务端最新值。
 *  2. **不重发原需求。** 请求体里只有身份与游标（见 `WpsPlanningContinueBody`）。
 *     带 `prompt` 会让服务端把它当成一次新规划，重复调用模型与工具。
 *  3. **写入次数为 0。** 续跑只推进"规划"，产出仍需经过既有的预览与用户确认
 *     才会写单元格。这条路径上没有、也不允许有任何宿主写动作。
 */
export async function resumeWpsPlanning(
  input: {
    turnId: string;
    modelId?: string | null;
    confirmModelSwitch?: true;
    signal?: AbortSignal;
    onProgress?: WpsReadonlyTurnInput["onProgress"];
    /**
     * 本次继续操作的**稳定幂等键**。
     *
     * 由调用方在"用户首次点击继续"时生成一次并固定；超时重试、网络抖动后的
     * 重发都必须**原样沿用同一个值**。服务端侧靠 `(turn_id, request_id)` 判重
     * （`claim_planning_continuation`）：同一个 requestId 的重复领取只会追加一次
     * 预算、只跑一次规划；换新 requestId 则被当成"另一次继续"，重复扣预算、
     * 重复调用模型。
     *
     * 省略时内部生成一个 —— 仅用于**单次调用**（例如测试或非交互路径）。
     * 交互式重试路径**必须**显式传入，否则重试会变成新一次继续。
     */
    requestId?: string;
    /**
     * 客户端已确认消费到的事件序号（`clientLastSequence` 与服务端
     * `lastSequence` 取小的那个，见 `planningAfterSequence`）。
     *
     * 省略时退回 `status.lastSequence` —— 那只适用于"客户端从未消费过这条
     * 任务的事件"的场景。**只要本端展示过进度**，就必须传真实游标，
     * 否则中途漏收的事件会被永久跳过。
     */
    afterSequence?: number;
    /**
     * 本轮**重新核实**过的现场证据。三项必填，与 Excel 端同一份结构与判据
     * （`buildPlanningContinuationEvidence`）：缺证据就不该走到这里，
     * 服务端也不再接受"省略即跳过比对"。
     */
    evidence: PlanningContinuationEvidence;
  },
  dependencies: WpsReadonlyAgentDependencies = DEFAULT_DEPENDENCIES
): Promise<WpsPlanningContinueResult> {
  const turnId = input.turnId.trim();
  if (!turnId) throw new Error("续跑需要一个已存在的规划任务身份。");

  /*
   * 幂等键在**查状态之前**就定下来。
   *
   * 放在查状态之后的话，"查状态成功、继续请求超时、用户重试"这条路径会拿到
   * 一个新的 requestId —— 服务端看到的是两次不同的继续，于是重复追一次预算。
   * 定在函数入口、"同一次继续操作 = 同一个 requestId"才成立。
   */
  const requestId = input.requestId?.trim() || newPlanningRequestId();

  // 第一步永远是查真实状态 —— 界面上的 revision/游标可能已经过期。
  const status = await dependencies.getPlanningStatus(turnId, input.signal);
  const decision = wpsPlanningResumeDecision(status);
  if (!decision.canContinue) {
    // 服务端说不能继续（已完成/已停止/没有可验证现场）：不硬续，
    // 把真实原因交给界面。这比"试一次再说"安全 —— 试一次可能重复规划。
    return { status, outcome: null as never };
  }
  if (decision.action !== "continue_checkpoint") {
    // `regenerate` 走的是"重新读源并生成预览"，属于新任务，不走这条续跑路径。
    return { status, outcome: null as never };
  }

  const outcome = await dependencies.continuePlanning(
    {
      turnId,
      expectedRevision: status.revision,
      requestId,
      /*
       * 游标只能交"客户端已确认消费到的位置"，**不能**直接用 `status.lastSequence`。
       *
       * 服务端补发的是 `> afterSequence` 的事件。若把服务端最新序号当游标，
       * 那些"已经提交到账本、但客户端在断线前没收到"的事件（序号 ≤ 最新）
       * 就被**永久跳过** —— 界面上步骤静默缺失，用户看到一份不完整的过程。
       * 传更小的值只会重复补发（幂等，可去重），传更大的值会丢事件（不可逆）。
       */
      afterSequence: Math.min(Number.isSafeInteger(input.afterSequence) && input.afterSequence! >= 0
        ? input.afterSequence! : 0, status.lastSequence),
      modelId: input.modelId ?? undefined,
      confirmModelSwitch: input.confirmModelSwitch,
      // 证据原样上交：这里**不**做"有就传、没有就省略"的条件拼装 ——
      // 那种写法正是上一版"三项全省略也能继续"的来源。
      workbookIdentity: input.evidence.workbookIdentity,
      authorizedSheetNames: [...input.evidence.authorizedSheetNames],
      sourceVerification: input.evidence.sourceVerification,
      ...(input.evidence.sourceVerification === "verified" &&
      input.evidence.sourceFingerprint
        ? { sourceFingerprint: input.evidence.sourceFingerprint }
        : {})
    },
    {
      signal: input.signal,
      onProgress: (event) => {
        // 与首次规划**同一份**映射：续跑时模型同样会发 `bridgeTool` 指针，
        // 漏传一次就等于这一轮续跑永远读不到表。
        input.onProgress?.(wpsProgressFromTaskEvent(turnId, event, "planning-continue", "正在继续规划"));
      }
    }
  );
  return { status, outcome };
}

export type { ModelOption };
