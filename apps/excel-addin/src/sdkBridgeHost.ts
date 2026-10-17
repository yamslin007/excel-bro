/**
 * Excel（Office.js）侧的桥宿主适配器。
 *
 * 把共享状态机的三个端口落到真实宿主上：
 * ``poll`` 拉取待执行调用 → ``verify`` 复核**当前**授权 → ``execute`` 用既有的
 * 确定性 ``query_table`` 读真实工作簿 → ``receipt`` 回执。
 *
 * 边界：
 * - 请求参数与值只走本机回环；模型、聊天历史与持久账本都看不到值。
 * - 执行前必须复核授权：工作簿身份、来源指纹、工作表范围任一变化都拒绝，
 *   否则会拿旧授权去读新数据。
 * - 这里**只读**：本模块不调用任何宿主写入 API。
 */

import {
  BRIDGE_UNVERIFIABLE_ERROR_CODE,
  bridgeWorkbookIdentity,
  verifyBridgeAuthorizationOrRefuse,
  type BridgeCallOutcome,
  type BridgeHostId,
  type BridgeHostPorts,
  type SdkBridgeBinding
} from "@excel-bro/core";

import {
  fetchSdkBridgeResult,
  pollSdkBridgeTool,
  postSdkBridgeReceipt
} from "./api";
import {
  DataToolExecutionError,
  executePreparePythonData,
  executeQueryTableTool,
  type PreparePythonDataSource,
  type QueryToolProgress
} from "./dataTools";
import { recordTrustedToolReads } from "./planSourceDependency";
import type {
  QueryTableArguments,
  SdkBridgeResultResponse,
  SdkBridgeToolReceipt
} from "./contracts";

/** 执行前重读的真实宿主现场。 */
export interface SdkBridgeLiveContext {
  workbookIdentity: string;
  sourceFingerprint: string | null;
  authorizedSheetNames: string[];
  activeWorksheet: string;
}

/** 只影响「绑定是什么」的部分。 */
export interface SdkBridgeBindingOptions {
  taskId: string;
  hostId: BridgeHostId;
  workbookName: string;
  workbookIdentity?: string | null;
  sourceFingerprint?: string | null;
  /** 本轮确认过的工作表范围；执行前复核要拿它跟当前授权比。 */
  authorizedSheetNames?: string[];
}

/** 只影响「怎么执行」的部分。 */
export interface SdkBridgeExecutionOptions {
  /**
   * 每次执行前调用一次，**重新读**当前授权、工作簿身份与来源指纹。
   *
   * 可以是异步的：读取真实宿主现场本身就是要 await 的。**绝不能**返回规划
   * 启动时的闭包快照 —— 那样等待期间用户收窄/撤销授权、换簿或改了源数据，
   * 复核都会拿旧现场"证明"没问题。
   */
  liveContext: () =>
    | SdkBridgeLiveContext
    | Promise<SdkBridgeLiveContext>;
  signal?: AbortSignal;
  onProgress?: (progress: QueryToolProgress) => void;
}

export type SdkBridgeHostOptions = SdkBridgeBindingOptions & SdkBridgeExecutionOptions;

export function createSdkBridgeBinding(
  options: SdkBridgeBindingOptions
): SdkBridgeBinding {
  return {
    taskId: options.taskId,
    hostId: options.hostId,
    workbookName: options.workbookName,
    workbookIdentity: bridgeWorkbookIdentity(
      options.workbookName,
      options.workbookIdentity
    ),
    sourceFingerprint: options.sourceFingerprint ?? null,
    authorizedSheetNames: [...(options.authorizedSheetNames ?? [])]
  };
}

/** 本地确定性查询失败 → 服务端固定错误码。绝不把异常文本回给模型。 */
export function bridgeErrorCode(reason: unknown): string {
  if (reason instanceof DataToolExecutionError) {
    switch (reason.code) {
      case "SCOPE_VIOLATION":
        return "unauthorized_scope";
      case "TOOL_LIMIT_EXCEEDED":
        return "result_too_large";
      case "CANCELLED":
        return "cancelled";
      case "UNSUPPORTED_TOOL":
      case "FIELD_NOT_FOUND":
        return "invalid_arguments";
      default:
        return "internal";
    }
  }
  if (reason instanceof DOMException && reason.name === "AbortError") {
    return "cancelled";
  }
  return "internal";
}

export interface SdkBridgeHostDependencies {
  poll?: typeof pollSdkBridgeTool;
  receipt?: typeof postSdkBridgeReceipt;
  fetchResult?: typeof fetchSdkBridgeResult;
  execute?: typeof executeQueryTableTool;
  prepareData?: typeof executePreparePythonData;
}

/** 组装共享状态机需要的三个端口。 */
export function createSdkBridgeHostPorts(
  options: SdkBridgeExecutionOptions,
  binding: SdkBridgeBinding,
  dependencies: SdkBridgeHostDependencies = {}
): BridgeHostPorts {
  const poll = dependencies.poll ?? pollSdkBridgeTool;
  const postReceipt = dependencies.receipt ?? postSdkBridgeReceipt;
  const execute = dependencies.execute ?? executeQueryTableTool;
  const prepareData = dependencies.prepareData ?? executePreparePythonData;

  return {
    poll: (current) => poll(current, options.signal),

    // 复核必须读**此刻**的宿主现场：`liveContext()` 会真的重读一次，不是取快照。
    // 读不到现场时走共享的 fail-closed 拒绝（必须回一张拒绝回执，不能抛给上层吞掉）。
    verify: (_current, call) =>
      verifyBridgeAuthorizationOrRefuse(binding, call, () => options.liveContext()),

    execute: async (_current, call): Promise<BridgeCallOutcome> => {
      let live: SdkBridgeLiveContext;
      try {
        /*
         * 现场读取单独接住：读不到 = **无法核实**授权/身份/来源，要与复核给同一个
         * 固定错误码（"请重新确认数据范围"），而不是混成 `internal` 让用户无从下手。
         * `verify` 正常会先拦住；这里是防御性的一致出口。
         */
        live = await options.liveContext();
      } catch {
        return { kind: "error", errorCode: BRIDGE_UNVERIFIABLE_ERROR_CODE };
      }
      try {
        if (call.tool === "prepare_python_data") {
          const sources = asPreparePythonSources(call.arguments);
          const blocks = await prepareData(
            sources,
            live.authorizedSheetNames,
            { signal: options.signal, onProgress: options.onProgress }
          );
          /*
           * 可信来源记账：数据登记也是"确定性地读过表"，读了谁由**实际读取成功
           * 的来源**回报，不是模型自报的参数集合。异常会在下面 `catch` 里变成
           * 固定错误码，此处只记账成功路径 —— 失败时不留记账（"没读到" ≠ "读过"）。
           */
          recordTrustedToolReads(
            binding.taskId,
            blocks.map((block) => sources[block.sourceIndex]?.sheet ?? "").filter(Boolean)
          );
          return { kind: "result", result: { datasets: blocks } };
        }
        const result = await execute(
          {
            // 用桥上的操作 ID 作为请求身份，便于把本地结果与那次调用对上。
            id: call.operationId,
            tool: "query_table",
            arguments: asQueryTableArguments(call.arguments)
          },
          live.authorizedSheetNames,
          live.activeWorksheet,
          { signal: options.signal, onProgress: options.onProgress }
        );
        /*
         * 可信来源记账：读了哪几张表由**确定性执行器**回报（`sourceSheets` 是它
         * 实际 addSheet 过的表），不是模型自报的参数。断点恢复的"实际来源依赖"
         * 只从这条记录来 —— 这里也是全应用唯一的记账入口。
         */
        recordTrustedToolReads(binding.taskId, result.sourceSheets ?? []);
        return { kind: "result", result };
      } catch (reason) {
        return { kind: "error", errorCode: bridgeErrorCode(reason) };
      }
    },

    receipt: async (current, call, outcome) => {
      const receipt: SdkBridgeToolReceipt =
        outcome.kind === "result"
          ? call.tool === "prepare_python_data"
            ? { datasets: (outcome.result as { datasets: SdkBridgeToolReceipt["datasets"] }).datasets }
            : { result: outcome.result as SdkBridgeToolReceipt["result"] }
          : { errorCode: outcome.errorCode };
      return postReceipt(
        {
          ...current,
          operationId: call.operationId,
          requestFingerprint: call.requestFingerprint,
          receipt
        },
        options.signal
      );
    }
  };
}

/**
 * 桥上的参数由服务端用 ``QueryTableArguments`` 校验过，这里只做形状闸：
 * 非对象直接按参数错误处理，真正的字段校验交给既有的确定性查询实现。
 */
function asQueryTableArguments(value: unknown): QueryTableArguments {
  if (value === null || typeof value !== "object" || Array.isArray(value)) {
    throw new DataToolExecutionError(
      "UNSUPPORTED_TOOL",
      "本地查询参数无效",
      false
    );
  }
  return value as QueryTableArguments;
}

/**
 * 桥上的 ``prepare_python_data`` 参数：`{ sources: [{sheet, range}] }`。
 * 这里只做形状闸，真正的范围核对在 ``executePreparePythonData`` 里做。
 */
function asPreparePythonSources(value: unknown): PreparePythonDataSource[] {
  if (value === null || typeof value !== "object" || Array.isArray(value)) {
    throw new DataToolExecutionError(
      "UNSUPPORTED_TOOL",
      "数据登记参数无效",
      false
    );
  }
  const sources = (value as { sources?: unknown }).sources;
  if (!Array.isArray(sources)) {
    throw new DataToolExecutionError(
      "UNSUPPORTED_TOOL",
      "数据登记来源无效",
      false
    );
  }
  return sources as PreparePythonDataSource[];
}

/** 取回结果值时回带的绑定。接受「可选身份」的宽松形状，便于消息里持久化的指针。 */
export interface BridgeResultLookup {
  taskId: string;
  hostId: BridgeHostId;
  workbookName: string;
  workbookIdentity?: string | null;
  sourceFingerprint?: string | null;
}

/**
 * 按引用取回结果值，供本地结果卡渲染。
 *
 * 此时会话通常已经关闭：服务端校验的是**授权归属**而不是活跃会话，所以要回带
 * 产生该结果那一轮的绑定（任务 + 工作簿身份），不能用「当前」绑定顶替。
 */
export async function loadBridgeResult(
  binding: BridgeResultLookup,
  resultRef: string,
  signal?: AbortSignal
): Promise<SdkBridgeResultResponse | null> {
  try {
    return await fetchSdkBridgeResult({ ...binding, resultRef }, signal);
  } catch {
    return null;
  }
}
