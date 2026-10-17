/**
 * 规划流 ↔ 本机桥的接线。
 *
 * 规划流里出现 ``bridgeTool`` 指针，就表示服务端的 SDK 正在等本机执行一次
 * 本地查询并回执。本 hook 把指针交给共享状态机（``runBridgeUntilIdle``），
 * 由宿主适配器完成「轮询 → 复核授权 → 真实读取 + 确定性计算 → 回执」。
 *
 * 三条纪律：
 * - 一轮内**串行**服务：指针可能密集到达，重入会让同一次调用被执行两次。
 * - 取消/失败即停止：不再轮询，也不接受迟到回执。
 * - 不碰宿主写入 API：本 hook 只读。
 */

import { useCallback, useEffect, useRef } from "react";
import {
  applyBridgeProgressEvent,
  cancelBridge,
  createSdkBridgeState,
  runBridgeUntilIdle,
  type BridgeHostId,
  type BridgeHostPorts,
  type SdkBridgeBinding,
  type SdkBridgeState
} from "@excel-bro/core";

import type { QueryToolProgress } from "../dataTools";
import {
  createSdkBridgeBinding,
  createSdkBridgeHostPorts,
  type SdkBridgeHostDependencies,
  type SdkBridgeLiveContext
} from "../sdkBridgeHost";

/**
 * 两端共用的可替换件。
 *
 * 宿主差异只有两处，都从这里进：
 *
 * - ``hostId``：绑定归属要如实写清是哪个宿主；
 * - ``createPorts``：怎么读宿主。缺省是 Excel（Office.js）；WPS 传自己的适配器。
 *
 * 其余（绑定、状态机、轮询/回执/取值）两端共用同一份实现 —— 各写一套迟早分叉。
 * 剩下的 ``poll``/``receipt``/``fetchResult``/``execute`` 只在**用缺省
 * createPorts 时**生效，WPS 不需要它们。
 */
export interface SdkBridgeHookDependencies extends SdkBridgeHostDependencies {
  hostId?: BridgeHostId;
  createPorts?: (
    options: SdkBridgeRunOptions,
    binding: SdkBridgeBinding
  ) => BridgeHostPorts;
}

export interface SdkBridgeRunOptions {
  taskId: string;
  workbookName: string;
  workbookIdentity?: string | null;
  sourceFingerprint?: string | null;
  /**
   * 每次执行前读一次**当前**授权与工作簿身份。
   *
   * 允许异步，也**必须真的重读宿主**：返回规划启动时的闭包快照，等于用一个
   * 静止的现场去"证明"等待期间什么都没变 —— 收窄/撤销授权、换簿、改源值都
   * 会漏过去。绑定时的范围快照取自这里，执行前的复核也取自这里，同一个读取器。
   */
  liveContext: () => SdkBridgeLiveContext | Promise<SdkBridgeLiveContext>;
  signal?: AbortSignal;
  onProgress?: (progress: QueryToolProgress) => void;
  /** 本地查询失败（固定错误码）时回调，用于在界面上如实说明。 */
  onError?: (errorCode: string) => void;
}

export interface SdkBridgeController {
  begin(options: SdkBridgeRunOptions): Promise<void>;
  /** 消费一条规划进度事件；带指针时触发服务。 */
  handleProgressEvent(
    event: { status?: string; stepId?: string; bridgeTool?: unknown } | null | undefined
  ): void;
  cancel(): void;
  end(): void;
  binding(): SdkBridgeBinding | null;
  failure(): string | null;
}

export function useSdkToolBridge(
  dependencies: SdkBridgeHookDependencies = {}
): SdkBridgeController {
  const stateRef = useRef<SdkBridgeState>(createSdkBridgeState(null));
  const optionsRef = useRef<SdkBridgeRunOptions | null>(null);
  const runningRef = useRef(false);
  const queuedRef = useRef(false);
  const failureRef = useRef<string | null>(null);
  const dependenciesRef = useRef(dependencies);
  dependenciesRef.current = dependencies;

  const drain = useCallback(async () => {
    if (runningRef.current) {
      // 已有一次服务在跑：记账，跑完再补一轮，避免并发重复执行。
      queuedRef.current = true;
      return;
    }
    runningRef.current = true;
    try {
      do {
        queuedRef.current = false;
        const options = optionsRef.current;
        const binding = stateRef.current.binding;
        if (options === null || binding === null || stateRef.current.cancelled) {
          return;
        }
        const ports = (
          dependenciesRef.current.createPorts ??
          ((options_, binding_) =>
            createSdkBridgeHostPorts(options_, binding_, dependenciesRef.current))
        )(options, binding);
        const result = await runBridgeUntilIdle(stateRef.current, ports, {
          maxRounds: 8
        });
        stateRef.current = result.state;
        for (const entry of result.outcomes) {
          if (entry.code !== null && entry.code !== undefined) {
            failureRef.current = entry.code;
            optionsRef.current?.onError?.(entry.code);
          }
        }
      } while (queuedRef.current && !stateRef.current.cancelled);
    } catch {
      // 桥失败不能拖垮整轮规划：服务端会超时并把固定错误码交回模型。
    } finally {
      runningRef.current = false;
    }
  }, []);

  const handleProgressEvent = useCallback(
    (event: { status?: string; stepId?: string; bridgeTool?: unknown } | null | undefined) => {
      const next = applyBridgeProgressEvent(stateRef.current, event);
      if (next === stateRef.current) return;
      stateRef.current = next;
      if (next.cancelled) return;
      void drain();
    },
    [drain]
  );

  const begin = useCallback(async (options: SdkBridgeRunOptions) => {
    failureRef.current = null;
    optionsRef.current = options;
    // 授权范围快照取自**与执行前复核同一个读取器**，避免调用方再传一份、
    // 两份迟早对不上。读不到现场就退回空范围 → 复核一律拒绝（fail-closed）。
    let authorizedSheetNames: string[] = [];
    try {
      authorizedSheetNames = [...(await options.liveContext()).authorizedSheetNames];
    } catch {
      authorizedSheetNames = [];
    }
    const binding = createSdkBridgeBinding({
      taskId: options.taskId,
      hostId: dependenciesRef.current.hostId ?? "excel-office",
      workbookName: options.workbookName,
      workbookIdentity: options.workbookIdentity,
      sourceFingerprint: options.sourceFingerprint,
      authorizedSheetNames
    });
    stateRef.current = createSdkBridgeState(binding);
    queuedRef.current = false;
  }, []);

  const cancel = useCallback(() => {
    stateRef.current = cancelBridge(stateRef.current);
    queuedRef.current = false;
  }, []);

  const end = useCallback(() => {
    // 停轮询但**保留绑定**：本轮的结果卡要在会话关闭之后按引用取值，
    // 服务端校验的是授权归属（claim）而不是活跃会话。
    stateRef.current = {
      ...stateRef.current,
      cancelled: true,
      pending: [],
      pollRequested: false
    };
    optionsRef.current = null;
    queuedRef.current = false;
  }, []);

  const binding = useCallback(() => stateRef.current.binding, []);

  useEffect(() => () => cancel(), [cancel]);

  return { begin, handleProgressEvent, cancel, end, binding, failure: () => failureRef.current };
}
