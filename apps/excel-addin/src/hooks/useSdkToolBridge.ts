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
  /**
   * 开一代新的桥。返回**这一代是否真的建立起来**：初始化（读宿主现场）期间被
   * `cancel()`/`end()`/另一次 `begin()` 取代时返回 false，且不写任何状态。
   */
  begin(options: SdkBridgeRunOptions): Promise<boolean>;
  /**
   * 保证"这一刻有一条为 `taskId` 服务的活桥"。
   *
   * 续跑路径必须走这个而不是 `begin`：
   * - 同一轮的多个指针**只初始化一次**（每来一个事件就 begin 会把 `delivered`
   *   清空，同一次调用被再执行一遍）；
   * - 上一轮留下的 `binding` **不等于**桥还活着（`end()` 之后绑定仍在，因为结果卡
   *   还要按引用取值）。只看 binding 是否存在，就会拿一条已收尾的桥去接新指针，
   *   结果是"继续分析之后桥又聋了" —— 服务端一直等到超时。
   *
   * 返回是否拿到了活桥；false 表示建不起来（读不到宿主现场等），调用方应停止继续。
   */
  ensureLive(options: SdkBridgeRunOptions): Promise<boolean>;
  /** Reconcile pending server calls after attaching, including pointers already received before pause. */
  pollPending(): Promise<void>;
  /** 消费一条规划进度事件；带指针时触发服务。 */
  handleProgressEvent(
    event: { status?: string; stepId?: string; bridgeTool?: unknown } | null | undefined
  ): void;
  cancel(): void;
  end(): void;
  binding(): SdkBridgeBinding | null;
  /** 桥此刻是否真的可服务（有绑定 + 未收尾 + 未取消）。 */
  isLive(): boolean;
  /** Pin the first tool's source without resetting this bridge or its delivered calls. */
  pinSourceFingerprint(taskId: string, live: SdkBridgeLiveContext): boolean;
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
  /**
   * 本地**代数**：每次 `begin` / `end` / `cancel` 都 +1。
   *
   * 只靠 `runningRef` 挡不住"旧 drain 走完再把结果提交回来"——它挡的是并发进入，
   * 不挡**迟到提交**。drain 每轮开始时记下代数，提交前再比一次；变了就整轮丢弃。
   */
  const epochRef = useRef(0);
  /** 桥是否处于"可服务"状态。`end()`/`cancel()` 之后为假，`begin()` 之后为真。 */
  const liveRef = useRef(false);
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
      /*
       * 用 `for(;;)` 而不是 `do…while(queued)`：**旧轮次退出的那一刻**必须自己
       * 判断"还有没有活儿要接"。
       *
       * 竞态很具体：旧轮次的 poll 还在飞，新一轮 `end()` + `ensureLive()` 之后
       * 收到指针 —— 此时 `runningRef` 仍是 true，`handleProgressEvent` 只能记账。
       * 旧轮次回来发现自己 stale 就 `return` 的话，账记着却没人兑现：桥是活的、
       * 状态里也有待办，但一个请求都不发，用户看到的是"继续分析之后一动不动"，
       * 必须再点一次/再等一个事件才恢复。
       *
       * 所以 stale 时**不 return**：只要此刻确实有一条活着的桥并且它排过队，
       * 就回到循环顶部，用**最新**的 options/binding/epoch 为它服务。
       *
       * 同一件事还有另一半：旧轮次未必**正常返回**。取消/断网会让 `poll` 或
       * `receipt` 抛出去，控制流直接进外层 `catch`；只在那里记账的话，排队的新
       * 指针同样没人接。所以"这一轮跑完了"这件事由**一个**出口判断（下面的
       * `if (!queuedRef.current …) return`），正常返回、失效返回、抛异常三条路
       * 走同一段代码 —— 收尾分两种，就一定有一种是漏的。
       */
      for (;;) {
        queuedRef.current = false;
        const options = optionsRef.current;
        const binding = stateRef.current.binding;
        if (options === null || binding === null || stateRef.current.cancelled || !liveRef.current) {
          return;
        }
        const epoch = epochRef.current;
        const isStale = () => epochRef.current !== epoch || !liveRef.current;
        const ports = (
          dependenciesRef.current.createPorts ??
          ((options_, binding_) =>
            createSdkBridgeHostPorts(options_, binding_, dependenciesRef.current))
        )(options, binding);
        let roundFailed = false;
        try {
          const result = await runBridgeUntilIdle(stateRef.current, ports, {
            maxRounds: 8,
            isStale
          });
          if (!isStale()) {
            stateRef.current = result.state;
            for (const entry of result.outcomes) {
              if (entry.code !== null && entry.code !== undefined) {
                failureRef.current = entry.code;
                optionsRef.current?.onError?.(entry.code);
              }
            }
          }
        } catch {
          /*
           * 这一轮**抛出来**了：控制流从 `poll` 或 `receipt` 一路到这里（取消 →
           * AbortError、断网 → fetch reject）。
           *
           * 以前这里只是「记账、然后 finally 放开 runningRef」，于是排过队的那条
           * 新指针永远没人兑现：桥活着、状态里有待办、一个请求都不发，表现和
           * "继续分析之后一动不动"完全一样。正常的 stale 分支已经会接续队列，
           * 异常分支也必须走**同一套**出口 —— 收尾不能分两种。
           *
           * 但"接续"不等于"重试"：见下面的 return 条件。
           */
          roundFailed = true;
        }
        if (!queuedRef.current || stateRef.current.cancelled || !liveRef.current) {
          return;
        }
        if (roundFailed && !isStale()) {
          /*
           * 本轮**自己**失败，且没有被收尾/被取代：**不重试**。
           *
           * 回执可能已经在路上（发出去之后才断的网），本地根本不知道服务端收没
           * 收到；当场再跑一遍就是擅自重放一次"状态未知"的写读操作。留待下一条
           * 真指针：那会**先问服务端**（`pollRequested` 还在），按服务端给的清单
           * 再决定执行什么，而不是我们替它重试。
           */
          return;
        }
        // 正常的 stale 分支与"旧轮次异常退出"的 stale 分支都在这里接续。
      }
    } catch {
      // 兜底：提交/回调这一段自己出错也不能拖垮整轮规划。
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

  const pollPending = useCallback(async () => {
    if (!liveRef.current || stateRef.current.cancelled || !stateRef.current.binding) return;
    stateRef.current = { ...stateRef.current, pollRequested: true };
    await drain();
  }, [drain]);

  const begin = useCallback(async (options: SdkBridgeRunOptions): Promise<boolean> => {
    failureRef.current = null;
    // 新一代开始：让所有在途 drain 的提交作废（它们跑的还是上一轮的绑定）。
    epochRef.current += 1;
    liveRef.current = false;
    // 这一代自己的编号：`await liveContext()` 之后拿它核验"我还是当前代吗"。
    const epoch = epochRef.current;
    optionsRef.current = options;
    // 授权范围快照取自**与执行前复核同一个读取器**，避免调用方再传一份、
    // 两份迟早对不上。读不到现场或范围为空时，不能向调用方报告桥已就绪。
    let authorizedSheetNames: string[] = [];
    try {
      authorizedSheetNames = [...(await options.liveContext()).authorizedSheetNames];
    } catch {
      authorizedSheetNames = [];
    }
    /*
     * 跨过 await 之后必须重新核验轮次：初始化不是瞬间完成的，读宿主是异步的。
     *
     * 期间的 `cancel()` / `end()` 已经把这轮判死，紧接着到来的 `begin()` 又开了新
     * 一代 —— 迟到的初始化若照旧写绑定并点亮 `liveRef`，就会把**已经取消的桥重新
     * 激活**（`isLive()` 从 false 变回 true），后续指针被交给一条本该作废的桥；
     * 反过来，迟到的 A 还会覆盖已经建好的 B。两种情况都只写一半状态：要么全写，
     * 要么一个字段都不碰。
     */
    if (epochRef.current !== epoch) {
      return false;
    }
    if (authorizedSheetNames.length === 0) {
      optionsRef.current = null;
      queuedRef.current = false;
      failureRef.current = "无法读取已授权工作表，未建立读表桥。";
      return false;
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
    liveRef.current = true;
    queuedRef.current = false;
    return true;
  }, []);

  const isLive = useCallback(() => liveRef.current && stateRef.current.binding !== null, []);

  const ensureLive = useCallback(
    async (options: SdkBridgeRunOptions) => {
      const existing = stateRef.current.binding;
      if (
        liveRef.current &&
        existing !== null &&
        existing.taskId === options.taskId
      ) {
        // 已经在为同一个任务服务：只更新读取器，**不**重建状态（重建会清空
        // `delivered`，同一次调用会被再执行一遍）。
        optionsRef.current = { ...options, taskId: existing.taskId };
        return true;
      }
      /*
       * 返回值必须来自**这一次建立**本身。
       *
       * 以前写的是 `await begin(options); return liveRef.current && …` —— 初始化
       * 期间被取消时，迟到的那次 `begin` 会把 `liveRef` 点亮，这里就把"别的轮次
       * 的活桥"当成自己的成功回报，调用方于是认为续跑有桥可用。
       */
      return begin(options);
    },
    [begin]
  );

  const cancel = useCallback(() => {
    epochRef.current += 1;
    liveRef.current = false;
    stateRef.current = cancelBridge(stateRef.current);
    queuedRef.current = false;
  }, []);

  const end = useCallback(() => {
    // 停轮询但**保留绑定**：本轮的结果卡要在会话关闭之后按引用取值，
    // 服务端校验的是授权归属（claim）而不是活跃会话。
    epochRef.current += 1;
    liveRef.current = false;
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
  const pinSourceFingerprint = useCallback((taskId: string, live: SdkBridgeLiveContext) => {
    const bound = stateRef.current.binding;
    if (!liveRef.current || !bound || bound.taskId !== taskId ||
        bound.workbookIdentity !== live.workbookIdentity || !live.sourceFingerprint ||
        bound.authorizedSheetNames.length !== live.authorizedSheetNames.length ||
        bound.authorizedSheetNames.some((name) => !live.authorizedSheetNames.includes(name))) return false;
    if (bound.sourceFingerprint !== null) return bound.sourceFingerprint === live.sourceFingerprint;
    stateRef.current = { ...stateRef.current, binding: { ...bound, sourceFingerprint: live.sourceFingerprint } };
    if (optionsRef.current) optionsRef.current = { ...optionsRef.current, sourceFingerprint: live.sourceFingerprint };
    return true;
  }, []);

  useEffect(() => () => cancel(), [cancel]);

  return {
    begin,
    ensureLive,
    pollPending,
    handleProgressEvent,
    cancel,
    end,
    binding,
    isLive,
    pinSourceFingerprint,
    failure: () => failureRef.current
  };
}
