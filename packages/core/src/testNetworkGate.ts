/**
 * 测试网络闸门：**未替身的网络出口一律让用例失败**。
 *
 * 存在理由（2026-10-27 诊断期间发现的真实事故）：`WpsAgentPanel.test.tsx` 只替掉了
 * `api.ts` 里的一部分函数，其余仍走真实实现；而 `api.ts` 的基地址是
 * `import.meta.env.VITE_API_BASE_URL ?? "http://127.0.0.1:8765"`。于是"跑一遍全量测试"
 * 就在用户正在使用的那个实例上留下了
 * `POST /api/wps/executions/<测试合成的 executionId>/host-writes` 的请求。
 *
 * 这不是环境噪声。它同时破坏三件事：
 *
 * 1. **用户实例被写**：测试合成的执行身份去碰真实账本，等于让测试污染现场；
 * 2. **结论不可信**：有没有请求到 8765 取决于"用户此刻的服务是否在跑"，于是同一份
 *    代码今天绿、明天红，和产品行为无关；
 * 3. **缺口被藏住**：`void report(...)` 这一类 best-effort 调用即使失败也被吞掉，
 *    于是"这里根本没替身"永远不会以失败的形式报出来。
 *
 * 所以闸门的规则只有两条，且**没有允许列表**（允许某个真实地址就是给下一次泄漏开口子）：
 *
 * 1. 任何未被 `vi.stubGlobal("fetch", ...)` 或明确的 HTTP 替身接管的请求，**立刻抛错**，
 *    错误原文带"网络闸门"，失败时一眼看出是谁在发请求；
 * 2. 即使调用方把错误 `catch` 掉了，`setupFiles` 里的 `afterEach` 也会因为落下了
 *    未处理的请求而让这条用例失败 —— 安静地失败是最坏的一种"通过"。
 *
 * 本模块只被两端测试的 `setupFiles` 引用，不进任何生产包。
 */

/** 一次未处理的网络尝试。`origin` 是发起处调用栈，用来定位是谁在偷偷上网。 */
export interface StrayNetworkAttempt {
  channel: "fetch" | "xhr" | "websocket" | "beacon";
  method: string;
  target: string;
  origin: string;
}

const attempts: StrayNetworkAttempt[] = [];

const CALL_STACK_LIMIT = 1200;

function originOf(): string {
  const stack = new Error("network-gate").stack ?? "";
  // 去掉头两行（Error 本身与闸门内部帧），只留调用方。
  return stack.split("\n").slice(2).join("\n").slice(0, CALL_STACK_LIMIT);
}

function record(channel: StrayNetworkAttempt["channel"], method: string, target: string): StrayNetworkAttempt {
  const attempt: StrayNetworkAttempt = { channel, method, target, origin: originOf() };
  attempts.push(attempt);
  return attempt;
}

function targetOf(input: unknown): string {
  if (typeof input === "string") return input;
  if (typeof URL !== "undefined" && input instanceof URL) return input.toString();
  const candidate = input as { url?: unknown } | null;
  if (candidate && typeof candidate.url === "string") return candidate.url;
  return String(input);
}

function messageOf(attempt: StrayNetworkAttempt): string {
  return (
    `测试网络闸门：拒绝未替身的请求 ${attempt.method} ${attempt.target}\n` +
    "所有网络调用必须有替身（`vi.mock(\"…/api\")` 或 `vi.stubGlobal(\"fetch\", …)`），" +
    "或指向隔离测试实例；测试不得访问用户正在使用的本机服务。"
  );
}

/** 当前用例里已经发生的未处理请求（`setupFiles` 负责按用例清空）。 */
export function strayNetworkAttempts(): readonly StrayNetworkAttempt[] {
  return attempts;
}

export function resetStrayNetworkAttempts(): void {
  attempts.length = 0;
}

export function describeStrayNetworkAttempts(list: readonly StrayNetworkAttempt[] = attempts): string {
  if (!list.length) return "没有任何未处理的网络请求。";
  return [
    `测试期间出现了 ${list.length} 个未替身的网络请求（必须替身或指向隔离实例）：`,
    ...list.map((item, index) => `${index + 1}. [${item.channel}] ${item.method} ${item.target}`)
  ].join("\n");
}

type Scope = typeof globalThis & {
  fetch: typeof fetch;
  XMLHttpRequest?: { prototype: XMLHttpRequest };
  WebSocket?: { new (url: string | URL, protocols?: string | string[]): unknown; prototype: unknown };
};

let installed = false;
let gateFetch: typeof fetch | null = null;

/**
 * 本用例是否自己装了网络替身。
 *
 * 判据只有一条：`globalThis.fetch` 已经不是闸门那个函数（用例 `vi.stubGlobal("fetch", …)`
 * 换过一次）。**没有**其它白名单，也不看 URL。
 */
export function networkDoubleInstalled(): boolean {
  return gateFetch !== null && globalThis.fetch !== gateFetch;
}

/**
 * 把一条**真实网络接缝**包成"只有本用例装了网络替身才真的调用"。
 *
 * 用在那些故意要跑真实实现的接缝上（桥轮询、暂停请求、计划校验）：它们只有在用例
 * 自己 `vi.stubGlobal("fetch", …)` 驱动时才有意义；否则每一次调用都只可能打到用户那台
 * `127.0.0.1:8765`。没有替身时按"连不上"失败 —— 与之前"看着像连不上"的偶然环境同义，
 * 但不再有字节出进程，也不再取决于用户的服务有没有在跑。
 */
export function guardNetworkSeam<A extends unknown[], R>(
  real: (...args: A) => Promise<R>
): (...args: A) => Promise<R> {
  return (...args) => {
    if (!networkDoubleInstalled()) {
      return Promise.reject(
        new TypeError(
          "测试网络闸门：这条接缝没有网络替身（请在本用例内 vi.stubGlobal(\"fetch\", …) 或替掉该函数）"
        )
      );
    }
    return real(...args);
  };
}

/**
 * 装上闸门（幂等）。`vi.stubGlobal` 之后仍然是替身的天下：它保存的是**装闸门之后**的
 * `fetch`，`vi.unstubAllGlobals()` 会把闸门放回来，而不是放回真实 `fetch`。
 */
export function installTestNetworkGate(): void {
  if (installed) return;
  installed = true;
  const scope = globalThis as Scope;

  scope.fetch = ((input: unknown, init?: RequestInit) => {
    const attempt = record("fetch", String(init?.method ?? "GET").toUpperCase(), targetOf(input));
    return Promise.reject(new Error(messageOf(attempt)));
  }) as typeof fetch;
  gateFetch = scope.fetch;

  const xhrPrototype = scope.XMLHttpRequest?.prototype;
  if (xhrPrototype) {
    /*
     * `open` 只**记账**、不转交真实实现：转交之后这个 XHR 就有了真实的连接状态，
     * 而真正危险的是 `send`。记账挂到实例上，好让 `send` 抛出带目标的那一条。
     */
    const noted = new WeakMap<object, StrayNetworkAttempt>();
    xhrPrototype.open = function openThroughGate(this: XMLHttpRequest, ...args: unknown[]) {
      noted.set(this, record("xhr", String(args[0] ?? "GET").toUpperCase(), targetOf(args[1])));
      return undefined;
    } as typeof xhrPrototype.open;
    xhrPrototype.send = function sendThroughGate(this: XMLHttpRequest) {
      const attempt = noted.get(this) ??
        record("xhr", "GET", "XMLHttpRequest（未调用 open）");
      throw new Error(messageOf(attempt));
    } as typeof xhrPrototype.send;
  }

  if (scope.WebSocket) {
    scope.WebSocket = function GatedWebSocket(this: unknown, url: string | URL) {
      const attempt = record("websocket", "CONNECT", targetOf(url));
      throw new Error(messageOf(attempt));
    } as unknown as Scope["WebSocket"];
    (scope.WebSocket as unknown as { prototype: unknown }).prototype = {};
  }

  const navigatorLike = (globalThis as { navigator?: { sendBeacon?: (url: string, data?: unknown) => boolean } })
    .navigator;
  if (navigatorLike && typeof navigatorLike.sendBeacon === "function") {
    navigatorLike.sendBeacon = (url: string) => {
      const attempt = record("beacon", "BEACON", targetOf(url));
      throw new Error(messageOf(attempt));
    };
  }
}
