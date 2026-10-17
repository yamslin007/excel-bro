import { useSyncExternalStore } from "react";

/**
 * WPS 任务窗格的哈希路由（响应式，共享单一 store）。
 *
 * 根因修复：之前「导航条」与「页面选择器」各自调用 hook、各自持有独立 state，
 * 改 hash 只更新前者自己的 state，页面选择组件读不到变化。这里改用
 * `useSyncExternalStore` 订阅一个**共享 store**：以 `window.location.hash`
 * 为唯一真源，`navigate` 写 hash 后通知所有订阅者重读。所有组件读同一份 hash，
 * 页面选择与（历史上那条）导航高亮不再各自为政。
 *
 * 任务窗格顶部那条导航条已整条移除（见 `AppRoutes`），但这个共享 store 仍然要留着：
 * DEV 诊断页靠**地址栏 hash** 进入，页面选择必须响应外部 `hashchange`。
 */

/** 归一化 hash：去掉开头的 `#`，空串/"/" 表示主面板（taskpane）。 */
function normalizeHash(raw: string): string {
  const value = raw.startsWith("#") ? raw.slice(1) : raw;
  if (value === "" || value === "/") return "/taskpane";
  return value;
}

function readCurrentHash(): string {
  return normalizeHash(typeof window !== "undefined" ? window.location.hash : "");
}

const listeners = new Set<() => void>();

function subscribe(listener: () => void): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

/** 快照：直接读 window.location.hash（唯一真源），归一后返回。 */
function getSnapshot(): string {
  return readCurrentHash();
}

/** hashchange 时通知所有订阅者重读。 */
function onHashChange() {
  for (const listener of listeners) listener();
}

// 只在浏览器环境订阅一次（模块加载时）。
if (typeof window !== "undefined") {
  window.addEventListener("hashchange", onHashChange);
}

export type WpsHashRoute = {
  /** 归一后的 hash（如 "/taskpane"、"/wide-view-poc"、"/dev-format-probe"）。 */
  hash: string;
  /** 导航到指定 hash。写 window.location.hash，触发 hashchange → 所有订阅者更新。 */
  navigate: (nextHash: string) => void;
};

export function useWpsHashRoute(): WpsHashRoute {
  const hash = useSyncExternalStore(subscribe, getSnapshot, getSnapshot);

  const navigate = (nextHash: string) => {
    const normalized = normalizeHash(nextHash);
    // 写 window.location.hash：真实浏览器会触发 hashchange → onHashChange 通知订阅者。
    window.location.hash = normalized;
    // jsdom 下设置 location.hash 不自动触发 hashchange，这里手动补一次通知（幂等）。
    onHashChange();
  };

  return { hash, navigate };
}
