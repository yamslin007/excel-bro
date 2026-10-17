import { act, type ReactElement } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

/**
 * 非 DEV 门控测试：生产环境（IS_DEV=false）下，
 * - 任务窗格无诊断导航条（该导航条已整条移除，见 `AppRoutes`）；
 * - 手改 hash 到诊断地址也打不开诊断页（兜底到主面板）。
 *
 * `import.meta.env.DEV` 是编译期内联常量（测试环境恒 true），无法用 stubEnv 模拟，
 * 因此这里 mock `isDev` 模块为 false，验证门控**逻辑**。
 */

vi.mock("./WpsProbePanel", () => ({
  default: () => <div data-panel="taskpane">主面板</div>
}));
vi.mock("./WideViewPocPanel", () => ({
  default: () => <div data-panel="wide-view-poc">宽幅审阅</div>
}));
vi.mock("./WpsFormatProbePanel", () => ({
  default: () => <div data-panel="dev-format-probe">格式探测</div>
}));

// 覆盖 isDev 模块为恒 false，AppRoutes 从该模块 import，mock 真正生效。
vi.mock("./isDev", () => ({
  isDev: () => false
}));

import { AppRoutes } from "./AppRoutes";
import { isDev } from "./isDev";

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let mounted: Array<{ unmount: () => void }> = [];
let host: HTMLElement | null = null;

async function mount(element: ReactElement): Promise<HTMLElement> {
  host = document.createElement("div");
  document.body.appendChild(host);
  const root = createRoot(host);
  await act(async () => { root.render(element); });
  mounted.push({ unmount: () => root.unmount() });
  return host;
}

afterEach(async () => {
  await act(async () => {
    for (const entry of mounted.splice(0)) entry.unmount();
    document.body.innerHTML = "";
  });
  host = null;
  window.location.hash = "";
});

describe("非 DEV 门控", () => {
  it("生产环境 isDev 为 false", () => {
    expect(isDev()).toBe(false);
  });

  it("生产环境任务窗格无诊断导航条", async () => {
    window.location.hash = "";
    await mount(<AppRoutes />);
    expect(host?.querySelector(".dev-nav")).toBeNull();
    expect(host?.querySelector("nav")).toBeNull();
    expect(host?.querySelector("[data-panel]")?.getAttribute("data-panel")).toBe("taskpane");
  });

  it("生产环境手改 hash 到诊断地址也打不开诊断页（兜底到主面板）", async () => {
    window.location.hash = "/dev-format-probe";
    await mount(<AppRoutes />);
    expect(host?.querySelector("[data-panel]")?.getAttribute("data-panel")).toBe("taskpane");
    expect(host?.querySelector(".dev-nav")).toBeNull();
  });
});
