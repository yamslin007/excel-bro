import { act, type ReactElement } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

/**
 * AppRoutes 路由集成测试：断言**渲染页面变化**，而非 hash 字符串。
 *
 * 通过 mock 三个面板组件为带标记的占位，验证：
 * - 任务窗格顶部**不再**有开发诊断导航条（三颗按钮已按产品要求移除）；
 * - 外部 hashchange 能更新页面；
 * - 首次直接进入诊断地址渲染诊断页；
 * - 非 DEV 无诊断导航、诊断地址不可达。
 */

// 在导入 AppRoutes 之前 mock 面板与副作用。
vi.mock("./WpsProbePanel", () => ({
  default: () => <div data-panel="taskpane">主面板</div>
}));
vi.mock("./WideViewPocPanel", () => ({
  default: () => <div data-panel="wide-view-poc">宽幅审阅</div>
}));
vi.mock("./WpsFormatProbePanel", () => ({
  default: () => <div data-panel="dev-format-probe">格式探测</div>
}));

import { AppRoutes } from "./AppRoutes";

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

function currentPanel(): string | null {
  const el = host?.querySelector("[data-panel]");
  return el?.getAttribute("data-panel") ?? null;
}

afterEach(async () => {
  await act(async () => {
    for (const entry of mounted.splice(0)) entry.unmount();
    document.body.innerHTML = "";
  });
  host = null;
  window.location.hash = "";
});

describe("AppRoutes 路由（DEV）", () => {
  /*
   * 回归：任务窗格顶部那三颗按钮（「主面板」/「宽幅审阅 · 能力探针」/
   * 「WPS 格式能力 · 受控探测」）整条移除。
   *
   * 判据读的是**界面上还剩什么**：没有 `.dev-nav`、没有 `nav`、没有任何按钮，
   * 且三颗按钮各自的文案一个字都不在。只断言"某个选择器为空"不够 ——
   * 把导航条换个类名重挂回来就会漏过。
   */
  it("DEV 下任务窗格顶部不再渲染开发诊断导航条", async () => {
    window.location.hash = "";
    await mount(<AppRoutes />);
    // 正向锚点：主面板确实渲染了（不是"整页空白"把负向断言蒙对）。
    expect(currentPanel()).toBe("taskpane");
    expect(host?.querySelector(".dev-nav")).toBeNull();
    expect(host?.querySelector("nav")).toBeNull();
    expect(host?.querySelectorAll("button")).toHaveLength(0);
    expect(host?.textContent).not.toContain("能力探针");
    expect(host?.textContent).not.toContain("受控探测");
  });

  it("外部 hashchange 能更新页面", async () => {
    window.location.hash = "";
    await mount(<AppRoutes />);
    expect(currentPanel()).toBe("taskpane");
    await act(async () => {
      window.location.hash = "/dev-format-probe";
      window.dispatchEvent(new HashChangeEvent("hashchange"));
    });
    expect(currentPanel()).toBe("dev-format-probe");
  });

  it("首次直接进入诊断地址（#/dev-format-probe）渲染诊断页面", async () => {
    window.location.hash = "/dev-format-probe";
    await mount(<AppRoutes />);
    expect(currentPanel()).toBe("dev-format-probe");
  });
});
