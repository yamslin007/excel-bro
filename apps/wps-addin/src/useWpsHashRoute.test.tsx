import { act, type ReactElement } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";
import { useWpsHashRoute } from "./useWpsHashRoute";

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let mounted: Array<{ unmount: () => void }> = [];

function Probe({ onRender }: { onRender: (hash: string) => void }) {
  const { hash, navigate } = useWpsHashRoute();
  onRender(hash);
  return <button onClick={() => navigate("/dev-format-probe")}>go</button>;
}

async function mountProbe(onRender: (hash: string) => void): Promise<HTMLElement> {
  const host = document.createElement("div");
  document.body.appendChild(host);
  const root = createRoot(host);
  await act(async () => { root.render(<Probe onRender={onRender} />); });
  mounted.push({ unmount: () => root.unmount() });
  return host;
}

afterEach(async () => {
  await act(async () => {
    for (const entry of mounted.splice(0)) entry.unmount();
    document.body.innerHTML = "";
  });
  // 清理 hash，避免污染后续用例。
  if (window.location.hash) {
    window.location.hash = "";
  }
});

describe("useWpsHashRoute", () => {
  it("初始读取当前 hash（空串归一为主面板路由）", async () => {
    window.location.hash = "";
    const seen: string[] = [];
    await mountProbe((hash) => seen.push(hash));
    expect(seen[0]).toBe("/taskpane");
  });

  it("navigate 写 hash 并触发状态更新", async () => {
    window.location.hash = "";
    const seen: string[] = [];
    const host = await mountProbe((hash) => seen.push(hash));
    const button = host.querySelector("button")!;
    await act(async () => { button.click(); });
    // 点击后 hash 状态应更新到 /dev-format-probe（渲染两次：初始 + 点击后）。
    expect(seen).toContain("/dev-format-probe");
  });

  it("外部 hashchange 能更新状态", async () => {
    window.location.hash = "";
    const seen: string[] = [];
    await mountProbe((hash) => seen.push(hash));
    await act(async () => {
      window.location.hash = "/wide-view-poc";
      window.dispatchEvent(new HashChangeEvent("hashchange"));
    });
    expect(seen).toContain("/wide-view-poc");
  });

  it("卸载时清理 hashchange 监听", async () => {
    window.location.hash = "";
    const seen: string[] = [];
    const host = await mountProbe((hash) => seen.push(hash));
    // 卸载前记一次监听器数量变化：卸载后 dispatch hashchange 不应再触发 onRender。
    await act(async () => {
      for (const entry of mounted.splice(0)) entry.unmount();
      document.body.innerHTML = "";
    });
    const before = seen.length;
    await act(async () => {
      window.dispatchEvent(new HashChangeEvent("hashchange"));
    });
    // 卸载后 hashchange 不应新增渲染。
    expect(seen.length).toBe(before);
    // 静默清理 host 引用。
    void host;
  });
});
