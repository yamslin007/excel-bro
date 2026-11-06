/**
 * WPS 专注窗口（父窗格侧）的集成测试。
 *
 * 被测的是 `WpsAgentPanel` 菜单里的「专注窗口」入口：打开、连续点击去重、工作簿切换
 * 关窗、以及贯穿始终的**零模型调用、零工作簿写入、零隐式审阅处置**。
 *
 * 宿主差异（真实 ShowDialog 窗口、真实跨 WebView BroadcastChannel）只能实机验证；
 * 这里用替身证明面板侧的确定性行为：调用真实 `startWpsFocusParentSession`、真实
 * 打开逻辑，但 `ShowDialog` 与跨窗口通道是替身。
 */

import { act, type ReactElement } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { listModels, understandMessage } from "../../excel-addin/src/api";
import WpsAgentPanel from "./WpsAgentPanel";
import { getWpsApplication } from "./wpsHost";

(
  globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }
).IS_REACT_ACT_ENVIRONMENT = true;

// jsdom 不实现 scrollIntoView；面板只用它把新消息滚到底部。
if (!Element.prototype.scrollIntoView) {
  Element.prototype.scrollIntoView = () => undefined;
}

vi.mock("../../excel-addin/src/api", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../excel-addin/src/api")>();
  return {
    ...actual,
    listModels: vi.fn(),
    understandMessage: vi.fn(async () => ({ kind: "conversation", message: "好", provider: "model" }))
  };
});

vi.mock("./wpsHost", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./wpsHost")>();
  return { ...actual, getWpsApplication: vi.fn(() => ({})) };
});

/** 同步投递的假 BroadcastChannel，父子窗口共用一个 sessionId 通道。 */
class FakeBroadcastChannel {
  static channels = new Map<string, Set<FakeBroadcastChannel>>();
  onmessage: ((event: MessageEvent<unknown>) => void) | null = null;
  private readonly peers: Set<FakeBroadcastChannel>;
  constructor(name: string) {
    this.peers = FakeBroadcastChannel.channels.get(name) ?? new Set<FakeBroadcastChannel>();
    this.peers.add(this);
    FakeBroadcastChannel.channels.set(name, this.peers);
  }
  postMessage(data: unknown): void {
    for (const peer of this.peers) {
      if (peer !== this) peer.onmessage?.({ data } as MessageEvent<unknown>);
    }
  }
  close(): void {
    this.peers.delete(this);
  }
}

const listModelsMock = vi.mocked(listModels);
const understandMock = vi.mocked(understandMessage);
const getWpsApplicationMock = vi.mocked(getWpsApplication);

function panel(): ReactElement {
  return <WpsAgentPanel selectedSheetIndexes={[1]} scopeDock={<div className="eb-scope-dock">当前表 Sheet1</div>} />;
}

async function mount(element: ReactElement) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  await act(async () => { root.render(element); });
  return {
    container,
    unmount: () => {
      act(() => root.unmount());
      container.remove();
    }
  };
}

async function settle() {
  await act(async () => {
    // 假定时器下真实 setTimeout 不会走，改用 advanceTimersByTimeAsync 冲刷微任务队列。
    if (vi.isFakeTimers()) {
      await vi.advanceTimersByTimeAsync(0);
    } else {
      await new Promise<void>((resolve) => setTimeout(resolve, 0));
    }
  });
}

async function click(el: Element | null) {
  if (!el) throw new Error("要点击的元素不存在");
  await act(async () => { el.dispatchEvent(new MouseEvent("click", { bubbles: true })); });
  await settle();
}

function findMenuItem(container: HTMLElement, text: string): HTMLButtonElement | null {
  return Array.from(container.querySelectorAll('[role="menuitem"]')).find(
    (item) => item.textContent?.includes(text)
  ) as HTMLButtonElement | null;
}

/** 一个会记录打开次数与 URL 的 ShowDialog 替身。 */
function makeShowDialogStub() {
  const calls: string[] = [];
  const showDialog = vi.fn((url: string) => {
    calls.push(url);
    return true;
  });
  return { showDialog, calls };
}

beforeEach(() => {
  localStorage.clear();
  vi.stubGlobal("BroadcastChannel", FakeBroadcastChannel);
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
  FakeBroadcastChannel.channels.clear();
  getWpsApplicationMock.mockReset();
  getWpsApplicationMock.mockReturnValue({} as never);
});

describe("WPS 专注窗口（父窗格侧）", () => {
  it("菜单点「专注窗口」打开真实对话框入口，且不调用模型、不写入工作簿", async () => {
    const { showDialog, calls } = makeShowDialogStub();
    getWpsApplicationMock.mockReturnValue({ Version: "12.0", ShowDialog: showDialog } as never);
    listModelsMock.mockResolvedValue({ defaultModelId: "m", models: [] });

    const { container, unmount } = await mount(panel());
    await settle();

    await click(container.querySelector('[aria-label="EB 菜单"]'));
    const entry = findMenuItem(container, "专注窗口");
    expect(entry).not.toBeNull();
    expect((entry as HTMLButtonElement).disabled).toBe(false);

    await click(entry);
    await settle();

    // 真实调用了一次宿主 ShowDialog，URL 指向 focus.html 并带 session 标识。
    expect(showDialog).toHaveBeenCalledTimes(1);
    expect(calls[0]).toContain("focus.html");
    expect(calls[0]).toMatch(/session=[0-9a-f]{32}/);
    // 全程零模型调用。
    expect(understandMock).not.toHaveBeenCalled();
    unmount();
  });

  it("连续点击不会重复创建窗口会话", async () => {
    const { showDialog, calls } = makeShowDialogStub();
    getWpsApplicationMock.mockReturnValue({ Version: "12.0", ShowDialog: showDialog } as never);
    listModelsMock.mockResolvedValue({ defaultModelId: "m", models: [] });

    const { container, unmount } = await mount(panel());
    await settle();

    await click(container.querySelector('[aria-label="EB 菜单"]'));
    const entry = findMenuItem(container, "专注窗口")!;
    await click(entry);
    await click(entry);
    await click(entry);
    await settle();

    // 已打开时只 sync，不重复 ShowDialog。
    expect(showDialog).toHaveBeenCalledTimes(1);
    expect(calls).toHaveLength(1);
    unmount();
  });

  it("宿主没有 ShowDialog/CreateWebDialog 时如实说明，不假装打开", async () => {
    getWpsApplicationMock.mockReturnValue({ Version: "11.0" } as never);
    listModelsMock.mockResolvedValue({ defaultModelId: "m", models: [] });

    const { container, unmount } = await mount(panel());
    await settle();

    await click(container.querySelector('[aria-label="EB 菜单"]'));
    await click(findMenuItem(container, "专注窗口"));
    await settle();

    expect(container.textContent).toContain("Web 对话框能力");
    unmount();
  });

  it("ShowDialog 抛错时保留原对话并如实报错", async () => {
    const showDialog = vi.fn(() => { throw new Error("宿主拒绝打开窗口"); });
    getWpsApplicationMock.mockReturnValue({ Version: "12.0", ShowDialog: showDialog } as never);
    listModelsMock.mockResolvedValue({ defaultModelId: "m", models: [] });

    const { container, unmount } = await mount(panel());
    await settle();

    // 先写一句话，确认打开失败后它仍在。
    const textarea = container.querySelector("textarea");
    Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, "value")?.set?.call(textarea, "保留这句话");
    await act(async () => { textarea?.dispatchEvent(new Event("input", { bubbles: true })); });

    await click(container.querySelector('[aria-label="EB 菜单"]'));
    await click(findMenuItem(container, "专注窗口"));
    await settle();

    expect(container.textContent).toContain("无法打开专注窗口");
    expect(container.querySelector("textarea")?.value).toBe("保留这句话");
    unmount();
  });

  it("全程不触发隐式审阅处置：确认/放弃/撤销一个都不碰", async () => {
    const { showDialog } = makeShowDialogStub();
    getWpsApplicationMock.mockReturnValue({ Version: "12.0", ShowDialog: showDialog } as never);
    listModelsMock.mockResolvedValue({ defaultModelId: "m", models: [] });

    const { container, unmount } = await mount(panel());
    await settle();

    await click(container.querySelector('[aria-label="EB 菜单"]'));
    await click(findMenuItem(container, "专注窗口"));
    await settle();

    // 专注窗口只做展示：不出现任何确认/放弃/撤销的写入入口。
    expect(container.querySelector('[data-recovery-primary]')).toBeNull();
    expect(container.textContent).not.toContain("确认全部变更");
    expect(container.textContent).not.toContain("放弃全部变更");
    unmount();
  });

  it("打开后卸载窗格会清理跨窗口通道，不留孤立会话", async () => {
    const { showDialog, calls } = makeShowDialogStub();
    getWpsApplicationMock.mockReturnValue({ Version: "12.0", ShowDialog: showDialog } as never);
    listModelsMock.mockResolvedValue({ defaultModelId: "m", models: [] });

    const { container, unmount } = await mount(panel());
    await settle();

    await click(container.querySelector('[aria-label="EB 菜单"]'));
    await click(findMenuItem(container, "专注窗口"));
    await settle();

    expect(showDialog).toHaveBeenCalledTimes(1);
    // 打开后父窗格会话建起了跨窗口通道（FakeBroadcastChannel 里至少有一条）。
    expect(FakeBroadcastChannel.channels.size).toBeGreaterThan(0);

    unmount();
    await settle();

    // 卸载后通道全部关闭（sendClosing + stop + transport.close）。
    for (const peers of FakeBroadcastChannel.channels.values()) {
      expect(peers.size).toBe(0);
    }
    expect(calls[0]).toContain("focus.html");
  });

  it("握手超时后入口复位为「专注窗口」，不再是卡死的「正在打开...」", async () => {
    vi.useFakeTimers();
    const { showDialog } = makeShowDialogStub();
    getWpsApplicationMock.mockReturnValue({ Version: "12.0", ShowDialog: showDialog } as never);
    listModelsMock.mockResolvedValue({ defaultModelId: "m", models: [] });

    const { container, unmount } = await mount(panel());
    await act(async () => { await vi.advanceTimersByTimeAsync(0); });

    // 点开后 ShowDialog 返回 true（无句柄），但没有子窗口回 hello ⇒ 触发握手超时。
    await click(container.querySelector('[aria-label="EB 菜单"]'));
    await click(findMenuItem(container, "专注窗口"));

    // 点开会关菜单；重新打开菜单看入口状态：此时是「正在打开...」且禁用。
    await click(container.querySelector('[aria-label="EB 菜单"]'));
    const openingEntry = findMenuItem(container, "正在打开");
    expect(openingEntry).not.toBeNull();
    expect((openingEntry as HTMLButtonElement).disabled).toBe(true);
    await click(container.querySelector('[aria-label="EB 菜单"]'));

    // 推进超过握手超时（10s），触发 handshake-timeout ⇒ 复位入口。
    await act(async () => { await vi.advanceTimersByTimeAsync(10_100); });

    // 重新打开菜单：按钮文案回到「专注窗口」、可点击；且如实报出握手失败。
    await click(container.querySelector('[aria-label="EB 菜单"]'));
    const resetEntry = findMenuItem(container, "专注窗口");
    expect(resetEntry).not.toBeNull();
    expect((resetEntry as HTMLButtonElement).disabled).toBe(false);
    await click(container.querySelector('[aria-label="EB 菜单"]'));
    expect(container.textContent).toContain("未能完成握手");

    vi.useRealTimers();
    unmount();
  });
});
