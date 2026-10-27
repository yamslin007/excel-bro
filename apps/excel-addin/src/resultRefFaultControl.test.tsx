// @vitest-environment jsdom
/**
 * A1-E1 验收控件：**两端共用的那一份**（`packages/core/src/ResultRefFaultControl.ts`）。
 *
 * 这个控件的全部作用范围就是"把用户的显式启用送到服务端"，因此这里断言的是它的
 * **通道边界**：
 *
 * - 非开发模式整块不渲染（不是渲染出来再禁用）；
 * - 挂载时不打请求（开发入口不该在每次开窗格时扰一次服务端）；
 * - 启用请求带的就是那三个字段，不多不少（多一个 taskId/resultRef 就等于把"让哪
 *   个引用失效"交给了调用方）；
 * - 它只经由传进来的 `load/arm/cancel` 与外界交互 —— 没有第二个通道，也就不可能
 *   顺手写宿主。
 */
import { cleanup, render } from "@testing-library/react";
import { act } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { ResultRefFaultControl } from "@excel-bro/core/ResultRefFaultControl";
import type { ResultRefFaultStatus, ResultRefFaultTarget } from "@excel-bro/core/resultRefFault";

// 本仓库不用全局 act 环境开关：显式开，否则 React 19 的更新不入 act 批次，
// 断言会在"界面还没提交"时抓到旧值（与其它 jsdom 用例同一约定）。
(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const TARGET: ResultRefFaultTarget = {
  hostId: "excel-office",
  workbookName: "测试.xlsx",
  workbookIdentity: "url:https://example/测试.xlsx"
};

function statusOf(
  phase: ResultRefFaultStatus["phase"],
  overrides: Partial<ResultRefFaultStatus> = {}
): ResultRefFaultStatus {
  return {
    enabled: true,
    phase,
    expiresInSeconds: phase === "armed" ? 590 : 0,
    taskBound: phase === "bound",
    refBound: phase === "bound",
    ...overrides
  };
}

type View = ReturnType<typeof render>;

/**
 * 展开 `<details>` 并把这次展开**触发的那次 `load()` 结算完**。
 *
 * jsdom 的 `click` 不会自己把 `open` 置上、也不派发 `toggle`，所以显式置位再派发
 * （组件的 `onToggle` 只读 `currentTarget.open`，与真实浏览器里点开那一行同义）。
 * `load` 是异步的：在同一个 act 里等它结算，否则它的 `setStatus` 落在 act 之外。
 */
async function expand(view: View) {
  const details = view.container.querySelector("details") as HTMLDetailsElement;
  details.open = true;
  await act(async () => {
    details.dispatchEvent(new Event("toggle"));
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
}

/**
 * 按下一个按钮，并等到它引发的那次请求（arm / cancel）**真的结算并提交到界面**。
 *
 * 只等"请求已发出"就断言运行态会偶发抓空 —— 那时 React 还没提交。
 */
async function press(element: HTMLElement) {
  await act(async () => {
    element.click();
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
}

function setup(options: {
  visible?: boolean;
  disabled?: boolean;
  target?: ResultRefFaultTarget | null;
  load?: () => Promise<ResultRefFaultStatus>;
  arm?: (target: ResultRefFaultTarget) => Promise<ResultRefFaultStatus>;
  cancel?: () => Promise<ResultRefFaultStatus>;
} = {}) {
  const load = vi.fn(options.load ?? (async () => statusOf("idle")));
  const arm = vi.fn(options.arm ?? (async () => statusOf("armed")));
  const cancel = vi.fn(options.cancel ?? (async () => statusOf("idle")));
  const view = render(
    <ResultRefFaultControl
      visible={options.visible ?? true}
      disabled={options.disabled ?? false}
      target={options.target === undefined ? TARGET : options.target}
      load={load}
      arm={arm}
      cancel={cancel}
    />
  );
  return { view, load, arm, cancel };
}

const armButton = (view: View) =>
  view.getByRole("button", { name: "启用下一次引用过期" }) as HTMLButtonElement;
const cancelButton = (view: View) =>
  view.getByRole("button", { name: "取消验收" }) as HTMLButtonElement;

afterEach(cleanup);

describe("可见性", () => {
  it("非开发模式整块不渲染 —— 不是渲染出来再禁用", () => {
    const { view, load } = setup({ visible: false });
    expect(view.container.querySelector(".result-ref-fault-control")).toBeNull();
    expect(view.container.textContent).toBe("");
    // 连状态都不该去读：看不见的入口不该产生请求。
    expect(load).not.toHaveBeenCalled();
  });

  it("开发模式才出现标题与两个按钮", () => {
    const { view } = setup();
    expect(view.container.querySelector("details.result-ref-fault-control")).not.toBeNull();
    expect(view.getByText("结果引用过期（开发专用）")).toBeTruthy();
    expect(armButton(view)).toBeTruthy();
    expect(cancelButton(view)).toBeTruthy();
  });

  it("收起时不播报状态 —— 折叠着的入口没有状态要对用户说", () => {
    /*
     * 面板里"这一轮还有没有进度提示"是按 `role="status"` 判的。一个常驻的活区会让
     * 那条判据失去意义（既有用例就是这么红的）。所以状态只在展开后出现。
     */
    const { view } = setup();
    expect(view.queryByRole("status")).toBeNull();
    expect(view.queryByRole("alert")).toBeNull();
  });
});

describe("读状态只在展开时发生一次", () => {
  it("挂载时不打请求，展开才读", async () => {
    const { view, load } = setup();
    expect(load).not.toHaveBeenCalled();
    await expand(view);
    expect(load).toHaveBeenCalledTimes(1);
  });

  it("展开后显示的是共享判据的文案（未启用）", async () => {
    const { view } = setup();
    await expand(view);
    expect(view.getByRole("status").textContent).toContain("未启用");
  });
});

describe("启用请求", () => {
  it("带的就是宿主与工作簿三项，不多不少", async () => {
    const { view, arm } = setup();
    await expand(view);
    await press(armButton(view));
    expect(arm).toHaveBeenCalledTimes(1);
    const sent = arm.mock.calls[0][0];
    expect(sent).toEqual(TARGET);
    expect(Object.keys(sent).sort()).toEqual(["hostId", "workbookIdentity", "workbookName"]);
    // 任务与引用由服务端自己按"哪一次任务"绑定，面板不许也无法掺一脚。
    expect(sent).not.toHaveProperty("taskId");
    expect(sent).not.toHaveProperty("resultRef");
  });

  it("启用成功后按钮锁死：绑到任务之后不能再改主意", async () => {
    const { view, arm } = setup({ arm: async () => statusOf("bound") });
    await expand(view);
    expect(armButton(view).disabled).toBe(false);
    await press(armButton(view));
    expect(arm).toHaveBeenCalledTimes(1);
    // "单次"的语义：绑上就不能换，也不能中途取消掉。
    expect(armButton(view).disabled).toBe(true);
    expect(cancelButton(view).disabled).toBe(true);
    expect(view.getByRole("status").textContent).toContain("已绑定");
  });

  it("执行中（忙）时启用按钮不可按", async () => {
    const { view } = setup({ disabled: true });
    await expand(view);
    expect(armButton(view).disabled).toBe(true);
  });

  it("没有可绑定的工作簿时启用按钮不可按", async () => {
    const { view } = setup({ target: null });
    await expand(view);
    expect(armButton(view).disabled).toBe(true);
  });

  it("服务端没开（enabled=false）时不可启用，并把原因说出来", async () => {
    const { view } = setup({ load: async () => statusOf("off", { enabled: false }) });
    await expand(view);
    expect(armButton(view).disabled).toBe(true);
    expect(view.getByRole("status").textContent).toContain("开发模式");
  });

  it("启用失败时把服务端的话原样显示，且不假装已启用", async () => {
    const { view } = setup({
      load: async () => statusOf("idle"),
      arm: async () => {
        throw new Error("结果引用过期入口未开启（需要开发模式）。");
      }
    });
    await expand(view);
    await press(armButton(view));
    const alert = view.getByRole("alert");
    expect(alert.textContent).toBe("结果引用过期入口未开启（需要开发模式）。");
    /*
     * 同一块 `<p>` 在报错时换成 `role="alert"`：所以此刻**不存在** `role="status"`。
     * 这正是"失败不许被当成状态播报"的判据 —— 若它同时还是 status，
     * 读屏用户会把一句失败当成"未启用"这条正常状态。
     */
    expect(view.queryByRole("status")).toBeNull();
    // 状态没变：按钮仍可按（用户可以重试）。
    expect(armButton(view).disabled).toBe(false);
  });
});

describe("取消", () => {
  it("只有处于「已启用」时才能取消", async () => {
    const { view, cancel } = setup({ load: async () => statusOf("armed") });
    await expand(view);
    expect(cancelButton(view).disabled).toBe(false);
    await press(cancelButton(view));
    expect(cancel).toHaveBeenCalledTimes(1);
    // 取消之后回到"未启用"，可以重新启用。
    expect(view.getByRole("status").textContent).toContain("未启用");
  });

  it("未启用时取消按钮不可按", async () => {
    const { view } = setup({ load: async () => statusOf("idle") });
    await expand(view);
    expect(cancelButton(view).disabled).toBe(true);
  });
});
