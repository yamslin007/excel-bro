/**
 * 定向测试：WPS 数据范围面板的**草稿 / 已提交**分离。
 *
 * 之前 `WpsProbePanel` 里勾选就是提交：`toggleSheet` 立刻改 `selectedIndexes`，
 * 而面板把它当授权范围传给 `WpsAgentPanel`。于是用户只要"顺手点一下另一张表
 * 看看"，就会触发面板里按授权 key 的失效副作用，把正在等回答的任务打掉 ——
 * 他并没有改变任何授权意图，只是想看一眼。
 *
 * 这里钉住三条：
 *  1. 勾选只改草稿；取消勾选再选回，已提交范围一字未动。
 *  2. 「确认」才提交，且**提交时重新读取真实宿主元数据**（不是拿开弹层时的快照）。
 *  3. 「取消」与关闭不改变已提交授权。
 *
 * 外加一条**层次**判据：动作条（`.wps-scope-actions`）是滚动容器的兄弟、
 * 固定页脚；Excel 端必须同层（见 excel-addin/src/scopeActionBarLayer.test.tsx）。
 */
import { act, useEffect, type ReactElement, type ReactNode } from "react";
import type { AuthorizedTarget } from "@excel-bro/core/authorizedTarget";
import { SCOPE_ADMISSION_MESSAGES } from "@excel-bro/core/scopeAdmission";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import WpsProbePanel from "./WpsProbePanel";
import { getWpsApplication } from "./wpsHost";
import { listWpsWorksheets, type WpsWorkbookMetadata } from "./wpsWorkbookReader";

vi.mock("./ribbon", () => ({ initializeRestoredTaskPaneWidth: () => undefined }));
vi.mock("./devProbe", () => ({ reportDevProbe: () => undefined }));
vi.mock("./wpsHost", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./wpsHost")>();
  return { ...actual, getWpsApplication: vi.fn(() => ({})) };
});
vi.mock("./wpsWorkbookReader", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./wpsWorkbookReader")>();
  return { ...actual, listWpsWorksheets: vi.fn() };
});
/**
 * 面板本体换成只暴露"它收到的已提交范围"的替身。
 *
 * 这样断言落在**授权的提交面**上，而不是被面板内部的渲染细节干扰 ——
 * 我们要证的是"勾选有没有提前提交授权"，不是面板长什么样。
 */
vi.mock("./WpsAgentPanel", () => ({
  default: ({
    selectedSheetIndexes,
    scopeDock,
    onScopeConfirmationReady
  }: {
    selectedSheetIndexes: readonly number[];
    scopeDock?: ReactNode;
    onScopeConfirmationReady?: (handler: ((target: AuthorizedTarget) => Promise<boolean>) | null) => void;
  }) => {
    useEffect(() => { onScopeConfirmationReady?.(scopeAuthorization); return () => onScopeConfirmationReady?.(null); }, [onScopeConfirmationReady]);
    return (
    <div>
      <div className="panel-committed">{selectedSheetIndexes.join(",")}</div>
      {scopeDock}
    </div>
    );
  }
}));

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
if (!Element.prototype.scrollIntoView) {
  Element.prototype.scrollIntoView = () => undefined;
}

const getWpsApplicationMock = vi.mocked(getWpsApplication);
const listWpsWorksheetsMock = vi.mocked(listWpsWorksheets);
const scopeAuthorization = vi.fn(async (_target: AuthorizedTarget) => true);

let mounted: Array<{ unmount: () => void }> = [];

async function act_(fn: () => void | Promise<void>) {
  await act(async () => {
    await fn();
  });
}

async function mount(element: ReactElement) {
  const host = document.createElement("div");
  document.body.appendChild(host);
  const root = createRoot(host);
  await act_(() => {
    root.render(element);
  });
  mounted.push({ unmount: () => root.unmount() });
  return { container: host };
}

async function click(target: HTMLElement | null) {
  if (!target) throw new Error("要点击的元素不存在");
  await act_(() => {
    target.dispatchEvent(new MouseEvent("click", { bubbles: true }));
  });
}

function byText(container: HTMLElement, selector: string, text: string): HTMLButtonElement {
  const found = Array.from(container.querySelectorAll<HTMLButtonElement>(selector)).find(
    (button) => (button.textContent ?? "").includes(text)
  );
  if (!found) throw new Error(`找不到「${text}」`);
  return found;
}

function sheetButton(container: HTMLElement, name: string): HTMLButtonElement {
  return byText(container, ".wps-sheet-picker-options button", name);
}

function committed(container: HTMLElement): string {
  return container.querySelector(".panel-committed")?.textContent ?? "";
}

/** 两张表的最小宿主替身（只有元数据走 mock，探测仍读真实宿主对象）。 */
function hostApplication(activeIndex = 1) {
  const makeSheet = (name: string, position: number) => ({
    Name: name,
    get Index() {
      return position + 1;
    }
  });
  const sheets = [makeSheet("订单", 0), makeSheet("明细", 1)];
  return {
    Version: "12.0",
    get ActiveWorkbook() {
      return {
        Name: "测试.xlsx",
        FullName: "C:\\数据\\测试.xlsx",
        Worksheets: {
          Count: sheets.length,
          Item: (index: number) => sheets[Number(index) - 1]
        },
        ActiveSheet: sheets[activeIndex - 1]
      };
    },
    get ActiveSheet() {
      return sheets[activeIndex - 1];
    },
    Selection: undefined,
    PluginStorage: { getItem: () => undefined, setItem: () => undefined },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  } as never;
}

function metadataWith(sheets: Array<{ index: number; name: string; isActive?: boolean }>): WpsWorkbookMetadata {
  return {
    name: "测试.xlsx",
    path: "C:\\数据\\测试.xlsx",
    fullName: "C:\\数据\\测试.xlsx",
    sheets: sheets.map((sheet) => ({ ...sheet, isActive: sheet.isActive ?? false }))
  } as never;
}

beforeEach(() => {
  vi.clearAllMocks();
  scopeAuthorization.mockResolvedValue(true);
  getWpsApplicationMock.mockReturnValue(hostApplication());
  listWpsWorksheetsMock.mockReturnValue(
    metadataWith([
      { index: 1, name: "订单", isActive: true },
      { index: 2, name: "明细" }
    ])
  );
});

afterEach(() => {
  act(() => { mounted.forEach((entry) => entry.unmount()); });
  mounted = [];
  document.body.innerHTML = "";
});

describe("数据范围：勾选只改草稿，确认才提交授权", () => {
  it("uses one compact status line and the existing refresh action", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    const summary = container.querySelector(".scope-admission-summary")!;
    expect(summary.textContent).toContain("已选 1 张表");
    expect(summary.querySelector("button")).toBeNull();
    expect(summary.textContent).not.toContain("仅核实可用性");
    expect(container.textContent).not.toContain("重新检查");
    expect(container.querySelectorAll(".scope-refresh")).toHaveLength(1);
  });

  it("opening automatic scope shows the same live sheet in both places without granting it", async () => {
    const { container } = await mount(<WpsProbePanel />);
    getWpsApplicationMock.mockReturnValue(hostApplication(2));
    listWpsWorksheetsMock.mockReturnValue(metadataWith([
      { index: 1, name: "订单" }, { index: 2, name: "明细", isActive: true }
    ]));
    await click(byText(container, ".scope-trigger", ""));
    expect(container.querySelector(".wps-current-file-card em")?.textContent).toBe("当前工作表 明细");
    expect(container.querySelector(".scope-trigger strong")?.textContent).toBe("当前表 明细");
    expect(committed(container)).toBe("1");
    expect(scopeAuthorization).not.toHaveBeenCalled();
  });

  it("refresh follows the active sheet only in the automatic draft, not committed authorization", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    getWpsApplicationMock.mockReturnValue(hostApplication(2));
    listWpsWorksheetsMock.mockReturnValue(metadataWith([
      { index: 1, name: "订单" }, { index: 2, name: "明细", isActive: true }
    ]));
    await click(byText(container, ".wps-scope-actions button", "刷新"));
    expect(container.querySelector(".wps-current-file-card em")?.textContent).toBe("当前工作表 明细");
    expect(container.querySelector(".scope-trigger strong")?.textContent).toBe("当前表 明细");
    expect(committed(container)).toBe("1");
    expect(scopeAuthorization).not.toHaveBeenCalled();
    await click(byText(container, ".wps-scope-actions button", "确认"));
    expect(scopeAuthorization.mock.calls[0][0].selected).toEqual(["2"]);
    expect(committed(container)).toBe("2");
  });

  it("refresh does not silently remove a missing manually selected sheet", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    await click(sheetButton(container, "明细"));
    await click(byText(container, ".wps-scope-actions button", "确认"));
    await click(byText(container, ".scope-trigger", ""));
    listWpsWorksheetsMock.mockReturnValue(metadataWith([{ index: 1, name: "订单", isActive: true }]));
    await click(byText(container, ".wps-scope-actions button", "刷新"));
    expect(committed(container)).toBe("1,2");
    expect(container.querySelector(".scope-admission-summary")?.textContent).toContain("已不存在");
    expect(byText(container, ".wps-scope-actions button", "确认").disabled).toBe(true);
    expect(scopeAuthorization).toHaveBeenCalledOnce();
  });

  it("a failed catalog refresh invalidates the visible check without publishing a new scope", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    listWpsWorksheetsMock.mockImplementationOnce(() => { throw new Error("PRIVATE_CATALOG_FAILURE"); });
    await click(byText(container, ".wps-scope-actions button", "刷新"));
    expect(byText(container, ".wps-scope-actions button", "确认").disabled).toBe(true);
    expect(container.querySelector(".scope-admission-summary")?.textContent).toContain("暂时无法核实");
    expect(container.textContent).not.toContain("PRIVATE_CATALOG_FAILURE");
    expect(committed(container)).toBe("1");
    expect(scopeAuthorization).not.toHaveBeenCalled();
    await click(byText(container, ".wps-scope-actions button", "刷新"));
    expect(byText(container, ".wps-scope-actions button", "确认").disabled).toBe(false);
    expect(committed(container)).toBe("1");
  });

  it("checks a picked scope inside the selector without granting authorization", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    await click(sheetButton(container, "明细"));
    expect(container.querySelector('.scope-admission-summary')?.textContent).toContain("可分析");
    // 这条判据原先钉的是"行内看得见写入状态"（行内徽标渲染 `write_unverified`
    // 那一句）。行内授权文案整条下线之后，判据的**去处**要跟着搬，而不是消失：
    // 状态仍然必须看得见 —— 在列表下方那一条汇总里；同时行里必须**没有**它。
    expect(container.querySelector('.scope-admission-summary')?.textContent).toContain("写入状态待核实");
    // 行内只剩「勾选框 + 表名」。用授权文案表**逐条**扫一遍，比只钉当前这一句稳：
    // 换一个 reason（例如文件只读）也得守住这条口径。
    const rowText = sheetButton(container, "明细").textContent ?? "";
    expect(rowText).toBe("✓明细");
    for (const message of Object.values(SCOPE_ADMISSION_MESSAGES)) {
      expect(rowText, `行内不该出现授权文案「${message}」`).not.toContain(message);
    }
    // 汇总行在列表**外面**：它不能挪进列表、变成"最后一行"。
    expect(container.querySelector(".wps-sheet-picker-options .scope-admission-summary")).toBeNull();
    expect(scopeAuthorization).not.toHaveBeenCalled();
    expect(committed(container)).toBe("1");
  });
  it("a failed selection check disables submission and offers an in-selector retry", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    getWpsApplicationMock.mockImplementation(() => { throw new Error("PRIVATE_HOST_FAILURE"); });
    await click(sheetButton(container, "明细"));
    expect(byText(container, ".wps-scope-actions button", "确认").disabled).toBe(true);
    expect(container.querySelector('.scope-admission-summary')?.textContent).toContain("暂时无法核实");
    expect(container.textContent).not.toContain("PRIVATE_HOST_FAILURE");
    expect(scopeAuthorization).not.toHaveBeenCalled();
    getWpsApplicationMock.mockReturnValue(hostApplication());
    await click(byText(container, ".wps-scope-actions button", "刷新"));
    expect(byText(container, ".wps-scope-actions button", "确认").disabled).toBe(false);
    expect(committed(container)).toBe("1");
  });
  it("explicit confirmation binds the exact live target before publishing the selected indexes", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    await click(sheetButton(container, "明细"));
    await click(byText(container, ".wps-scope-actions button", "确认"));
    expect(scopeAuthorization).toHaveBeenCalledOnce();
    expect(scopeAuthorization.mock.calls[0][0]).toMatchObject({
      fileKey: "C:\\数据\\测试.xlsx", selected: ["1", "2"]
    });
    expect(committed(container)).toBe("1,2");
  });
  it("failed target verification preserves the original authorization and scope draft", async () => {
    scopeAuthorization.mockResolvedValue(false);
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    await click(sheetButton(container, "明细"));
    await click(byText(container, ".wps-scope-actions button", "确认"));
    expect(scopeAuthorization).toHaveBeenCalledOnce();
    expect(committed(container)).toBe("1");
    expect(container.querySelector(".scope-popover")).not.toBeNull();
    expect(container.querySelector('[role="alert"]')?.textContent).toContain("未变更授权");
  });
  it("取消勾选再选回，已提交范围一字未动（净变化为零时不算草稿）", async () => {
    const { container } = await mount(<WpsProbePanel />);
    expect(committed(container)).toBe("1");

    await click(byText(container, ".scope-trigger", ""));
    // 切到「选择多个」才能手动勾表。
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    await click(sheetButton(container, "明细"));
    // 勾选只是草稿：授权范围没有提前改变。
    expect(committed(container)).toBe("1");
    expect(container.querySelector(".scope-confirm")).not.toBeNull();

    // 再取消选回 —— 勾选集合回到已提交范围。
    await click(sheetButton(container, "明细"));
    expect(sheetButton(container, "明细").className).not.toContain("selected");
    // 全程已提交范围一字未动：用户只是"看一眼另一张表"，任务不该因此被销毁。
    expect(committed(container)).toBe("1");
  });

  it("确认才提交，且提交时重新读取真实宿主元数据", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    await click(sheetButton(container, "明细"));
    expect(committed(container)).toBe("1");

    // 弹层开着的时候，宿主那边发生了变化：索引 2 已经不存在了。
    listWpsWorksheetsMock.mockReturnValue(
      metadataWith([{ index: 1, name: "订单", isActive: true }])
    );

    await click(byText(container, ".wps-scope-actions button", "确认"));
    // 缺失项不能被静默剔除：留在选择器中修正，不提交另一份范围。
    expect(committed(container)).toBe("1");
    expect(container.querySelector(".scope-popover")).not.toBeNull();
    expect(scopeAuthorization).not.toHaveBeenCalled();
    expect(container.querySelector(".scope-admission-summary")?.textContent).toContain("已不存在");
  });

  it("确认提交新增的工作表，并且只在确认之后才生效", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    await click(sheetButton(container, "明细"));
    await click(byText(container, ".wps-scope-actions button", "确认"));
    expect(committed(container)).toBe("1,2");
  });

  it("取消不改变已提交授权，重开弹层时草稿回到已提交范围", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    await click(sheetButton(container, "明细"));
    await click(byText(container, ".wps-scope-actions button", "取消"));
    expect(committed(container)).toBe("1");

    // 重开：草稿必须是已提交范围（明细未被选中）。
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    expect(sheetButton(container, "明细").className).not.toContain("selected");
  });

  it("手动模式下没有任何勾选时不能确认", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    await click(byText(container, ".sheet-picker-toolbar button", "清空"));
    const confirm = byText(container, ".wps-scope-actions button", "确认");
    expect(confirm.disabled).toBe(true);
    // 已提交范围不受草稿清空影响。
    expect(committed(container)).toBe("1");
  });

  it("子表选择面板紧跟「选择多个」，并由 open 类驱动展开/收起", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));

    const modeList = container.querySelector(".scope-mode-list");
    expect(modeList).not.toBeNull();
    // 断言读的是**实际 DOM 位置**：子表面板必须在模式列表**内部**，
    // 而不是列表外/列表之后的兄弟块 —— 否则它会掉到「工作目录」下面（修复点一）。
    const picker = modeList!.querySelector(".wps-manual-sheet-picker");
    expect(picker).not.toBeNull();
    expect(modeList!.contains(picker)).toBe(true);
    const reveal = modeList!.querySelector(".scope-picker-reveal");
    expect(reveal).not.toBeNull();
    expect(picker!.closest(".scope-picker-reveal")).toBe(reveal);

    // 收起态：面板**常挂载**（否则切走时直接卸载，收缩无从播放），
    // 只是没有 `open`（CSS：0fr + opacity:0 + visibility:hidden，不进 Tab 序）。
    expect(reveal!.classList.contains("open")).toBe(false);

    await click(byText(container, ".scope-mode-list button", "选择多个"));
    expect(reveal!.classList.contains("open")).toBe(true);

    // 切回「跟随当前」：必须有收起态，而不是把节点从 DOM 里摘掉。
    await click(byText(container, ".scope-mode-list button", "跟随当前"));
    expect(reveal!.classList.contains("open")).toBe(false);
    expect(modeList!.contains(picker)).toBe(true);

    const manual = byText(container, ".scope-mode-list button", "选择多个");
    const folder = byText(container, ".scope-mode-list button", "工作目录");
    const precedes = (before: Node, after: Node) =>
      Boolean(before.compareDocumentPosition(after) & Node.DOCUMENT_POSITION_FOLLOWING);
    // 顺序：选择多个 → 子表面板 → 工作目录。
    expect(precedes(manual, picker!)).toBe(true);
    expect(precedes(picker!, folder)).toBe(true);
  });

  it("操作区用「确认 / 取消」，去掉草稿提示，刷新图标点击转一圈", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));

    // 那段"勾选还只是草稿…"的提醒已移除（用户反馈多余）。
    expect(container.querySelector(".wps-scope-draft-hint")).toBeNull();

    // 刷新图标：默认不转；点一下才挂上 spinning（组件用点击计数当 key，
    // 换 key 即重挂载 → 动画重放一次）。
    const glyph = () => container.querySelector(".scope-refresh-glyph")!;
    expect(glyph()).not.toBeNull();
    expect(glyph().classList.contains("spinning")).toBe(false);
    await click(byText(container, ".wps-scope-actions button", "刷新"));
    expect(glyph().classList.contains("spinning")).toBe(true);

    // 造出草稿差异，操作区应同时出现「确认」与「取消」两个按钮。
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    await click(sheetButton(container, "明细"));
    const labels = Array.from(container.querySelectorAll<HTMLButtonElement>(".wps-scope-actions button"))
      .map((button) => (button.textContent ?? "").trim());
    expect(labels).toContain("确认");
    expect(labels).toContain("取消");
    // 负向：旧文案不许复活。
    expect(container.textContent).not.toContain("确认范围");
    expect(container.textContent).not.toContain("放弃修改");
  });

  it("动作条是滚动容器的兄弟：固定页脚，不跟着内容滚走", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));

    const popover = container.querySelector<HTMLElement>(".scope-popover")!;
    const scroll = popover.querySelector<HTMLElement>(".scope-popover-scroll")!;
    const actions = popover.querySelector<HTMLElement>(".wps-scope-actions")!;
    expect(scroll).not.toBeNull();
    expect(actions).not.toBeNull();

    // 正向：直接住在弹层上，并且**紧跟在滚动容器之后** —— 只有这个位置才被
    // `.scope-popover`（column flex + overflow: hidden）钉成固定页脚。
    expect(actions.parentElement).toBe(popover);
    expect(actions.previousElementSibling).toBe(scroll);

    // 反向：不在滚动容器里。放进滚动容器，`flex: 0 0 auto` 只表示"不许被压缩"，
    // 挡不住容器整体滚动，动作条会跟内容一起滚出视口 —— 这正是
    // 「Excel 的完成按钮所在的层和 WPS 不一样」的根因，所以两端的判据必须成对存在。
    expect(scroll.contains(actions)).toBe(false);
    expect(scroll.querySelector(".wps-scope-actions")).toBeNull();
    expect(scroll.querySelector(".scope-confirm")).toBeNull();

    /*
      工作目录来源下同一件事也要成立 —— 这是**同一个调用点**，只是换了个来源。
      那一边**没有** `.wps-scope-actions`：它的固定页脚是 `.folder-selection-footer`
      （与 Excel 同构：计数与确认同一行，承认汇总落在页脚下方）。
      页脚自身的层次判据在这一份夹具里钉不了 —— 没有目录时它根本不渲染 ——
      所以钉在带目录夹具的 `wpsWorkDirectoryScope.test.tsx` 里（「固定页脚的层与格式」）。
      这里能钉、也必须钉的是：换来源之后**不许**冒出第二个固定页脚，
      而且承认汇总仍在滚动容器**外面**（它决定能不能提交，不能被滚走）。
    */
    await click(byText(container, ".scope-mode-list button", "工作目录"));
    const folderPopover = container.querySelector<HTMLElement>(".scope-popover")!;
    const folderScroll = folderPopover.querySelector<HTMLElement>(".scope-popover-scroll")!;
    // 负向：工作簿那条动作条不许跟到工作目录模式里来（否则一个面板会有两条固定页脚）。
    expect(folderPopover.querySelector(".wps-scope-actions")).toBeNull();
    // 判据用 `querySelector(...).toBeNull()` 而不是 `contains(...)` 取反：失败原文里
    // 会带上被卷进去的那个元素本身，比一句 `expected true to be false` 有用得多。
    expect(container.querySelector(".scope-admission-summary")).not.toBeNull();
    expect(folderScroll.querySelector(".scope-admission-summary")).toBeNull();
    expect(folderScroll.querySelector(".folder-selection-footer")).toBeNull();

    // Excel 端工作簿动作条的同一判据见 apps/excel-addin/src/scopeActionBarLayer.test.tsx；
    // 两端读的也是同一份声明（`.scope-action-buttons, .wps-scope-actions`）。
  });
});
