/**
 * 定向测试：WPS 数据范围面板的**第三种来源 —— 工作目录**。
 *
 * 在这之前，WPS 的「工作目录」卡是一颗 `disabled` 的死按钮，副标题写着
 * 「WPS 当前仅支持打开的工作簿」。本轮把它按 Excel 的同一套口径打开：
 * 选目录 → 文件/表**两级勾选** → 承认核验 → 自动提交授权 → 收起保留选择 →
 * 刷新增表不作废、换目录/目标消失即作废。
 *
 * 这份用例钉住的是**授权语义**，不是面板长什么样，所以 `WpsAgentPanel` 换成
 * 只暴露"它收到的来源与工作目录范围"的替身（`panel-scope-source` /
 * `panel-folder-scope`）—— 被证的是"有没有把工作目录当成这次的授权"。
 *
 * 判据之间刻意**读不同的数据源**，免得一条坏掉另一条跟着空过：
 *  - 「勾选后自动提交」读替身面板收到的来源；
 *  - 「勾选即授权」读 `createFolderSnapshot` 收到的真实 selections；
 *  - 「刷新不是同意」读刷新后替身面板收到的来源**是否仍然**是 folder；
 *  - 「目标消失 / 换目录 ⇒ 授权作废」读 `panel-folder-scope` 是否变回 `none`。
 */
import { act, useEffect, type ReactElement, type ReactNode } from "react";
import type { AuthorizedTarget } from "@excel-bro/core/authorizedTarget";
import { admissionItem } from "@excel-bro/core/scopeAdmission";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { FolderCatalog, FolderSelection } from "../../excel-addin/src/contracts";
import WpsProbePanel from "./WpsProbePanel";
import { getWpsApplication } from "./wpsHost";
import { listWpsWorksheets, type WpsWorkbookMetadata } from "./wpsWorkbookReader";
import {
  checkFolderAdmission,
  createFolderSnapshot,
  scanFolderByPath,
  selectFolder
} from "../../excel-addin/src/api";

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
 * 工作目录的四个接口全部替换成替身。
 *
 * 目录框是**服务端**弹的（`POST /api/folders/select`），jsdom 里没有本机服务端，
 * 所以这里必须换掉；但替身按**服务端的真实契约**造返回：承认项的 `id` 就是
 * `fileId + NUL + 表名`（`folder_admission.py` 的拼法），与勾选一一对应。
 */
vi.mock("../../excel-addin/src/api", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../excel-addin/src/api")>();
  return {
    ...actual,
    selectFolder: vi.fn(),
    scanFolderByPath: vi.fn(),
    createFolderSnapshot: vi.fn(),
    checkFolderAdmission: vi.fn()
  };
});

vi.mock("./WpsAgentPanel", () => ({
  default: function WpsDirectoryScopeHarness({
    scopeSource,
    folderScope,
    selectedSheetIndexes,
    scopeDock,
    scopeOverlay,
    onScopeConfirmationReady
  }: {
    scopeSource?: "workbook" | "folder";
    folderScope?: { folderName: string; sheetCount: number } | null;
    selectedSheetIndexes: readonly number[];
    scopeDock?: ReactNode;
    scopeOverlay?: ReactNode;
    onScopeConfirmationReady?: (handler: ((target: AuthorizedTarget) => Promise<boolean>) | null) => void;
  }) {
    // 工作簿模式的确认要经过"绑定真实目标"这一步；替身按真实面板的做法登记回调，
    // 否则工作簿侧的确认永远走不到，判据会指向一个不存在的缺陷。
    useEffect(() => {
      onScopeConfirmationReady?.(scopeAuthorization);
      return () => onScopeConfirmationReady?.(null);
    }, [onScopeConfirmationReady]);
    return (
      <div>
        <div className="panel-scope-source">{scopeSource ?? "workbook"}</div>
        <div className="panel-folder-scope">
          {folderScope ? `${folderScope.folderName}|${folderScope.sheetCount}` : "none"}
        </div>
        <div className="panel-committed">{selectedSheetIndexes.join(",")}</div>
        {scopeDock}
        {scopeOverlay ? <div className="scope-layer">{scopeOverlay}</div> : null}
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
const selectFolderMock = vi.mocked(selectFolder);
const scanFolderByPathMock = vi.mocked(scanFolderByPath);
const createFolderSnapshotMock = vi.mocked(createFolderSnapshot);
const checkFolderAdmissionMock = vi.mocked(checkFolderAdmission);
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
  if (!found) throw new Error(`找不到「${text}」（${selector}）`);
  return found;
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

/** f1 带一张 EB 系统遗留表 `#EB_RULES`：它能被扫描到，但**不可勾选**。 */
function folderCatalog(overrides: Partial<FolderCatalog> = {}): FolderCatalog {
  return {
    sessionId: "folder-session-1",
    folderName: "数据",
    folderPath: "C:\\工作\\数据",
    files: [
      {
        id: "f1",
        name: "销售.xlsx",
        relativePath: "销售.xlsx",
        worksheets: [
          { name: "明细", rowCount: 8, columnCount: 3 },
          { name: "#EB_RULES", rowCount: 1, columnCount: 2 }
        ]
      },
      {
        id: "f2",
        name: "库存.xlsx",
        relativePath: "库存.xlsx",
        worksheets: [{ name: "入库", rowCount: 4, columnCount: 2 }]
      }
    ],
    totalFiles: 2,
    truncated: false,
    expiresAt: "2026-10-16T23:00:00.000Z",
    ...overrides
  };
}

/** 与服务端同构的承认项：一项 = 一个 `fileId\0表名`。 */
function admissionFor(selections: readonly FolderSelection[]) {
  return {
    items: selections.flatMap((selection) =>
      selection.sheets.map((name) => admissionItem(`${selection.fileId}\u0000${name}`, "none"))
    ),
    fingerprint: "folder-fingerprint"
  };
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
  selectFolderMock.mockResolvedValue(folderCatalog());
  scanFolderByPathMock.mockResolvedValue(folderCatalog());
  checkFolderAdmissionMock.mockImplementation(async (_sessionId, selections) =>
    admissionFor(selections) as never
  );
  createFolderSnapshotMock.mockImplementation(async (_sessionId, selections) => ({
    name: "工作目录",
    capturedAt: "2026-10-16T00:00:00.000Z",
    activeWorksheet: selections[0]?.sheets[0] ?? "",
    worksheets: selections.flatMap((selection) =>
      selection.sheets.map((name) => ({ name, rowCount: 1, columnCount: 1 }))
    )
  }) as never);
});

afterEach(() => {
  act(() => { mounted.forEach((entry) => entry.unmount()); });
  mounted = [];
  document.body.innerHTML = "";
});

function openPopover(container: HTMLElement) {
  return click(byText(container, ".scope-trigger", ""));
}

/** 打开弹层 → 切到「工作目录」→ 选好目录，停在"目录已就绪、一个勾选都没有"。 */
async function openWorkDirectory(container: HTMLElement) {
  await openPopover(container);
  await click(byText(container, ".scope-mode-list button", "工作目录"));
  await click(byText(container, ".work-directory-empty button", "选择工作目录"));
}

function scopeSourceOf(container: HTMLElement): string {
  return container.querySelector(".panel-scope-source")?.textContent ?? "";
}

function folderScopeOf(container: HTMLElement): string {
  return container.querySelector(".panel-folder-scope")?.textContent ?? "";
}

function footerCount(container: HTMLElement): string {
  return container.querySelector(".folder-selection-footer strong")?.textContent ?? "";
}

function triggerText(container: HTMLElement): string {
  return container.querySelector(".scope-trigger")?.textContent ?? "";
}

/** 打开某个文件（一级），返回它二级列表里每个工作表按钮的 `<strong>` 文字。 */
async function openFile(container: HTMLElement, relativePath: string): Promise<string[]> {
  await click(byText(container, ".folder-file-option", relativePath));
  return Array.from(container.querySelectorAll<HTMLElement>(".folder-sheet-list button strong"))
    .map((node) => node.textContent ?? "");
}

/** 勾上当前文件的某张表（二级）。 */
function folderSheetButton(container: HTMLElement, name: string): HTMLButtonElement {
  return byText(container, ".folder-sheet-list button", name);
}

/** 目录工具条上的按钮（与工作簿侧的 `.sheet-picker-toolbar` 分开，避免点错）。 */
function folderToolbarButton(container: HTMLElement, text: string): HTMLButtonElement {
  return byText(container, ".folder-sheet-detail .sheet-picker-toolbar button", text);
}

function closeButton(container: HTMLElement): HTMLButtonElement {
  // 「确认」现在有两个可能的家：工作目录模式下它在固定页脚 `.folder-selection-footer` 里
  // （与 Excel 同构），工作簿模式下它在动作条 `.wps-scope-actions` 里。同一时刻只会有一个。
  return byText(container, ".folder-selection-footer button, .wps-scope-actions button", "确认");
}

describe("WPS 工作目录：卡片放开与文案口径", () => {
  it("「工作目录」不再禁用，副标题对齐 Excel，旧的不可用说明一个字都不在", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await openPopover(container);
    const folder = byText(container, ".scope-mode-list button", "工作目录");
    expect(folder.disabled).toBe(false);
    expect(folder.querySelector("small")?.textContent).toBe("批量处理多个本地工作簿");
    // 负向：那句"WPS 当前仅支持打开的工作簿"是这一轮要拆掉的**能力说明**，
    // 不是换个说法留着 —— 近义写法也不许在。
    expect(container.textContent).not.toContain("WPS 当前仅支持");
    expect(container.textContent).not.toContain("仅支持打开的工作簿");
    // 三种来源都可以来回切。
    expect(byText(container, ".scope-mode-list button", "跟随当前").disabled).toBe(false);
    expect(byText(container, ".scope-mode-list button", "选择多个").disabled).toBe(false);
  });

  it("还没选目录时不显示勾选区，并明确说取消目录框不改范围", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await openPopover(container);
    await click(byText(container, ".scope-mode-list button", "工作目录"));
    expect(container.querySelector(".work-directory-empty")).not.toBeNull();
    expect(container.querySelector(".folder-master-detail")).toBeNull();
    expect(container.textContent).toContain("取消不会改变当前范围");
    // 目录框返回 null（用户点了取消）⇒ 一个字段都不许变。
    selectFolderMock.mockResolvedValue(null);
    await click(byText(container, ".work-directory-empty button", "选择工作目录"));
    expect(container.querySelector(".work-directory-empty")).not.toBeNull();
    expect(scopeSourceOf(container)).toBe("folder");
    expect(folderScopeOf(container)).toBe("none");
  });
});

describe("WPS 工作目录：两级勾选、承认汇总、勾选即授权", () => {
  it("文件/表两级勾选；勾选后自动授权工作目录", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await openWorkDirectory(container);

    // 一级：两个文件；可勾选表数**不含** EB 系统遗留表（f1 是 1 张而不是 2 张）。
    expect(container.querySelectorAll(".folder-file-option")).toHaveLength(2);
    expect(byText(container, ".folder-file-option", "销售.xlsx").querySelector("b")?.textContent).toBe("0/1");
    expect(container.querySelector(".folder-selection-filter span")?.textContent).toBe("2 个文件 · 2 张表");

    // 二级：进入 f1，工作表里**只有**「明细」——`#EB_RULES` 连名字都不出现。
    expect(await openFile(container, "销售.xlsx")).toEqual(["明细"]);
    expect(container.textContent).not.toContain("#EB_RULES");

    await click(folderSheetButton(container, "明细"));
    expect(footerCount(container)).toBe("已选 1 个文件 · 1 张表");
    expect(container.querySelector(".scope-admission-summary")?.textContent).toContain("已选 1 张表 · 可分析");
    expect(scopeSourceOf(container)).toBe("folder");
    expect(folderScopeOf(container)).toBe("数据|1");
    expect(container.querySelector(".scope-confirm")).toBeNull();
    await click(closeButton(container));
    // 确认走的是**工作目录通道**：快照按勾选算，与勾选一一对应。
    expect(createFolderSnapshotMock).toHaveBeenCalledTimes(1);
    expect(createFolderSnapshotMock.mock.calls[0][1]).toEqual([{ fileId: "f1", sheets: ["明细"] }]);
    expect(scopeSourceOf(container)).toBe("folder");
    expect(folderScopeOf(container)).toBe("数据|1");
    // 工作簿侧已提交的范围没有被顺手改掉（两种来源各持自己的授权）。
    expect(container.querySelector(".panel-committed")?.textContent).toBe("1");
  });

  it("「只看已选」只留勾过的文件；清空本文件把勾选清零并挡住提交", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await openWorkDirectory(container);
    await openFile(container, "销售.xlsx");
    await click(folderSheetButton(container, "明细"));
    await click(byText(container, ".folder-selection-filter button", "只看已选"));
    const shown = Array.from(container.querySelectorAll<HTMLButtonElement>(".folder-file-option"))
      .map((file) => file.textContent ?? "");
    expect(shown).toHaveLength(1);
    expect(shown[0]).toContain("销售.xlsx");

    await click(folderToolbarButton(container, "清空本文件"));
    expect(footerCount(container)).toBe("已选 0 个文件 · 0 张表");
    // 负向：勾选归零后不能提交（空范围不是一次授权）。
    expect(folderScopeOf(container)).toBe("none");
  });

  it("「清空」一次撤销整个目录的勾选授权", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await openWorkDirectory(container);
    await openFile(container, "销售.xlsx");
    await click(folderSheetButton(container, "明细"));
    await click(byText(container, ".folder-file-option", "库存.xlsx"));
    await click(folderSheetButton(container, "入库"));
    expect(footerCount(container)).toBe("已选 2 个文件 · 2 张表");

    await click(byText(container, ".folder-selection-filter button", "清空"));
    expect(footerCount(container)).toBe("已选 0 个文件 · 0 张表");
    expect(folderScopeOf(container)).toBe("none");
  });

  it("清空立即撤销目录授权，关闭重开不恢复旧勾选", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await openWorkDirectory(container);
    await openFile(container, "销售.xlsx");
    await click(folderSheetButton(container, "明细"));
    expect(folderScopeOf(container)).toBe("数据|1");
    await click(byText(container, ".folder-selection-filter button", "清空"));
    expect(folderScopeOf(container)).toBe("none");
    await click(closeButton(container));
    await openPopover(container);
    expect(footerCount(container)).toBe("已选 0 个文件 · 0 张表");
    expect(folderScopeOf(container)).toBe("none");
  });

  it("承认检查逐项对齐勾选：检查不通过就不能提交，也不产生快照", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await openWorkDirectory(container);
    await openFile(container, "销售.xlsx");
    checkFolderAdmissionMock.mockImplementation(async (_sessionId, selections) => ({
      items: selections.flatMap((selection) =>
        selection.sheets.map((name) => admissionItem(`${selection.fileId}\u0000${name}`, "unreadable"))
      ),
      fingerprint: "folder-fingerprint"
    }) as never);
    await click(folderSheetButton(container, "明细"));
    expect(container.querySelector(".scope-admission-summary")?.textContent).toContain("1 张不可用");
    expect(folderScopeOf(container)).toBe("none");
    await click(closeButton(container));
    expect(createFolderSnapshotMock).not.toHaveBeenCalled();
    expect(scopeSourceOf(container)).toBe("folder");
  });
});

describe("WPS 工作目录：刷新不是同意", () => {
  it("刷新后同根同会话、勾选都还在 ⇒ 保有授权，新增表不自动勾上", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await openWorkDirectory(container);
    await openFile(container, "销售.xlsx");
    await click(folderSheetButton(container, "明细"));
    await click(closeButton(container));
    expect(folderScopeOf(container)).toBe("数据|1");

    // 目录里多了一张表（在另一个文件里），原有的「明细」还在。
    scanFolderByPathMock.mockResolvedValue(folderCatalog({
      files: [
        { id: "f1", name: "销售.xlsx", relativePath: "销售.xlsx",
          worksheets: [{ name: "明细", rowCount: 8, columnCount: 3 }] },
        { id: "f2", name: "库存.xlsx", relativePath: "库存.xlsx",
          worksheets: [
            { name: "入库", rowCount: 4, columnCount: 2 },
            { name: "新增表", rowCount: 2, columnCount: 2 }
          ] }
      ]
    }));

    // 打开弹层时会做一次**静默刷新** —— 那也是一次刷新，同样不许作废授权。
    await openPopover(container);
    expect(scopeSourceOf(container)).toBe("folder");
    expect(folderScopeOf(container)).toBe("数据|1");
    expect(container.querySelector(".folder-refresh-notice")?.textContent).toContain("发现 1 张新增工作表");
    // 新增表**不自动勾选**（自动勾 = 替用户扩权）。
    expect(footerCount(container)).toBe("已选 1 个文件 · 1 张表");

    // 再显式刷一次：这次没有新增，授权仍然不许被动。
    await click(byText(container, ".folder-refresh-button", "刷新"));
    expect(scopeSourceOf(container)).toBe("folder");
    expect(folderScopeOf(container)).toBe("数据|1");
    expect(container.querySelector(".folder-refresh-notice")?.textContent).toContain("已刷新");
    expect(createFolderSnapshotMock).toHaveBeenCalledTimes(1);
  });

  it("已勾的表在刷新后消失 ⇒ 授权作废（保留来源，但授权归零，发送侧继续拦）", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await openWorkDirectory(container);
    await openFile(container, "销售.xlsx");
    await click(folderSheetButton(container, "明细"));
    await click(closeButton(container));
    expect(folderScopeOf(container)).toBe("数据|1");

    // 那张表被改名/删掉了：勾选已经不是事实，授权不能靠它继续成立。
    scanFolderByPathMock.mockResolvedValue(folderCatalog({
      files: [
        { id: "f1", name: "销售.xlsx", relativePath: "销售.xlsx",
          worksheets: [{ name: "明细改名", rowCount: 8, columnCount: 3 }] }
      ]
    }));
    await openPopover(container);
    expect(folderScopeOf(container)).toBe("none");
    expect(footerCount(container)).toBe("已选 0 个文件 · 0 张表");
    // 来源**不偷偷改回工作簿**：用户选的还是工作目录，只是现在什么都没授权。
    // 触发器如实说"已选 0 张表"，而不是转身去用当前打开的工作簿。
    expect(triggerText(container)).toContain("工作目录 数据 · 已选 0 张表");
    expect(triggerText(container)).not.toContain("测试.xlsx");
    expect(folderScopeOf(container)).toBe("none");
  });

  it("换成另一个根目录 ⇒ 旧勾选与旧授权一起作废", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await openWorkDirectory(container);
    await openFile(container, "销售.xlsx");
    await click(folderSheetButton(container, "明细"));
    await click(closeButton(container));
    expect(folderScopeOf(container)).toBe("数据|1");

    selectFolderMock.mockResolvedValue(folderCatalog({
      sessionId: "folder-session-2",
      folderName: "别的数据",
      folderPath: "D:\\其他\\数据",
      files: [
        { id: "g1", name: "报表.xlsx", relativePath: "报表.xlsx",
          worksheets: [{ name: "汇总", rowCount: 3, columnCount: 2 }] }
      ]
    }));
    await openPopover(container);
    await click(byText(container, ".folder-toolbar button", "更换工作目录"));
    expect(folderScopeOf(container)).toBe("none");
    expect(footerCount(container)).toBe("已选 0 个文件 · 0 张表");
    expect(container.querySelector(".folder-selection-filter span")?.textContent).toBe("1 个文件 · 1 张表");
    // 旧根上的表一个都不许跟过来。
    expect(container.querySelector(".folder-file-list")?.textContent).not.toContain("销售.xlsx");
    expect(container.querySelector(".folder-file-list")?.textContent).toContain("报表.xlsx");
  });
});

describe("WPS 工作目录：收起保留授权", () => {
  it("关闭保留两次勾选立即生效的授权", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await openWorkDirectory(container);
    await openFile(container, "销售.xlsx");
    await click(folderSheetButton(container, "明细"));
    await openFile(container, "库存.xlsx");
    await click(folderSheetButton(container, "入库"));
    expect(folderScopeOf(container)).toBe("数据|2");
    expect(createFolderSnapshotMock.mock.lastCall?.[1]).toEqual([
      { fileId: "f1", sheets: ["明细"] }, { fileId: "f2", sheets: ["入库"] }
    ]);
    await click(closeButton(container));
    await openPopover(container);
    expect(footerCount(container)).toBe("已选 2 个文件 · 2 张表");
    expect(folderScopeOf(container)).toBe("数据|2");
  });

  it("触发器在工作目录模式下只说工作目录，切回工作簿后照旧", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await openWorkDirectory(container);
    await openFile(container, "销售.xlsx");
    await click(folderSheetButton(container, "明细"));
    await click(closeButton(container));

    const trigger = container.querySelector(".scope-trigger")!;
    expect(trigger.querySelector("strong")?.textContent).toBe("工作目录 数据 · 已选 1 张表");
    // 工作目录模式下**不许**出现当前工作簿名：这一行是"这次用哪些数据"。
    expect(trigger.querySelector("small")).toBeNull();
    expect(trigger.textContent).not.toContain("测试.xlsx");
    expect(trigger.textContent).not.toContain("当前表");

    // 切回「跟随当前」并确认：那一行照旧显示工作簿与当前表。
    await openPopover(container);
    await click(byText(container, ".scope-mode-list button", "跟随当前"));
    await click(closeButton(container));
    expect(container.querySelector(".scope-trigger strong")?.textContent).toBe("当前表 订单");
    expect(container.querySelector(".scope-trigger small")?.textContent).toBe("测试.xlsx");
  });
});

describe("WPS 工作目录：固定页脚的层与格式（对齐 Excel）", () => {
  it("页脚住在弹层上、不在滚动容器里，收起就在页脚里，承认汇总落在它下方", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await openWorkDirectory(container);

    const popover = container.querySelector<HTMLElement>(".scope-popover")!;
    const scroll = popover.querySelector<HTMLElement>(".scope-popover-scroll")!;
    const footer = popover.querySelector<HTMLElement>(".folder-selection-footer")!;
    expect(footer).not.toBeNull();

    // 反向（先判，失败原文里才会带出是哪个元素被卷进去了）：页脚与承认汇总都不在滚动容器里。
    // 进去之后它就只是"内容末尾的一行 flex"，外层容器一滚就跟着滚走 —— 这正是
    // 「Excel 的完成按钮所在的层和 WPS 不一样」的根因，两端必须成对存在。
    expect(scroll.querySelector(".folder-selection-footer")).toBeNull();
    expect(scroll.querySelector(".folder-selection-summary")).toBeNull();
    expect(scroll.querySelector(".scope-admission-summary")).toBeNull();

    // 正向：它直接住在弹层上。
    expect(footer.parentElement).toBe(popover);

    // 格式对齐 Excel：确认按钮就在这条页脚里（`.folder-selection-footer .scope-confirm`），
    // 承认汇总紧跟在页脚**下方** —— 不是"计数一行、承认一行、动作条再一行"的三行堆叠。
    expect(footer.querySelectorAll("button.scope-close")).toHaveLength(1);
    expect(footer.nextElementSibling?.classList.contains("scope-admission-summary")).toBe(true);
    // 工作目录模式**没有**独立的 `.wps-scope-actions`（那条是工作簿模式的）。
    expect(container.querySelector(".wps-scope-actions")).toBeNull();
  });
});
