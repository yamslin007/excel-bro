// @vitest-environment jsdom
import { cleanup, fireEvent, render, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import App from "./App";
import { checkFolderAdmission, checkIntent, createFolderSnapshot, refreshFolder, scanFolderByPath, selectFolder, understandMessage } from "./api";
import { captureWorkbook, captureWorkbookCatalog, captureWorkbookStructure } from "./excel";
import { inspectOfficeScopeAdmission } from "./officeAuthorizedTarget";
import type { FolderCatalog, WorkbookSnapshot } from "./contracts";

const fixtures = vi.hoisted(() => ({
  catalog: null as FolderCatalog | null,
  office: {
    name: "book.xlsx", capturedAt: "2026-10-08T00:00:00Z", activeWorksheet: "Unselected",
    worksheets: ["Selected", "Unselected"].map((name) => ({
      name, usedRange: "A1:B3", rowCount: 3, columnCount: 2,
      headers: ["Category", "Value"], dataRows: [], truncated: false
    }))
  } as WorkbookSnapshot
}));

vi.mock("./api", async (original) => ({
  ...await original<typeof import("./api")>(),
  checkHealth: vi.fn(async () => ({ status: "ok", modelConfigured: true })),
  listModels: vi.fn(async () => ({ defaultModelId: "model-1", models: [{ id: "model-1", label: "模型", modelId: "test-model", baseUrl: "http://localhost", supportsVision: true, available: true }] })),
  getModelSettings: vi.fn(async () => ({ connectionCount: 1 })),
  selectFolder: vi.fn(async () => fixtures.catalog),
  refreshFolder: vi.fn(async () => fixtures.catalog!),
  checkFolderAdmission: vi.fn(async (_session, selections) => ({ fingerprint: "selection-check",
    items: selections.flatMap((s: {fileId: string; sheets: string[]}) => s.sheets.map((name) => ({ id: `${s.fileId}\0${name}`, status: "ready", reason: "none" }))) })),
  scanFolderByPath: vi.fn(),
  createFolderSnapshot: vi.fn(async () => ({
    name: "文件夹：Data", capturedAt: "2026-10-08T00:00:00Z", activeWorksheet: "book.xlsx › Selected",
    sourceFingerprint: "synthetic-fingerprint", sourceFingerprintSheets: ["book.xlsx › Selected"],
    worksheets: [{ ...fixtures.office.worksheets[0], name: "book.xlsx › Selected",
      sourceFileId: "same-relative-file-id", sourceSheetId: "same-sheet-id", sourceFile: "book.xlsx", sourceSheet: "Selected" }]
  })),
  understandMessage: vi.fn(), checkIntent: vi.fn()
}));
vi.mock("./excel", async (original) => ({
  ...await original<typeof import("./excel")>(),
  isRunningInExcel: () => true,
  captureWorkbookCatalog: vi.fn(async () => ({ ...fixtures.office, catalogOnly: true })),
  captureWorkbook: vi.fn(async () => fixtures.office),
  captureWorkbookStructure: vi.fn(async () => fixtures.office),
  watchWorkbookStructureChanges: vi.fn(async () => () => {}),
  currentOfficeWorkbookIdentity: () => "office:book-1"
}));
vi.mock("./officeAuthorizedTarget", async (original) => ({
  ...await original<typeof import("./officeAuthorizedTarget")>(),
  inspectOfficeAuthorizedTarget: vi.fn(async (names: readonly string[] | null) => ({ fileKey: "office:book-1", name: "book.xlsx",
    sheets: [{ id: "selected", name: "Selected" }, { id: "unselected", name: "Unselected" }], selected: names === null ? ["unselected"] : names.map((name) => name.toLowerCase()) })),
  inspectOfficeScopeAdmission: vi.fn(async (names: readonly string[]) => ({
    items: names.map((id) => ({ id, status: "ready", reason: "none" })),
    target: { fileKey: "office:book-1", name: "book.xlsx", sheets: [{ id: "selected", name: "Selected" }, { id: "unselected", name: "Unselected" }], selected: names.map((name) => name.toLowerCase()) }
  })),
  confirmOfficeAuthorizedScope: vi.fn(async (_names, confirm, expected) => confirm(expected)),
  officeAnalysisTargetEvidence: vi.fn(async () => ({ key: "target-1", conclusive: true }))
}));

function catalog(path = "C:/One/Data", sessionId = "session-one"): FolderCatalog {
  return { sessionId, folderName: "Data", folderPath: path, totalFiles: 1, truncated: false,
    expiresAt: "2030-01-01T00:00:00Z", files: [{ id: "same-relative-file-id", name: "book.xlsx",
      relativePath: "book.xlsx", worksheets: ["Selected", "Unselected"].map((name) => ({ name, rowCount: 3, columnCount: 2 })) }] };
}

beforeEach(() => {
  localStorage.clear(); vi.clearAllMocks();
  fixtures.office.activeWorksheet = "Unselected";
  fixtures.catalog = catalog();
  vi.mocked(refreshFolder).mockImplementation(async () => fixtures.catalog!);
  vi.mocked(scanFolderByPath).mockRejectedValue(new Error("合成目录已删除"));
  vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "需要表格", provider: "model" });
  vi.mocked(checkIntent).mockResolvedValue({ kind: "conversation", message: "字段已核对。", provider: "model" });
  vi.spyOn(console, "log").mockImplementation(() => {});
  vi.stubGlobal("Office", { context: { document: { url: "file:///C:/Current/book.xlsx" } }, onReady: (cb: () => void) => cb() });
  Element.prototype.scrollIntoView = vi.fn();
});
afterEach(() => { cleanup(); vi.unstubAllGlobals(); vi.restoreAllMocks(); });

async function boot() {
  const view = render(<App />);
  await waitFor(() => expect(captureWorkbookCatalog).toHaveBeenCalled());
  await waitFor(() => expect(view.container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title).toContain("当前模型"));
  return view;
}
async function pickFolder(view: ReturnType<typeof render>, next: FolderCatalog) {
  if (!view.container.querySelector(".scope-popover")) fireEvent.click(view.container.querySelector(".scope-trigger")!);
  if (!view.container.querySelector(".folder-scope-picker")) {
    fireEvent.click(view.getByRole("button", { name: /工作目录.*批量处理/ }));
  }
  const browse = view.getByRole("button", { name: /^选择工作目录$|^更换工作目录/ });
  await waitFor(() => expect((browse as HTMLButtonElement).disabled).toBe(false));
  fixtures.catalog = next;
  fireEvent.click(browse);
  await waitFor(() => expect(view.container.querySelector(".folder-sheet-list")).not.toBeNull());
  await waitFor(() => expect((view.getByRole("button", { name: /^Selected/ }) as HTMLButtonElement).disabled).toBe(false));
  fireEvent.click(view.getByRole("button", { name: /^Selected/ }));
  await waitFor(() => expect(view.container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("false"));
  fireEvent.click(view.container.querySelector(".scope-close")!);
  await waitFor(() => expect(view.container.querySelector(".scope-popover")).toBeNull());
}
function send(view: ReturnType<typeof render>, prompt: string) {
  fireEvent.change(view.container.querySelector(".composer textarea")!, { target: { value: prompt } });
  /*
   * 有待回答的问题时，右下角那颗按钮的可访问名是「发送回答」——这时这句话就是
   * 对那道题的回答（旧问题只在它自己还 pending 时才接受答案，身份绑定管住这一点）。
   */
  fireEvent.click(view.getByRole("button", { name: /^(发送|发送回答)$/ }));
}

describe("actual Excel folder scope safety", () => {
  it("failed selection checks stay in the existing picker, without data reads or a second dialog", async () => {
    const view = await boot();
    vi.mocked(checkFolderAdmission).mockResolvedValueOnce({ fingerprint: null,
      items: [{ id: "same-relative-file-id\0Selected", status: "unavailable", reason: "unreadable" }] });
    fireEvent.click(view.container.querySelector(".scope-trigger")!);
    fireEvent.click(view.getByRole("button", { name: /工作目录.*批量处理/ }));
    fireEvent.click(view.getByRole("button", { name: "选择工作目录" }));
    await waitFor(() => expect(view.container.querySelector(".folder-sheet-list")).not.toBeNull());
    fireEvent.click(view.getByRole("button", { name: /^Selected/ }));
    await waitFor(() => expect(view.container.textContent).toContain("无法读取文件"));
    expect(view.container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("true");
    expect(view.queryByRole("dialog")).toBeNull();
    expect(createFolderSnapshot).not.toHaveBeenCalled();
    expect(captureWorkbook).not.toHaveBeenCalled();
    fireEvent.click(view.container.querySelector(".folder-refresh-button")!);
    await waitFor(() => expect(view.container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("false"));
  });
  it("opening and closing an unchanged folder picker does not demand authorization again", async () => {
    const view = await boot();
    await pickFolder(view, catalog());
    fireEvent.click(view.container.querySelector(".scope-trigger")!);
    await waitFor(() => expect(refreshFolder).toHaveBeenCalled());
    await waitFor(() => expect(view.container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("false"));
    fireEvent.click(view.getByRole("button", { name: "关闭数据范围选择" }));
    send(view, "检查已授权表");
    await waitFor(() => expect(checkIntent).toHaveBeenCalledOnce());
    expect(view.container.textContent).not.toContain("请先在上方「数据范围」中选择文件夹和工作表");
    expect(createFolderSnapshot).toHaveBeenCalledOnce();
    expect(JSON.stringify(vi.mocked(checkIntent).mock.calls)).not.toContain("C:/One/Data");
  });

  it("a follow-up cannot answer an old question against another same-named directory", async () => {
    const view = await boot();
    await pickFolder(view, catalog());
    vi.mocked(checkIntent).mockResolvedValueOnce({ kind: "clarification", provider: "model", clarification: {
      id: "question-one", summary: "请选择处理规则", question: "逐行还是汇总？", reason: "需要明确规则",
      scopeLabel: "book.xlsx › Selected", options: []
    } });
    send(view, "检查已授权表");
    await waitFor(() => expect(view.container.textContent).toContain("逐行还是汇总？"));
    await pickFolder(view, catalog("D:/Two/Data", "session-two"));
    send(view, "逐行处理");
    await waitFor(() => expect(view.container.textContent).toContain("数据范围已经变化，请检查范围后重新发送需求。"));
    expect(checkIntent).toHaveBeenCalledOnce();
  });

  it("failed session recovery removes consent and stops before field analysis", async () => {
    const view = await boot();
    await pickFolder(view, catalog());
    vi.mocked(checkIntent).mockResolvedValueOnce({ kind: "clarification", provider: "model", clarification: {
      id: "recover-question", summary: "请选择处理规则", question: "逐行还是汇总？", reason: "需要明确规则",
      scopeLabel: "book.xlsx › Selected", options: []
    } });
    send(view, "检查已授权表");
    await waitFor(() => expect(view.container.textContent).toContain("逐行还是汇总？"));
    vi.mocked(refreshFolder).mockRejectedValueOnce(new Error("文件夹会话已失效"));
    fireEvent.click(view.container.querySelector(".scope-trigger")!);
    await waitFor(() => expect(scanFolderByPath).toHaveBeenCalled());
    await waitFor(() => expect(view.container.textContent).toContain("文件夹会话已失效且无法自动重建"));
    fireEvent.click(view.getByRole("button", { name: "关闭数据范围选择" }));
    send(view, "检查已授权表");
    await waitFor(() => expect(view.container.textContent).toContain("请先在上方「数据范围」中选择文件夹和工作表。"));
    expect(checkIntent).toHaveBeenCalledOnce();
  });

  it("session recovery retains the latest checkbox authorization without a second confirmation", async () => {
    const view = await boot();
    await pickFolder(view, catalog());
    fireEvent.click(view.container.querySelector(".scope-trigger")!);
    await waitFor(() => expect(view.container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("false"));
    vi.mocked(refreshFolder).mockRejectedValueOnce(new Error("文件夹会话已失效"));
    vi.mocked(scanFolderByPath).mockResolvedValueOnce(catalog("C:/One/Data", "session-rebuilt"));
    fireEvent.click(view.container.querySelector(".folder-refresh-button")!);
    await waitFor(() => expect(createFolderSnapshot).toHaveBeenCalledTimes(2));
    expect(vi.mocked(createFolderSnapshot).mock.calls[1]).toEqual([
      "session-rebuilt", [{ fileId: "same-relative-file-id", sheets: ["Selected"] }]
    ]);
    fireEvent.click(view.getByRole("button", { name: "关闭数据范围选择" }));
    send(view, "检查已授权表");
    await waitFor(() => expect(checkIntent).toHaveBeenCalledOnce());
  });

  it("adopting a folder file already open in Excel reads only explicitly selected data", async () => {
    vi.stubGlobal("Office", { context: { document: { url: "file:///C:/One/Data/book.xlsx" } }, onReady: (cb: () => void) => cb() });
    const view = await boot();
    await pickFolder(view, catalog());
    // Boot, opening the selector, and explicit adoption all re-read names only.
    expect(captureWorkbookCatalog).toHaveBeenCalledTimes(3);
    expect(vi.mocked(captureWorkbook).mock.calls).toEqual([[['Selected']]]);
    expect(createFolderSnapshot).not.toHaveBeenCalled();
  });

  /*
   * ② 从文件入口选中"就是当前打开的这个工作簿"之后，已提交授权必须**当场**就位。
   *
   * 采用动作会切回工作簿来源并启用目标守卫；如果已提交范围要等一个被动 effect
   * 才落地，守卫启用的第一帧读到的是旧/空名单 —— 于是紧接着的一句需求被
   * "工作表范围发生变化"拦下，用户被要求把刚确认过的范围再确认一次。
   */
  it("②从文件入口采用当前工作簿后：下一条需求不再要求确认", async () => {
    vi.stubGlobal("Office", { context: { document: { url: "file:///C:/One/Data/book.xlsx" } }, onReady: (cb: () => void) => cb() });
    const view = await boot();
    await pickFolder(view, catalog());
    expect(view.container.textContent).toContain("所选文件就是当前工作簿");

    send(view, "检查已授权表");
    await waitFor(() => expect(checkIntent).toHaveBeenCalledOnce());
    // 需求真的跑通了（不是被拦下），而且没有出现范围提示。
    await waitFor(() => expect(view.container.textContent).toContain("字段已核对。"));
    expect(view.container.querySelector(".target-scope-notice")).toBeNull();
    expect(view.container.textContent).not.toContain("工作表范围发生变化");
  });

  it("从工作目录切回「选择多个」：面板回到真实当前工作簿，不再显示文件夹复合名", async () => {
    /*
     * 用户实测的缺陷：文件夹模式下确认授权后，`workbook` 已被换成**文件夹快照**
     * （表名是「文件.xlsx › Sheet」复合名），名字草稿也是它们。这时点「选择多个」，
     * 手工列表拿这份快照当当前工作簿渲染 —— 三张全是「文件或工作表已不存在」。
     * 修法：切回工作簿来源时重新读一次真实宿主，`workbook` 与名字草稿都拉回当前工作簿。
     * 已提交的授权不动（要换成工作簿范围仍需按「确认」）。
     */
    const view = await boot();
    await pickFolder(view, catalog());
    fireEvent.click(view.container.querySelector(".scope-trigger")!);

    fireEvent.click(view.getByRole("button", { name: /选择多个.*跨表/ }));

    // 判据一：手工列表必须是**真实当前工作簿**的表（Selected / Unselected）。
    // 列表会先用文件夹快照渲染一帧（重读宿主是异步的），所以必须**等到**真实表出现 ——
    // 在坏代码里它们永远不出现，这条 waitFor 会以超时变红。
    await waitFor(() => {
      const current = [...view.container.querySelectorAll(".workbook-sheet-list strong")]
        .map((node) => node.textContent ?? "");
      expect(current).toContain("Selected");
      expect(current).toContain("Unselected");
    });
    const names = [...view.container.querySelectorAll(".workbook-sheet-list strong")]
      .map((node) => node.textContent ?? "");
    expect(names.some((name) => name.includes("›"))).toBe(false);

    // 判据二：承认汇总不许再喊「文件或工作表已不存在」。
    await waitFor(() => {
      const summary = view.container.querySelector(".scope-admission-summary")?.textContent ?? "";
      expect(summary).not.toContain("已不存在");
      expect(summary).not.toContain("不可用");
    });

    // 判据三：草稿重置为当前工作簿的当前表（1 个工作表，不是文件夹那 3 张）。
    expect(view.container.querySelector(".scope-trigger")?.textContent).toContain("已固定 1 个工作表");
  });
});

describe("actual Excel workbook scope panel layout", () => {
  it("refresh rechecks even unchanged metadata and preserves the manual draft", async () => {
    const view = await boot();
    fireEvent.click(view.container.querySelector(".scope-trigger")!);
    await waitFor(() => expect(view.container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("false"));
    fireEvent.click(view.getByRole("button", { name: /选择多个.*跨表/ }));
    fireEvent.click(view.getByRole("button", { name: /^Selected/ }));
    await waitFor(() => expect(view.container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("false"));
    const before = vi.mocked(inspectOfficeScopeAdmission).mock.calls.length;
    fireEvent.click(view.container.querySelector(".scope-refresh")!);
    await waitFor(() => expect(vi.mocked(inspectOfficeScopeAdmission).mock.calls.length).toBeGreaterThan(before));
    expect(view.container.querySelector(".scope-admission-summary")?.textContent).toContain("已选 2 张表");
    expect(view.container.querySelector(".scope-admission-summary button")).toBeNull();
    fireEvent.click(view.getByRole("button", { name: "关闭数据范围选择" }));
    send(view, "检查已授权表");
    await waitFor(() => expect(view.container.textContent).toContain("字段已核对。"));
    expect(captureWorkbook).not.toHaveBeenCalled();
    expect(captureWorkbookStructure).toHaveBeenCalledOnce();
    expect(checkIntent).toHaveBeenCalledOnce();
    expect((view.container.querySelector(".composer textarea") as HTMLTextAreaElement).value).toBe("");
  });

  it("automatic refresh updates both labels and checks the new draft without reading fields", async () => {
    const view = await boot();
    fireEvent.click(view.container.querySelector(".scope-trigger")!);
    await waitFor(() => expect(view.container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("false"));
    fixtures.office.activeWorksheet = "Selected";
    fireEvent.click(view.container.querySelector(".scope-refresh")!);
    await waitFor(() => expect(view.container.querySelector(".workbook-identity em")?.textContent).toContain("当前工作表 Selected"));
    expect(view.container.querySelector(".scope-trigger strong")?.textContent).toBe("当前表 Selected");
    await waitFor(() => expect(vi.mocked(inspectOfficeScopeAdmission).mock.lastCall?.[0]).toEqual(["Selected"]));
    expect(captureWorkbook).not.toHaveBeenCalled();
    expect(checkIntent).not.toHaveBeenCalled();
  });

  it("子表选择面板紧跟「选择多个」，并由 open 类驱动展开/收起", async () => {
    const view = await boot();
    fireEvent.click(view.container.querySelector(".scope-trigger")!);
    await waitFor(() => expect(view.container.querySelector(".scope-popover")).not.toBeNull());

    const modeList = view.container.querySelector(".scope-mode-list")!;
    const modeButtons = () => Array.from(modeList.querySelectorAll<HTMLButtonElement>(":scope > button"));
    const manual = modeButtons().find((button) => (button.textContent ?? "").includes("选择多个"))!;
    const folder = modeButtons().find((button) => (button.textContent ?? "").includes("工作目录"))!;

    // 断言读的是**实际 DOM 位置**：子表面板必须在模式列表**内部**，
    // 而不是列表之后/之外的兄弟块 —— 否则它会掉到「工作目录」下面（修复点一）。
    const picker = modeList.querySelector(".workbook-sheet-picker");
    expect(picker).not.toBeNull();
    expect(modeList.contains(picker)).toBe(true);
    const reveal = modeList.querySelector(".scope-picker-reveal");
    expect(reveal).not.toBeNull();
    expect(picker!.closest(".scope-picker-reveal")).toBe(reveal);
    // 收起态：面板**常挂载**（否则切走时直接卸载，收缩无从播放），
    // 只是没有 `open`（CSS：0fr + opacity:0 + visibility:hidden，不进 Tab 序）。
    expect(reveal!.classList.contains("open")).toBe(false);

    fireEvent.click(manual);
    expect(reveal!.classList.contains("open")).toBe(true);

    // 切回「跟随当前」：必须有收起态，而不是把节点从 DOM 里摘掉。
    fireEvent.click(modeButtons().find((button) => (button.textContent ?? "").includes("跟随当前"))!);
    expect(reveal!.classList.contains("open")).toBe(false);
    expect(modeList.contains(picker)).toBe(true);

    const precedes = (before: Node, after: Node) =>
      Boolean(before.compareDocumentPosition(after) & Node.DOCUMENT_POSITION_FOLLOWING);
    // 顺序：选择多个 → 子表面板 → 工作目录。
    expect(precedes(manual, picker!)).toBe(true);
    expect(precedes(picker!, folder)).toBe(true);
  });
});

/*
 * Excel 工作目录的**固定页脚的层与格式** —— 与 WPS 的
 * 「WPS 工作目录：固定页脚的层与格式（对齐 Excel）」是同一件事的两个调用点。
 *
 * 这一条是结构判据，不是像素判据：
 *   · 正向：页脚是 `.scope-popover` 的**直接子节点**，确认与（可选的）取消都在页脚里，
 *           承认汇总紧跟在页脚**下方**；
 *   · 反向：页脚**不在**滚动容器里（`scroll.contains(footer) === false`）——
 *           放进滚动容器它就只是"内容末尾的一行"，内容一滚就跟着滚走。
 *
 * 为什么取消也要在这里钉：工作目录模式的授权也有"草稿 vs 已提交"（草稿 = 勾选，
 * 已提交 = 「确认」时才写进 `confirmedFolderScopeRef`）。没有出口时，误勾一张就只能
 * 关掉弹层 —— 那是用户**猜**出来的语义，不是界面上明示的。
 */
describe("Excel 工作目录：固定页脚的层与格式（对齐 WPS）", () => {
  it("页脚住在弹层上、不在滚动容器里，确认在页脚里，承认汇总落在它下方", async () => {
    const view = await boot();
    await pickFolder(view, catalog());
    fireEvent.click(view.container.querySelector(".scope-trigger")!);

    const popover = await waitFor(() => {
      const node = view.container.querySelector<HTMLElement>(".scope-popover");
      if (!node) throw new Error("数据范围弹层没有打开");
      return node;
    });
    const scroll = popover.querySelector<HTMLElement>(".scope-popover-scroll")!;
    const footer = popover.querySelector<HTMLElement>(".folder-selection-footer")!;
    expect(footer).not.toBeNull();

    // 反向（先判：失败原文里才会带出到底是谁被卷进了滚动容器）
    expect(scroll.querySelector(".folder-selection-footer")).toBeNull();
    expect(scroll.querySelector(".folder-selection-summary")).toBeNull();
    expect(scroll.querySelector(".scope-admission-summary")).toBeNull();

    // 正向：它直接住在弹层上。
    expect(footer.parentElement).toBe(popover);

    // 格式对齐 WPS：确认就在这条页脚里，承认汇总紧跟在页脚**下方**。
    expect(footer.querySelectorAll("button.scope-close")).toHaveLength(1);
    expect(footer.nextElementSibling?.classList.contains("scope-admission-summary")).toBe(true);
    // 工作目录模式**没有**工作簿那条动作条（否则一个面板会出现两条固定页脚）。
    expect(popover.querySelector(".scope-action-buttons")).toBeNull();

    // 刚确认过 ⇒ 草稿 == 已提交 ⇒ 页脚里还不该有取消。
    expect(footer.querySelector("button.scope-discard")).toBeNull();
  });

  it("勾选新增表立即生效，收起与重开保留两张表", async () => {
    const view = await boot();
    await pickFolder(view, catalog());
    fireEvent.click(view.container.querySelector(".scope-trigger")!);
    await waitFor(() => expect(view.container.querySelector(".folder-sheet-list")).not.toBeNull());
    fireEvent.click(view.getByRole("button", { name: /^Unselected/ }));
    await waitFor(() => expect(view.container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("false"));
    expect(vi.mocked(createFolderSnapshot).mock.lastCall?.[1]).toEqual([
      { fileId: "same-relative-file-id", sheets: ["Selected", "Unselected"] }
    ]);
    expect(view.container.querySelector(".scope-discard")).toBeNull();
    fireEvent.click(view.container.querySelector(".scope-close")!);
    fireEvent.click(view.container.querySelector(".scope-trigger")!);
    await waitFor(() => expect(view.container.querySelector(".folder-sheet-list")).not.toBeNull());
    expect(view.getByRole("button", { name: /^Unselected/ }).getAttribute("aria-pressed")).toBe("true");
  });

  it("「清空」一次清掉整个目录的草稿勾选，按钮在勾选归零后禁用", async () => {
    const view = await boot();
    await pickFolder(view, catalog()); // 已提交授权 = book.xlsx 的 Selected
    fireEvent.click(view.container.querySelector(".scope-trigger")!);

    await waitFor(() => expect(view.container.querySelector(".folder-sheet-list")).not.toBeNull());

    // 勾上第二张表，确保不止一张被清。
    fireEvent.click(view.getByRole("button", { name: /^Unselected/ }));
    await waitFor(() =>
      expect(view.container.querySelector(".folder-selection-footer strong")?.textContent)
        .toContain("2 张表")
    );

    fireEvent.click(view.getByRole("button", { name: "清空" }));

    // 草稿归零：页脚计数回到 0、清空按钮自身也因没得清而禁用。
    await waitFor(() =>
      expect(view.container.querySelector(".folder-selection-footer strong")?.textContent)
        .toContain("0 张表")
    );
    expect((view.getByRole("button", { name: "清空" }) as HTMLButtonElement).disabled).toBe(true);
  });
});
