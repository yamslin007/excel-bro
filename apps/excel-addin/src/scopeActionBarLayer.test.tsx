// @vitest-environment jsdom
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { cleanup, fireEvent, render, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import App from "./App";
import { captureWorkbookCatalog, captureWorkbookStructure } from "./excel";
import { inspectOfficeAuthorizedTarget, officeAnalysisTargetEvidence } from "./officeAuthorizedTarget";
import { checkIntent, understandMessage } from "./api";

/*
 * 数据范围弹层的**动作条层次**：Excel 与 WPS 的「完成 / 确认」按钮必须住在同一层。
 *
 * 这一条是结构判据，不是像素判据 —— 也正因为如此，jsdom 就够用：
 *   · 正向：动作条是 `.scope-popover` 的**直接子节点**，且紧跟在 `.scope-popover-scroll` 之后；
 *   · 反向：动作条**不在**滚动容器里面（`scroll.contains(bar) === false`）。
 *
 * 为什么非要两边都测：层的差别只由"在不在滚动容器里"决定。放进滚动容器，
 * `flex: 0 0 auto` 挡不住容器整体滚动，动作条会跟着内容滚出视口 —— 这正是
 * 「Excel 的完成按钮所在的层和 WPS 不一样」的根因。拆掉其中任一方向，
 * 这组用例必须变红，否则它就是恒真的装饰。
 */

const HERE = dirname(fileURLToPath(import.meta.url));

const fixtures = vi.hoisted(() => ({
  snapshot: {
    name: "book.xlsx",
    capturedAt: "2026-10-08T00:00:00Z",
    activeWorksheet: "Unselected",
    worksheets: ["Selected", "Unselected"].map((name) => ({
      name, usedRange: "A1:B3", rowCount: 3, columnCount: 2,
      headers: ["Category", "Value"], dataRows: [], truncated: false
    }))
  },
  target: {
    fileKey: "office:book-1",
    name: "book.xlsx",
    sheets: [{ id: "selected", name: "Selected" }, { id: "unselected", name: "Unselected" }],
    selected: ["unselected"]
  }
}));

vi.mock("./api", async (importOriginal) => ({
  ...await importOriginal<typeof import("./api")>(),
  checkHealth: vi.fn(async () => ({ status: "ok", modelConfigured: true })),
  listModels: vi.fn(async () => ({ defaultModelId: "model-1", models: [{ id: "model-1", label: "模型", modelId: "test-model", baseUrl: "http://localhost", supportsVision: true, available: true }] })),
  getModelSettings: vi.fn(async () => ({ connectionCount: 1 })),
  understandMessage: vi.fn(),
  checkIntent: vi.fn()
}));
vi.mock("./excel", async (importOriginal) => ({
  ...await importOriginal<typeof import("./excel")>(),
  isRunningInExcel: () => true,
  captureWorkbookCatalog: vi.fn(async () => ({
    ...fixtures.snapshot, catalogOnly: true,
    worksheets: fixtures.snapshot.worksheets.map((sheet) => ({ ...sheet, headers: [], dataRows: [] }))
  })),
  captureWorkbook: vi.fn(async () => fixtures.snapshot),
  captureWorkbookStructure: vi.fn(async () => fixtures.snapshot),
  watchWorkbookStructureChanges: vi.fn(async () => () => {}),
  currentOfficeWorkbookIdentity: () => "office:book-1"
}));
vi.mock("./officeAuthorizedTarget", async (importOriginal) => ({
  ...await importOriginal<typeof import("./officeAuthorizedTarget")>(),
  inspectOfficeAuthorizedTarget: vi.fn(async (names: readonly string[] | null) => ({ ...fixtures.target, selected: names === null ? ["unselected"] : names.map((name) => name.toLowerCase()) })),
  inspectOfficeScopeAdmission: vi.fn(async (names: readonly string[]) => ({
    items: names.map((id) => ({ id, status: "ready", reason: "none" })),
    target: { ...fixtures.target, selected: names.map((name) => name.toLowerCase()) }
  })),
  confirmOfficeAuthorizedScope: vi.fn(async (_names, confirm, expected) => confirm(expected)),
  officeAnalysisTargetEvidence: vi.fn(async () => ({ key: "target-1", conclusive: true }))
}));

beforeEach(() => {
  localStorage.clear();
  vi.clearAllMocks();
  vi.mocked(inspectOfficeAuthorizedTarget).mockImplementation(async (names) => ({ ...fixtures.target, selected: names === null ? ["unselected"] : names.map((name) => name.toLowerCase()) }));
  vi.mocked(officeAnalysisTargetEvidence).mockResolvedValue({ key: "target-1", conclusive: true });
  vi.mocked(understandMessage).mockResolvedValue({ kind: "conversation", message: "已收到。", provider: "model" });
  vi.mocked(checkIntent).mockResolvedValue({ kind: "conversation", message: "已收到。", provider: "model" });
  vi.spyOn(console, "log").mockImplementation(() => {});
  vi.stubGlobal("Office", { context: { document: { url: "file:///C:/Data/book.xlsx" } }, onReady: (cb: () => void) => cb() });
  Element.prototype.scrollIntoView = vi.fn();
});
afterEach(() => { cleanup(); vi.unstubAllGlobals(); vi.restoreAllMocks(); });

async function boot() {
  const view = render(<App />);
  await waitFor(() => expect(captureWorkbookCatalog).toHaveBeenCalled());
  await waitFor(() => expect(view.container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title).toContain("当前模型"));
  return view;
}

function openPicker(view: ReturnType<typeof render>): HTMLElement {
  if (!view.container.querySelector(".scope-popover")) {
    fireEvent.click(view.container.querySelector(".scope-trigger")!);
  }
  const popover = view.container.querySelector<HTMLElement>(".scope-popover");
  if (!popover) throw new Error("数据范围弹层没有打开");
  return popover;
}

function barOf(popover: HTMLElement): HTMLElement {
  const bar = popover.querySelector<HTMLElement>(".scope-action-buttons");
  if (!bar) throw new Error("动作条不在弹层里");
  return bar;
}

function labelsIn(root: ParentNode): string[] {
  return Array.from(root.querySelectorAll<HTMLButtonElement>("button"))
    .map((button) => (button.textContent ?? "").replace(/\s+/g, " ").trim());
}

function sheetRow(container: HTMLElement, name: string): HTMLButtonElement {
  const row = Array.from(container.querySelectorAll<HTMLButtonElement>(".workbook-sheet-list > button"))
    .find((candidate) => candidate.querySelector("strong")?.textContent === name);
  if (!row) throw new Error(`工作表行「${name}」不存在`);
  return row;
}

function css(path: string): string {
  return readFileSync(path, "utf8");
}

/* 注释里的同类选择器不算数：判据只对**生效的声明**负责。 */
function withoutCssComments(source: string): string {
  return source.replace(/\/\*[\s\S]*?\*\//g, "");
}

describe("数据范围动作条的层：Excel 必须与 WPS 同层", () => {
  it("动作条是滚动容器的兄弟，不是它的内容", async () => {
    const view = await boot();
    const popover = openPicker(view);
    const scroll = popover.querySelector<HTMLElement>(".scope-popover-scroll");
    const bar = barOf(popover);
    expect(scroll).not.toBeNull();

    // 正向：直接住在弹层上，并且**紧跟在滚动容器后面** —— 只有这个位置才被
    // `.scope-popover`（column flex + overflow: hidden）钉成固定页脚。
    expect(bar.parentElement).toBe(popover);
    expect(bar.previousElementSibling).toBe(scroll);

    // 反向：不在滚动容器里面。这是"层不一样"的唯一判别点。
    expect(scroll!.contains(bar)).toBe(false);
    expect(scroll!.querySelector(".scope-action-buttons")).toBeNull();
    expect(scroll!.querySelector(".scope-confirm")).toBeNull();
  });

  it("两端读同一份页脚声明，WPS 侧不再自己写一遍", () => {
    const shared = withoutCssComments(css(join(HERE, "styles.css")));
    const wps = withoutCssComments(css(join(HERE, "..", "..", "wps-addin", "src", "styles.css")));

    const rule = shared.match(/\.scope-action-buttons,\s*\n\.wps-scope-actions\s*\{([^}]*)\}/)?.[1] ?? "";
    expect(rule, "共享样式表里应有 `.scope-action-buttons, .wps-scope-actions` 这一条").not.toBe("");
    for (const declaration of [
      "display: flex",
      "flex: 0 0 auto",
      "justify-content: flex-end",
      "gap: 7px",
      "border-top: 1px solid var(--color-border)",
      "padding: 10px var(--spacing-md)"
    ]) {
      expect(rule, declaration).toContain(declaration);
    }

    // 反向：WPS 侧不许再留一条同名规则，否则两端又变成各写一套、各自漂移。
    expect(wps).not.toMatch(/^\s*\.wps-scope-actions\s*\{/m);
    // 正向锚点：同一个文件里别的规则照常在，证明上面的"不匹配"不是因为读错文件。
    expect(wps).toMatch(/^\s*\.wps-backup-card\s*\{/m);
    expect(shared).toMatch(/^\.scope-action-buttons,$/m);
  });

  it("授权说明在固定标题，动作条只有刷新与确认，勾选仍立即生效", async () => {
    const view = await boot();
    const popover = openPicker(view);
    expect(labelsIn(barOf(popover))).toEqual(["↻ 刷新", "确认"]);
    fireEvent.click(view.getByRole("button", { name: /选择多个/ }));
    fireEvent.click(sheetRow(view.container, "Selected"));
    await waitFor(() => expect(view.container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("false"));
    expect(labelsIn(barOf(popover))).toEqual(["↻ 刷新", "确认"]);
    expect(popover.querySelector(":scope > .scope-popover-header .scope-authorization-hint")?.textContent).toContain("勾选即授权");
    expect(barOf(popover).textContent).not.toContain("勾选即授权");
    expect(view.container.querySelector(".scope-confirm")).toBeNull();
  });

  it("确认关闭面板并保留勾选和固定模式，重开不撤销授权", async () => {
    const view = await boot();
    openPicker(view);
    fireEvent.click(view.getByRole("button", { name: /选择多个/ }));
    fireEvent.click(sheetRow(view.container, "Selected"));
    await waitFor(() => expect(view.container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("false"));
    fireEvent.click(view.container.querySelector(".scope-close")!);
    await waitFor(() => expect(view.container.querySelector(".scope-popover")).toBeNull());
    expect(view.container.querySelector(".scope-trigger")?.textContent).toContain("已固定 2 个工作表");
    openPicker(view);
    expect(sheetRow(view.container, "Selected").getAttribute("aria-pressed")).toBe("true");
    expect(sheetRow(view.container, "Unselected").getAttribute("aria-pressed")).toBe("true");
  });

  it("错误条排在页脚之后，与 WPS 的 scopeError 同位置", async () => {
    const view = await boot();
    const popover = openPicker(view);
    fireEvent.click(view.getByRole("button", { name: /选择多个/ }));
    await waitFor(() => expect(view.container.querySelector(".workbook-sheet-list")).not.toBeNull());
    fireEvent.click(sheetRow(view.container, "Selected"));

    vi.mocked(officeAnalysisTargetEvidence).mockResolvedValue(null);
    const error = await waitFor(() => {
      const node = view.container.querySelector<HTMLElement>(".authorized-target-error");
      if (!node) throw new Error("错误条尚未出现");
      return node;
    });
    expect(captureWorkbookStructure).not.toHaveBeenCalled();
    expect(view.queryByRole("dialog")).toBeNull();

    // 判据：错误条是弹层的**直接子节点**，且紧跟在动作条之后 ——
    // 它在固定页脚**外面**、下面，不会被页脚盖住，也不跟着内容滚。
    // 直接查 `.authorized-target-error`（不写 `>`）是刻意的：写 `>` 时
    // "被塞进页脚里"这个回归只会让选择器落空、判据以超时方式变红，
    // 看不出是哪条口径坏了；不写 `>` 才会落在下面这三条断言上。
    expect(error.parentElement).toBe(view.container.querySelector(".scope-popover"));
    expect(error.closest(".scope-action-buttons")).toBeNull();
    expect(error.previousElementSibling?.classList.contains("scope-action-buttons")).toBe(true);
    expect(error.nextElementSibling).toBeNull();
  });

  it("文件夹主从区是进入式单栏：宽屏两栏覆盖不许复活，返回按钮不许被藏", () => {
    /*
     * 「点文件 → 进入该文件 → 勾它的子表（← 文件 返回）」靠的是**窄屏形态**那几条基础规则：
     * `.folder-picker-view-sheets .folder-file-list { display: none }` 把文件列表藏掉、
     * `.folder-back-button` 露出来。原来还有一段 `@media (min-width: 520px)` 的覆盖在宽面板
     * （WPS 任务窗格）把它翻成两栏 —— 文件列表一直挂着、返回按钮被藏，于是 WPS 是
     * "左边勾文件、右边勾子表"，和 Excel 用户的交互对不上。2026-10-16 按用户口径整段删掉，
     * 两端同一份声明、任何宽度都走进入式单栏。这条判据读的是**剥掉注释后的生效文本**，
     * 防止那段覆盖换了个写法悄悄回来。
     */
    const shared = withoutCssComments(css(join(HERE, "styles.css")));
    const wps = withoutCssComments(css(join(HERE, "..", "..", "wps-addin", "src", "styles.css")));

    // 反向：两栏覆盖不许复活 —— 无论是那段媒体查询还是它的两栏声明。
    expect(shared).not.toContain("@media (min-width: 520px)");
    expect(shared).not.toContain("minmax(160px, 0.8fr)");
    // 反向：返回按钮不许再被藏（否则进了文件却回不去，只能靠重开面板）。
    expect(shared).not.toMatch(/\.folder-back-button\s*\{[^}]*display:\s*none/);
    // 反向：WPS 侧不许自己写一遍同样的覆盖（两端一份声明，不许分叉）。
    expect(wps).not.toContain("@media (min-width: 520px)");
    expect(wps).not.toMatch(/\.folder-master-detail\s*\{/);

    // 正向锚点：证明上面的"不匹配"不是因为读错了文件或删错了机制 ——
    // 进入式单栏**依赖**的那条规则必须还在。
    expect(shared).toMatch(
      /\.folder-master-detail\.folder-picker-view-sheets \.folder-file-list\s*\{[^}]*display:\s*none/
    );
    expect(shared).toMatch(/\.folder-back-button\s*\{/);
  });
});
