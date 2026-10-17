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
  inspectOfficeAuthorizedTarget: vi.fn(async () => fixtures.target),
  inspectOfficeScopeAdmission: vi.fn(async (names: readonly string[]) => ({
    items: names.map((id) => ({ id, status: "ready", reason: "none" })),
    target: fixtures.target
  })),
  confirmOfficeAuthorizedScope: vi.fn(async (_names, confirm, expected) => confirm(expected)),
  officeAnalysisTargetEvidence: vi.fn(async () => ({ key: "target-1", conclusive: true }))
}));

beforeEach(() => {
  localStorage.clear();
  vi.clearAllMocks();
  vi.mocked(inspectOfficeAuthorizedTarget).mockResolvedValue(fixtures.target);
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

  it("按钮集合对齐 WPS：刷新 / 确认；草稿与已提交授权不一致时才出现取消", async () => {
    const view = await boot();
    const popover = openPicker(view);

    // 初始草稿 == 已提交授权 ⇒ 只有刷新 / 确认。
    expect(labelsIn(barOf(popover))).toEqual(["↻ 刷新", "确认"]);
    // 负向：旧文案「完成」不许复活在动作条里（它与「确认」是同一个动作，
    // 两端叫法必须一致，否则用户看到的是两个不同层次的东西）。
    expect(labelsIn(barOf(popover))).not.toContain("完成");
    expect(view.container.textContent).not.toContain("确认范围");
    expect(view.container.textContent).not.toContain("放弃修改");

    // 造出草稿差异：手动模式下多勾一张表。
    fireEvent.click(view.getByRole("button", { name: /选择多个/ }));
    await waitFor(() => expect(view.container.querySelector(".workbook-sheet-list")).not.toBeNull());
    fireEvent.click(sheetRow(view.container, "Selected"));
    await waitFor(() => expect(labelsIn(barOf(popover))).toContain("取消"));
    expect(labelsIn(barOf(popover))).toEqual(["↻ 刷新", "确认", "取消"]);
  });

  it("取消把草稿退回已提交授权并关闭弹层：已提交的范围一字不改", async () => {
    const view = await boot();
    const popover = openPicker(view);
    fireEvent.click(view.getByRole("button", { name: /选择多个/ }));
    await waitFor(() => expect(view.container.querySelector(".workbook-sheet-list")).not.toBeNull());

    // 已提交授权仍是"当前表 Unselected"（跟随当前表建立的）。
    const committedLabel = view.container.querySelector(".scope-trigger")?.textContent ?? "";
    expect(committedLabel).toContain("已固定 1 个工作表");
    expect(sheetRow(view.container, "Unselected").getAttribute("aria-pressed")).toBe("true");

    fireEvent.click(sheetRow(view.container, "Selected"));
    expect(sheetRow(view.container, "Selected").getAttribute("aria-pressed")).toBe("true");
    expect(view.container.querySelector(".scope-trigger")?.textContent).toContain("已固定 2 个工作表");

    const discard = barOf(popover).querySelector<HTMLButtonElement>(".scope-discard");
    expect(discard).not.toBeNull();
    fireEvent.click(discard!);

    // 判据一：弹层关闭 —— 与 WPS 的 `discardScopeEdit` 同一条出口语义（用户裁决
    // 2026-10-16：取消即离开，不是"我点开只是想看看"）。坏代码里弹层还开着，
    // 这条 waitFor 以超时变红。
    await waitFor(() => expect(view.container.querySelector(".scope-popover")).toBeNull());

    // 判据二：草稿退回已提交授权 —— 弹层关了，触发器（一直挂着）退回"已固定 1 个"。
    expect(view.container.querySelector(".scope-trigger")?.textContent).toContain("已固定 1 个工作表");

    // 判据三：重开面板，勾选状态也要退回去 —— 不只按钮消失，Selected 不再勾着。
    fireEvent.click(view.container.querySelector(".scope-trigger")!);
    await waitFor(() => expect(view.container.querySelector(".workbook-sheet-list")).not.toBeNull());
    expect(sheetRow(view.container, "Selected").getAttribute("aria-pressed")).toBe("false");
    expect(sheetRow(view.container, "Unselected").getAttribute("aria-pressed")).toBe("true");
  });

  it("错误条排在页脚之后，与 WPS 的 scopeError 同位置", async () => {
    const view = await boot();
    const popover = openPicker(view);
    fireEvent.click(view.getByRole("button", { name: /选择多个/ }));
    await waitFor(() => expect(view.container.querySelector(".workbook-sheet-list")).not.toBeNull());
    fireEvent.click(sheetRow(view.container, "Selected"));

    // 带着未确认的范围修改发送 ⇒ 停在原地给出错误条（不读字段、不弹第二个对话框）。
    // 注意顺序：产品是先理解消息、再在"要不要读字段"这道闸前面停下，
    // 所以这里必须让理解结果落在"需要工作簿上下文"上，才真的走到那道闸。
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "看看这列", provider: "model" });
    fireEvent.change(view.container.querySelector(".composer textarea")!, { target: { value: "看看这列" } });
    fireEvent.click(view.getByRole("button", { name: /^发送$/ }));

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
