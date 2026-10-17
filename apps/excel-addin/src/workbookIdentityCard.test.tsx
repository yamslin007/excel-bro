// @vitest-environment jsdom
import { cleanup, fireEvent, render, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import App from "./App";
import { captureWorkbookCatalog } from "./excel";
import { inspectOfficeAuthorizedTarget } from "./officeAuthorizedTarget";
import { checkIntent, understandMessage } from "./api";

/*
 * 「不显示报表日期」——**识别这件事本身已经下线**（2026-10-16 产品决定：用户原话「不需要 识别日期」）。
 *
 * 这条口径走过两步，别把第二步读成第一步的重复：
 *   ① 用户先说「文件名中未识别到报表日期，这个文字不应该出现呀」——
 *      那一版只删了**否定式说明**，文件名里有日期时**照旧显示**「报表日期 2026-03-15」；
 *   ② 用户接着说「不需要 识别日期」—— `extractWorkbookDataPeriod` 与它派生出的
 *      `workbookDataPeriod` 因此**整条删除**，两个调用点（当前文件卡 / 数据范围触发器）
 *      都只剩"哪个文件、哪张表"。
 * 所以判据不能只测"没有日期的那支"：那样把 `extractWorkbookDataPeriod` 请回来、
 * 只保留"有则显示"也照样绿。**必须拿一个文件名里真有日期的表来测"不显示"**。
 *
 * 两个调用点都要在场，缺一个就是漏了半边：
 *   ① 弹层里的当前文件卡（`div.workbook-identity em`）；
 *   ② 收起来的触发器（`.scope-trigger strong`，文案是「当前表 X」/「已固定 N 个工作表」）。
 * 只钉住①，把日期加回②的变异会全身而过。
 *
 * 正向锚点在**每一条负向断言旁边**：卡片仍在、文件名的 `strong` 仍在、`em` 仍在、
 * 所述工作表名仍在 —— 否则"删掉整行"就能把负向断言骗绿。
 */

const fixtures = vi.hoisted(() => ({
  snapshot: {
    name: "评分表.xlsx",
    capturedAt: "2026-10-08T00:00:00Z",
    activeWorksheet: "Sheet2",
    worksheets: ["Sheet1", "Sheet2"].map((name) => ({
      name, usedRange: "A1:B3", rowCount: 3, columnCount: 2,
      headers: ["Category", "Value"], dataRows: [], truncated: false
    }))
  },
  target: {
    fileKey: "office:book-1",
    name: "评分表.xlsx",
    sheets: [{ id: "sheet-1", name: "Sheet1" }, { id: "sheet-2", name: "Sheet2" }],
    selected: ["sheet-2"]
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
  fixtures.snapshot.name = "评分表.xlsx";
  vi.mocked(inspectOfficeAuthorizedTarget).mockResolvedValue(fixtures.target);
  vi.mocked(understandMessage).mockResolvedValue({ kind: "conversation", message: "已收到。", provider: "model" });
  vi.mocked(checkIntent).mockResolvedValue({ kind: "conversation", message: "已收到。", provider: "model" });
  vi.spyOn(console, "log").mockImplementation(() => {});
  vi.stubGlobal("Office", { context: { document: { url: "file:///C:/Data/%E8%AF%84%E5%88%86%E8%A1%A8.xlsx" } }, onReady: (cb: () => void) => cb() });
  Element.prototype.scrollIntoView = vi.fn();
});
afterEach(() => { cleanup(); vi.unstubAllGlobals(); vi.restoreAllMocks(); });

/** 打开数据范围弹层，返回当前文件卡。触发器文案在打开**之前**就已经可读。 */
async function openCard(view: ReturnType<typeof render>): Promise<HTMLElement> {
  await waitFor(() => expect(captureWorkbookCatalog).toHaveBeenCalled());
  await waitFor(() => expect(view.container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title).toContain("当前模型"));
  fireEvent.click(view.container.querySelector(".scope-trigger")!);
  return waitFor(() => {
    const node = view.container.querySelector<HTMLElement>(".workbook-identity");
    if (!node) throw new Error("当前文件卡没有出现");
    return node;
  });
}

/** 两个调用点共用的断言：卡片第二行 + 触发器文案，都不含日期。 */
function expectNoDateAnywhere(view: ReturnType<typeof render>, card: HTMLElement, name: string) {
  // 正向锚点：该在的都还在，负向断言才不是"整行/整块被删掉"的假绿。
  expect(card.querySelector("strong")?.textContent).toBe(name);
  const line = card.querySelector("em");
  expect(line).not.toBeNull();
  expect(line!.textContent).toBe("当前工作表 Sheet2");
  const trigger = view.container.querySelector(".scope-trigger strong")!.textContent;
  expect(trigger).toBe("当前表 Sheet2");

  // ① 当前文件卡：不许出现日期、不许出现"未识别到"、不许留下悬空的「·」。
  expect(card.textContent).not.toContain("报表日期");
  expect(card.textContent).not.toContain("未识别到");
  expect(line!.textContent).not.toContain("·");
  // ② 触发器：日期一旦被拼回来，这里会立刻出现 `2026-03-15 · ` 这种前缀。
  expect(trigger).not.toMatch(/\d{4}/);
  // 正向锚点②：整块面板里也不许有这句话（它可能被搬到别的行继续存在）。
  expect(view.container.textContent).not.toContain("报表日期");
  expect(view.container.textContent).not.toContain("未识别到报表日期");
}

describe("当前文件卡与触发器：不显示、也不识别文件名里的日期", () => {
  it("文件名里没有日期：卡片与触发器都只有文件与工作表", async () => {
    fixtures.snapshot.name = "评分表.xlsx";
    const view = render(<App />);
    const card = await openCard(view);
    expectNoDateAnywhere(view, card, "评分表.xlsx");
  });

  it("文件名里有日期也不显示 —— 识别整条已下线，不是「有则显示」", async () => {
    // 这个名字按老口径会被识别成 `2026-03-15`；新口径下它和普通文件名没有区别。
    fixtures.snapshot.name = "评分表-2026-03-15.xlsx";
    const view = render(<App />);
    const card = await openCard(view);
    expectNoDateAnywhere(view, card, "评分表-2026-03-15.xlsx");
  });
});
