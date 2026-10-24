/**
 * 定向测试：WPS 数据范围列表的**选项行**该长什么样。
 *
 * 三件事（都是用户直接看出来的）：
 *  ① 行内**不再**出现授权状态文案（「可分析 · 修改仍需预览确认」这一串）。
 *     它是 `SCOPE_ADMISSION_MESSAGES["none"]`，原先只由行内徽标渲染；
 *     授权状态是**整次选择**的属性，现在只在列表下方的 `ScopeAdmissionSummary`
 *     汇总一次。判据不是"文案被删掉了"，而是"名字还在、汇总还在、行内没有它"。
 *  ② 行内勾选框必须用**共享**的 `.sheet-checkbox`（与 Excel 同一份声明）。
 *     WPS 侧若再写一条 `.wps-sheet-picker-options > button i`
 *     （特异度 0-1-2 高过 `.sheet-checkbox` 的 0-1-0），共享声明会被整个盖掉，
 *     两端又变成各写一套 —— 那正是观感漂移的来源。
 *  ③ 行高与模式卡（跟随当前 / 选择多个 / 工作目录）读**同一个** token。
 *     原先列表行自定 38px，比上面那几张卡矮一大截，用户口径是"太细了"。
 *
 * 判据分两层（jsdom 没有布局引擎，量不出像素高度）：
 *  ① 这里钉**行结构**与**样式契约**；
 *  ② **真实像素高度**由 `scripts/check_scope_option_rows.mjs` 在无头浏览器里
 *     挂**同一个面板**实测，断言"行高 == 工作目录框高"，并带 `--legacy`
 *     反向对照证明测量有判别力。**真机 WPS 的像素验收未执行。**
 */
import { act, useEffect, type ReactElement, type ReactNode } from "react";
import type { AuthorizedTarget } from "@excel-bro/core/authorizedTarget";
import { createRoot } from "react-dom/client";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
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
/** 面板本体换替身：只留出范围草稿面板那条槽位，避免把聊天面板一起拉起来。 */
/*
 * 替身必须是**有名字的大写函数**，不能写成匿名 `default: (...) => ...`：
 * 里面调了 `useEffect`，匿名箭头挂在 `default` 上时 rules-of-hooks 认不出它是组件，
 * lint 直接判 error（"called in function 'default'"）。名字只是给 linter 看的，
 * 行为一个字都没变。
 */
function MockWpsAgentPanel({
  scopeDock,
  scopeOverlay,
  onScopeConfirmationReady
}: {
  selectedSheetIndexes?: readonly number[];
  scopeDock?: ReactNode;
  scopeOverlay?: ReactNode;
  onScopeConfirmationReady?: (handler: ((target: AuthorizedTarget) => Promise<boolean>) | null) => void;
}) {
  useEffect(() => { onScopeConfirmationReady?.(scopeAuthorization); return () => onScopeConfirmationReady?.(null); }, [onScopeConfirmationReady]);
  return (
    <div>
      {scopeDock}
      {scopeOverlay ? <div className="scope-layer">{scopeOverlay}</div> : null}
    </div>
  );
}
vi.mock("./WpsAgentPanel", () => ({ default: MockWpsAgentPanel }));

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

function sheetRow(container: HTMLElement, name: string): HTMLButtonElement {
  const row = Array.from(
    container.querySelectorAll<HTMLButtonElement>(".wps-sheet-picker-options button")
  ).find((button) => button.querySelector(":scope > span")?.textContent === name);
  if (!row) throw new Error(`工作表「${name}」的选项行尚未出现`);
  return row;
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

const here = dirname(fileURLToPath(import.meta.url));
const wpsStyles = readFileSync(join(here, "styles.css"), "utf8");
/** 跨端读另一端的样式表。`new URL("../..")` 会吞掉一层，必须用 `join` 逐段拼。 */
const sharedStyles = readFileSync(
  join(here, "..", "..", "excel-addin", "src", "styles.css"),
  "utf8"
);

/** 取一条规则的声明块。用 `{`/`}` 定界，断言不跨规则。 */
function ruleBody(source: string, selector: string): string {
  const start = source.indexOf(`${selector} {`);
  if (start < 0) throw new Error(`styles.css 里找不到规则：${selector}`);
  const end = source.indexOf("}", start);
  if (end < 0) throw new Error(`规则没有闭合：${selector}`);
  return source.slice(start, end);
}

/**
 * 剥掉注释再断言"某条规则不存在"。
 * 注释里为了讲清楚"为什么不能这么写"会引用选择器名，不剥注释就会把
 * **解释**当成**实现**判红（或反之放行）。
 */
function withoutComments(source: string): string {
  return source.replace(/\/\*[\s\S]*?\*\//g, "");
}

/** 读一个自定义属性的值（剥注释后取第一处声明）。`{`/`}` 定界在 `:root` 上不可靠，故用它。 */
function tokenValue(source: string, name: string): string | undefined {
  const hit = withoutComments(source).match(new RegExp(`${name}\\s*:\\s*([^;]+);`));
  return hit?.[1]?.trim();
}

/** 打开范围面板、切到「选择多个」。返回容器。 */
async function openManualPicker() {
  const { container } = await mount(<WpsProbePanel />);
  await click(byText(container, ".scope-trigger", ""));
  await click(byText(container, ".scope-mode-list button", "选择多个"));
  return container;
}

describe("WPS 工作表列表：行内只有勾选框与表名", () => {
  it("勾选前后行内都不出现授权文案，但表名与汇总行都还在", async () => {
    const container = await openManualPicker();

    const before = sheetRow(container, "明细");
    // 行内不许有徽标元素，也不许有那句话的任何一个字。
    expect(before.querySelector(".scope-admission-badge")).toBeNull();
    expect(before.textContent ?? "").not.toContain("可分析");
    expect(before.textContent ?? "").not.toContain("修改仍需预览确认");
    // 基线锚点：这一行本来就没有徽标，所以"勾选后没有"不能靠"本来就没有"混过去。
    expect(before.getAttribute("aria-pressed")).toBe("false");

    await click(before);
    const after = sheetRow(container, "明细");

    expect(after.getAttribute("aria-pressed")).toBe("true");
    expect(after.querySelector(".scope-admission-badge")).toBeNull();
    expect(after.textContent ?? "").not.toContain("可分析");
    expect(after.textContent ?? "").not.toContain("修改仍需预览确认");

    // 正向锚点：不是把行内容删空换来的"没有文案"。
    const name = after.querySelector<HTMLElement>(":scope > span");
    expect(name?.textContent).toBe("明细");
    // 勾选框仍然给出了勾选反馈（空 → ✓）。
    const checkbox = after.querySelector<HTMLElement>(":scope > i");
    expect(checkbox?.classList.contains("sheet-checkbox")).toBe(true);
    expect(checkbox?.textContent).toBe("✓");

    // 授权状态整体没有消失，只是下沉到列表下方的汇总行 —— 而且**不在行内**。
    // 只钉"汇总行还在 + 它不在列表里"，不钉汇总的具体措辞：
    // 措辞随核实状态变化（正在检查…／不可用／可分析），那不是这条判据管的事。
    const summary = container.querySelector<HTMLElement>(".scope-admission-summary");
    expect(summary).not.toBeNull();
    expect(summary!.textContent ?? "").toMatch(/已选 \d+ 张表/);
    expect(container.querySelector(".wps-sheet-picker-options .scope-admission-summary")).toBeNull();
  });

  it("行内勾选框只认共享的 .sheet-checkbox，WPS 侧不再各写一套", () => {
    // 行结构与列宽的第一列都是 20px（勾选框宽度）。
    const row = ruleBody(wpsStyles, ".wps-sheet-picker-options > button");
    expect(row).toContain("grid-template-columns: 20px minmax(0, 1fr) auto");

    // WPS 侧那条高特异度的行内 `i` 规则必须**不存在**：它会把共享声明整个盖掉。
    const wpsRules = withoutComments(wpsStyles);
    expect(wpsRules).not.toContain(".wps-sheet-picker-options > button i {");
    expect(wpsRules).not.toContain(".wps-sheet-picker-options > button.selected i {");
    // 行内标签/徽标的三个覆盖规则也随行内容一起下线。
    expect(wpsRules).not.toContain(".wps-sheet-option-label");
    expect(wpsRules).not.toContain(".wps-sheet-option-name");
    expect(wpsRules).not.toContain("wps-sheet-picker-options > button .scope-admission-badge");

    // 共享声明留在另一端：空框不填色（填了会隐进行底色里），勾上才整块换品牌绿。
    const box = ruleBody(sharedStyles, ".sheet-checkbox");
    expect(box).toContain("width: 20px");
    expect(box).toContain("border: 1px solid var(--color-border-strong)");
    expect(box).toContain("background: transparent");
    const picked = ruleBody(sharedStyles, "button.selected > .sheet-checkbox");
    expect(picked).toContain("background: var(--color-brand-text)");
    expect(picked).toContain("color: var(--color-background)");
  });

  it("行高与模式卡读同一个 token（不再各自写 38px）", () => {
    expect(tokenValue(sharedStyles, "--scope-option-height")).toBe("59px");

    const row = ruleBody(wpsStyles, ".wps-sheet-picker-options > button");
    expect(row).toContain("min-height: var(--scope-option-height)");
    expect(row).not.toContain("min-height: 38px");

    // 两端模式卡：同一个 token，所以"统一"是被一份声明保证的。
    expect(ruleBody(sharedStyles, ".scope-mode-list > button")).toContain(
      "min-height: var(--scope-option-height)"
    );
    expect(ruleBody(sharedStyles, ".workbook-sheet-list > button")).toContain(
      "min-height: var(--scope-option-height)"
    );
  });
});
