// @vitest-environment jsdom
//
// 这条用例专门盯**真实出站请求体**：`local_plan_rule` 路径算出的计划不能以任何
// 形式离开本机。与 `WpsAgentPanel.test.tsx` 里 `expect(validatePlanMock).not
// .toHaveBeenCalled()` 的区别在于：
//
// - 那种断言只证明"没调用某个函数"，挡不住实现改成直接 `fetch`、或在别的
//   路径上把同样的矩阵发出去；
// - 这里在 `globalThis.fetch` 这一层拦截，抓的是**真正上线的字节**。
//
// 所以本文件刻意**不 mock `../../excel-addin/src/api`** —— 让真实的
// `checkIntent` / `createAssistantResponse` / `validatePlan` 跑起来，只把
// `fetch` 换成记录器。这样 `/api/turn` 的请求体是真实序列化结果。
//
// ---------------------------------------------------------------------------
// 本文件分两条**语义互斥**的用例，各自对应一种真实产物（必须分开写，不能合并）：
//
//   A. 未知私有值 —— 源区域里存在一个**规则没有列进映射表**的非空值。
//      共享的 `mapColumnValues` 刻意"拒绝而非兜底"（抛 `UnmappedCategoryError`），
//      `blankLabel` 只对**真空格**生效。所以本地计划整体失败、**没有写入预览**。
//      这条路径要钉的是"失败时的出站面"：连一个"算不出来"的值都不会被送出去。
//
//   B. 已知映射值 —— 规则列出了源区域里**全部**非空来源类别，本地计划算出并
//      生成两行预览。这条路径要钉的是"成功时的出站面"：算好的目标矩阵、源快照
//      与完整 actions[].values 都不上传，且首次规划与编辑后都不走校验接口。
//
// 早先版本把这两件事塞进同一条用例：让哨兵同时出现在 `mapping.source`
// （好让孩子能算出来、预览能生成），却仍称它是"服务端/规则/提示都不知道的私有值"。
// 那两句话不能同时为真 —— 规则里白纸黑字写了它，出站负向断言就退化成自查自证。
// 更糟的是，为了让断言看着成立，会诱使生产代码把未命中值悄悄兜底成 `blankLabel`，
// 那是**降低**安全语义去迁就测试。现在拆开：未知值就必须失败，已知值才谈成功。
import { readFileSync } from "node:fs";
import { act, type ReactElement } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import WpsAgentPanel from "./WpsAgentPanel";
import { getWpsApplication } from "./wpsHost";

vi.mock("./wpsHost", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./wpsHost")>();
  return { ...actual, getWpsApplication: vi.fn(() => ({})) };
});

(
  globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }
).IS_REACT_ACT_ENVIRONMENT = true;

if (!Element.prototype.scrollIntoView) {
  Element.prototype.scrollIntoView = () => undefined;
}

// 样式只为贴合真实 DOM 结构，不参与断言。
readFileSync("src/styles.css", "utf8");

const getWpsApplicationMock = vi.mocked(getWpsApplication);

/**
 * 私有哨兵值：**只**存在于工作簿的一条数据行里。
 *
 * 它必须满足三个"不知道"：
 * - 服务端规则（`localRuleTurnResponseWithoutSentinel().rule.mapping`）不知道它；
 * - 服务端响应不知道它；
 * - 用户提示（prompt）不知道它。
 *
 * 三者缺一，"请求体不含哨兵"就会退化成必然成立的空过断言。注意：**它绝不能出现
 * 在 `mapping.source` 里** —— 那会让"规则不知道它"当场变成假话。需要它参与计算
 * 才能生成预览时，用另一条已知映射值用例，不要把这个哨兵塞进规则。
 */
const PRIVATE_SENTINEL = "内部代号-7Q4X-仅存在于数据行";
/** 规则真正会命中的那条已知类别（表头里就有，属于可出站的字段/值域）。 */
const KNOWN_CATEGORY = "办公用品";
/** 第二条已知类别 —— 规则必须列全所有非空来源值，否则本地计划会拒绝。 */
const SECOND_CATEGORY = "清洁用品";
const TARGET_VALUE = "目标值-已算好-3B8Z";
const SECOND_TARGET_VALUE = "目标值-已算好-5C1M";

/** 规则声明的源 / 目标区域 —— 多条用例共用，避免第二处写死字符串绕过变异。 */
const SOURCE_SHEET = "Sheet1";
const RULE_SOURCE_RANGE = "A2:A3";
const RULE_TARGET_RANGE = "B2:B3";

type CapturedRequest = { url: string; body: string };

/** 把 `"A2:A2"` / `"B2"` 解析成行列矩形；只服务替身，不做生产级校验。 */
function parseA1Rect(address: string) {
  const toColumn = (letters: string): number =>
    letters
      .toUpperCase()
      .split("")
      .reduce((total, char) => total * 26 + (char.charCodeAt(0) - 64), 0);
  const match = /^\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?$/.exec(address.trim());
  if (!match) return null;
  const startColumn = toColumn(match[1]);
  const startRow = Number(match[2]);
  const endColumn = match[3] ? toColumn(match[3]) : startColumn;
  const endRow = match[4] ? Number(match[4]) : startRow;
  return { startRow, startColumn, endRow, endColumn };
}

let captured: CapturedRequest[] = [];
let mounted: Array<{ container: HTMLElement; unmount: () => void }> = [];

function mountPanel() {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  return { container, root };
}

async function mount(element: ReactElement) {
  const { container, root } = mountPanel();
  await act(async () => {
    root.render(element);
  });
  const entry = {
    container,
    unmount: () => {
      act(() => {
        root.unmount();
      });
      container.remove();
    }
  };
  mounted.push(entry);
  return entry;
}

async function settle(): Promise<void> {
  await act(async () => {
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
  });
}

async function until(check: () => boolean, message: string): Promise<void> {
  for (let index = 0; index < 80; index += 1) {
    if (check()) return;
    await settle();
  }
  throw new Error(message);
}

function textOf(container: HTMLElement): string {
  return container.textContent ?? "";
}

function findButton(container: HTMLElement, name: string): HTMLButtonElement | null {
  return (
    Array.from(container.querySelectorAll("button")).find(
      (item) => (item.textContent ?? "").trim() === name
    ) ?? null
  );
}

async function click(element: Element | null): Promise<void> {
  if (!element) throw new Error("要点击的元素不存在");
  await act(async () => {
    element.dispatchEvent(new MouseEvent("click", { bubbles: true }));
  });
  await settle();
}

async function type(element: HTMLTextAreaElement | null, value: string): Promise<void> {
  if (!element) throw new Error("没有找到输入框");
  const setter = Object.getOwnPropertyDescriptor(
    HTMLTextAreaElement.prototype,
    "value"
  )?.set;
  await act(async () => {
    setter?.call(element, value);
    element.dispatchEvent(new Event("input", { bubbles: true }));
  });
}

async function submit(element: HTMLTextAreaElement | null): Promise<void> {
  const form = element?.closest("form");
  if (!form) throw new Error("没有找到输入区表单");
  await act(async () => {
    form.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true }));
  });
  await settle();
}

function jsonResponse(payload: unknown): Response {
  return {
    ok: true,
    status: 200,
    json: async () => payload,
    text: async () => JSON.stringify(payload)
  } as unknown as Response;
}

/**
 * 造一台只含一个"已用区域 = A1:B3"的 WPS 宿主。
 *
 * 布局刻意分成表头 + 两条数据行，因为：
 * - **表头会被合法地发出去**：第 1 行是字段名（`类别` / `归属`），意图阶段只读
 *   字段名，这部分允许出站 —— 它是正向锚点，证明负向断言不是"整个请求体为空"；
 * - 第 2 / 3 行是**落在规则源范围内**的数据行。A2 恒为已知类别；A3 由调用方
 *   决定 —— 可以是"已知映射值"（成功路径），也可以是"未知私有值"（失败路径）。
 *
 * 需要实现的接口（见 `wpsWorkbookReader.readChunkRange`）：
 * - `sheet.UsedRange`：整块已用区域；
 * - `usedRange.Item(row, col)`：定位到单元格；
 * - `sheet.Range(first, last)`：取矩形。
 * 三者缺一，读取器会抛"没有提供所需的分块读取接口"。
 */
function makeApplication(row3Value: string) {
  const headerRow = ["类别", "归属"];
  const cellAt = (row: number, col: number): string => {
    if (row === 1) return headerRow[col - 1] ?? "";
    // 第 2 行：已知类别，目标列读取时为空。
    if (row === 2 && col === 1) return KNOWN_CATEGORY;
    // 第 3 行：由调用方指定，是第二行数据（已知或未知）。
    if (row === 3 && col === 1) return row3Value;
    return "";
  };

  function makeRange(startRow: number, startCol: number, endRow: number, endCol: number) {
    const values: string[][] = [];
    for (let row = startRow; row <= endRow; row += 1) {
      const line: string[] = [];
      for (let col = startCol; col <= endCol; col += 1) line.push(cellAt(row, col));
      values.push(line);
    }
    const flat = values.length === 1 && values[0].length === 1 ? values[0][0] : values;
    const range = {
      Rows: { Count: endRow - startRow + 1 },
      Columns: { Count: endCol - startCol + 1 },
      MergeCells: false,
      // `readChunkRange` 读 Value2；`prepareWpsPlan` 复核源值时读 Formula。
      Formula: flat,
      Value2: flat
    } as Record<string, unknown>;
    return range;
  }

  const sheet = {
    Name: "Sheet1",
    ProtectContents: false,
    // 已用区域：A1:B3（表头 + 两行数据）。
    UsedRange: makeRange(1, 1, 3, 2),
    /**
     * WPS 的 `Range` 有**两种**调用形态，这里都要支持 —— 少一种就会静默回落到
     * A1，让执行预检拿 A1 的值去比 A2 的快照，报出"源区域在计算后发生变化"
     * 这种与真实缺陷无关的假失败：
     * - `Range(cellA, cellB)`：读取器 `readChunkRange` 用，参数是 `Item` 出来的单元格；
     * - `Range("A2:A2")`：执行预检 `rangeFor` 用，参数是 **A1 字符串**。
     */
    Range: (
      first: { __cell?: [number, number] } | string,
      last?: { __cell?: [number, number] }
    ) => {
      if (typeof first === "string") {
        const rect = parseA1Rect(first);
        if (!rect) throw new Error(`替身无法解析区域「${first}」`);
        return makeRange(rect.startRow, rect.startColumn, rect.endRow, rect.endColumn);
      }
      const [sr, sc] = first?.__cell ?? [1, 1];
      const [er, ec] = last?.__cell ?? [sr, sc];
      return makeRange(sr, sc, er, ec);
    }
  } as Record<string, unknown> & { Name: string };

  // 给 Item 出来的单元格标注坐标，Range(first,last) 才能拼回矩形。
  (sheet.UsedRange as { Item: (r: number, c: number) => unknown }).Item = (
    row: number,
    col: number
  ) => {
    const cell = makeRange(row, col, row, col);
    (cell as { __cell?: [number, number] }).__cell = [row, col];
    return cell;
  };

  const storage = new Map<string, unknown>();
  return {
    application: {
      Version: "12.0",
      ActiveWorkbook: {
        Name: "测试.xlsx",
        FullName: "C:\\Data\\测试.xlsx",
        Path: "C:\\Data",
        ReadOnly: false,
        Worksheets: { Count: 1, Item: () => sheet },
        ActiveSheet: sheet
      },
      ActiveSheet: sheet,
      PluginStorage: {
        getItem: (key: string) => storage.get(key),
        setItem: (key: string, item: unknown) => {
          storage.set(key, item);
        }
      },
      CreateTaskPane: () => ({ ID: 1, Visible: true }),
      GetTaskPane: () => undefined
    } as never
  };
}

/**
 * 服务端返回的本地计算规则：**不含任何目标 values**，只有规则本身。
 *
 * 源区域是 **A2:A3**（两行），目标是 **B2:B3**：A2 = `KNOWN_CATEGORY`，
 * A3 = `row3Category`。**两条用例共用这一份构造**，只有映射表不同 ——
 * 这样"未知值必须失败 / 已知值必须成功"的差异只来自映射表本身，不来自替身差异。
 *
 * `mapping` 由调用方给出，且**只在已知值用例里才列全**。私有哨兵永远不进这里。
 */
function localRuleTurnResponse(row3Category: string, includeRow3: boolean) {
  const mapping = [{ source: KNOWN_CATEGORY, target: TARGET_VALUE }];
  if (includeRow3) {
    // 规则列全了非空来源类别。注意这是**测试替身规则自己声明的类别**，
    // 与"服务端知不知道某个私有值"无关 —— 它的取值就是工作簿里真实存在的那个。
    mapping.push({ source: row3Category, target: SECOND_TARGET_VALUE });
  }
  return {
    kind: "local_plan_rule",
    provider: "model",
    rule: {
      kind: "columnMapping",
      sourceSheet: SOURCE_SHEET,
      targetSheet: SOURCE_SHEET,
      sourceRange: RULE_SOURCE_RANGE,
      targetRange: RULE_TARGET_RANGE,
      mapping,
      blankLabel: "未分类"
    },
    summary: "按映射规则写入。",
    warnings: [],
    assumptions: []
  };
}

/**
 * 意图阶段的服务端响应。
 *
 * 注意 `checkIntent` 与 `createAssistantResponse` **打的是同一个
 * `POST /api/turn`**，靠响应区分：这里返回 `proceed`，面板随后才会发起
 * 规划请求并拿到 `local_plan_rule`。两者都由同一个 fetch 替身按调用次序返回。
 */
function intentProceedResponse() {
  return {
    kind: "proceed",
    summary: "列映射",
    confirmedPrompt: "把类别列翻译成归属",
    provider: "model"
  };
}

const modelCatalog = {
  defaultModelId: "connection:first",
  models: [
    {
      id: "connection:first",
      label: "第一个连接",
      provider: "model" as const,
      available: true,
      supportsVision: false
    }
  ]
};

/** 每个用例注入自己的规划响应；fetch 替身共用。 */
let planResponseFactory: () => unknown = () => ({});

beforeEach(() => {
  localStorage.clear();
  captured = [];
  vi.clearAllMocks();
  // `POST /api/turn` 被复用两次：先意图、再规划。用计数器区分，避免把
  // 意图请求误判成规划请求。
  let turnCallCount = 0;
  vi.stubGlobal(
    "fetch",
    vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = typeof input === "string" ? input : String(input);
      captured.push({ url, body: typeof init?.body === "string" ? init.body : "" });
      // 模型目录：真实 api 模块会来读，返回一份可用连接。
      if (url.includes("/api/models")) return jsonResponse(modelCatalog);
      if (url.includes("/api/settings/model")) {
        return jsonResponse({ connectionCount: 1 });
      }
      if (url.includes("/api/turn")) {
        turnCallCount += 1;
        // 第一次是意图判断，第二次才是规划。
        return jsonResponse(
          turnCallCount === 1 ? intentProceedResponse() : planResponseFactory()
        );
      }
      if (url.includes("/api/plans/validate")) {
        // 真走到这里就是违约：本地计划不该来校验。返回原样，让断言能给出
        // 更接近真实的行为对比（而不是因为 404 提前崩）。
        return jsonResponse({});
      }
      return jsonResponse({});
    })
  );
});

afterEach(() => {
  mounted.forEach((entry) => entry.unmount());
  mounted = [];
  document.body.innerHTML = "";
  vi.unstubAllGlobals();
});

function panel() {
  return (
    <WpsAgentPanel
      selectedSheetIndexes={[1]}
      scopeDock={<div className="eb-scope-dock">当前表 Sheet1</div>}
    />
  );
}

/** 从捕获结果里取所有请求体的合并文本，便于做"不含某值"的断言。 */
function allBodies(): string {
  return captured.map((item) => `${item.url}\n${item.body}`).join("\n");
}

/** 跑一次完整的"输入 → 提交"，停在结果出现之前。 */
async function ask(
  container: HTMLElement,
  prompt = "把类别列翻译成归属"
): Promise<void> {
  const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
  await type(textarea, prompt);
  await submit(textarea);
}

describe("local_plan_rule 的真实出站请求体", () => {
  describe("未知私有值：规则里没有它，本地计划必须拒绝而不是兜底", () => {
    beforeEach(() => {
      // 规则只列了 A2 的已知类别；A3 的私有值**不在映射表内**。
      planResponseFactory = () =>
        localRuleTurnResponse(PRIVATE_SENTINEL, false);
    });

    it("拒绝映射、不生成写入预览，且请求体里没有该值也没有源快照", async () => {
      const { application } = makeApplication(PRIVATE_SENTINEL);
      getWpsApplicationMock.mockReturnValue(application);

      const { container } = await mount(panel());
      await until(() => textOf(container).includes("模型：第一个连接"), "应加载模型");

      // prompt 里**不含**私有值：它只能从 A3 读出来。
      await ask(container);
      // 失败路径的锚点：等错误提示出现。`mapColumnValues` 对未命中值抛
      // `UnmappedCategoryError`，面板把它转成系统消息。
      await until(
        () => textOf(container).includes("不在映射表内"),
        "未知来源值应被拒绝并给出提示"
      );

      // 正向锚点 1：确实发生过出站（意图 + 规划），否则下面的负向断言是空过。
      expect(captured.length).toBeGreaterThan(0);
      const turnRequests = captured.filter((item) => item.url.includes("/api/turn"));
      expect(turnRequests.length).toBeGreaterThanOrEqual(2);

      // 正向锚点 2：**没有任何写入预览**（没有"确认写入"按钮，也没有目标区域）。
      // 这正是"拒绝而非兜底"的可观测后果 —— 若生产代码把未命中值兜成 blankLabel，
      // 这里会冒出一个预览，用例立刻变红。
      expect(findButton(container, "确认写入")).toBeNull();
      expect(textOf(container)).not.toContain(RULE_TARGET_RANGE);
      expect(textOf(container)).not.toContain(TARGET_VALUE);
      expect(textOf(container)).not.toContain("已按映射处理");
      // 错误文案本身提及未命中的来源值（这是要让用户知道去补映射表），
      // 但那只发生在**界面**上；下面的负向断言针对的是出站字节。
      expect(textOf(container)).toContain(PRIVATE_SENTINEL);

      const outbound = allBodies();
      // 1) 私有值不出站 —— 规则/响应/提示都不知道它，所以这是真断言。
      expect(outbound).not.toContain(PRIVATE_SENTINEL);
      expect(outbound).not.toContain("7Q4X");
      // 2) 源快照（A2:A3 的完整值，含那行私有值）不能上传。
      expect(outbound).not.toContain("sourceSnapshots");
      // 3) 本地计划**不上传校验接口** —— 在 URL 层面确认，比"函数没被调用"更硬。
      expect(captured.some((item) => item.url.includes("/api/plans/validate"))).toBe(false);
      // 4) 已知类别的目标矩阵同样不该出现（本次根本没算出成功计划）。
      expect(outbound).not.toContain(TARGET_VALUE);
      expect(outbound).not.toContain("3B8Z");

      // 5) 表头合法出站这件事本身也在断言范围内：`类别` 是字段名，允许上传。
      //    这样才能证明第 1 条不是「整个请求体为空」造成的空过。
      expect(turnRequests[0].body).toContain("类别");
    });
  });

  describe("已知映射值：规则列全非空来源类别，本地计划成功且出站面不变", () => {
    beforeEach(() => {
      // 规则把 A2、A3 两条已知类别都列进映射表。
      planResponseFactory = () =>
        localRuleTurnResponse(SECOND_CATEGORY, true);
    });

    it("生成两行预览，且私有源值、源快照与目标矩阵都不出站、不走校验接口", async () => {
      const { application } = makeApplication(SECOND_CATEGORY);
      getWpsApplicationMock.mockReturnValue(application);

      const { container } = await mount(panel());
      await until(() => textOf(container).includes("模型：第一个连接"), "应加载模型");

      await ask(container);
      await until(() => textOf(container).includes("写入前确认"), "应展示写入计划");

      // 正向锚点 1：确实发生过出站。
      const turnRequests = captured.filter((item) => item.url.includes("/api/turn"));
      expect(turnRequests.length).toBeGreaterThanOrEqual(2);

      // 正向锚点 2：A3 **在源范围内**（A2:A3），预览完整生成了两行目标 ——
      // 说明负向断言不是因为流程提前失败或该行被跳过而空过。
      expect(textOf(container)).toContain(`${SOURCE_SHEET}!${RULE_TARGET_RANGE}`);
      expect(textOf(container)).toContain(TARGET_VALUE);
      expect(textOf(container)).toContain(SECOND_TARGET_VALUE);

      const outbound = allBodies();
      // 1) 算好的目标矩阵（两行）都不出站。
      expect(outbound).not.toContain(TARGET_VALUE);
      expect(outbound).not.toContain("3B8Z");
      expect(outbound).not.toContain(SECOND_TARGET_VALUE);
      expect(outbound).not.toContain("5C1M");
      // 2) 源快照本身（A2:A3 的完整值）也不能上传。
      expect(outbound).not.toContain("sourceSnapshots");
      // 3) 完整写入矩阵字段名不该出现。
      expect(outbound).not.toContain('"values"');
      // 4) 本地计划**不上传校验接口**。
      expect(captured.some((item) => item.url.includes("/api/plans/validate"))).toBe(false);

      // 关于动作名 `writeValues`：它**会合法地**出现在
      // `selfReportedActions`（能力自报清单，见 `wpsHostAdapter`）里，所以不能拿
      // "整个请求体不含 writeValues" 当泄漏判据 —— 那是把白名单字段误判成泄漏。
      // 真正该钉的是：没有任何请求把**计划本体**送出去（动作数组 / 计划 id /
      // 目标区域成对出现）。用这些计划特有的结构做断言，既避开能力自报，
      // 又比"某个字符串没出现"更贴近"计划没有出站"这件事本身。
      expect(outbound).not.toContain('"actions"');
      expect(outbound).not.toContain("wps-map-column-");
      expect(outbound).not.toContain(`${SOURCE_SHEET}!${RULE_TARGET_RANGE}`);

      // 5) 表头合法出站 —— 证明上面的负向断言不是"整个请求体为空"。
      expect(turnRequests[0].body).toContain("类别");
      expect(KNOWN_CATEGORY).not.toBe(SECOND_CATEGORY);
    });

    it("编辑后重新预览：仍未把源快照或完整 actions[].values 发出去", async () => {
      const { application } = makeApplication(SECOND_CATEGORY);
      getWpsApplicationMock.mockReturnValue(application);

      const { container } = await mount(panel());
      await until(() => textOf(container).includes("模型：第一个连接"), "应加载模型");

      await ask(container);
      await until(() => textOf(container).includes("写入前确认"), "应展示写入计划");

      // 记账：编辑动作之前的出站量。
      const beforeEdit = captured.length;

      await click(findButton(container, "编辑计划"));
      const editBox = container.querySelector<HTMLTextAreaElement>(
        '[aria-label="编辑写入计划 JSON"]'
      );
      const editedPlan = {
        id: "wps-map-column-edited",
        title: `按映射表填写 ${SOURCE_SHEET}!${RULE_TARGET_RANGE}`,
        summary: "按映射规则写入。",
        assumptions: [],
        warnings: [],
        actions: [
          {
            type: "writeValues",
            sheet: SOURCE_SHEET,
            range: RULE_TARGET_RANGE,
            values: [[TARGET_VALUE], [SECOND_TARGET_VALUE]]
          }
        ]
      };
      await type(editBox, JSON.stringify(editedPlan));
      await click(findButton(container, "应用修改并重新预览"));
      await until(
        () => findButton(container, "确认写入") !== null,
        "编辑后应重新预览"
      );

      // 编辑后是否还有请求？**本地计算计划这一支是零出站** ——
      // `applyEditedPlan` 看到 `pendingPlan.sourceSnapshots.length > 0` 就直接
      // 复用编好的计划，绕开 `validatePlan`（见 WpsAgentPanel 编辑分支）。
      // 所以这里不能拿"总出站量 > 0"做正向锚点：真正该钉住的是
      // "编辑后**仍然**没有请求可以把源值带出去"，以及"本地计划确实没走校验"。
      const editRequests = captured.slice(beforeEdit);
      const editOutbound = editRequests
        .map((item) => `${item.url}\n${item.body}`)
        .join("\n");

      // 锚点 1：本地计算计划在编辑路径上不产生任何服务端校验请求。
      expect(
        editRequests.filter((item) => item.url.includes("/api/plans/validate")).length
      ).toBe(0);
      // 锚点 2：预览确实被重建了（否则上面一条会因为"什么都没做"而空过）。
      // 覆盖范围刻意**从规则里取**而不是写死：这样"源范围退回 A2:A2"这类变异
      // 会在这里一并变红，而不是被一个恰好匹配的硬编码字符串糊过去。
      expect(findButton(container, "确认写入")).not.toBeNull();
      expect(textOf(container)).toContain(`${SOURCE_SHEET}!${RULE_TARGET_RANGE}`);
      expect(textOf(container)).toContain(TARGET_VALUE);
      expect(textOf(container)).toContain(SECOND_TARGET_VALUE);
      // 锚点 3：编辑路径出站量要么是 0，要么每一笔都不含目标矩阵、源快照或
      // 完整 actions[].values —— 逐条看，不用"合并文本不含"这种一旦为空
      // 就恒真的写法。
      //
      // 这里同样**不用 `writeValues`** 做判据：它是能力自报清单里的合法字段
      // （`selfReportedActions`），拿它当泄漏信号会把白名单内容误报成缺陷。
      for (const item of editRequests) {
        expect(item.body).not.toContain(TARGET_VALUE);
        expect(item.body).not.toContain("3B8Z");
        expect(item.body).not.toContain(SECOND_TARGET_VALUE);
        expect(item.body).not.toContain("5C1M");
        expect(item.body).not.toContain("sourceSnapshots");
        expect(item.body).not.toContain('"values"');
        // 计划本体的结构（动作数组 / 计划 id）不该出现在编辑后的任何请求里。
        expect(item.body).not.toContain('"actions"');
        expect(item.body).not.toContain("wps-map-column-");
        expect(item.body).not.toContain(PRIVATE_SENTINEL);
      }
      expect(editOutbound).not.toContain(TARGET_VALUE);
      expect(editOutbound).not.toContain(SECOND_TARGET_VALUE);
    });
  });
});
