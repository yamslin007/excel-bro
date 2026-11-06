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
//   A. 未知私有值 —— 先停止并本地询问，绝不兜底；用户逐项明确映射后
//      重新读源、生成新预览。整个过程不把来源值或目标矩阵发给服务端。
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
import type { WpsApplication } from "./wpsHost";
import { prepareImageFile } from "../../excel-addin/src/imageAttachments";

vi.mock("../../excel-addin/src/imageAttachments", async (original) => ({
  ...await original<typeof import("../../excel-addin/src/imageAttachments")>(),
  prepareImageFile: vi.fn(async (file: File) => ({
    id: file.name, name: file.name, mediaType: "image/png" as const,
    data: "ZXhwbGljaXQtaW1hZ2Utc2VudGluZWw=",
    previewUrl: "data:image/png;base64,ZXhwbGljaXQtaW1hZ2Utc2VudGluZWw="
  }))
}));

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

async function typeInput(element: HTMLInputElement | null, value: string): Promise<void> {
  if (!element) throw new Error("没有找到映射输入框");
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set;
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
function makeApplication(initialRow3Value: string) {
  let row3Value = initialRow3Value;
  let writes = 0;
  const headerRow = ["类别", "归属"];
  /**
   * 真实取值面的记账：**谁被取过值就记谁**，不看调用方贴的 `purpose` 标签。
   *
   * 为什么不能按目的记账：`purpose` 只是 `readChunkRange` 传下来的标记，
   * 它不改变 `Value2` 取到哪块矩形。按"取数据目的才算"记数，就会得到
   * "数据读取 = 0"这种与真实访问范围不符的结论 —— 那是**假证据**。
   *
   * - `headerValues` —— 第 1 行（表头）被取过值的格。
   * - `dataValues` —— 第 2 行及以后（数据行）被取过值的格。这条必须为空。
   * - `sourceColumnWholeReads` —— 整片源列（跨表头与数据行）被一次性取走。
   */
  const headerValues = new Set<string>();
  const dataValues = new Set<string>();
  const sourceColumnWholeReads: string[] = [];
  /*
   * **真实写入面**：执行器写完之后会逐格回读并与计划比对。替身若只把 setter
   * 记成"写了一次"而不改值，回读永远等于原值 —— 那会把"执行链路正常"误判成
   * "回读不符"，测到的是替身缺陷而不是产品行为。所以这里按坐标真正存值。
   */
  const written = new Map<string, unknown>();
  const cellAt = (row: number, col: number): string => {
    const key = `${row},${col}`;
    if (written.has(key)) return written.get(key) as string;
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
    /** 把标量或二维矩阵按本区域坐标写进存储（与宿主语义一致：左上角对齐）。 */
    const store = (value: unknown) => {
      writes += 1;
      const matrix = Array.isArray(value) ? value : [[value]];
      for (let r = 0; r < endRow - startRow + 1; r += 1) {
        for (let c = 0; c < endCol - startCol + 1; c += 1) {
          const row = Array.isArray(matrix[r]) ? (matrix[r] as unknown[])[c] : undefined;
          if (row === undefined && c > 0) continue;
          written.set(`${startRow + r},${startCol + c}`, row ?? null);
        }
      }
    };
    const range = {
      Rows: { Count: endRow - startRow + 1 },
      Columns: { Count: endCol - startCol + 1 },
      MergeCells: false,
      // `readChunkRange` 读 Value2；`prepareWpsPlan` 复核源值时读 Formula。
      get Formula() { return values.length === 1 && values[0].length === 1 ? cellAt(startRow, startCol) : values.map((row, r) => row.map((_v, c) => cellAt(startRow + r, startCol + c))); },
      set Formula(value: unknown) { store(value); },
      get Value2() { return values.length === 1 && values[0].length === 1 ? cellAt(startRow, startCol) : values.map((row, r) => row.map((_v, c) => cellAt(startRow + r, startCol + c))); },
      set Value2(value: unknown) { store(value); },
      // `clearRange`（contents）走 ClearContents；与"写入空串"是两回事。
      ClearContents: () => {
        writes += 1;
        for (let r = startRow; r <= endRow; r += 1) {
          for (let c = startCol; c <= endCol; c += 1) written.delete(`${r},${c}`);
        }
      }
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
      last?: { __cell?: [number, number] },
      // 真宿主签名契约：`Worksheet.Range` 只接受两个参数，多传抛
      // `Too many parameters.`（2026-09-30 真机证据）。
      ...extra: unknown[]
    ) => {
      if (extra.length > 0) throw new Error("Too many parameters.");
      if (typeof first === "string") {
        const rect = parseA1Rect(first);
        if (!rect) throw new Error(`替身无法解析区域「${first}」`);
        return makeRange(rect.startRow, rect.startColumn, rect.endRow, rect.endColumn);
      }
      const [sr, sc] = first?.__cell ?? [1, 1];
      const [er, ec] = last?.__cell ?? [sr, sc];
      // 逐格扫一遍 —— 与真实 `Value2` 矩阵等价：取到哪块，那块的值就进了内存。
      // 记账**只看坐标**：第 1 行是表头，第 2 行及以后是数据行。
      for (let row = sr; row <= er; row += 1) {
        for (let col = sc; col <= ec; col += 1) {
          if (row === 1) headerValues.add(`${row},${col}`);
          else dataValues.add(`${row},${col}`);
        }
      }
      if (er > 1) sourceColumnWholeReads.push(`${sr},${sc}:${er},${ec}`);
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
    } as never,
    /** 表头行（第 1 行）被取过值的格数 —— 结构扫描确实生效的正向锚点。 */
    headerValueCount: () => headerValues.size,
    /** 数据行（第 2 行及以后）被取过值的格数 —— 必须为 0。 */
    dataValueCount: () => dataValues.size,
    /** 一次性跨过表头与数据行的读取次数 —— 必须为 0（没有整片源列读取）。 */
    wholeReads: () => sourceColumnWholeReads.length,
    setRow3Value: (value: string) => { row3Value = value; },
    writeCount: () => writes
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
    turnId: "turn-local-outbound",
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
  modelCatalog.models[0].supportsVision = false;
  // `POST /api/turn` 被复用两次：先意图、再规划。用计数器区分，避免把
  // 意图请求误判成规划请求。
  let turnCallCount = 0;
  /*
   * 账本在**登记**时报出恢复描述，之后每个检查点/完成回执都会带上同一份描述。
   * 替身要像真实服务端那样把它记住，否则后续动作回执里的空描述会被真实的
   * `assertWpsRecoveryDescriptor` 拒掉 —— 那是替身不真实，不是产品问题。
   */
  let ledgerDescriptor: unknown = null;
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
      if (url.includes("/api/intent/message")) {
        return jsonResponse({ kind: "needs_workbook_context", summary: "需要字段结构", provider: "model" });
      }
      if (url.includes("/api/turn")) {
        turnCallCount += 1;
        if (url.includes("/stream")) {
          return new Response(`event: result\ndata: ${JSON.stringify(planResponseFactory())}\n\n`, {
            headers: { "Content-Type": "text/event-stream" }
          });
        }
        // Distinguish by the real request shape, including subsequent clarification turns.
        return jsonResponse(
          JSON.parse(String(init?.body ?? "{}")).scope ? intentProceedResponse() : planResponseFactory()
        );
      }
      /*
       * CAP-03 账本端点：真实实现会在**首写之前**登记、每个动作前后推进检查点。
       * 这里必须给真实形状的答复，否则登记失败会让写入被安全地中止 ——
       * 那样测到的是"替身缺端点"，不是产品行为。
       */
      if (url.includes("/api/wps/executions")) {
        const request = init?.body ? JSON.parse(String(init.body)) : {};
        if (url.includes("/register") && request?.descriptor) ledgerDescriptor = request.descriptor;
        const revision = Number(request?.expectedRevision ?? 0) + 1;
        const operationId = decodeURIComponent(url.split("/executions/")[1] ?? "")
          .split("/")[0] || "wps-execution-test";
        return jsonResponse({
          operationId,
          status: url.endsWith("/complete") ? "completed" : "running",
          revision,
          workbookIdentity: "wps-path:c:////data////测试.xlsx",
          taskId: "task-test",
          // 回显请求里的描述：真实服务端会校验它（不允许空 actions、不允许矩阵）。
          descriptor: request?.descriptor ?? ledgerDescriptor ?? { actions: [] },
          authorizedSheetNames: ["Sheet1"],
          // 进度必须与描述里的动作数一致（真实服务端就这么校验）。
          nextActionIndex: url.includes("/actions/complete") ? Math.min(1, (ledgerDescriptor as { actions?: unknown[] })?.actions?.length ?? 1) : 0,
          totalActions: (ledgerDescriptor as { actions?: unknown[] })?.actions?.length ?? 1,
          inFlightActionIndex: null,
          beforeSnapshot: null,
          checkpoint: null,
          result: null,
          error: null,
          canResume: true,
          ownerActive: false,
          requiresReconciliation: false,
          ownerToken: "owner-token-0123456789abcdef0123456789abcdef"
        });
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

/**
 * 只统计**会携带计划语义**的出站：意图、规划、计划校验。
 *
 * 为什么不数 `captured.length`：CAP-03 的账本登记（`/api/wps/executions/register`）
 * 发生在**执行前**，是恢复证据 —— 它只带动作类型/表/区域/计数与两个指纹，
 * 不含任何单元格内容。新流程"计划就绪即执行"，所以它必然出现在这一轮里；
 * 把它算进"计划出站"会让"本地计划不出站"这条判据失去判别力（每次都被它顶红）。
 * 真正的泄漏判据仍然是"目标值/源快照/`values` 不出现在任何请求体里"。
 */
function planCalls() {
  return captured.filter((item) =>
    item.url.includes("/api/intent") ||
    item.url.includes("/api/turn") ||
    item.url.includes("/api/plans/validate"));
}

describe("图片附件的真实 fetch 链路", () => {
  async function attach(container: HTMLElement) {
    const input = container.querySelector<HTMLInputElement>('input[aria-label="选择图片附件"]')!;
    const file = new File(["explicit image"], "user-screenshot.png", { type: "image/png" });
    Object.defineProperty(input, "files", { configurable: true, value: [file] });
    await act(async () => { input.dispatchEvent(new Event("change", { bubbles: true })); });
    await until(() => Boolean(container.querySelector(".composer-attachments img")), "应出现图片预览");
    return container.querySelector<HTMLTextAreaElement>(".composer textarea");
  }

  it("先确认目标；意图只发数量，规划发真实图片；澄清续答仍携带原图，不落聊天存储", async () => {
    const host = makeApplication(PRIVATE_SENTINEL);
    getWpsApplicationMock.mockReturnValue(host.application);
    modelCatalog.models[0].supportsVision = true;
    let plans = 0;
    planResponseFactory = () => ++plans === 1 ? {
      kind: "clarification", provider: "model", clarification: {
        id: "image-question", summary: "补充图片目标", question: "如何使用图片？",
        reason: "缺目标", scopeLabel: "Sheet1", options: [
          { id: "analyze", label: "只分析图片", description: "不修改原表", resolution: "只分析图片", resolutionMode: "append" },
          { id: "preview", label: "参考图片生成预览", description: "先预览", resolution: "参考图片生成预览", resolutionMode: "append" }
        ], fields: []
      }
    } : { kind: "answer", provider: "model", message: "已结合附件图片分析", resultContext: null };
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("第一个连接") ?? false, "模型加载");
    const textarea = await attach(container);
    expect(prepareImageFile).toHaveBeenCalledOnce();
    expect(container.querySelector<HTMLImageElement>(".composer-attachments img")?.src).toContain("ZXhwbGljaXQ");
    expect(captured.filter((r) => r.url.includes("/api/turn"))).toHaveLength(0);
    await submit(textarea);
    if (container.querySelector('[role="dialog"]')) await confirmBeforeAnalysis(container);
    await until(() => textOf(container).includes("如何使用图片？"), "应出现图片任务澄清");
    // 点选项只选中；提交是右下角那颗「发送回答」。
    await click(Array.from(container.querySelectorAll("button")).find((button) =>
      button.textContent?.includes("只分析图片")) ?? null);
    await click(container.querySelector<HTMLButtonElement>('button[aria-label="发送回答"]'));
    await until(() => textOf(container).includes("已结合附件图片分析"), "图片续答应完成");
    const bodies = captured.filter((r) => r.url.includes("/api/turn")).map((r) => JSON.parse(r.body));
    const intents = bodies.filter((b) => b.imageCount !== undefined);
    const planning = bodies.filter((b) => b.workbook);
    expect(intents.length).toBeGreaterThan(0);
    expect(intents.every((b) => b.imageCount === 1 && !b.images)).toBe(true);
    expect(planning).toHaveLength(2);
    for (const request of planning) {
      expect(request.images).toEqual([{ name: "user-screenshot.png", mediaType: "image/png", data: "ZXhwbGljaXQtaW1hZ2Utc2VudGluZWw=" }]);
    }
    expect(allBodies()).not.toContain(PRIVATE_SENTINEL);
    expect(allBodies()).not.toContain((host.application as WpsApplication).ActiveWorkbook!.FullName);
    const stored = localStorage.getItem("excel-bro.wps.chat.v1") ?? "";
    expect(stored).toContain("user-screenshot.png");
    expect(stored).not.toContain("ZXhwbGljaXQ");
    expect(host.writeCount()).toBe(0);
  });

  it("非视觉模型允许添加和删除，但阻止发送，不丢附件", async () => {
    const host = makeApplication(PRIVATE_SENTINEL);
    getWpsApplicationMock.mockReturnValue(host.application);
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("第一个连接") ?? false, "模型加载");
    const textarea = await attach(container);
    await submit(textarea);
    expect(textOf(container)).toContain("当前模型不支持图片");
    expect(captured.filter((r) => r.url.includes("/api/turn"))).toHaveLength(0);
    expect(container.querySelector(".composer-attachments img")).not.toBeNull();
    await click(container.querySelector('button[aria-label="移除图片 user-screenshot.png"]'));
    expect(container.querySelector(".composer-attachments")).toBeNull();
    expect(host.writeCount()).toBe(0);
  });

  it("图片规划连接失败后保留图片与原需求，不能冒充成功或写入", async () => {
    const host = makeApplication(PRIVATE_SENTINEL);
    getWpsApplicationMock.mockReturnValue(host.application);
    modelCatalog.models[0].supportsVision = true;
    planResponseFactory = () => { throw new TypeError("Failed to fetch"); };
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("第一个连接") ?? false, "模型加载");
    const textarea = await attach(container);
    await type(textarea, "按截图生成预览，不执行写入");
    await submit(textarea);
    if (container.querySelector('[role="dialog"]')) await confirmBeforeAnalysis(container);
    // Stream transport failure may use the existing one-shot fallback as well.
    await until(() => captured.filter((r) => r.url.includes("/api/turn")).length >= 2 &&
      textarea?.value === "按截图生成预览，不执行写入", "失败后应恢复原需求");
    expect(container.querySelector(".composer-attachments img")).not.toBeNull();
    expect(localStorage.getItem("excel-bro.wps.chat.v1")).not.toContain("ZXhwbGljaXQ");
    expect(host.writeCount()).toBe(0);
  });
});

/** 跑一次完整的"输入 → 提交"，停在结果出现之前。 */
async function ask(
  container: HTMLElement,
  prompt = "把类别列翻译成归属"
): Promise<void> {
  const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
  await type(textarea, prompt);
  await submit(textarea);
}

async function confirmBeforeAnalysis(container: HTMLElement) {
  await until(() => Boolean(container.querySelector('[role="dialog"]')), "需求分析前应先显示独立目标确认");
  const dialog = container.querySelector('[role="dialog"]')!;
  expect(dialog.querySelector('input[type="radio"]')).toBeNull();
  expect(dialog.textContent).toContain("确认本次操作范围");
  await click(findButton(container, "确认范围并继续"));
}

describe("local_plan_rule 的真实出站请求体", () => {
  describe("未知私有值：规则里没有它，本地计划必须拒绝而不是兜底", () => {
    beforeEach(() => {
      // 规则只列了 A2 的已知类别；A3 的私有值**不在映射表内**。
      planResponseFactory = () =>
        localRuleTurnResponse(PRIVATE_SENTINEL, false);
    });

    it("每次返回新包装对象：分析前确认目标，随后业务补全直接预览，不再混入授权", async () => {
      const host = makeApplication(PRIVATE_SENTINEL);
      const app = host.application as WpsApplication;
      const book = app.ActiveWorkbook!;
      Object.defineProperty(app, "ActiveWorkbook", { get: () => ({
        ...book, Worksheets: { ...book.Worksheets, Item: (index: number) => ({ ...book.Worksheets.Item(index) }) }
      }) });
      expect(app.ActiveWorkbook).not.toBe(app.ActiveWorkbook);
      expect(app.ActiveWorkbook!.Worksheets.Item(1)).not.toBe(app.ActiveWorkbook!.Worksheets.Item(1));
      getWpsApplicationMock.mockReturnValue(host.application);
      const { container } = await mount(panel());
      await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
      await ask(container);
      await until(() => Boolean(container.querySelector('[role="dialog"]')), "分析前先确认目标");
      expect(captured.filter((item) => item.url.includes("/api/turn"))).toHaveLength(0);
      expect(host.headerValueCount()).toBe(0);
      expect(host.dataValueCount()).toBe(0);
      expect(host.writeCount()).toBe(0);
      expect(textOf(container)).toContain("理解消息");
      expect(textOf(container)).not.toContain("读取已授权表结构");
      const firstRouting = captured.filter((item) => item.url.includes("/api/intent/message"));
      expect(firstRouting).toHaveLength(1);
      expect(Object.keys(JSON.parse(firstRouting[0].body)).sort()).toEqual(["conversation", "imageCount", "modelId", "prompt"]);
      expect(textOf(container)).not.toContain(PRIVATE_SENTINEL);
      await confirmBeforeAnalysis(container);
      await until(() => textOf(container).includes(PRIVATE_SENTINEL), "应展示本地未知来源值");
      const requestCount = planCalls().length;
      const target = "用户本地答案-6K9B";
      expect(findButton(container, "按补充映射重新生成预览")!.disabled).toBe(true); // Missing input, not missing authorization.
      expect(container.querySelector('.local-mapping-scope-confirmation')).toBeNull();
      expect(textOf(container)).not.toContain("确认当前操作目标：");
      await typeInput(container.querySelector<HTMLInputElement>('.wps-local-mapping-values input[type="text"]'), target);
      await until(() => !findButton(container, "按补充映射重新生成预览")!.disabled, "填写完整后应可预览");
      await click(findButton(container, "按补充映射重新生成预览"));
      await until(() => textOf(container).includes("本次执行计划"), "确认后必须进入新预览，而不是重问同一道题");
      expect(textOf(container)).not.toContain("无法确认当前工作簿和授权工作表");
      expect(textOf(container)).not.toContain("已重新确认当前数据范围");
      expect(findButton(container, "确认此范围并继续")).toBeNull();
      expect(planCalls().length).toBe(requestCount);
      expect(findButton(container, "按补充映射重新生成预览")).toBeNull();
      expect(textOf(container)).toContain(target);
      expect(allBodies()).not.toContain(PRIVATE_SENTINEL);
      expect(allBodies()).not.toContain(target);
      const stored = localStorage.getItem("excel-bro.wps.chat.v1") ?? "";
      expect(stored).not.toContain(PRIVATE_SENTINEL);
      expect(stored).not.toContain(target);
      /*
       * 「再确认一次目标」**不是**一次新授权：授权范围一个字没变，
       * `authorizationRevision` 就不许前进一格 —— 版本号是"已提交范围的身份"，
       * 不是"用户点过确认"的计数。前进一格会把仍在进行的审阅判成"授权已变化"
       * （真机症状：确认/放弃反复报"工作簿或授权发生变化"）。
       */
      expect(stored).toContain('"authorizationRevision":1');
      expect(stored).not.toContain('"authorizationRevision":2');
      expect(stored).not.toContain(book.FullName);
      expect(allBodies()).not.toContain(book.FullName);
      // 补全后计划**已经执行**（新流程是"先执行、后审阅"）：这里校验的是
      // "只写了这一处、且私有值仍不出站/不落盘"，而不是"零写入" ——
      // 零写入的正确观测点是执行开始之前，由本用例前半段断言。
      expect(host.writeCount()).toBe(1);
    });

    it("分析前授权后换到同名异路径文件：不得把私有补全规则用于另一份文件", async () => {
      const host = makeApplication(PRIVATE_SENTINEL);
      const app = host.application as WpsApplication;
      const book = app.ActiveWorkbook!;
      let fullName = book.FullName;
      Object.defineProperty(app, "ActiveWorkbook", { get: () => ({
        ...book, FullName: fullName,
        Worksheets: { ...book.Worksheets, Item: (index: number) => ({ ...book.Worksheets.Item(index) }) }
      }) });
      getWpsApplicationMock.mockReturnValue(host.application);
      const { container } = await mount(panel());
      await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
      await ask(container);
      await confirmBeforeAnalysis(container);
      await until(() => textOf(container).includes(PRIVATE_SENTINEL), "应展示本地澄清");
      const target = "禁止跨文件挪用-4Q1C";
      await typeInput(container.querySelector<HTMLInputElement>('.wps-local-mapping-values input[type="text"]'), target);
      await until(() => !findButton(container, "按补充映射重新生成预览")!.disabled, "已有前置授权且输入有效");
      const count = planCalls().length;
      fullName = "C:\\Other\\测试.xlsx";
      await click(findButton(container, "按补充映射重新生成预览"));
      await until(() => Boolean(container.querySelector('[role="dialog"]')), "跨文件应先要求选择目标，不能先重规划");
      expect(planCalls().length).toBe(count);
      expect(host.writeCount()).toBe(0);
      expect(container.querySelector<HTMLInputElement>('.wps-local-mapping-values input[type="text"]')!.value).toBe(target);
      const dialog = container.querySelector('[role="dialog"]')!;
      await click(dialog.querySelectorAll('input[type="radio"]')[1]);
      await click(dialog.querySelector('input[type="checkbox"]'));
      await click(findButton(container, "确认目标并继续"));
      expect(planCalls().length).toBe(count); // Scope confirmation must not auto-replay old mapping or planning.
      expect(textOf(container)).not.toContain(target);
      expect(textOf(container)).not.toContain("本次执行计划");
      expect(allBodies()).not.toContain(target);
    });

    it("新包装对象且回答期间源值改变：确认目标后仍读真实新值，不能用旧快照生成预览", async () => {
      const host = makeApplication(PRIVATE_SENTINEL);
      const app = host.application as WpsApplication;
      const book = app.ActiveWorkbook!;
      Object.defineProperty(app, "ActiveWorkbook", { get: () => ({ ...book,
        Worksheets: { ...book.Worksheets, Item: (index: number) => ({ ...book.Worksheets.Item(index) }) }
      }) });
      getWpsApplicationMock.mockReturnValue(app);
      const { container } = await mount(panel());
      await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
      await ask(container);
      await confirmBeforeAnalysis(container);
      await until(() => textOf(container).includes(PRIVATE_SENTINEL) &&
        Boolean(container.querySelector('.wps-local-mapping-values input[type="text"]')), "前置确认后应展示业务缺口");
      const count = planCalls().length;
      host.setRow3Value("新的私有来源-8F2H");
      await typeInput(container.querySelector<HTMLInputElement>('.wps-local-mapping-values input[type="text"]'), "旧值目标-4J2H");
      await click(findButton(container, "按补充映射重新生成预览"));
      await until(() => textOf(container).includes("新的私有来源-8F2H"), "应读取当前数据并在原卡询问新缺口");
      expect(findButton(container, "确认全部变更")).toBeNull();
      expect(planCalls().length).toBe(count);
      expect(host.writeCount()).toBe(0);
      expect(allBodies()).not.toContain("新的私有来源-8F2H");
      expect(localStorage.getItem("excel-bro.wps.chat.v1") ?? "").not.toContain("新的私有来源-8F2H");
    });

    it("主动询问未知值；补全后重新预览，全程不把源值或目标矩阵出站/落盘", async () => {
      const host = makeApplication(PRIVATE_SENTINEL);
      getWpsApplicationMock.mockReturnValue(host.application);

      const { container } = await mount(panel());
      await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");

      // prompt 里**不含**私有值：它只能从 A3 读出来。
      await ask(container);
      // 本地计算发现缺口后必须成为可操作的澄清卡，而不是终止错误。
      await until(
        () => textOf(container).includes("本地映射需要你决定"),
        "未知来源值应触发本地澄清"
      );

      // 正向锚点 1：确实发生过出站（意图 + 规划），否则下面的负向断言是空过。
      expect(captured.length).toBeGreaterThan(0);
      const turnRequests = captured.filter((item) => item.url.includes("/api/turn"));
      expect(turnRequests.length).toBeGreaterThanOrEqual(2);

      // 正向锚点 2：**没有任何写入预览**（没有"生成右侧副本并排审阅"按钮，也没有目标区域）。
      // 这正是"拒绝而非兜底"的可观测后果 —— 若生产代码把未命中值兜成 blankLabel，
      // 这里会冒出一个预览，用例立刻变红。
      expect(findButton(container, "确认全部变更")).toBeNull();
      expect(textOf(container)).not.toContain(RULE_TARGET_RANGE);
      expect(textOf(container)).not.toContain(TARGET_VALUE);
      expect(textOf(container)).not.toContain("已按映射处理");
      // 只在当前 DOM 展示，不能跟消息历史一起落盘。
      expect(textOf(container)).toContain(PRIVATE_SENTINEL);
      expect(localStorage.getItem("excel-bro.wps.chat.v1") ?? "").not.toContain(PRIVATE_SENTINEL);

      const target = "用户明确目标-8L2P";
      const fetchesBeforeAnswer = planCalls().length;
      const sourceReadsBeforeAnswer = host.wholeReads();
      await typeInput(container.querySelector<HTMLInputElement>(`.wps-local-mapping-values input`), target);
      await click(findButton(container, "按补充映射重新生成预览"));
      await until(() => textOf(container).includes("本次执行计划"), "补充后应生成新的写入预览");
      expect(textOf(container)).toContain(target);
      expect(findButton(container, "确认全部变更")).not.toBeNull();
      expect(host.wholeReads()).toBeGreaterThan(sourceReadsBeforeAnswer);
      expect(planCalls().length).toBe(fetchesBeforeAnswer);
      expect(localStorage.getItem("excel-bro.wps.chat.v1") ?? "").not.toContain(PRIVATE_SENTINEL);
      expect(localStorage.getItem("excel-bro.wps.chat.v1") ?? "").not.toContain(target);

      const outbound = allBodies();
      // 1) 私有值不出站 —— 规则/响应/提示都不知道它，所以这是真断言。
      expect(outbound).not.toContain(PRIVATE_SENTINEL);
      expect(outbound).not.toContain("7Q4X");
      expect(outbound).not.toContain(target);
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

    it("回答期间源值变化时重新读取并继续询问，不拿旧值生成预览", async () => {
      const host = makeApplication(PRIVATE_SENTINEL);
      getWpsApplicationMock.mockReturnValue(host.application);
      const { container } = await mount(panel());
      await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
      await ask(container);
      await until(() => textOf(container).includes(PRIVATE_SENTINEL), "应展示首个未知值");

      host.setRow3Value("新出现的本地值-5D8T");
      await typeInput(container.querySelector<HTMLInputElement>(".wps-local-mapping-values input"), "原值的目标");
      await click(findButton(container, "按补充映射重新生成预览"));
      await until(() => textOf(container).includes("新出现的本地值-5D8T"), "应询问重新读取后发现的新值");
      expect(findButton(container, "确认全部变更")).toBeNull();
      expect(allBodies()).not.toContain("新出现的本地值-5D8T");
      expect(localStorage.getItem("excel-bro.wps.chat.v1") ?? "").not.toContain("新出现的本地值-5D8T");
    });

    it("未声明分隔符时留在原卡补全规则，不取消卡片或重启模型", async () => {
      const compound = "甲，乙-仅本地";
      const host = makeApplication(compound);
      getWpsApplicationMock.mockReturnValue(host.application);
      planResponseFactory = () => {
        const response = localRuleTurnResponse(compound, false);
        return { ...response, rule: { ...response.rule, delimiters: ["；"], undeclaredSeparators: ["，"] } };
      };
      const { container } = await mount(panel());
      await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
      await ask(container);
      await until(() => textOf(container).includes(compound), "应展示无法归因的整格值");
      expect(findButton(container, "确认全部变更")).toBeNull();
      const fetchesBefore = planCalls().length;
      await click(findButton(container, "改为说明拆分规则"));
      expect(container.querySelector<HTMLInputElement>('input[aria-label="拆分符"]')?.value).toBe("，");
      expect(textOf(container)).toContain("多个类别同时出现时");
      expect(container.querySelector(".clarification-card.pending")).not.toBeNull();
      expect(container.querySelector<HTMLTextAreaElement>(".composer textarea")?.value).toBe("");
      expect(planCalls().length).toBe(fetchesBefore);
      expect(allBodies()).not.toContain(compound);
    });

    it("已拆分但缺类别时需显式确认优先级，补全后本地重读生成预览", async () => {
      const compound = `${KNOWN_CATEGORY}；${PRIVATE_SENTINEL}`;
      const host = makeApplication(compound);
      getWpsApplicationMock.mockReturnValue(host.application);
      planResponseFactory = () => {
        const response = localRuleTurnResponse(compound, false);
        return { ...response, rule: { ...response.rule, delimiters: ["；"] } };
      };
      const { container } = await mount(panel());
      await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
      await ask(container);
      await until(() => textOf(container).includes("拆分项未被规则覆盖"), "应主动询问缺失拆分项");
      expect(textOf(container)).toContain(PRIVATE_SENTINEL);
      expect(findButton(container, "按补充映射重新生成预览")?.hasAttribute("disabled")).toBe(true);
      const fetchesBefore = planCalls().length;
      const target = "补充目标-8P2L";
      await typeInput(container.querySelector<HTMLInputElement>(`input[aria-label="拆分项 ${PRIVATE_SENTINEL} 的目标值"]`), target);
      await act(async () => {
        container.querySelector<HTMLInputElement>('.local-mapping-priority input[type="checkbox"]')?.click();
      });
      await click(findButton(container, "按补充映射重新生成预览"));
      await until(() => textOf(container).includes("本次执行计划"), "补全优先级后应生成新预览");
      expect(planCalls().length).toBe(fetchesBefore);
      expect(allBodies()).not.toContain(target);
      expect(localStorage.getItem("excel-bro.wps.chat.v1") ?? "").not.toContain(target);
      expect(findButton(container, "确认全部变更")).not.toBeNull();
      expect(allBodies()).not.toContain(PRIVATE_SENTINEL);
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
      await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");

      await ask(container);
      try {
        await until(() => textOf(container).includes("本次执行计划"), "应展示写入计划");
      } catch (error) {
        console.log("Full text content:", textOf(container));
        console.log("Pending plan elements:", container.querySelectorAll('.wps-plan-preview, .wps-plan-outcome, [class*="plan"]').length);
        throw error;
      }

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
      /*
       * `"actions"` 这个键**会合法地**出现在 CAP-03 的账本登记里：那是恢复描述
       * （动作类型/表/区域/计数 + 两个指纹），不含任何单元格内容。新流程是
       * "计划就绪即执行"，所以登记必然发生 —— 拿整串关键字判泄漏会把合法描述
       * 判成缺陷。真正要钉的是"**没有计划本体**出站"：账本描述里不许出现
       * 计划 id、目标值矩阵或源快照。
       */
      for (const item of captured) {
        expect(item.body).not.toContain("wps-map-column-");
        expect(item.body).not.toContain('"values"');
        expect(item.body).not.toContain("sourceSnapshots");
      }

      // 5) 表头合法出站 —— 证明上面的负向断言不是"整个请求体为空"。
      expect(turnRequests[0].body).toContain("类别");
      expect(KNOWN_CATEGORY).not.toBe(SECOND_CATEGORY);
    });

    it("业务补全后重新生成：本地重算不走校验接口，也不把源快照或目标矩阵发出去", async () => {
      /*
       * 这条原来走「编辑计划 → 应用修改并重新预览」。新流程取消了编辑入口
       * （计划一经校验即执行，编辑一个已执行的动作没有意义），但它要钉的性质
       * 仍然有效且重要：**本地重新生成计划不经过服务端校验，也不上传矩阵**。
       * 所以改用仍然可达的"业务补全后重新生成"路径，保持同一个判别点。
       */
      const host = makeApplication(PRIVATE_SENTINEL);
      getWpsApplicationMock.mockReturnValue(host.application);

      const { container } = await mount(panel());
      await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");

      await ask(container);
      // 目标确认是否出现取决于本次授权现场：可能先弹确认，也可能已有前置授权直接进入业务缺口。
      await until(
        () => Boolean(container.querySelector('[role="dialog"]')) ||
          textOf(container).includes("本地映射需要你决定"),
        "应进入目标确认或本地澄清"
      );
      if (container.querySelector('[role="dialog"]')) await confirmBeforeAnalysis(container);
      await until(() => textOf(container).includes(PRIVATE_SENTINEL), "应展示本地未知来源值");

      // 记账：重新生成之前的出站量（只数会携带计划语义的那几类）。
      const beforeRegenerate = planCalls().length;
      const capturedBefore = captured.length;
      const target = "补全后目标-2R8K";
      await typeInput(container.querySelector<HTMLInputElement>('.wps-local-mapping-values input[type="text"]'), target);
      await until(() => !findButton(container, "按补充映射重新生成预览")!.disabled, "填写完整后应可生成");
      await click(findButton(container, "按补充映射重新生成预览"));
      await until(() => textOf(container).includes("本次执行计划"), "补全后应生成新的写入计划");

      const regenerateRequests = captured.slice(capturedBefore);
      // 锚点 1：这次重新生成**没有再打模型/校验接口**（本地规则重算）。
      expect(planCalls().length).toBe(beforeRegenerate);
      expect(
        regenerateRequests.filter((item) => item.url.includes("/api/plans/validate")).length
      ).toBe(0);
      // 锚点 2：新计划确实被重建了（否则上面一条会因"什么都没做"而空过）。
      expect(findButton(container, "确认全部变更")).not.toBeNull();
      expect(textOf(container)).toContain(target);
      // 锚点 3：重新生成期间产生的任何请求都不带源值、源快照或完整矩阵。
      for (const item of regenerateRequests) {
        expect(item.body).not.toContain(PRIVATE_SENTINEL);
        expect(item.body).not.toContain(target);
        expect(item.body).not.toContain("sourceSnapshots");
        expect(item.body).not.toContain('"values"');
        expect(item.body).not.toContain("wps-map-column-");
      }
    });

    it("算得档必须真的读映射源列 —— 对照锚点，证明澄清档的零读取是收敛结果", async () => {
      // 这条与澄清档的 `dataValueCount() === 0` 成对：同一个工作簿、同一份规则形状，
      // 只差"服务端有没有下发 clarification"。算得档需要源值，所以**必须**读到
      // A2:A3；澄清档不需要，所以**一格都不该**读到。
      //
      // 少了这条对照，澄清档的零读取断言可能因为"整条链路压根没接读取器"而空过。
      const { application, headerValueCount, dataValueCount, wholeReads } =
        makeApplication(SECOND_CATEGORY);
      getWpsApplicationMock.mockReturnValue(application);

      const { container } = await mount(panel());
      await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");

      await ask(container);
      await until(() => textOf(container).includes("本次执行计划"), "应展示写入计划");

      expect(headerValueCount()).toBeGreaterThan(0);
      expect(dataValueCount()).toBeGreaterThan(0);
      expect(wholeReads()).toBeGreaterThan(0);
      // 读取过 ≠ 出站过：源值与算好的目标矩阵都不离开本机。
      const outbound = allBodies();
      expect(outbound).not.toContain(KNOWN_CATEGORY);
      expect(outbound).not.toContain(SECOND_CATEGORY);
      expect(outbound).not.toContain(TARGET_VALUE);
      expect(outbound).not.toContain(SECOND_TARGET_VALUE);
      expect(outbound).not.toContain("sourceSnapshots");
    });
  });

  describe("规则自带澄清：缺类别时必须先问，且不许读源值", () => {
    beforeEach(() => {
      // 服务端在规则上挂了 `clarification`：它已判定"这条规则还缺类别，算不出正确答案"。
      planResponseFactory = () => ({
        ...localRuleTurnResponse(SECOND_CATEGORY, false),
        rule: {
          ...localRuleTurnResponse(SECOND_CATEGORY, false).rule,
          clarification: {
            reason: "matchedDelimiterMissingCategories",
            message: "有 2 个来源格拆出的类别不在映射表里。",
            missingCategories: [SECOND_CATEGORY],
            unresolvedCellCount: 0,
            resolvedBy: ["addCategories"]
          }
        }
      });
    });

    it("展示澄清卡、不生成预览，且没有整片源列读取、数据行值不出站", async () => {
      const { application, headerValueCount, dataValueCount, wholeReads } =
        makeApplication(SECOND_CATEGORY);
      getWpsApplicationMock.mockReturnValue(application);

      const { container } = await mount(panel());
      await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");

      await ask(container);
      await until(
        () => Boolean(container.querySelector(".composer-box .clarification-composer")),
        "缺类别时应先弹澄清卡，而不是算出一个猜的矩阵"
      );

      // 正向锚点 1：确实发生过出站（意图 + 规划）。
      const turnRequests = captured.filter((item) => item.url.includes("/api/turn"));
      expect(turnRequests.length).toBeGreaterThanOrEqual(2);
      // 正向锚点 2：选项是"补充类别映射"（归因确定 → 只有这一条路），
      // 且缺哪个类别写进了选项文案 —— 与 Excel 端走同一个
      // `columnMappingClarificationMessage`，所以这里钉的是"WPS 确实复用了同一份
      // 卡片翻译"。类别名在 `choices[].resolution` 里（`summary` 只是标题、
      // `question` 是规则给的 message），所以按**动作原文链路**断言。
      const stored = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1") ?? "null");
      const cardMessage = stored?.conversations?.[0]?.messages?.find(
        (message: { clarification?: { status?: string } }) =>
          message.clarification?.status === "pending"
      );
      expect(cardMessage).toBeTruthy();
      expect(cardMessage.clarification.summary).toBe("映射表缺这几个类别");
      // 选项文案（含缺失类别）持久化在卡片的 `answer` 之外的原始选项里，
      // 通过 pending 澄清的 options 校验"补充类别映射"存在。
      expect(
        (cardMessage.clarification.options as { label: string }[]).map((option) => option.label)
      ).toContain("补充类别映射");

      // ---- 真实读取面（本地）：按 `Value2` 访问范围取证，不认 purpose 标签 ----
      //
      // 正向锚点 3：结构扫描确实取到了表头 —— 否则下面的负向断言是空过。
      expect(headerValueCount()).toBeGreaterThan(0);
      // 判据 1：数据行（第 2 行及以后）的值**一格都没进内存**。
      // 澄清档排在 `readMappingSource` 之前，所以除结构扫描外没有任何源列读取；
      // 而结构扫描的矩形也必须止步于表头行。
      expect(dataValueCount()).toBe(0);
      // 判据 2：没有一次读取跨过表头与数据行 —— 没有整片源列读取。
      expect(wholeReads()).toBe(0);

      // ---- 出站面（向模型/服务端发送）与本地读取是两个独立证据面 ----
      // 表头会合法出站（字段名），所以先确认请求体非空，负向断言才不是空过。
      expect(turnRequests[0].body).toContain("类别");
      const outbound = allBodies();
      expect(outbound).not.toContain("sourceSnapshots");
      expect(outbound).not.toContain(PRIVATE_SENTINEL);
      expect(outbound).not.toContain(SECOND_CATEGORY);
      expect(outbound).not.toContain(TARGET_VALUE);
      // 映射源列地址与规则本体都不该离开本机。
      expect(outbound).not.toContain(RULE_SOURCE_RANGE);
      expect(outbound).not.toContain('"mapping"');
      // 零写入：没有预览、没有确认按钮、也不走校验接口。
      expect(findButton(container, "确认全部变更")).toBeNull();
      expect(captured.some((item) => item.url.includes("/api/plans/validate"))).toBe(false);
    });
  });

  describe("空白：真读到空白才问，选完就地重算新预览，全程零新增出站", () => {
    beforeEach(() => {
      // A3 是**空的**；规则只列了 A2 的已知类别，并且**没有**给空白策略 ——
      // 组合起来正是旧实现会"还没读就先问空白怎么办"的那条路。
      planResponseFactory = () => {
        const response = localRuleTurnResponse("", false);
        return { ...response, rule: { ...response.rule, blankLabel: undefined } };
      };
    });

    /**
     * 读出**本次执行计划**的真实 JSON —— 断言的是真正被执行的那些动作。
     *
     * 以前借"编辑计划"的文本框，现在执行即落地、不再可编辑；但同一份 JSON
     * 仍以只读方式挂在执行摘要卡里（`本次执行计划 JSON`），判据不变、入口更稳。
     */
    async function planUnderPreview(container: HTMLElement) {
      const raw = container.querySelector<HTMLPreElement>(
        '[aria-label="本次执行计划 JSON"]'
      )?.textContent;
      if (!raw) throw new Error("没有取到本次执行计划 JSON");
      return JSON.parse(raw) as {
        actions: { type: string; range: string; applyTo?: string }[];
      };
    }

    async function reachBlankCard(container: HTMLElement, host: ReturnType<typeof makeApplication>) {
      await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
      await ask(container);
      await until(() => textOf(container).includes("个空白格"), "应展示空白处理卡");
      // 计数来自真实读取，不是一句"可能有空白"：源列确实被读过。
      expect(host.wholeReads()).toBeGreaterThan(0);
      // 停在澄清卡上：没有预览、没有确认写入的入口。
      expect(findButton(container, "确认全部变更")).toBeNull();
    }

    it("保留目标原值：那一格不进任何写入动作，且全程零新增出站", async () => {
      const host = makeApplication("");
      getWpsApplicationMock.mockReturnValue(host.application);
      const { container } = await mount(panel());
      await reachBlankCard(container, host);

      const fetchesBefore = planCalls().length;
      await click(findButton(container, "按选定的空白处理重新生成预览"));
      await until(() => textOf(container).includes("本次执行计划"), "选定后应生成新预览");

      const plan = await planUnderPreview(container);
      // B3 一次都不写：公式、值、数字格式、条件格式都不动 —— 它不是"读回原值
      // 再写回"，而是压根不在动作里。
      expect(plan.actions.map((action) => action.range)).toEqual(["B2:B2"]);
      expect(JSON.stringify(plan.actions)).not.toContain("B3");

      // 重算在本机完成：没有再打一次模型/服务端。
      expect(planCalls().length).toBe(fetchesBefore);
      // 持久化与出站都不含源值或算好的目标矩阵。
      const stored = localStorage.getItem("excel-bro.wps.chat.v1") ?? "";
      expect(stored).not.toContain(TARGET_VALUE);
      expect(stored).not.toContain("3B8Z");
      const outbound = allBodies();
      expect(outbound).not.toContain(TARGET_VALUE);
      expect(outbound).not.toContain("sourceSnapshots");
      // 卡片与预览都不落盘源值：A3 是空的，A2 的已知类别也不该进历史。
      expect(stored).not.toContain(KNOWN_CATEGORY);
    });

    it("清空目标内容：走 clearRange 而不是写空串", async () => {
      const host = makeApplication("");
      getWpsApplicationMock.mockReturnValue(host.application);
      const { container } = await mount(panel());
      await reachBlankCard(container, host);

      await click(
        screen$ByLabelText(container, "清空目标内容 —— 清成真空单元格，不是写一个空字符串")
      );
      await click(findButton(container, "按选定的空白处理重新生成预览"));
      await until(() => textOf(container).includes("本次执行计划"), "选定后应生成新预览");

      const plan = await planUnderPreview(container);
      const cleared = plan.actions.find((action) => action.type === "clearRange");
      expect(cleared).toMatchObject({ range: "B3:B3", applyTo: "contents" });
      // 写空串是**另一种**意图，不会在这里出现：B3 上只有 clearRange。
      expect(plan.actions.filter((action) => action.range === "B3:B3")).toHaveLength(1);
    });
  });
});

/** 按 label 文本找元素（本文件用原生 DOM，不走 testing-library）。 */
function screen$ByLabelText(container: HTMLElement, label: string): HTMLElement {
  const match = Array.from(container.querySelectorAll("label")).find((item) =>
    (item.textContent ?? "").includes(label)
  );
  if (!match) throw new Error(`没有找到标签为「${label}」的选项`);
  return match.querySelector("input") ?? match;
}
