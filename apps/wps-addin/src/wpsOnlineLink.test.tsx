// @vitest-environment jsdom
/**
 * **双端在线连续集成（WPS 端）**：真实面板发起的同一次任务，从头到尾不换手。
 *
 * 与 `wpsContinuousPythonChain.test.tsx` 的区别（这是本文件的全部意义）：那个文件
 * 是**工件回放集成** —— 白名单计划、结果矩阵、异常说明都是测试里的预制常量，
 * `fetch` 与服务端规划流都被桩顶替，面板自己的 `prepare_python_data` 回执
 * **没有参与**那次计算。本文件把两半接在**同一条链**上：
 *
 *   真实 `WpsAgentPanel` 发起任务 → 真实 `../../excel-addin/src/api` 走真实 HTTP
 *   → 真实规划流（SSE）→ 规划流下发 `bridgeTool` 指针 → 面板自己的桥轮询
 *   → 读**有状态宿主替身** → 真实 `/receipt` 回执 → 服务端登记数据集
 *   → `run_python_analysis` → **真实验证过的 Docker 隔离执行器**算出结果
 *   → 登记结果引用 → 真实 `submit_result_plan` 编译白名单计划
 *   → 规划流把计划交回面板 → 面板预览（确认前零写入）
 *   → 用户点「生成右侧副本并排审阅」→「确认写回原表」→ 产品执行器写回
 *   → 产品回读核对 → 用户点「↶ 撤销本次执行」→ 目标区域逐格还原。
 *
 * 允许的替身（仅此两类）：
 *
 * 1. **模型推理** —— 由 `server/tests/online_entry.py`（显式测试入口）注入固定工具
 *    调用序列。它只产出工具调用，不调用任何内部函数绕过分派/登记/编译/执行。
 * 2. **WPS 原生对象** —— 本文件底部的有状态宿主替身（`Range().Value2` 读写同一份
 *    矩阵，`Worksheet.Copy()` 真的生成一张**单表**未保存副本）。它被注入的位置
 *    只有一个：`./wpsHost` 的 `getWpsApplication()` —— 也就是"拿到原生 Application
 *    对象"那一步。数据与计划路径上的任何环节都没有被替换：
 *    `./wpsReadonlyAgent`、`./wpsNativeComparison`、`./wpsPlanExecutor`、
 *    `./wpsWorkbookReader`、`../../excel-addin/src/api` 全部跑真实实现。
 *
 * 明确**不做**的事：不为 poll/receipt/result/规划流摆 fetch 桩；不预制结果矩阵、
 * 异常说明或最终计划；不从旧 JSON 工件回放；测试代码不直接给 `Range().Value2`
 * 赋值来代替产品写回（宿主替身里唯一的"赋值"发生在 `Value2` 的 setter 上，
 * 也就是产品自己调用的那一处）。
 *
 * 两端共用 `config/online_link_scenario.json` 的同一份数据、同一套规则与同一套
 * 期望。期望矩阵只用于断言，绝不用于构造产品返回值。
 *
 * 环境阻断即不报成功：没有 `EB_ONLINE=1` + `VITE_API_BASE_URL` 时整个套件跳过，
 * 由 `scripts/run_online_link.mjs` 判定"未完成"。
 */
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import scenario from "../../../config/online_link_scenario.json";
import WpsAgentPanel from "./WpsAgentPanel";
import {
  getWpsApplication,
  type WpsApplication,
  type WpsRange,
  type WpsWorkbook,
  type WpsWorksheet
} from "./wpsHost";

const ONLINE = process.env.EB_ONLINE === "1";
const API_BASE = process.env.VITE_API_BASE_URL ?? "";
if (ONLINE && !API_BASE) {
  throw new Error("EB_ONLINE=1 但没有 VITE_API_BASE_URL：在线用例不能指向真实服务端。");
}

const MEMBERS = scenario.memberSheet;
const ORDERS = scenario.orderSheet;
const TARGET = scenario.targetSheet;
const ALL_SHEETS = [MEMBERS, ORDERS, TARGET];
const SHEET_INDEXES = [1, 2, 3];

/**
 * 回读比对前先折空白：计划值来自**服务端编译层**，分析产出的 `""` 在那里被折成
 * `null`（写入意义上的空白只有"空单元格"一种）。宿主读回来仍是 `""` —— 两者是
 * 同一个空白的两种写法。归一后仍逐单元格严格：少写一格、值不同都会被抓到。
 */
function foldBlank(matrix: ReadonlyArray<ReadonlyArray<unknown>>) {
  return matrix.map((row) => row.map((cell) => (cell === "" ? null : cell)));
}

type Matrix = (string | number | null)[][];
type Rect = { startRow: number; startColumn: number; endRow: number; endColumn: number };

/** 唯一被顶替的入口：拿到原生 WPS Application 对象。 */
vi.mock("./wpsHost", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./wpsHost")>()),
  getWpsApplication: () => currentHost.application()
}));

let currentHost: WpsHostDouble;

type WpsHostDouble = {
  application: () => WpsApplication;
  /** 读出真实矩阵副本（供"来源表一格未动"这类断言）。 */
  matrix: (sheet: string) => Matrix;
  cell: (sheet: string, row: number, column: number) => string | number | null;
  /** 直接改宿主里的某一格 —— 因果验证里"改键 7 的第一次出现"就走这里。 */
  setCell: (sheet: string, row: number, column: number, value: string | number) => void;
  /** 原表写入次数（审阅副本上的试写**不计入**，另见 `copyWrites`）。 */
  writes: () => number;
  copyWrites: () => number;
  /** 每次 `Worksheet.Copy()` 复制出去的工作表名 —— 用来证明副本只含一张表。 */
  copiedSheets: () => string[];
};

/** 极小 A1 解析（只支持 `A1`、`A1:B4` 这类无 `$` 的地址）。 */
function parseA1(address: string): Rect {
  const [start, end = start] = address.split(":");
  const toCell = (text: string) => {
    const match = /^([A-Z]+)(\d+)$/i.exec(text.trim());
    if (!match) throw new Error(`无法解析区域：${address}`);
    let column = 0;
    for (const char of match[1].toUpperCase()) column = column * 26 + (char.charCodeAt(0) - 64);
    return { column, row: Number(match[2]) };
  };
  const from = toCell(start);
  const to = toCell(end);
  return { startRow: from.row, startColumn: from.column, endRow: to.row, endColumn: to.column };
}

function createWpsHostDouble(memberRows: Matrix, orderRows: Matrix): WpsHostDouble {
  const grids: Record<string, Matrix> = {
    [MEMBERS]: memberRows.map((row) => [...row]),
    [ORDERS]: orderRows.map((row) => [...row]),
    [TARGET]: []
  };
  let originalWrites = 0;
  let copyWrites = 0;
  const copiedSheets: string[] = [];
  const storage = new Map<string, unknown>();
  let active: WpsWorkbook;

  /**
   * 一个矩形区域。值矩阵**每次现读**（写入后目标表会长大），单格返回标量、
   * 多格返回矩阵 —— 与真宿主一致。
   */
  function rectangle(read: () => Matrix, isOriginal: boolean, rect: Rect): WpsRange {
    const rows = rect.endRow - rect.startRow + 1;
    const columns = rect.endColumn - rect.startColumn + 1;
    const slice = () =>
      Array.from({ length: rows }, (_, rowOffset) =>
        Array.from({ length: columns }, (_, columnOffset) =>
          read()[rect.startRow - 1 + rowOffset]?.[rect.startColumn - 1 + columnOffset] ?? null
        )
      );
    const store = (incoming: unknown): void => {
      if (isOriginal) originalWrites += 1;
      else copyWrites += 1;
      const normalized = Array.isArray(incoming) ? (incoming as unknown[][]) : [[incoming]];
      const grid = read();
      for (let rowOffset = 0; rowOffset < rows; rowOffset += 1) {
        const rowIndex = rect.startRow - 1 + rowOffset;
        if (!grid[rowIndex]) grid[rowIndex] = [];
        for (let columnOffset = 0; columnOffset < columns; columnOffset += 1) {
          const value = normalized[rowOffset]?.[columnOffset];
          if (value !== undefined) grid[rowIndex][rect.startColumn - 1 + columnOffset] = value as never;
        }
      }
    };
    const range: Record<string, unknown> = {
      Rows: { Count: rows },
      Columns: { Count: columns },
      MergeCells: false,
      // 数据登记要取 `Text`/`NumberFormat` 才能保住前导零与日期形态。
      Interior: { Color: 0 },
      Item: (row: number, column = 1) =>
        rectangle(read, isOriginal, {
          startRow: rect.startRow + row - 1,
          startColumn: rect.startColumn + column - 1,
          endRow: rect.startRow + row - 1,
          endColumn: rect.startColumn + column - 1
        })
    };
    Object.defineProperties(range, {
      Value2: {
        get: () => {
          const current = slice();
          return rows === 1 && columns === 1 ? current[0][0] : current;
        },
        set: store,
        enumerable: true
      },
      Formula: {
        get: () => {
          const current = slice();
          return rows === 1 && columns === 1 ? current[0][0] : current;
        },
        set: store,
        enumerable: true
      },
      Text: { get: () => slice().map((row) => row.map((value) => (value == null ? "" : String(value)))), enumerable: true },
      NumberFormat: { get: () => slice().map((row) => row.map(() => "General")), enumerable: true }
    });
    return range as unknown as WpsRange;
  }

  function buildSheet(name: string, read: () => Matrix, isOriginal: boolean, onCopy?: () => void): WpsWorksheet {
    const grid = () => read();
    const sheet: Record<string, unknown> = {
      Name: name,
      CodeName: `Code_${name}`,
      ProtectContents: false,
      Range: (first: unknown, last?: unknown) => {
        if (typeof first === "string") return rectangle(grid, isOriginal, parseA1(first));
        const start = first as { Row?: number; Column?: number };
        const end = (last ?? first) as { Row?: number; Column?: number };
        return rectangle(grid, isOriginal, {
          startRow: Number(start.Row ?? 1),
          startColumn: Number(start.Column ?? 1),
          endRow: Number(end.Row ?? start.Row ?? 1),
          endColumn: Number(end.Column ?? start.Column ?? 1)
        });
      }
    };
    Object.defineProperties(sheet, {
      UsedRange: {
        get: () => {
          const current = grid();
          const columnCount = Math.max(0, ...current.map((row) => row.length));
          const used: Record<string, unknown> = {
            Row: 1,
            Column: 1,
            Rows: { Count: current.length },
            Columns: { Count: columnCount },
            Item: (row: number, column = 1) => ({ Row: row, Column: column }),
            MergeCells: false
          };
          Object.defineProperty(used, "Value2", { get: () => current.map((row) => [...row]), enumerable: true });
          return used;
        },
        enumerable: true
      }
    });
    if (onCopy) sheet.Copy = onCopy;
    return sheet as unknown as WpsWorksheet;
  }

  const originalSheets = new Map<string, WpsWorksheet>();
  function originalSheet(name: string): WpsWorksheet {
    const cached = originalSheets.get(name);
    if (cached) return cached;
    const built = buildSheet(name, () => grids[name] ?? [], true, () => {
      copiedSheets.push(name);
      const snapshot = (grids[name] ?? []).map((row) => [...row]);
      active = buildCopyBook(name, () => snapshot);
    });
    originalSheets.set(name, built);
    return built;
  }

  function buildCopyBook(sheetName: string, read: () => Matrix): WpsWorkbook {
    const sheet = buildSheet(sheetName, read, false);
    const book: WpsWorkbook = {
      // 未保存副本：没有盘符路径，`workbookFullName` 与 `Path` 都为空。
      Name: "Book1",
      FullName: "",
      Path: "",
      ReadOnly: false,
      Worksheets: { Count: 1, Item: () => sheet },
      ActiveSheet: sheet,
      Windows: { Count: 1, Item: () => ({ Caption: "Book1" }) },
      Activate: () => {
        active = book;
      },
      Close: (saveChanges?: boolean) => {
        if (saveChanges !== false) throw new Error("副本不得保存");
        if (active === book) active = originalBook;
        return true;
      }
    };
    return book;
  }

  const originalBook: WpsWorkbook = {
    Name: scenario.workbookName,
    // 必须是真实 Windows 盘符反斜杠路径：写入闸只对已保存的本地 .xlsx 开放。
    FullName: "C:\\Data\\会员.xlsx",
    Path: "C:\\Data",
    ReadOnly: false,
    // 真宿主对不存在的名字/越界序号会抛；替身返回 undefined，由
    // `authorizeWpsSheets` 判成"授权前提不成立"。故整体收窄成集合类型。
    Worksheets: {
      get Count() {
        return ALL_SHEETS.length;
      },
      Item: (key: number | string) => {
        if (typeof key === "string" && !/^\d+$/.test(key)) {
          return ALL_SHEETS.includes(key) ? originalSheet(key) : undefined;
        }
        const name = ALL_SHEETS[Number(key) - 1];
        return name ? originalSheet(name) : undefined;
      }
    } as unknown as WpsWorkbook["Worksheets"],
    ActiveSheet: originalSheet(MEMBERS),
    Windows: { Count: 1, Item: () => ({ Caption: scenario.workbookName }) },
    Activate: () => {
      active = originalBook;
    }
  };
  active = originalBook;

  const application = {
    Version: "12.0",
    get ActiveWorkbook() {
      return active;
    },
    get ActiveSheet() {
      return active.ActiveSheet;
    },
    Windows: {
      Count: 1,
      Item: () => ({ Caption: active.Name }),
      CompareSideBySideWith: () => true,
      SyncScrollingSideBySide: false
    },
    PluginStorage: {
      getItem: (key: string) => storage.get(key),
      setItem: (key: string, value: unknown) => {
        storage.set(key, value);
      }
    },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  } as unknown as WpsApplication;

  return {
    application: () => application,
    matrix: (sheet: string) => (grids[sheet] ?? []).map((row) => [...row]),
    cell: (sheet: string, row: number, column: number) => grids[sheet]?.[row - 1]?.[column - 1] ?? null,
    setCell: (sheet: string, row: number, column: number, value: string | number) => {
      const grid = grids[sheet];
      if (!grid) throw new Error(`宿主里没有「${sheet}」`);
      if (!grid[row - 1]) grid[row - 1] = [];
      grid[row - 1][column - 1] = value;
    },
    writes: () => originalWrites,
    copyWrites: () => copyWrites,
    copiedSheets: () => [...copiedSheets]
  };
}

// 本文件用 `createRoot` 直接挂真实面板，必须打开 React 的 act 环境，
// 否则 `act(...)` 只警告不生效，状态更新会在断言之后才落地。
(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

type Mounted = { container: HTMLElement; unmount: () => void };
let mounted: Mounted[] = [];

if (!Element.prototype.scrollIntoView) {
  Element.prototype.scrollIntoView = () => undefined;
}

async function settle(): Promise<void> {
  await act(async () => {
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
  });
}

/** 轮询等待。message 在调用时即求值，所以现场要落在 `check()` 里。 */
async function until(check: () => boolean, message: () => string): Promise<void> {
  for (let index = 0; index < 600; index += 1) {
    if (check()) return;
    await settle();
  }
  throw new Error(message());
}

async function mount(element: React.ReactElement): Promise<HTMLElement> {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  await act(async () => {
    root.render(element);
  });
  await settle();
  mounted.push({ container, unmount: () => act(() => root.unmount()) });
  return container;
}

function panel(): React.ReactElement {
  return (
    <WpsAgentPanel
      selectedSheetIndexes={SHEET_INDEXES}
      scopeDock={<div className="eb-scope-dock">当前表 会员表、订单表、汇总表</div>}
    />
  );
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

/** 结果卡里的撤销按钮。 */
function cardUndoButton(container: HTMLElement): HTMLButtonElement | null {
  return container.querySelector<HTMLButtonElement>(
    '.wps-verification-card button[aria-label="撤销本次执行"]'
  );
}

/**
 * 审阅卡里的「我已核对副本」确认项。
 *
 * 写回入口在勾选前是**禁用**的（确认框是真闸）：不勾就点，写入会被拦下、
 * 原表一格不写 —— 所以走完写回链路前必须先勾这里。
 */
async function acknowledgeCopy(container: HTMLElement): Promise<void> {
  const review = container.querySelector<HTMLElement>(".wps-native-review-status");
  const checkbox = review?.querySelector<HTMLInputElement>('input[type="checkbox"]') ?? null;
  if (!checkbox) throw new Error("审阅卡里应有「我已核对副本」确认项");
  if (!checkbox.checked) await click(checkbox);
}

async function type(element: HTMLTextAreaElement | null, value: string): Promise<void> {
  if (!element) throw new Error("没有找到输入框");
  const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, "value")?.set;
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

function snapshotText(container: HTMLElement): string {
  return (container.textContent ?? "").slice(-600);
}

beforeEach(() => {
  localStorage.clear();
  vi.clearAllMocks();
  currentHost = createWpsHostDouble(
    scenario.memberRows as unknown as Matrix,
    scenario.orderRows as unknown as Matrix
  );
  vi.spyOn(console, "log").mockImplementation(() => {});
});

afterEach(() => {
  mounted.forEach((entry) => entry.unmount());
  mounted = [];
  document.body.innerHTML = "";
  vi.restoreAllMocks();
});

describe.skipIf(!ONLINE)("WPS 在线连续集成", () => {
  it("宿主读取 → 隔离容器计算 → 白名单计划 → 确认写回 → 回读 → 撤销；改宿主数据后结果随真实来源变化", async () => {
    // ---------------- 基线：金卡 400 / 银卡 200 / 未匹配 400 ----------------
    const container = await mount(panel());
    await until(
      () =>
        container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型") ??
        false,
      () => "应加载模型目录：" + snapshotText(container)
    );

    const textarea = container.querySelector("textarea");
    await type(textarea, scenario.prompt);
    await submit(textarea);

    // 预览期：来源与目标一个单元格都没动（审阅副本上的试写另计）。
    await until(
      () => findButton(container, "生成右侧副本并排审阅") !== null,
      () => "应出现写入计划预览（生成右侧副本并排审阅）：" + snapshotText(container)
    );
    expect(currentHost.writes()).toBe(0);
    expect(currentHost.copiedSheets()).toHaveLength(0);
    expect(currentHost.matrix(MEMBERS)).toEqual(scenario.memberRows);
    expect(currentHost.matrix(ORDERS)).toEqual(scenario.orderRows);
    expect(currentHost.matrix(TARGET).flat().every((value) => value === null || value === "")).toBe(true);
    // 计划里的值来自服务端隔离执行（不是测试常量）：预览文案带目标表名。
    expect(container.textContent).toContain(TARGET);

    // 副本只承载**写入目标**那一张表，绝不复制整本工作簿。
    await click(findButton(container, "生成右侧副本并排审阅"));
    await until(
      () => findButton(container, "确认写回原表") !== null,
      () => "审阅副本应生成并给出确认写回入口：" + snapshotText(container)
    );
    expect(currentHost.copiedSheets()).toEqual([TARGET]);
    expect(currentHost.writes()).toBe(0);
    expect(currentHost.copyWrites()).toBeGreaterThan(0);
    // 两张来源表在预览后仍未改变。
    expect(currentHost.matrix(MEMBERS)).toEqual(scenario.memberRows);
    expect(currentHost.matrix(ORDERS)).toEqual(scenario.orderRows);

    // 真实确认 → 产品执行器写回 → 回读核对。
    await acknowledgeCopy(container);
    await click(findButton(container, "确认写回原表"));
    await until(
      () => container.querySelector(".wps-verification-card") !== null,
      () => "写入后应出现执行结果卡：" + snapshotText(container)
    );
    await until(
      () => currentHost.matrix(TARGET).length === scenario.baselineExpected.length,
      () => "目标区域应被写入：" + JSON.stringify(currentHost.matrix(TARGET))
    );
    expect(foldBlank(currentHost.matrix(TARGET))).toEqual(scenario.baselineExpected);
    // 只写目标表；两张来源表一格未变。
    expect(currentHost.matrix(MEMBERS)).toEqual(scenario.memberRows);
    expect(currentHost.matrix(ORDERS)).toEqual(scenario.orderRows);
    expect(currentHost.writes()).toBeGreaterThan(0);
    // 回读核对通过（产品自己的判据，不是测试替它下的结论）。
    expect(container.textContent).toContain("完成目标区域回读核对");

    // 撤销：走真实按钮，目标区域回到执行前的空表。
    const undo = cardUndoButton(container);
    expect(undo, "执行结果卡里必须挂可用的撤销按钮").not.toBeNull();
    await click(undo);
    await until(
      () => container.textContent?.includes("已撤销并核对恢复") ?? false,
      () => "撤销后应显示已撤销并核对恢复：" + snapshotText(container)
    );
    expect(currentHost.matrix(TARGET).flat().every((value) => value === null || value === "")).toBe(true);
    expect(currentHost.matrix(MEMBERS)).toEqual(scenario.memberRows);
    expect(currentHost.matrix(ORDERS)).toEqual(scenario.orderRows);

    // ---------------- 因果验证：改**宿主**里键 7 的第一次出现 ----------------
    // 这一次改动只发生在 TS 宿主替身上：服务端夹具、mock、期望都没有为它改动。
    mounted.forEach((entry) => entry.unmount());
    mounted = [];
    document.body.innerHTML = "";
    // 新对话，避免把上一轮的历史卡片带进来（断言必须落在本轮新计划上）。
    localStorage.clear();
    const change = scenario.causalChange;
    currentHost.setCell(change.sheet, 2, 3, change.to);
    expect(currentHost.cell(change.sheet, 2, 3)).toBe(change.to);

    const second = await mount(panel());
    await until(
      () =>
        second.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型") ??
        false,
      () => "（第二轮）应加载模型目录：" + snapshotText(second)
    );
    await type(second.querySelector("textarea"), scenario.prompt);
    await submit(second.querySelector("textarea"));
    await until(
      () => findButton(second, "生成右侧副本并排审阅") !== null,
      () => "（第二轮）应出现写入计划预览：" + snapshotText(second)
    );
    await click(findButton(second, "生成右侧副本并排审阅"));
    await until(
      () => findButton(second, "确认写回原表") !== null,
      () => "（第二轮）审阅副本应生成：" + snapshotText(second)
    );
    await acknowledgeCopy(second);
    await click(findButton(second, "确认写回原表"));

    await until(
      () => second.querySelector(".wps-verification-card") !== null,
      () => "（第二轮）写入后应出现执行结果卡：" + snapshotText(second)
    );
    // 撤销后目标区留着 6×3 的空矩阵，所以判别点必须是**表头真的写进来了**，
    // 不能只看行数。
    await until(
      () => currentHost.matrix(TARGET)[0]?.[0] === scenario.afterCausalChangeExpected[0][0],
      () => "（第二轮）目标区域应被写入：" + JSON.stringify(currentHost.matrix(TARGET))
    );
    // 金卡 300 / 银卡 300 / 未匹配 400 —— 只能来自宿主里那一格的真实改动。
    expect(foldBlank(currentHost.matrix(TARGET))).toEqual(
      scenario.afterCausalChangeExpected
    );
    // 来源数据本身只改了那一格；其余一格未动。
    expect(currentHost.matrix(MEMBERS)).toEqual(change.memberRows);
    expect(currentHost.matrix(ORDERS)).toEqual(scenario.orderRows);
    await until(
      () => second.textContent?.includes("完成目标区域回读核对") ?? false,
      () => "（第二轮）应完成回读核对：" + snapshotText(second)
    );
  }, 240_000);
});
