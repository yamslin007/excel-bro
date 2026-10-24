// @vitest-environment jsdom
/**
 * **双端在线连续集成（Excel 端）**：真实面板发起的同一次任务，从头到尾不换手。
 *
 * 与 `continuousPythonChain.test.tsx` / `crossProcessChain.test.tsx` 的区别（这是
 * 本文件的全部意义）：那两个文件是**工件回放集成** —— 面板消费服务端算好的结果，
 * 面板自己的 `prepare_python_data` 回执**没有参与**那次计算，`/result` 也由 fetch
 * 桩顶替。本文件把两半**接在同一条链上**：
 *
 *   真实 `<App/>` 发起任务 → 真实 `./api` 走真实 HTTP → 真实规划流（SSE）
 *   → 规划流下发 `bridgeTool` 指针 → 面板自己的桥轮询 → 读**有状态宿主替身**
 *   → 真实 `/receipt` 回执 → 服务端登记数据集 → `run_python_analysis`
 *   → **真实验证过的 Docker 隔离执行器**算出结果 → 登记结果引用
 *   → 真实 `submit_result_plan` 编译白名单计划 → 规划流把计划交回面板
 *   → 面板预览（确认前零写入）→ 用户点击「执行到 Excel」→ 产品执行器写回
 *   → 产品回读核对 → 用户点击「撤销本次执行」→ 目标区域逐格还原。
 *
 * 允许的替身（仅此两类）：
 * 1. **模型推理** —— 由 `server/tests/online_entry.py`（显式测试入口）注入固定工具
 *    调用序列。它只产出工具调用，不调用任何内部函数绕过分派/登记/编译/执行。
 * 2. **Office.js 宿主对象** —— `officeTestHost.ts` 的有状态替身（原表与副本分账）。
 *    授权范围读取（`./officeAuthorizedTarget`）与既有写侧端到端
 *    （`excelWriteEndToEnd.test.tsx`）一致地替换为宿主元数据替身：替身不模拟
 *    Office 的持久化范围/对话框表面。**数据与计划路径上的任何环节都没有被替换。**
 *
 * 明确**不做**的事：不为 poll/receipt/result/规划流摆 fetch 桩；不预制结果矩阵、
 * 异常说明或最终计划；不从旧 JSON 工件回放；测试代码不直接给 `range.values` 赋值。
 *
 * 环境阻断即不报成功：没有 `EB_ONLINE=1` + `VITE_API_BASE_URL` 时整个套件跳过，
 * 由 `scripts/run_online_link.mjs` 判定"未完成"。
 */
import { cleanup, fireEvent, render, waitFor } from "@testing-library/react";
import { transferableAbortController } from "node:util";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import scenario from "../../../config/online_link_scenario.json";
import App from "./App";
import { createOfficeTestHost, installOfficeTestHost, type TestOfficeHost } from "./officeTestHost";
import { currentOfficeWorkbookIdentity, invalidateWorkbookStructureCache } from "./excel";
import { resetTrustedSourceReads } from "./planSourceDependency";

const ONLINE = process.env.EB_ONLINE === "1";
const API_BASE = process.env.VITE_API_BASE_URL ?? "";
if (ONLINE && !API_BASE) {
  throw new Error("EB_ONLINE=1 但没有 VITE_API_BASE_URL：在线用例不能指向真实服务端。");
}

const MEMBERS = scenario.memberSheet;
const ORDERS = scenario.orderSheet;
const TARGET = scenario.targetSheet;
const ALL_SHEETS = [MEMBERS, ORDERS, TARGET];
const SHEET_IDS: Record<string, string> = {
  [MEMBERS]: "sheet-1",
  [ORDERS]: "sheet-2",
  [TARGET]: "sheet-3"
};

/**
 * 回读比对前先折空白：计划值来自**服务端编译层**，分析产出的 `""` 在那里被折成
 * `null`（写入意义上的空白只有"空单元格"一种）。宿主读回来仍是 `""` —— 两者是
 * 同一个空白的两种写法。归一后仍逐单元格严格：少写一格、值不同都会被抓到。
 */
function foldBlank(matrix: ReadonlyArray<ReadonlyArray<unknown>>) {
  return matrix.map((row) => row.map((cell) => (cell === "" ? null : cell)));
}

/**
 * 授权目标替身：只提供"文件里有哪些表 / 本次选了哪些"，与真实
 * `inspectOfficeAuthorizedTarget` 同语义（不能只返回被选中的表）。
 *
 * 身份两维（fileKey / fileRef）取**当前宿主现场**而不是常量：`officeAnalysisTargetEvidence`
 * 是真实实现，会核对 `target.fileRef === Office.context.document` 且
 * `identity === target.fileKey`。把证据键换成固定常量会让"同一组表两个键"的
 * 回归在这条集成链路里永远抓不到。
 */
function targetFor(names: readonly string[] | null) {
  const sheets = ALL_SHEETS.map((name) => ({ id: SHEET_IDS[name], name }));
  const chosen = names === null ? [...ALL_SHEETS] : [...names];
  return {
    fileKey: currentOfficeWorkbookIdentity() ?? null,
    fileRef: Office.context.document,
    name: scenario.workbookName,
    sheets,
    selected: chosen.map((name) => SHEET_IDS[name]).filter(Boolean)
  };
}

vi.mock("./officeAuthorizedTarget", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./officeAuthorizedTarget")>()),
  inspectOfficeAuthorizedTarget: vi.fn(async (names: readonly string[] | null) =>
    targetFor(names)
  ),
  inspectOfficeScopeAdmission: vi.fn(async (names: readonly string[]) => ({
    items: names.map((id) => ({ id, status: "ready", reason: "none" })),
    target: targetFor(names)
  })),
  confirmOfficeAuthorizedScope: vi.fn(async (_names, confirm, expected) =>
    confirm(expected)
  )
}));

let host: TestOfficeHost;

function createHost(memberRows: ReadonlyArray<ReadonlyArray<string | number>>) {
  return createOfficeTestHost({
    url: scenario.workbookIdentity.replace(/^url:/, ""),
    selectedRange: scenario.memberRange,
    activeSheet: MEMBERS,
    sheets: [
      { name: MEMBERS, values: memberRows },
      { name: ORDERS, values: scenario.orderRows },
      { name: TARGET, values: [] }
    ]
  });
}

beforeEach(() => {
  // Real Node fetch must receive Node signals, not jsdom's separate DOM realm.
  vi.stubGlobal('AbortController', transferableAbortController().constructor);
  localStorage.clear();
  vi.clearAllMocks();
  invalidateWorkbookStructureCache();
  resetTrustedSourceReads();
  host = createHost(scenario.memberRows);
  installOfficeTestHost(host);
  vi.spyOn(console, "log").mockImplementation(() => {});
  Element.prototype.scrollIntoView = vi.fn();
});
afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/** 走真实范围选择弹层授权三张表（不靠改 mock 造授权）。 */
async function setAuthorizedScope(view: ReturnType<typeof render>) {
  const trigger = view.container.querySelector<HTMLButtonElement>(".scope-trigger");
  if (!trigger) throw new Error("范围选择入口尚未出现");
  fireEvent.click(trigger);
  await waitFor(() =>
    expect(view.container.querySelector(".scope-popover")).not.toBeNull()
  );
  const manual = [
    ...view.container.querySelectorAll<HTMLButtonElement>(".scope-mode-list button")
  ].find((candidate) => candidate.textContent?.includes("选择多个"));
  if (manual) fireEvent.click(manual);
  await waitFor(() =>
    expect(view.container.querySelector(".workbook-sheet-list")).not.toBeNull()
  );
  // 按**非目录顺序**点选：授权集合的身份是成员的集合，不是点击先后。
  // 目录顺序（会员表 → 订单表 → 汇总表）恰好等于 adopt 重写后的顺序，
  // 那样"顺序差异"这条缺陷在这个用例里永远暴露不出来。
  const buttons = () =>
    [...view.container.querySelectorAll<HTMLButtonElement>(".workbook-sheet-list button")];
  for (const name of [ORDERS, TARGET, MEMBERS]) {
    const button = buttons().find(
      (candidate) => (candidate.querySelector("strong")?.textContent ?? "") === name
    );
    if (!button) throw new Error(`范围列表里没有工作表：${name}`);
    if (button.getAttribute("aria-pressed") !== "true") fireEvent.click(button);
  }
  for (const button of buttons()) {
    const name = button.querySelector("strong")?.textContent ?? "";
    if (button.getAttribute("aria-pressed") === "true" && !ALL_SHEETS.includes(name)) {
      fireEvent.click(button);
    }
  }
  const commit = view.container.querySelector<HTMLButtonElement>(".scope-confirm");
  if (!commit) throw new Error("范围「确认」按钮尚未出现");
  await waitFor(() => expect(commit.disabled).toBe(false));
  fireEvent.click(commit);
  await waitFor(() => expect(view.container.querySelector(".scope-popover")).toBeNull());
}

/** 发一条需求，等到白名单计划回到面板（真实 SSE + 真实桥），返回视图。 */
async function runTurn(count: number) {
  const view = render(<App />);
  await waitFor(() =>
    expect(
      view.container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title
    ).toContain("当前模型")
  );
  await setAuthorizedScope(view);
  if (count === 1) expect(host.originalWrites()).toHaveLength(0);
  expect(host.matrix(TARGET, scenario.targetRange).flat().every(v => v === "")).toBe(true);
  fireEvent.change(view.container.querySelector(".composer textarea")!, {
    target: { value: scenario.prompt }
  });
  fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
  await waitFor(
    () => expect(findButton(view, "确认变更"), view.container.textContent ?? '').toBeDefined(),
    { timeout: 60_000 }
  );
  return view;
}

function findButton(view: ReturnType<typeof render>, label: string) {
  // 取**最后一个**匹配：同一视图里历史消息会各自渲染自己的操作卡，
  // 要断言的是本轮新计划那张。
  return [...view.container.querySelectorAll("button")]
    .filter((candidate) => candidate.textContent === label)
    .at(-1) as HTMLButtonElement | undefined;
}

describe.skipIf(!ONLINE)("Excel 在线连续集成", () => {
  it("宿主读取 → 隔离容器计算 → 白名单计划 → 确认写回 → 回读 → 撤销；改宿主数据后结果随真实来源变化", async () => {
    // ---------------- 基线：金卡 400 / 银卡 200 / 未匹配 400 ----------------
    const view = await runTurn(1);

    // Results are already on the original; confirmation must not execute again.
    expect(host.matrix(MEMBERS, scenario.memberRange)).toEqual(scenario.memberRows);
    expect(host.matrix(ORDERS, scenario.orderRange)).toEqual(scenario.orderRows);
    // 计划里的值来自服务端隔离执行（不是测试常量）：预览文案带目标区域。
    expect(view.container.textContent).toContain(scenario.targetSheet);

    await waitFor(
      () =>
        expect(foldBlank(host.matrix(TARGET, scenario.targetRange))).toEqual(
          scenario.baselineExpected
        ),
      { timeout: 30_000 }
    );
    // 只写目标表；两张来源表一格未变。
    expect(host.matrix(MEMBERS, scenario.memberRange)).toEqual(scenario.memberRows);
    expect(host.matrix(ORDERS, scenario.orderRange)).toEqual(scenario.orderRows);
    expect(host.originalWrites().length).toBeGreaterThan(0);

    await waitFor(() => expect(view.container.textContent).toContain("并通过独立验证"), {
      timeout: 30_000
    });
    const writesBeforeAccept = host.originalWrites().length;
    const accept = findButton(view, "确认变更")!;
    await waitFor(() => expect(accept.disabled).toBe(false));
    fireEvent.click(accept);
    await waitFor(() => expect(findButton(view, "确认变更")).toBeUndefined());
    expect(host.originalWrites()).toHaveLength(writesBeforeAccept);
    // 确认只清标记并开放保存工具，不是第二次执行。
    expect(
      view.container.querySelector("button[title='保存为可重复使用的个人工具']")
    ).not.toBeNull();

    const undoButton = findButton(view, "↶ 撤销本次执行");
    expect(undoButton, "执行后状态已完整核对，撤销按钮必须可用").toBeDefined();
    expect(undoButton!.disabled).toBe(false);
    fireEvent.click(undoButton!);
    await waitFor(() => expect(view.container.textContent).toContain("已撤销上一次执行"));
    expect(
      host.matrix(TARGET, scenario.targetRange).flat().every((v) => v === "")
    ).toBe(true);
    expect(host.matrix(MEMBERS, scenario.memberRange)).toEqual(scenario.memberRows);
    expect(host.matrix(ORDERS, scenario.orderRange)).toEqual(scenario.orderRows);

    // ---------------- 因果验证：改**宿主**里键 7 的第一次出现 ----------------
    // 这一次改动只发生在 TS 宿主替身上：服务端夹具、mock、期望都没有为它改动。
    cleanup();
    // 新对话，避免把上一轮的历史卡片带进来（断言必须落在本轮新计划上）。
    localStorage.clear();
    const change = scenario.causalChange;
    host.setCell(change.sheet, change.cell, change.to);
    expect(host.cell(change.sheet, change.cell)).toBe(change.to);
    invalidateWorkbookStructureCache();
    resetTrustedSourceReads();

    const second = await runTurn(2);

    await waitFor(
      () =>
        expect(foldBlank(host.matrix(TARGET, scenario.targetRange))).toEqual(
          scenario.afterCausalChangeExpected
        ),
      { timeout: 30_000 }
    );
    // 来源数据本身只改了那一格；其余一格未动。
    expect(host.matrix(MEMBERS, scenario.memberRange)).toEqual(
      change.memberRows
    );
    expect(host.matrix(ORDERS, scenario.orderRange)).toEqual(scenario.orderRows);

    await waitFor(() => expect(second.container.textContent).toContain("并通过独立验证"), {
      timeout: 30_000
    });
    const secondWrites = host.originalWrites().length;
    const secondAccept = findButton(second, "确认变更")!;
    await waitFor(() => expect(secondAccept.disabled).toBe(false));
    fireEvent.click(secondAccept);
    await waitFor(() => expect(findButton(second, "确认变更")).toBeUndefined());
    expect(host.originalWrites()).toHaveLength(secondWrites);
    const secondUndo = findButton(second, "↶ 撤销本次执行")!;
    await waitFor(() => expect(secondUndo.disabled).toBe(false));
    fireEvent.click(secondUndo);
    await waitFor(() => expect(host.matrix(TARGET, scenario.targetRange).flat().every(v => v === "")).toBe(true));
  }, 180_000);
});
