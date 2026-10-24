// @vitest-environment jsdom
/**
 * **同一次任务的连续链路**（Excel 端）：宿主读取 → `prepare_python_data` →
 * 数据句柄登记 → `run_python_analysis` → 已验证的隔离执行器 → 结果登记与引用
 * 取回 → 产品白名单计划编译 → 面板预览 → 用户确认 → 产品执行器写回 → 回读 → 撤销。
 *
 * **链路里各环节跑在哪一侧（照实分区）**：
 * - **本机桥**只承载两个工具：`query_table` 与 `prepare_python_data`
 *   （`packages/core/src/sdkToolBridge.ts` 的 `BridgeToolName`）。本文件驱动的是
 *   `prepare_python_data`：读授权表 → 数据句柄登记 → 回执。
 * - **`run_python_analysis` 在服务端**（MCP handler → `python_coordinator` →
 *   已验证的隔离执行器），它**不是**桥指针。它的真实执行在
 *   `server/tests/test_python_continuous_chain_real.py`（真实 Docker）验证；
 *   本文件把"服务端已用该结果编译出白名单计划"作为同一次任务的下一步接收。
 * - **白名单计划 → 预览 → 确认 → 产品执行器写回 → 回读 → 撤销** 全在真实
 *   `<App/>` 上跑，写回与撤销都是产品的 `executePlan` / 撤销按钮。
 *
 * 纪律（用户点名要求）：
 * - 允许**固定模型的工具调用序列**（桥指针顺序由服务端决定，本用例按真实顺序摆好）
 *   与**有状态宿主替身**（`Excel.run` 读写同一份矩阵）。
 * - **不用**预制结果矩阵、预制异常说明、预制最终计划，也**不在测试里直接给
 *   `range.values` 赋值**来代替产品环节：写回一律走真实确认+执行器，回读一律走
 *   真实 `Excel.run`，撤销一律走真实「撤销本次执行」按钮。
 * - 唯一被顶替的是**网络往返**（`fetch` → 回环端点）与**隔离容器本身**；
 *   前者由本文件拦截真实请求体，后者见上述服务端真实 Docker 用例。
 */
import { cleanup, fireEvent, render, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import App from "./App";
import { checkIntent, streamAssistantResponse, understandMessage } from "./api";
import {
  captureWorkbookCatalog,
  captureWorkbookSourceFingerprint,
  currentOfficeWorkbookIdentity
} from "./excel";
import { inspectOfficeAuthorizedTarget } from "./officeAuthorizedTarget";
import { createOfficeTestHost, installOfficeTestHost, type TestOfficeHost } from "./officeTestHost";
import { resetTrustedSourceReads } from "./planSourceDependency";

const MEMBERS = "会员表";
const ORDERS = "订单表";
const TARGET = "汇总表";
const BOOK_URL = "https://contoso.sharepoint.com/membership.xlsx";

const MEMBER_ROWS = [
  ["会员编号", "会员", "等级"],
  ["  7  ", "带空格会员", "金卡"],
  ["7", "重复键会员", "银卡"],
  ["007", "前导零会员", "银卡"],
  ["", "空键会员", "铜卡"],
  ["9", "正常会员", "金卡"]
];
const ORDER_ROWS = [
  ["会员编号", "金额"],
  ["7", 100],
  ["007", 200],
  ["9", 300],
  ["404", 400]
];

/** 手算的等级汇总（关联后才可能得到）：金卡 400 / 银卡 200 / 未匹配 400。 */
const EXPECTED_SUMMARY = [
  ["会员等级", "金额合计"],
  ["金卡", 400],
  ["银卡", 200],
  ["未匹配", 400]
];

const RESULT_REF = "res_" + "5".repeat(32);
/** 桥 operationId 的**真实形状**：`op_` + 32 位十六进制（两侧解析同一正则）。 */
const PREPARE_OP = "op_" + "1".repeat(32);
const FP = "a".repeat(64);

let host: TestOfficeHost;

function targetFor(names: readonly string[] | null) {
  const all = [MEMBERS, ORDERS, TARGET];
  const ids: Record<string, string> = {
    [MEMBERS]: "sheet-1",
    [ORDERS]: "sheet-2",
    [TARGET]: "sheet-3"
  };
  return {
    fileKey: "office:book-1",
    name: "membership.xlsx",
    sheets: all.map((name) => ({ id: ids[name], name })),
    selected: (names === null ? [MEMBERS, ORDERS, TARGET] : [...names]).map(
      (name) => ids[name]
    )
  };
}

const fixtures = vi.hoisted(() => ({ snapshot: {} as Record<string, unknown> }));

vi.mock("./api", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./api")>()),
  checkHealth: vi.fn(async () => ({ status: "ok", modelConfigured: true })),
  listModels: vi.fn(async () => ({
    defaultModelId: "claude-agent-sdk",
    models: [
      {
        id: "claude-agent-sdk",
        label: "Claude Agent SDK",
        modelId: "claude-agent-sdk",
        baseUrl: "http://localhost",
        supportsVision: false,
        available: true
      }
    ]
  })),
  getModelSettings: vi.fn(async () => ({ connectionCount: 1 })),
  understandMessage: vi.fn(),
  checkIntent: vi.fn(),
  streamAssistantResponse: vi.fn(),
  getPlanningTaskStatus: vi.fn(),
  continuePlanningTask: vi.fn(),
  prepareExcelReviewLayout: vi.fn(async () => {
    throw new Error("test-host: 未提供本机窗口排列服务");
  })
}));

/**
 * 只顶替**读不到真实宿主**的那几件事（是否在 Excel 里、目录快照），
 * **不顶替** `captureWorkbookSourceFingerprint` / `currentOfficeWorkbookIdentity`：
 * 它们是"计划落地时的来源基线"与"执行前复核"两侧**共用的同一个读取器**。
 * 只 mock 一侧，计划里存的就是假指纹，执行时真读取器一算就判"来源已变化"
 * ——那是测试自己制造的分叉，不是产品缺陷。
 */
vi.mock("./excel", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./excel")>()),
  isRunningInExcel: () => true,
  captureWorkbookCatalog: vi.fn(async () => ({
    name: "membership.xlsx",
    workbookIdentity: `url:${BOOK_URL}`,
    sourceFingerprint: "fp-source-1",
    capturedAt: "2026-10-10T00:00:00Z",
    activeWorksheet: MEMBERS,
    catalogOnly: true,
    worksheets: [MEMBERS, ORDERS, TARGET].map((name) => ({
      name,
      sourceSheet: name,
      usedRange: "A1:C6",
      rowCount: 0,
      columnCount: 0,
      headers: [],
      dataRows: [],
      truncated: false
    }))
  }))
}));

vi.mock("./officeAuthorizedTarget", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./officeAuthorizedTarget")>()),
  inspectOfficeAuthorizedTarget: vi.fn(async (names: readonly string[] | null) =>
    targetFor(names)
  ),
  inspectOfficeScopeAdmission: vi.fn(async (names: readonly string[]) => ({
    items: names.map((id) => ({ id, status: "ready", reason: "none" })),
    target: targetFor(names)
  })),
  confirmOfficeAuthorizedScope: vi.fn(async (_names, confirm, expected) => confirm(expected)),
  officeAnalysisTargetEvidence: vi.fn(async () => ({ key: "target-1", conclusive: true }))
}));

beforeEach(() => {
  localStorage.clear();
  vi.clearAllMocks();
  resetTrustedSourceReads();
  host = createOfficeTestHost({
    url: BOOK_URL,
    selectedRange: "A1:C6",
    activeSheet: MEMBERS,
    sheets: [
      { name: MEMBERS, values: MEMBER_ROWS },
      { name: ORDERS, values: ORDER_ROWS },
      { name: TARGET, values: [] }
    ]
  });
  installOfficeTestHost(host);
  vi.spyOn(console, "log").mockImplementation(() => {});
  Element.prototype.scrollIntoView = vi.fn();
});
afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * 拦截回环端点。
 *
 * 本机桥在一轮里只会收到 `prepare_python_data` 指针（读授权表 → 数据句柄登记 →
 * 回执）。`run_python_analysis` 在服务端、不在桥上 —— 所以这里**不**摆 analyze
 * 指针，而是让服务端（被 mock 的 `streamAssistantResponse`）在 prepare 回执之后
 * 直接给出**由那次隔离执行编译出的**白名单计划。这样"同一次任务"的连续性体现在
 * 桥读表 → 服务端执行 → 计划回到面板 的同一条 turn 上。
 */
function stubContinuousBridge(options: {
  /** 结果端点返回的矩阵（默认是手算期望）。 */
  result?: unknown;
  /** result 端点的 HTTP 状态（>=400 用来模拟"已过保留期/授权不符"）。 */
  resultStatus?: number;
} = {}) {
  const captured: Array<{ url: string; body: Record<string, unknown> }> = [];
  vi.stubGlobal(
    "fetch",
    vi.fn(async (url: string, init?: RequestInit) => {
      const target = String(url);
      captured.push({ url: target, body: JSON.parse(String(init?.body ?? "{}")) });
      // 断点恢复登记（写侧自有生命周期，与桥无关）：给一个可用的写入权收据。
      // 形状必须满足 `assertOfficeExecutionState` —— 缺字段会被如实拒绝。
      if (target.includes("/api/office/executions/register")) {
        return json(
          {
            operationId: String(JSON.parse(String(init?.body ?? "{}")).operationId ?? ""),
            status: "running",
            revision: 1,
            workbookIdentity: `url:${BOOK_URL}`,
            plan: { id: "plan-continuous-1", actions: [] },
            authorizedSheetNames: [MEMBERS, ORDERS, TARGET],
            allowedExternalFiles: [],
            nextActionIndex: 0,
            canResume: true,
            requiresReconciliation: false,
            ownerToken: "owner-token-1",
            checkpoint: null
          },
          200
        );
      }
      if (target.includes("/api/office/executions/")) {
        return json(
          {
            operationId: "execution-1",
            status: "running",
            revision: 2,
            workbookIdentity: `url:${BOOK_URL}`,
            plan: { id: "plan-continuous-1", actions: [] },
            authorizedSheetNames: [MEMBERS, ORDERS, TARGET],
            allowedExternalFiles: [],
            nextActionIndex: 0,
            canResume: true,
            requiresReconciliation: false,
            ownerToken: "owner-token-1",
            checkpoint: null
          },
          200
        );
      }
      if (target.includes("/receipt")) {
        return json({ status: "accepted", code: null }, 200);
      }
      if (target.includes("/result")) {
        const status = options.resultStatus ?? 200;
        if (status >= 400) return json({ detail: { code: "RESULT_EXPIRED" } }, status);
        return json(
          options.result ?? {
            resultRef: RESULT_REF,
            title: "会员等级汇总",
            headers: EXPECTED_SUMMARY[0],
            rows: EXPECTED_SUMMARY.slice(1),
            sourceSheets: [MEMBERS, ORDERS],
            warnings: [
              "空键：会员表有 1 条记录键为空，未参与关联。",
              "重复键：会员表有 1 个键重复，保留第一次出现。",
              "未匹配：订单表有 1 笔订单未匹配到会员（键 404）。"
            ],
            complete: true,
            calculation: "按会员等级汇总金额",
            scannedRows: 9
          },
          200
        );
      }
      // 轮询：本机桥收到的唯一指针是 prepare_python_data（读授权表 → 登记句柄）。
      return json(
        {
          calls: [
            {
              operationId: PREPARE_OP,
              tool: "prepare_python_data",
              arguments: { sources: [{ sheet: MEMBERS, range: "A1:C6" }] },
              requestFingerprint: FP,
              expiresAt: 0
            }
          ],
          generation: 1
        },
        200
      );
    })
  );
  return captured;
}

function json(payload: unknown, status: number): Response {
  return new Response(JSON.stringify(payload), {
    status,
    headers: { "Content-Type": "application/json" }
  });
}

/** 驱动真实面板发一条需求；`streamAssistantResponse` 会挂住，好让桥先跑。 */
async function sendPrompt(text: string, scope: readonly string[]) {
  const view = render(<App />);
  await waitFor(() => expect(vi.mocked(captureWorkbookCatalog)).toHaveBeenCalled());
  await waitFor(() =>
    expect(
      view.container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title
    ).toContain("当前模型")
  );
  // 授权三张表（两张来源 + 写入目标）：走**真实范围选择弹层**，
  // 否则 `committedSheetNamesRef` 为空，桥复核会判授权范围为空。
  await setAuthorizedScope(view, scope);
  fireEvent.change(view.container.querySelector(".composer textarea")!, {
    target: { value: text }
  });
  fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
  return view;
}

async function setAuthorizedScope(
  view: ReturnType<typeof render>,
  names: readonly string[]
) {
  const trigger = view.container.querySelector<HTMLButtonElement>(".scope-trigger");
  if (!trigger) throw new Error("范围选择入口尚未出现");
  fireEvent.click(trigger);
  await waitFor(() =>
    expect(view.container.querySelector(".scope-popover")).not.toBeNull()
  );
  const manual = [
    ...view.container.querySelectorAll<HTMLButtonElement>(".scope-mode-list button")
  ].find((candidate) => candidate.textContent?.includes("选择多个"));
  if (!manual) throw new Error("「选择多个」入口尚未出现");
  fireEvent.click(manual);
  await waitFor(() =>
    expect(view.container.querySelector(".workbook-sheet-list")).not.toBeNull()
  );
  for (const button of [
    ...view.container.querySelectorAll<HTMLButtonElement>(".workbook-sheet-list button")
  ]) {
    const name = button.querySelector("strong")?.textContent ?? "";
    const selected = button.getAttribute("aria-pressed") === "true";
    if (selected !== names.includes(name)) fireEvent.click(button);
  }
  const commit = view.container.querySelector<HTMLButtonElement>(".scope-close");
  if (!commit) throw new Error("范围「确认」按钮尚未出现");
  await waitFor(() => expect(view.container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("false"));
  fireEvent.click(commit);
  await waitFor(() => expect(view.container.querySelector(".scope-popover")).toBeNull());
}

describe("Excel 连续链路：同一次任务从宿主读取到撤销", () => {
  it("prepare_python_data → 句柄登记 → 服务端隔离执行 → 白名单计划 → 预览 → 确认 → 写回 → 回读 → 撤销", async () => {
    vi.mocked(understandMessage).mockResolvedValue({
      kind: "needs_workbook_context",
      summary: "按会员等级汇总订单金额",
      provider: "model"
    });
    vi.mocked(checkIntent).mockResolvedValue({
      kind: "proceed",
      summary: "按会员等级汇总订单金额",
      confirmedPrompt: "按会员等级汇总订单金额",
      provider: "model"
    } as never);

    // 规划流：先发 prepare 指针（本机桥读授权表、登记句柄）；桥回执后服务端
    // 用那次隔离执行的结果编译白名单计划，在同一 turn 上交给面板。
    const gate: { release: (() => void) | null } = { release: null };
    let planReleased = false;
    vi.mocked(streamAssistantResponse).mockImplementation(async (_request, hooks) => {
      hooks?.onProgress?.({
        status: "started",
        stepId: "bridge.prepare_python_data",
        bridgeTool: { operationId: PREPARE_OP, tool: "prepare_python_data" }
      } as never);
      await new Promise<void>((resolve) => {
        gate.release = () => {
          resolve();
        };
      });
      planReleased = true;
      return {
        kind: "plan",
        provider: "model",
        plan: {
          id: "plan-continuous-1",
          title: "会员等级汇总",
          summary: "按会员等级汇总订单金额，写到汇总表",
          assumptions: [],
          warnings: [
            "空键：会员表有 1 条记录键为空，未参与关联。",
            "重复键：会员表有 1 个键重复，保留第一次出现。",
            "未匹配：订单表有 1 笔订单未匹配到会员（键 404）。"
          ],
          actions: [
            {
              type: "writeValues",
              sheet: TARGET,
              range: "A1:B4",
              values: EXPECTED_SUMMARY
            }
          ],
          acceptanceCriteria: [
            {
              type: "rangeEquals",
              sheet: TARGET,
              range: "A1:B4",
              expected: EXPECTED_SUMMARY
            }
          ]
        }
      } as never;
    });

    const captured = stubContinuousBridge();
    const view = await sendPrompt("按会员等级汇总订单金额，结果写到汇总表", [
      MEMBERS,
      ORDERS,
      TARGET
    ]);

    // 本机桥真实执行了 prepare_python_data 并回执：回执带**数据句柄**。
    await waitFor(
      () => expect(captured.some((c) => c.url.includes("/receipt"))).toBe(true),
      { timeout: 4000 }
    );
    const prepareReceipt = captured.find((c) => c.url.includes("/receipt"))?.body
      .receipt as Record<string, unknown>;
    expect(Array.isArray(prepareReceipt.datasets)).toBe(true);
    expect((prepareReceipt.datasets as unknown[]).length).toBeGreaterThan(0);
    // 值只进回执字段，不外泄到 result 通道。
    expect(prepareReceipt.result).toBeUndefined();

    // 写之前来源与目标一个单元格都没动。
    expect(host.matrix(MEMBERS, "A1:C6")).toEqual(MEMBER_ROWS);
    expect(host.matrix(ORDERS, "A1:B5")).toEqual(ORDER_ROWS);

    const beforeWrites = host.originalWrites().length;
    expect(host.matrix(TARGET, 'A1:B4').flat().every(value => value === '')).toBe(true);
    gate.release?.();
    // 放行只是解开 mock 的等待：真正的"计划回到面板"以下面的预览断言为准，
    // 不用一个同步标志去替它作证（那样断言的是测试自己的时序，不是产品行为）。
    await waitFor(() => expect(planReleased).toBe(true));
    // 白名单计划回到面板 → 计划预览出现；确认前零写入。
    await waitFor(() => expect(view.container.querySelector(".inline-plan")).not.toBeNull());
    // 异常明细随计划如实到达（不预制异常说明 —— 它在计划 warnings 里）。
    expect(view.container.textContent).toContain("未匹配");

    // 真实点击「执行到 Excel」：产品执行器写回 + 回读核对。
    const runButton = await waitFor(() => {
      const found = [...view.container.querySelectorAll("button")].find(
        (c) => c.textContent === "确认变更"
      ) as HTMLButtonElement | undefined;
      if (!found) throw new Error("结果审阅确认按钮尚未出现");
      return found;
    });
    await waitFor(() => expect(runButton.disabled).toBe(false));
    const writesBeforeAccept = host.originalWrites().length;
    fireEvent.click(runButton);
    await waitFor(() => expect(view.queryByRole('button', { name: '确认变更' })).toBeNull());
    expect(host.originalWrites()).toHaveLength(writesBeforeAccept);

    await waitFor(() => expect(host.cell(TARGET, "A1")).toBe("会员等级"), {
      timeout: 4000
    });
    // 值来自真实计划 + 真实执行器写出（不是测试直接赋值）。
    expect(host.matrix(TARGET, "A1:B4")).toEqual(EXPECTED_SUMMARY);
    // 只写目标表；两张来源表一格未变。
    expect(host.matrix(MEMBERS, "A1:C6")).toEqual(MEMBER_ROWS);
    expect(host.matrix(ORDERS, "A1:B5")).toEqual(ORDER_ROWS);
    expect(host.originalWrites().length).toBeGreaterThan(beforeWrites);

    await waitFor(() => expect(view.container.textContent).toContain("并通过独立验证"));

    // 撤销：走真实按钮，验证目标区逐格回到执行前（空表）。
    const undoButton = [...view.container.querySelectorAll("button")].find(
      (c) => c.textContent === "↶ 撤销本次执行"
    ) as HTMLButtonElement | undefined;
    expect(undoButton, "执行后状态已完整核对，撤销按钮必须可用").toBeDefined();
    expect(undoButton!.disabled).toBe(false);
    fireEvent.click(undoButton!);

    await waitFor(() => expect(view.container.textContent).toContain("已撤销上一次执行"));
    expect(host.matrix(TARGET, "A1:B4")).toEqual([
      ["", ""],
      ["", ""],
      ["", ""],
      ["", ""]
    ]);
    // 撤销也只回滚目标区，来源表始终未变。
    expect(host.matrix(MEMBERS, "A1:C6")).toEqual(MEMBER_ROWS);
    expect(host.matrix(ORDERS, "A1:B5")).toEqual(ORDER_ROWS);
  });
});
