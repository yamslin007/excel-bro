// @vitest-environment jsdom
/**
 * **跨前后端同一次任务的数据流**（Excel 端消费侧）。
 *
 * 与既有两端面板用例的区别（这正是用户点名要补的那一环）：
 * - `continuousPythonChain.test.tsx` / `wpsContinuousPythonChain.test.tsx`：面板是真的，
 *   但**结果矩阵、异常明细、最终计划都是测试里预制的常量**（`EXPECTED_SUMMARY`、
 *   `CONTINUOUS_PLAN`），`/result` 端点由 fetch mock 返回固定成功。
 * - **本文件**读 `config/cross_process_chain.json` —— 它由
 *   `server/tests/test_cross_process_chain_real.py` 在**真实 ASGI 服务端 +
 *   真实 Docker 隔离容器**里跑同一条链路后导出（结果引用 / 真实矩阵 / 异常明细 /
 *   服务端 `compile_result_plan` 编译的 `AnalysisPlan`）。
 *
 * 所以本文件里：
 * - `RESULT`（矩阵 + 异常明细 + resultRef）来自工件，**不是**手写常量；
 * - `PLAN`（写入坐标 + 值 + 验收条件）来自工件里**服务端编译**的计划，
 *   **不是** `CONTINUOUS_PLAN` 这类预制对象；
 * - 桥 `/result` 端点返回的就是工件里的真实响应体（仍拦截请求体以断言真实请求）。
 *
 * 链路：桥读授权表 → `prepare_python_data` 句柄登记 → 服务端隔离执行（工件即其产物）
 * → 白名单计划回到面板 → 预览 → 用户确认 → 产品执行器写回 → 回读 → 撤销。
 * 唯一被顶替的是**网络往返**（fetch → 回环端点）与**隔离容器本身**（其真实执行见
 * `server/tests/test_cross_process_chain_real.py`）—— 而本文件消费的正是后者的产物。
 *
 * 环境阻断即不报成功：工件由生产侧（真实服务端 + 真实 Docker 隔离容器）写出，
 * **不存在于仓库里**。因此这里既不静态 `import` 它（缺工件时会让默认 `vitest run`
 * 收集失败、`tsc --noEmit` 报 TS2307），也不预制替身。改为：
 * - 运行时按需读取工件（`EB_CROSS_PROCESS=1` 才读）；
 * - 没有该环境变量时整套用例跳过（与 `onlineLink.test.tsx` 同一编排），
 *   由 `scripts/run_cross_process_chain.mjs` 判「未完成」并以非零退出。
 */
import { existsSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
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
// 服务端真实执行的产物（由 `test_cross_process_chain_real.py` 导出）：
// 结果引用 + 矩阵 + 异常明细 + 服务端编译的白名单计划，全部来自一次真实执行。
//
// **不能**静态 `import` 这份 JSON：它不在仓库里（生产侧每次真实跑完才写），静态
// import 会在缺工件时把默认 `vitest run`（收集失败）与 `tsc --noEmit`（TS2307）
// 一起打崩 —— 那是**环境缺失**污染了构建，而不是产品缺陷。改为运行时按需读取，
// 并用 `EB_CROSS_PROCESS=1` 显式门控（只有生产侧确证产出工件的那个脚本会带上它）。
const CROSS_PROCESS = process.env.EB_CROSS_PROCESS === "1";
/*
 * 工件路径按 **vitest 的工作目录**（本 workspace 根）解析，不用 `import.meta.url`：
 * jsdom 环境下 `import.meta.url` 不是 file: URL，`fileURLToPath` 会抛
 * "The URL must be of scheme file"，而 `existsSync(URL)` 又一律返回 false ——
 * 两种写法都会把"工件就在仓库里"误报成"生产侧没有产出工件"。
 */
const ARTIFACT_PATH = resolve(process.cwd(), "..", "..", "config", "cross_process_chain.json");
if (CROSS_PROCESS && !existsSync(ARTIFACT_PATH)) {
  throw new Error(
    "EB_CROSS_PROCESS=1 但没有 config/cross_process_chain.json：" +
      "生产侧没有产出工件，拒绝用旧工件或替身回放。"
  );
}

/** 工件的形状（生产侧 `_export_fixture` 写出；全部字段都来自服务端）。 */
type CrossProcessArtifact = {
  workbookName: string;
  memberSheet: string;
  orderSheet: string;
  targetSheet: string;
  resultRef: string;
  result: { resultRef: string; headers: string[]; rows: unknown[][] };
  plan: Record<string, unknown>;
};

const artifact: CrossProcessArtifact = CROSS_PROCESS
  ? (JSON.parse(readFileSync(ARTIFACT_PATH, "utf8")) as CrossProcessArtifact)
  : {
      workbookName: "",
      memberSheet: "",
      orderSheet: "",
      targetSheet: "",
      resultRef: "",
      result: { resultRef: "", headers: [], rows: [] },
      plan: {}
    };

const MEMBERS = artifact.memberSheet;
const ORDERS = artifact.orderSheet;
const TARGET = artifact.targetSheet;
const BOOK_URL = "https://contoso.sharepoint.com/membership.xlsx";

/** 与工件逐字同一批脏数据（两份来源表；分析在服务端执行，这里只负责读）。 */
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

/**
 * 工件里的**真实服务端结果**（不是手写常量）：
 * - `resultRef`：服务端登记的结果引用（`res_<32hex>`）；
 * - `headers` / `rows`：`POST /api/sdk/tool-bridge/result` 的真实响应体；
 * - 异常明细（空键数 / 重复键数 / 未匹配键 404）就在 `rows` 里。
 */
const RESULT = artifact.result;
const RESULT_REF = artifact.resultRef;

/**
 * 工件里的**服务端编译计划**（不是预制 `CONTINUOUS_PLAN`）。
 * 写入坐标、值、验收条件都由服务端 `compile_result_plan` 给出。
 */
const PLAN = {
  ...artifact.plan,
  // 面板把计划当作 `AnalysisPlan` 接收；字段形状由服务端决定。
  provider: "model"
} as unknown as Record<string, unknown>;

/** 桥 operationId 的真实形状：`op_` + 32 位十六进制（两侧解析同一正则）。 */
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
 * - `/result`：返回**工件里的真实响应体**（服务端那次执行的产物），不是手写常量。
 * - 轮询：本机桥收到的唯一指针是 `prepare_python_data`（读授权表 → 登记句柄）。
 *   `run_python_analysis` 在服务端、不在桥上 —— 所以这里不摆 analyze 指针，
 *   而是让服务端（被 mock 的 `streamAssistantResponse`）在 prepare 回执之后直接
 *   给出**工件里那份由那次隔离执行编译出的**白名单计划。
 */
function stubCrossProcessBridge() {
  const captured: Array<{ url: string; body: Record<string, unknown> }> = [];
  vi.stubGlobal(
    "fetch",
    vi.fn(async (url: string, init?: RequestInit) => {
      const target = String(url);
      captured.push({ url: target, body: JSON.parse(String(init?.body ?? "{}")) });
      if (target.includes("/api/office/executions/register")) {
        return json(
          {
            operationId: String(JSON.parse(String(init?.body ?? "{}")).operationId ?? ""),
            status: "running",
            revision: 1,
            workbookIdentity: `url:${BOOK_URL}`,
            plan: { id: String(PLAN.id ?? "plan-1"), actions: [] },
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
            plan: { id: String(PLAN.id ?? "plan-1"), actions: [] },
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
        // **工件里的真实服务端响应体**（矩阵 + 异常明细 + 结果引用都来自它）。
        return json(RESULT, 200);
      }
      // 轮询：唯一指针是 prepare_python_data（读授权表 → 登记句柄）。
      return json(
        {
          calls: [
            {
              operationId: PREPARE_OP,
              tool: "prepare_python_data",
              arguments: {
                sources: [
                  { sheet: MEMBERS, range: "A1:C6" },
                  { sheet: ORDERS, range: "A1:B5" }
                ]
              },
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

async function sendPrompt(text: string, scope: readonly string[]) {
  const view = render(<App />);
  await waitFor(() => expect(vi.mocked(captureWorkbookCatalog)).toHaveBeenCalled());
  await waitFor(() =>
    expect(
      view.container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title
    ).toContain("当前模型")
  );
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
  const commit = view.container.querySelector<HTMLButtonElement>(".scope-confirm");
  if (!commit) throw new Error("范围「确认」按钮尚未出现");
  await waitFor(() => expect(commit.disabled).toBe(false));
  fireEvent.click(commit);
  await waitFor(() => expect(view.container.querySelector(".scope-popover")).toBeNull());
}

describe.skipIf(!CROSS_PROCESS)("跨前后端连续链路：消费服务端真实产物跑完整写回", () => {
  it("工件与真实执行自洽：矩阵/异常明细/计划都来自服务端（不是预制常量）", () => {
    // 结果引用是服务端登记句柄的形状。
    expect(RESULT_REF).toMatch(/^res_[0-9a-f]{32}$/);
    expect(RESULT.resultRef).toBe(RESULT_REF);
    // 矩阵第一行是分析代码的表头；异常明细就在 rows 里（不是另写的说明）。
    expect(RESULT.rows[0]).toEqual(["会员等级", "金额合计", ""]);
    expect(RESULT.rows).toContainEqual(["未匹配键", 1, "404"]);
    // 计划是服务端编译出来的：坐标、值、验收条件齐备。
    const action = (PLAN.actions as { type: string; sheet: string; range: string }[])[0];
    expect(action.type).toBe("writeValues");
    expect(action.sheet).toBe(TARGET);
    expect(PLAN.localResultRef).toBe(RESULT_REF);
    expect(PLAN.sourceDependencySheets).toEqual([MEMBERS, ORDERS]);
  });

  it("prepare_python_data → 服务端隔离执行（工件）→ 白名单计划 → 预览 → 确认 → 写回 → 回读 → 撤销", async () => {
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
    // 把**工件里那份由隔离执行编译出的**白名单计划在同一 turn 上交给面板。
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
        // **工件里的服务端编译计划**（不是测试预制的 `CONTINUOUS_PLAN`）。
        plan: PLAN
      } as never;
    });

    const captured = stubCrossProcessBridge();
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
    expect(prepareReceipt.result).toBeUndefined();

    // 写之前来源与目标一个单元格都没动。
    expect(host.matrix(MEMBERS, "A1:C6")).toEqual(MEMBER_ROWS);
    expect(host.matrix(ORDERS, "A1:B5")).toEqual(ORDER_ROWS);

    // 计划还没到，原表写入次数必须是 0；自动执行发生在计划到达之后。
    const beforeWrites = host.originalWrites().length;
    gate.release?.();
    await waitFor(() => expect(planReleased).toBe(true));
    await waitFor(() => expect(view.container.querySelector(".inline-plan")).not.toBeNull());
    /*
     * 异常明细随计划如实到达。
     *
     * 注意口径（照实写，不美化）：在**写入计划**这条通路上，空键/重复键/未匹配
     * 是以**结果矩阵的行**抵达的（`["未匹配键",1,"404"]` 就在计划的写入值里），
     * **不是**计划 warnings —— 那是服务端真实编译的产物形状，与两端既有面板用例
     * 里"异常写在预制 warnings 里"不同（那正是本用例要纠正的地方）。
     * 这里断言的是**工件计划**里的真实值，逐单元格。
     */
    const planActionForDetail = (PLAN.actions as { values: unknown[][] }[])[0];
    expect(planActionForDetail.values).toContainEqual(["未匹配键", 1, "404"]);
    // `null` 而不是 `""`：计划值经过**服务端编译层**的空白折叠
    // （`_blank_cell_of_write` 把分析结果里的 `""` 折成 `None`）—— 写入意义上的
    // 空白只有一种，即空单元格。这里断言的是**工件里真实编译出来的**值。
    expect(planActionForDetail.values).toContainEqual(["空键数", 1, null]);
    expect(planActionForDetail.values).toContainEqual(["重复键数", 1, null]);

    // 当前审阅契约：计划到达后由产品**自动执行一次**，不再有「执行到 Excel」
    // 手动入口，也不以双视图作为执行前置。这里等的是执行完成后的「确认变更」。
    const confirm = await waitFor(() => {
      const found = [...view.container.querySelectorAll("button")].find(
        (c) => c.textContent === "确认变更"
      ) as HTMLButtonElement | undefined;
      if (!found) throw new Error("计划到达后没有进入待确认状态");
      return found;
    }, { timeout: 15_000 });
    expect(host.originalWrites().length).toBeGreaterThan(beforeWrites);

    // 断言到**具体单元格**：值与坐标都来自工件里的服务端计划。
    //
    // 回读比对先折空白：计划值经**服务端编译层**折过（`""` → `null`，写入意义上
    // 的空白只有空单元格一种），而宿主读回来的是 `""`。两者是**同一个空白**的
    // 两种写法 —— 归一后再比，且仍逐单元格严格（少写一格、值不同都会被抓到）。
    const foldBlank = (cell: unknown) => (cell === "" ? null : cell);
    const planAction = (PLAN.actions as { range: string; values: unknown[][] }[])[0];
    const [endCell] = planAction.range.split(":").slice(-1);
    await waitFor(() => expect(host.cell(TARGET, "A1")).toBe("列1"), { timeout: 4000 });
    expect(host.matrix(TARGET, planAction.range).map((row) => row.map(foldBlank))).toEqual(
      planAction.values
    );
    expect(endCell).toBe("C8");
    // 只写目标表；两张来源表一格未变。
    expect(host.matrix(MEMBERS, "A1:C6")).toEqual(MEMBER_ROWS);
    expect(host.matrix(ORDERS, "A1:B5")).toEqual(ORDER_ROWS);
    expect(host.originalWrites().length).toBeGreaterThan(beforeWrites);

    await waitFor(() => expect(view.container.textContent).toContain("并通过独立验证"));

    // 确认只清标记、不第二次写入，并开放保存工具。
    const writesAfterExecution = host.originalWrites().length;
    fireEvent.click(confirm);
    await waitFor(() => expect(
      [...view.container.querySelectorAll("button")].some((c) => c.textContent === "确认变更")
    ).toBe(false));
    expect(host.originalWrites()).toHaveLength(writesAfterExecution);
    expect(
      view.container.querySelector("button[title='保存为可重复使用的个人工具']"),
      "确认后应开放保存工具"
    ).not.toBeNull();

    // 撤销：走真实按钮，验证目标区逐格回到执行前（空表）。
    const undoButton = [...view.container.querySelectorAll("button")].find(
      (c) => c.textContent === "↶ 撤销本次执行"
    ) as HTMLButtonElement | undefined;
    expect(undoButton, "执行后状态已完整核对，撤销按钮必须可用").toBeDefined();
    expect(undoButton!.disabled).toBe(false);
    fireEvent.click(undoButton!);

    await waitFor(() => expect(view.container.textContent).toContain("已撤销上一次执行"));
    const blank = planAction.values.map((row) => row.map(() => ""));
    expect(host.matrix(TARGET, planAction.range)).toEqual(blank);
    // 撤销也只回滚目标区，来源表始终未变。
    expect(host.matrix(MEMBERS, "A1:C6")).toEqual(MEMBER_ROWS);
    expect(host.matrix(ORDERS, "A1:B5")).toEqual(ORDER_ROWS);
  });
});
