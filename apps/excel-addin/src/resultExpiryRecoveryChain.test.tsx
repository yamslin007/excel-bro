// @vitest-environment jsdom
/**
 * **A1-E2 消费侧**（Excel 端）：消费真实链跑出来的"引用过期恢复"工件，把恢复后的
 * 那份**服务端编译计划**在真实 `<App/>` 上执行，验证"执行一次 → 回读 → 审阅"。
 *
 * 工件从哪来：`scripts/drive_real_sdk_result_expiry.py` 在**真实模型 + 真实服务端 +
 * 真实登记器 + 真实 Docker 隔离执行**上跑完整条链后写出 `outputs/a1-e2-recovery.json`。
 * 里面同时留着：
 * - `registeredResultRefs[0]`：被 A1-E1 精确失效的**旧引用**（服务端已删，取回 404）；
 * - `registeredResultRefs[1]`：模型重算后登记的**新引用**（可读）；
 * - `plan`：服务端用**新引用**编译出的白名单计划；
 * - `fault.timeline`：armed → bound → consumed 的时序（"确实在那一刻失效"）。
 *
 * 所以本文件**不预制**矩阵、异常明细与计划，也不摆"第二次 analyze"的假指针；
 * 唯一顶替的是网络往返（fetch → 回环端点）与隔离容器本身。
 *
 * 环境阻断即不报成功：工件不在仓库里（生产侧每跑一次才写），所以运行时按需读取，
 * 用 `EB_A1_E2=1` 显式门控；`scripts/run_a1_e2_chain.mjs` 把"跳过"判成未完成。
 */
import { existsSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { cleanup, fireEvent, render, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import App from "./App";
import { checkIntent, streamAssistantResponse, understandMessage } from "./api";
import { captureWorkbookCatalog } from "./excel";
import { createOfficeTestHost, installOfficeTestHost, type TestOfficeHost } from "./officeTestHost";
import { resetTrustedSourceReads } from "./planSourceDependency";

const A1_E2 = process.env.EB_A1_E2 === "1";
/*
 * 路径按 **vitest 的工作目录**（`apps/excel-addin`）解析 —— 不用 `import.meta.url`：
 * jsdom 下它不是 file: URL，`fileURLToPath` 会抛错、`existsSync(URL)` 又恒 false，
 * 两种写法都会把"工件就在仓库里"误报成"生产侧没产出"。
 */
const ARTIFACT_PATH = resolve(
  process.cwd(),
  "..",
  "..",
  // 默认消费 WPS 宿主那一轮；Excel 宿主那一轮由 `EB_A1_E2_ARTIFACT` 指过来 ——
  // 两端各消费自己那一轮的真实产物，不拿一端的工件替另一端作证。
  process.env.EB_A1_E2_ARTIFACT ?? "outputs/a1-e2-recovery.json"
);
if (A1_E2 && !existsSync(ARTIFACT_PATH)) {
  throw new Error(
    "EB_A1_E2=1 但没有 outputs/a1-e2-recovery.json：" +
      "生产侧没有产出工件，拒绝用旧工件或替身回放。"
  );
}

type PlanAction = {
  type: string;
  sheet: string;
  range: string;
  values: (string | number | null)[][];
};

type Artifact = {
  scenario: string;
  turnId: string;
  workbook: { name: string; identity: string; hostId: string };
  fault: {
    armed: boolean;
    timeline: Array<{ phase: string; taskBound: boolean; refBound: boolean }>;
    final: { phase: string; taskBound: boolean; refBound: boolean };
  };
  bridge: {
    served: Array<{ index: number; tool: string; refused: boolean }>;
    prepareCount: number;
    refusals: Array<{ index: number; tool: string; code: string }>;
  };
  registeredResultRefs: string[];
  probes: Array<{ index: number; resultRef: string; ok: boolean; status: number }>;
  repeatOldRef: { ok: boolean; status: number } | null;
  plan: {
    id: string;
    localResultRef: string;
    sourceDependencySheets: string[];
    actions: PlanAction[];
    acceptanceCriteria: Array<{ type: string; expected?: (string | number | null)[][] }>;
  } | null;
  planActionCount: number;
};

const artifact: Artifact = A1_E2
  ? (JSON.parse(readFileSync(ARTIFACT_PATH, "utf8")) as Artifact)
  : {
      scenario: "recovery",
      turnId: "",
      workbook: { name: "", identity: "", hostId: "" },
      fault: { armed: false, timeline: [], final: { phase: "", taskBound: false, refBound: false } },
      bridge: { served: [], prepareCount: 0, refusals: [] },
      registeredResultRefs: [],
      probes: [],
      repeatOldRef: null,
      plan: null,
      planActionCount: 0
    };

const PLAN = (artifact.plan ?? {}) as unknown as Record<string, unknown> & {
  actions: PlanAction[];
  localResultRef: string;
  acceptanceCriteria: Array<{ type: string; expected?: (string | number | null)[][] }>;
};
const REFS = artifact.registeredResultRefs;
/** 被 A1-E1 精确失效的那一个：旧引用必须永远取不回来。 */
const OLD_REF = REFS[0] ?? "";
/** 模型重算后登记、也是计划实际引用的那一个。 */
const NEW_REF = REFS[1] ?? "";

const WORKBOOK_NAME = "membership.xlsx";
const BOOK_URL = "https://contoso.sharepoint.com/membership.xlsx";
/**
 * 工件里的计划目标是 `汇总表`；来源是 `会员表` / `订单表`。
 * 这三个名字来自工件（不是测试里另起的名字），否则"来源未变"就无法对账。
 */
const MEMBERS = "会员表";
const ORDERS = "订单表";
const TARGET = "汇总表";

/** 与生产侧驱动逐字同一批脏数据（分析在服务端算，本端只负责读）。 */
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
    name: WORKBOOK_NAME,
    sheets: all.map((name) => ({ id: ids[name], name })),
    selected: (names === null ? all : [...names]).map((name) => ids[name])
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

/** 只顶替"读不到真实宿主"的两件事；来源指纹与身份仍走产品自己的读取器。 */
vi.mock("./excel", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./excel")>()),
  isRunningInExcel: () => true,
  captureWorkbookCatalog: vi.fn(async () => ({
    name: WORKBOOK_NAME,
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

function json(payload: unknown, status: number): Response {
  return new Response(JSON.stringify(payload), {
    status,
    headers: { "Content-Type": "application/json" }
  });
}

/**
 * 回环端点。
 *
 * - `/result`：**新引用**返回工件计划里那份结果矩阵（表头 + 数据行，就是被写进
 *   目标区的那些值）；**旧引用**返回服务端真实给过的 404 固定码 —— 面板此时拿不到
 *   正文，只能重新查询。
 * - 轮询：给一次 `prepare_python_data` 指针（本机桥真读两张授权表），这样计划落地时
 *   的"来源依赖"有本端凭据，与真实面板一致。
 */
function stubBridge(resultBody: unknown) {
  const captured: Array<{ url: string; body: Record<string, unknown> }> = [];
  vi.stubGlobal(
    "fetch",
    vi.fn(async (url: string, init?: RequestInit) => {
      const target = String(url);
      const body = JSON.parse(String(init?.body ?? "{}")) as Record<string, unknown>;
      captured.push({ url: target, body });
      if (target.includes("/api/office/executions/register")) {
        return json(
          {
            operationId: String(body.operationId ?? ""),
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
      if (target.includes("/receipt")) return json({ status: "accepted", code: null }, 200);
      if (target.includes("/result")) {
        if (body.resultRef === OLD_REF) {
          // 服务端对过期引用给的就是这个：404 + 固定码，正文不存在。
          return json(
            {
              detail: {
                code: "BRIDGE_RESULT_UNAVAILABLE",
                message: "本地结果不存在或已过期，请重新查询。",
                retryable: false
              }
            },
            404
          );
        }
        return json(resultBody, 200);
      }
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
  // 动作条只有「刷新」与「确认」两个按钮（`styles.css` 的 `.scope-action-buttons`
  // 是两端同一份声明）。「确认」按钮的类是 `.scope-close`。
  const commit = view.container.querySelector<HTMLButtonElement>(".scope-close");
  if (!commit) throw new Error("范围「确认」按钮尚未出现");
  await waitFor(() => expect(commit.disabled).toBe(false));
  fireEvent.click(commit);
  await waitFor(() => expect(view.container.querySelector(".scope-popover")).toBeNull());
}

describe.skipIf(!A1_E2)("A1-E2 消费侧：恢复后的计划在真实 <App/> 上执行一次", () => {
  it("工件自洽：旧引用已失效、新引用可读、计划引用的是新引用", () => {
    expect(artifact.scenario).toBe("recovery");
    expect(REFS).toHaveLength(2);
    for (const ref of REFS) expect(ref).toMatch(/^res_[0-9a-f]{32}$/);
    expect(OLD_REF).not.toBe(NEW_REF);

    // 故障时序必须是 单次、绑到任务、且真的消费过 —— 这三条缺一条就说明"没在那一刻失效"。
    const phases = artifact.fault.timeline.map((item) => item.phase);
    expect(phases[0]).toBe("armed");
    expect(phases).toContain("bound");
    expect(phases[phases.length - 1]).toBe("consumed");
    expect(artifact.fault.timeline.find((item) => item.phase === "bound")?.taskBound).toBe(true);
    expect(artifact.fault.final.phase).toBe("consumed");

    // 旧引用取不回来、新引用可读，重复取旧引用仍然不可读（不因缓存/重试复活）。
    const old = artifact.probes.find((probe) => probe.resultRef === OLD_REF);
    const fresh = artifact.probes.find((probe) => probe.resultRef === NEW_REF);
    expect(old?.ok).toBe(false);
    expect(old?.status).toBe(404);
    expect(fresh?.ok).toBe(true);
    expect(artifact.repeatOldRef?.ok).toBe(false);

    // 计划是服务端用**新引用**编译的，不是旧引用。
    expect(PLAN.localResultRef).toBe(NEW_REF);
    expect(PLAN.actions).toHaveLength(artifact.planActionCount);
    expect(PLAN.actions[0].type).toBe("writeValues");
    expect(PLAN.actions[0].sheet).toBe(TARGET);

    // 恢复不是"重放读表"：过期后模型重算（新 py 步骤 + 新引用），桥只读过一次表。
    expect(artifact.bridge.prepareCount).toBe(1);
    expect(artifact.bridge.refusals).toHaveLength(0);
  });

  it("恢复后的计划 → 执行一次 → 逐格回读 → 确认不重执行 → 撤销回执行前", async () => {
    const action = PLAN.actions[0];
    // `/result` 返回的就是被写进目标区的那些值（表头 + 数据行）。
    const expected = PLAN.acceptanceCriteria.find((item) => item.type === "rangeEquals")
      ?.expected as (string | number | null)[][];
    const resultBody = {
      resultRef: NEW_REF,
      title: "Python 分析结果",
      headers: expected[0],
      rows: expected.slice(1),
      sourceSheets: [MEMBERS, ORDERS],
      warnings: [],
      complete: true,
      calculation: "隔离容器内执行的分析代码",
      scannedRows: expected.length
    };

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

    const gate: { release: (() => void) | null } = { release: null };
    vi.mocked(streamAssistantResponse).mockImplementation(async (_request, hooks) => {
      hooks?.onProgress?.({
        status: "started",
        stepId: "bridge.prepare_python_data",
        bridgeTool: { operationId: PREPARE_OP, tool: "prepare_python_data" }
      } as never);
      await new Promise<void>((done) => {
        gate.release = () => done();
      });
      return { kind: "plan", provider: "model", plan: PLAN } as never;
    });

    const captured = stubBridge(resultBody);
    const view = await sendPrompt("按会员等级汇总订单金额，结果写到汇总表", [
      MEMBERS,
      ORDERS,
      TARGET
    ]);

    // 桥真的读过两张授权表并回执（计划的来源依赖要靠这条凭据）。
    await waitFor(() => expect(captured.some((c) => c.url.includes("/receipt"))).toBe(true), {
      timeout: 4000
    });
    expect(host.matrix(MEMBERS, "A1:C6")).toEqual(MEMBER_ROWS);
    expect(host.matrix(ORDERS, "A1:B5")).toEqual(ORDER_ROWS);

    const beforeWrites = host.originalWrites().length;
    gate.release?.();
    await waitFor(() => expect(view.container.querySelector(".inline-plan")).not.toBeNull());

    const confirm = await waitFor(() => {
      const found = [...view.container.querySelectorAll("button")].find(
        (c) => c.textContent === "确认变更"
      ) as HTMLButtonElement | undefined;
      if (!found) throw new Error("计划到达后没有进入待确认状态");
      return found;
    }, { timeout: 15_000 });

    // 执行一次：写入次数 = 计划动作数（不是"大于 0"就放过）。
    expect(host.originalWrites()).toHaveLength(beforeWrites + PLAN.actions.length);

    // 逐格回读：空白只有一种写法（`""` ⇔ `null`），归一后仍逐格严格比较。
    const foldBlank = (cell: unknown) => (cell === "" ? null : cell);
    await waitFor(() => expect(host.cell(TARGET, "A1")).toBe(action.values[0][0]), {
      timeout: 4000
    });
    expect(host.matrix(TARGET, action.range).map((row) => row.map(foldBlank))).toEqual(
      action.values
    );
    // 只写目标表；两张来源表一格未变（撤销前也不许动）。
    expect(host.matrix(MEMBERS, "A1:C6")).toEqual(MEMBER_ROWS);
    expect(host.matrix(ORDERS, "A1:B5")).toEqual(ORDER_ROWS);

    // 确认只清标记，不第二次写入。
    const writesAfterExecution = host.originalWrites().length;
    fireEvent.click(confirm);
    await waitFor(() =>
      expect(
        [...view.container.querySelectorAll("button")].some((c) => c.textContent === "确认变更")
      ).toBe(false)
    );
    expect(host.originalWrites()).toHaveLength(writesAfterExecution);

    // 撤销：目标区逐格回到执行前（空表），来源表始终未变。
    const undoButton = [...view.container.querySelectorAll("button")].find(
      (c) => c.textContent === "↶ 撤销本次执行"
    ) as HTMLButtonElement | undefined;
    expect(undoButton).toBeDefined();
    expect(undoButton!.disabled).toBe(false);
    fireEvent.click(undoButton!);
    await waitFor(() => expect(view.container.textContent).toContain("已撤销上一次执行"));
    expect(host.matrix(TARGET, action.range)).toEqual(
      action.values.map((row) => row.map(() => ""))
    );
    expect(host.matrix(MEMBERS, "A1:C6")).toEqual(MEMBER_ROWS);
    expect(host.matrix(ORDERS, "A1:B5")).toEqual(ORDER_ROWS);
  });

  it("旧引用在面板侧也取不回正文（404），面板不会把它当成可用结果", async () => {
    // 直接走产品的读取出口：`loadBridgeResult` 是结果卡唯一的取数通道。
    const { loadBridgeResult } = await import("./sdkBridgeHost");
    stubBridge({
      resultRef: NEW_REF,
      title: "Python 分析结果",
      headers: ["列1"],
      rows: [["x"]],
      sourceSheets: [MEMBERS],
      warnings: [],
      complete: true,
      calculation: "c",
      scannedRows: 1
    });
    const binding = {
      taskId: artifact.turnId,
      hostId: "excel-office" as const,
      workbookName: WORKBOOK_NAME,
      workbookIdentity: `url:${BOOK_URL}`,
      sourceFingerprint: null
    };
    expect(await loadBridgeResult(binding, OLD_REF)).toBeNull();
    expect(await loadBridgeResult(binding, NEW_REF)).not.toBeNull();
  });

  /**
   * **A 路：模型没有再次读取宿主。**
   *
   * 时序全部由**屏障**决定，不用 sleep 抢：等桥的读表回执到齐（首次读取真的完成）
   * 之后才改来源单元格，改完再放行"计划到达"。所以"计划里的固定值来自旧来源"
   * 不是假设，是这段顺序的结果。
   *
   * 判据：自动执行必须**一次都不写**，并且界面要说清来源变了（不是静默不动）。
   * 旧结果已经拿到新 `resultRef`、也成功编译成计划了 —— 写前复核是最后一道闸。
   */
  it("来源在分析期间变化（模型没有再次读取）：自动执行写入 0 次", async () => {
    const action = PLAN.actions[0];
    const expected = PLAN.acceptanceCriteria.find((item) => item.type === "rangeEquals")
      ?.expected as (string | number | null)[][];
    const resultBody = {
      resultRef: NEW_REF,
      title: "Python 分析结果",
      headers: expected[0],
      rows: expected.slice(1),
      sourceSheets: [MEMBERS, ORDERS],
      warnings: [],
      complete: true,
      calculation: "隔离容器内执行的分析代码",
      scannedRows: expected.length
    };

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

    const gate: { release: (() => void) | null } = { release: null };
    vi.mocked(streamAssistantResponse).mockImplementation(async (_request, hooks) => {
      hooks?.onProgress?.({
        status: "started",
        stepId: "bridge.prepare_python_data",
        bridgeTool: { operationId: PREPARE_OP, tool: "prepare_python_data" }
      } as never);
      await new Promise<void>((done) => {
        gate.release = () => done();
      });
      return { kind: "plan", provider: "model", plan: PLAN } as never;
    });

    const captured = stubBridge(resultBody);
    const view = await sendPrompt("按会员等级汇总订单金额，结果写到汇总表", [
      MEMBERS,
      ORDERS,
      TARGET
    ]);
    await waitFor(() => expect(captured.some((c) => c.url.includes("/receipt"))).toBe(true), {
      timeout: 4000
    });
    const beforeWrites = host.originalWrites().length;

    // ── 屏障 ── 首次读取已完成（回执到齐）；模型仍按旧来源算结果；此刻外部改源。
    host.setCell(ORDERS, "B2", 250);
    expect(host.cell(ORDERS, "B2")).toBe(250);

    gate.release?.();
    await waitFor(() => expect(view.container.querySelector(".inline-plan")).not.toBeNull());

    await waitFor(
      () => {
        const text = view.container.textContent ?? "";
        if (!text.includes("数据来源在预览后发生了变化")) {
          throw new Error(
            `面板没有按"来源变化"拦下这份计划；新增写入 ${
              host.originalWrites().length - beforeWrites
            } 次`
          );
        }
      },
      { timeout: 12_000 }
    );

    // 一次都不许写：目标区保持执行前的空表，来源表只保留刚才那次外部修改。
    expect(host.originalWrites()).toHaveLength(beforeWrites);
    expect(
      host
        .matrix(TARGET, action.range)
        .every((row) => row.every((cell) => cell === "" || cell === null))
    ).toBe(true);
    expect(host.matrix(MEMBERS, "A1:C6")).toEqual(MEMBER_ROWS);
    expect(host.cell(ORDERS, "B2")).toBe(250);
  }, 30_000);
});

/**
 * **A1-E2 B 路（Excel 端）：模型**确实**再次读取来源，被生产桥拒交值。**
 *
 * A 路（上面那一节）证明的是"模型没再读时，只能靠写前复核拦住旧结果"。
 * B 路要证明的是另一件事：**模型再读一次**时，这次读取走的是生产桥的授权/来源复核
 * —— 判定来自 `packages/core/src/sdkToolBridge.ts` 的 `verifyBridgeAuthorization`
 * （驱动通过 `scripts/bridge_authorization_gate.mjs` 直接跑同一份 TS），
 * 回给模型的是固定码 `source_changed`，不是"再给你一份新值"。
 *
 * 工件由 `scripts/drive_real_sdk_result_expiry.py --scenario source-change-reread
 * --host excel-office` 在真实模型 + 真实服务端 + 真实 Docker 上跑出来：
 * 需求里写明"提交写入预览前再读取一次核对"，于是第二次 `prepare_python_data`
 * 是模型**自己**发起的真实工具调用；屏障保证改源发生在第一次交值之后。
 *
 * 这一轮**不启用**引用过期：旧引用仍然有效、计划仍可能被编译出来 —— 否则
 * "迟到的旧计划不得写入"就没有东西可验。
 */
type ReReadEvent = { stepId: string | null; status: string | null; tool?: string | null };
type ReReadArtifact = {
  scenario: string;
  turnId: string;
  promptLabel: string;
  fault: { armed: boolean };
  branch: string;
  bridge: {
    served: Array<{
      index: number;
      operationId: string;
      stepId: string;
      tool: string;
      refused: boolean;
      code: string | null;
      sources: Array<{ sheet: string; range: string }> | null;
      gate: { ok: boolean; errorCode?: string };
    }>;
    prepareCount: number;
    prepareAttempts: Array<{ index: number; operationId: string; stepId: string; served: boolean; code: string | null }>;
    refusals: Array<{ index: number; tool: string; code: string; operationId: string; stepId: string }>;
  };
  sourceChange: {
    afterServedCall: number;
    sheet: string;
    cell: string;
    before: number;
    after: number;
    fingerprintBefore: string;
    fingerprintAfter: string;
  };
  secondRead: {
    branch: string;
    attempted: boolean;
    refused: boolean;
    code: string | null;
    operationId: string | null;
    stepId: string | null;
    sameScope: boolean;
  };
  planSourceBasis: string;
  planActionCount: number;
  plan: {
    id: string;
    localResultRef: string;
    actions: Array<{ type: string; sheet: string; range: string; values: (string | number | null)[][] }>;
  } | null;
  kind: string | null;
  hostWritesDuringPlanning: number;
  events: ReReadEvent[];
};

/** 惰性读：缺工件时由"读"这一下失败，而不是在收集用例阶段把整套带崩。 */
function readReReadArtifact(): ReReadArtifact {
  const path = resolve(
    process.cwd(),
    "..",
    "..",
    // 默认消费 Excel 宿主那一轮；留了 `EB_A1_E2_REREAD_ARTIFACT` 只为**反向对照**：
    // 拿 A 路工件来跑这一组用例必须变红（否则这些断言就只是"看起来验过"）。
    process.env.EB_A1_E2_REREAD_ARTIFACT ??
      "outputs/a1-e2-source-change-reread-excel.json"
  );
  if (!existsSync(path)) {
    throw new Error(
      `EB_A1_E2=1 但没有 ${path}：` +
        "B 路这一轮没有产出工件，拒绝用 A 路工件或替身顶替。"
    );
  }
  return JSON.parse(readFileSync(path, "utf8")) as ReReadArtifact;
}

describe.skipIf(!A1_E2)("A1-E2 B 路：模型再次读取被 source_changed 拒绝", () => {
  it("工件自洽：第二次读取是真实工具调用、被生产复核拒绝、拒绝前后写入 0 次", () => {
    const a = readReReadArtifact();
    expect(a.scenario).toBe("source-change-reread");
    expect(a.promptLabel).toBe("two-read-verification");
    // 单独隔离"来源变化"这一个因素：这一轮不启用引用过期。
    expect(a.fault.armed).toBe(false);
    expect(a.branch).toBe("B");

    // 屏障：改的是哪个格子、改在首次**成功交值**之后（不是"计划到达之后"）。
    const change = a.sourceChange;
    expect(change.before).not.toBe(change.after);
    expect(change.fingerprintBefore).not.toBe(change.fingerprintAfter);
    const firstServed = a.bridge.served.find(
      (item) => item.tool === "prepare_python_data" && !item.refused
    );
    expect(firstServed).toBeDefined();
    expect(change.afterServedCall).toBe(firstServed!.index);

    // 真的读了第二次：独立操作身份 + 独立账本步骤 + 生产判据给的固定码。
    const second = a.secondRead;
    expect(second.attempted).toBe(true);
    expect(second.refused).toBe(true);
    expect(second.code).toBe("source_changed");
    // 两次读的是同一批来源 ⇒ 两次之间唯一的变量是"数据变了"。
    expect(second.sameScope).toBe(true);
    expect(second.operationId).toMatch(/^op_[0-9a-f]{32}$/);
    expect(second.operationId).not.toBe(firstServed!.operationId);
    expect(second.stepId).toBe(
      `bridge.prepare_python_data:${String(second.operationId).slice(3, 15)}`
    );

    // 那次调用的裁决来自共享核心（不是驱动自己拼的字符串）。
    const refusedCall = a.bridge.served.find((item) => item.operationId === second.operationId);
    expect(refusedCall?.refused).toBe(true);
    expect(refusedCall?.gate).toEqual({ ok: false, errorCode: "source_changed" });

    // 服务端账本记下了这次调用，并把它记成 warning（source_changed 是"没做成"）。
    const ledger = a.events.filter((event) => event.stepId === second.stepId);
    expect(ledger.length).toBeGreaterThan(0);
    expect(ledger.map((event) => event.status)).toContain("warning");
    // 正向锚点：第一次读取在账本里是 completed（否则"第二次"就没有参照）。
    const firstLedger = a.events.filter((event) => event.stepId === firstServed!.stepId);
    expect(firstLedger.map((event) => event.status)).toContain("completed");

    // 变更之后再也没有成功交值 ⇒ 计划（若有）只可能建立在旧那一版数据上。
    expect(a.planSourceBasis).toBe("old");
    // 规划期宿主写入必须是 0：拒绝发生在任何写入之前。
    expect(a.hostWritesDuringPlanning).toBe(0);
  });

  it("迟到的旧计划：真实 <App/> 上写前复核拒绝，写入 0 次，目标区保持空", async () => {
    const a = readReReadArtifact();
    /*
     * 这一轮模型有可能在第二次读取被拒后**仍然**按旧来源提交计划（B 路挖缺陷要看的
     * 正是这种情形）；也可能如实停在"来源已变化"不再提交。两种都记下来：
     * 没提交就断言"planActionCount = 0、kind 不是 plan"，绝不假装验过写入。
     */
    if (!a.plan) {
      expect(a.planActionCount).toBe(0);
      expect(a.kind).not.toBe("plan");
      return;
    }
    const plan = a.plan;
    const action = plan.actions[0];
    expect(action.type).toBe("writeValues");
    expect(action.sheet).toBe(TARGET);
    const expected = action.values;
    const resultBody = {
      resultRef: plan.localResultRef,
      title: "Python 分析结果",
      headers: expected[0],
      rows: expected.slice(1),
      sourceSheets: [MEMBERS, ORDERS],
      warnings: [],
      complete: true,
      calculation: "隔离容器内执行的分析代码",
      scannedRows: expected.length
    };

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

    const gate: { release: (() => void) | null } = { release: null };
    vi.mocked(streamAssistantResponse).mockImplementation(async (_request, hooks) => {
      // 首次受控读取：锚点就在这一步固定下来（面板拿的是**此刻**的现场）。
      hooks?.onProgress?.({
        status: "started",
        stepId: "bridge.prepare_python_data",
        bridgeTool: { operationId: PREPARE_OP, tool: "prepare_python_data" }
      } as never);
      await new Promise<void>((done) => {
        gate.release = () => done();
      });
      return { kind: "plan", provider: "model", plan } as never;
    });

    const captured = stubBridge(resultBody);
    const view = await sendPrompt("按会员等级汇总订单金额，结果写到汇总表", [
      MEMBERS,
      ORDERS,
      TARGET
    ]);
    await waitFor(() => expect(captured.some((c) => c.url.includes("/receipt"))).toBe(true), {
      timeout: 4000
    });
    const beforeWrites = host.originalWrites().length;

    // ── 屏障 ── 首次读取完成；此刻外部改源（与工件里同一格、同一个值）。
    host.setCell(ORDERS, a.sourceChange.cell, a.sourceChange.after);
    expect(host.cell(ORDERS, a.sourceChange.cell)).toBe(a.sourceChange.after);

    gate.release?.();
    await waitFor(() => expect(view.container.querySelector(".inline-plan")).not.toBeNull());

    await waitFor(
      () => {
        const text = view.container.textContent ?? "";
        if (!text.includes("数据来源在预览后发生了变化")) {
          throw new Error(
            `面板没有按"来源变化"拦下这份迟到的旧计划；新增写入 ${
              host.originalWrites().length - beforeWrites
            } 次`
          );
        }
      },
      { timeout: 12_000 }
    );

    expect(host.originalWrites()).toHaveLength(beforeWrites);
    expect(
      host
        .matrix(TARGET, action.range)
        .every((row) => row.every((cell) => cell === "" || cell === null))
    ).toBe(true);
    expect(host.matrix(MEMBERS, "A1:C6")).toEqual(MEMBER_ROWS);
    expect(host.cell(ORDERS, a.sourceChange.cell)).toBe(a.sourceChange.after);
  }, 30_000);
});
