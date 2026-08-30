import fs from "node:fs";
import path from "node:path";
import { performance } from "node:perf_hooks";
import { JSDOM } from "jsdom";
import type { ProfilerOnRenderCallback } from "react";
import type { Root } from "react-dom/client";
import type { SavedTool } from "../src/storage";
import type { AnalysisPlan } from "../src/contracts";
import type { ChatHistoryState, ChatMessage } from "../src/types/chat";

const SAMPLE_COUNT = 12;
const OUTPUT_DIR = path.resolve(
  process.cwd(),
  "docs/performance/render-profile"
);

interface CommitSample {
  phase: "mount" | "update" | "nested-update";
  actualDuration: number;
  baseDuration: number;
  startTime: number;
  commitTime: number;
}

interface RunSample {
  wallMs: number;
  commitCount: number;
  totalActualMs: number;
  maxCommitMs: number;
  totalBaseMs: number;
}

interface ScenarioResult {
  id: string;
  label: string;
  scale: string;
  runs: RunSample[];
  median: RunSample;
  p95: RunSample;
}

const dom = new JSDOM(
  "<!doctype html><html><body><div id=\"profile-root\"></div></body></html>",
  { url: "https://localhost:3000/", pretendToBeVisual: true }
);

const exposedGlobals = {
  window: dom.window,
  document: dom.window.document,
  navigator: dom.window.navigator,
  localStorage: dom.window.localStorage,
  HTMLElement: dom.window.HTMLElement,
  HTMLTextAreaElement: dom.window.HTMLTextAreaElement,
  Element: dom.window.Element,
  Node: dom.window.Node,
  Event: dom.window.Event,
  MouseEvent: dom.window.MouseEvent,
  KeyboardEvent: dom.window.KeyboardEvent,
  PointerEvent: dom.window.PointerEvent ?? dom.window.MouseEvent,
  File: dom.window.File,
  Blob: dom.window.Blob,
  DOMParser: dom.window.DOMParser,
  MutationObserver: dom.window.MutationObserver,
  getComputedStyle: dom.window.getComputedStyle.bind(dom.window)
};

for (const [name, value] of Object.entries(exposedGlobals)) {
  Object.defineProperty(globalThis, name, {
    configurable: true,
    writable: true,
    value
  });
}

Object.defineProperty(dom.window.Element.prototype, "scrollIntoView", {
  configurable: true,
  value: () => undefined
});
Object.defineProperty(dom.window, "matchMedia", {
  configurable: true,
  value: () => ({
    matches: false,
    addEventListener: () => undefined,
    removeEventListener: () => undefined
  })
});
Object.defineProperty(dom.window, "close", {
  configurable: true,
  value: () => undefined
});

// Keep network-dependent work pending so the measured commits only contain the
// user-triggered local state transition, not variable backend response time.
globalThis.fetch = (() => new Promise<Response>(() => undefined)) as typeof fetch;

// React DOM must initialize after the synthetic DOM exists; otherwise its event
// feature detection runs against Node globals and controlled input events do not fire.
const React = await import("react");
const { Profiler } = React;
const { createRoot } = await import("react-dom/client");
const { act, fireEvent } = await import("@testing-library/react");
const { default: App } = await import("../src/App");
const { default: PlanEditorDialog } = await import(
  "../src/components/PlanEditorDialog"
);
const { createTool } = await import("../src/storage");
const { demoWorkbook } = await import("../src/demo");

const plan: AnalysisPlan = {
  id: "profile-plan",
  title: "性能采样计划",
  summary: "用于渲染采样的合法白名单计划",
  assumptions: [],
  warnings: [],
  actions: [
    {
      type: "autofit",
      sheet: demoWorkbook.activeWorksheet,
      range: "A:D"
    }
  ]
};

function makeMessages(count: number): ChatMessage[] {
  return Array.from({ length: count }, (_, index) => ({
    id: `profile-message-${index}`,
    role: index % 2 === 0 ? "user" : "assistant",
    text:
      index % 2 === 0
        ? `第 ${index + 1} 条用户问题：请分析当前工作表并给出简洁结论。`
        : `第 ${index + 1} 条分析结果：这里是用于模拟真实对话长度的回复内容。`,
    provider: index % 2 === 0 ? undefined : "local",
    createdAt: new Date(2026, 7, 30, 0, index).toISOString()
  }));
}

function makeChatHistory(messageCount: number): ChatHistoryState {
  const messages = makeMessages(messageCount);
  const timestamp = new Date(2026, 7, 30).toISOString();
  return {
    activeConversationId: "profile-conversation",
    conversations: [
      {
        id: "profile-conversation",
        title: "性能采样对话",
        messages,
        createdAt: timestamp,
        updatedAt: timestamp
      }
    ]
  };
}

function makeTools(count: number): SavedTool[] {
  return Array.from({ length: count }, (_, index) => ({
    ...createTool(
      { ...plan, id: `profile-plan-${index}`, title: `性能工具 ${index + 1}` },
      `性能工具 ${index + 1}`,
      "重复执行当前工作表的列宽调整，用于工具箱上限规模渲染采样。",
      [demoWorkbook.activeWorksheet]
    ),
    id: `profile-tool-${index}`
  }));
}

const tools = makeTools(50);

function seedStorage(messageCount = 100): void {
  localStorage.clear();
  localStorage.setItem(
    "excel-bro.chat.v4",
    JSON.stringify(makeChatHistory(messageCount))
  );
  localStorage.setItem("excel-bro.tools.v2", JSON.stringify(tools));
  localStorage.setItem("excel-bro.query-tools.v1", "[]");
  localStorage.setItem("excel-bro.pet.visibility.v1", "hidden");
  localStorage.setItem("excel-bro.model-guide-dismissed.v1", "1");
}

function profileCallback(commits: CommitSample[]): ProfilerOnRenderCallback {
  return (_id, phase, actualDuration, baseDuration, startTime, commitTime) => {
    commits.push({
      phase,
      actualDuration,
      baseDuration,
      startTime,
      commitTime
    });
  };
}

async function flush(): Promise<void> {
  await act(async () => {
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
  });
}

async function mountApp(): Promise<{
  root: Root;
  container: HTMLElement;
  commits: CommitSample[];
}> {
  seedStorage();
  Reflect.deleteProperty(globalThis, "Office");
  const container = document.getElementById("profile-root")!;
  container.replaceChildren();
  const root = createRoot(container);
  const commits: CommitSample[] = [];
  await act(async () => {
    root.render(
      <Profiler id="App" onRender={profileCallback(commits)}>
        <App />
      </Profiler>
    );
  });
  await flush();
  return { root, container, commits };
}

async function dispose(root: Root): Promise<void> {
  await act(async () => root.unmount());
}

function summarize(wallMs: number, commits: CommitSample[]): RunSample {
  return {
    wallMs,
    commitCount: commits.length,
    totalActualMs: commits.reduce((sum, item) => sum + item.actualDuration, 0),
    maxCommitMs: Math.max(0, ...commits.map((item) => item.actualDuration)),
    totalBaseMs: commits.reduce((sum, item) => sum + item.baseDuration, 0)
  };
}

function percentile(runs: RunSample[], ratio: number): RunSample {
  const metric = (key: keyof RunSample) => {
    const values = runs.map((run) => run[key]).sort((a, b) => a - b);
    return values[Math.min(values.length - 1, Math.ceil(values.length * ratio) - 1)];
  };
  return {
    wallMs: metric("wallMs"),
    commitCount: metric("commitCount"),
    totalActualMs: metric("totalActualMs"),
    maxCommitMs: metric("maxCommitMs"),
    totalBaseMs: metric("totalBaseMs")
  };
}

async function collectScenario(
  id: string,
  label: string,
  scale: string,
  run: () => Promise<RunSample>
): Promise<ScenarioResult> {
  await run(); // warm-up: module/JIT setup is not user-perceived steady-state cost.
  const runs: RunSample[] = [];
  for (let index = 0; index < SAMPLE_COUNT; index += 1) {
    runs.push(await run());
  }
  return {
    id,
    label,
    scale,
    runs,
    median: percentile(runs, 0.5),
    p95: percentile(runs, 0.95)
  };
}

async function initialRender(): Promise<RunSample> {
  const started = performance.now();
  const mounted = await mountApp();
  const result = summarize(performance.now() - started, mounted.commits);
  await dispose(mounted.root);
  return result;
}

async function sendMessage(): Promise<RunSample> {
  const mounted = await mountApp();
  const textarea = mounted.container.querySelector<HTMLTextAreaElement>(
    'textarea[aria-label="给 Excel Bro 发消息"]'
  )!;
  await act(async () => {
    fireEvent.input(textarea, { target: { value: "分析当前工作表" } });
  });
  mounted.commits.length = 0;
  const send = mounted.container.querySelector<HTMLButtonElement>(
    'button[aria-label="发送"]'
  )!;
  if (send.disabled) {
    throw new Error(`发送按钮仍被禁用，输入值为「${textarea.value}」`);
  }
  const started = performance.now();
  await act(async () => fireEvent.click(send));
  const result = summarize(performance.now() - started, mounted.commits);
  await dispose(mounted.root);
  return result;
}

async function typingBurst(): Promise<RunSample> {
  const mounted = await mountApp();
  const textarea = mounted.container.querySelector<HTMLTextAreaElement>(
    'textarea[aria-label="给 Excel Bro 发消息"]'
  )!;
  mounted.commits.length = 0;
  const text = "分析当前工作表";
  const started = performance.now();
  for (let index = 1; index <= text.length; index += 1) {
    await act(async () => {
      fireEvent.input(textarea, { target: { value: text.slice(0, index) } });
    });
  }
  const result = summarize(performance.now() - started, mounted.commits);
  await dispose(mounted.root);
  return result;
}

async function switchWorksheet(): Promise<RunSample> {
  const mounted = await mountApp();
  const trigger = mounted.container.querySelector<HTMLButtonElement>(
    ".scope-trigger"
  )!;
  await act(async () => fireEvent.click(trigger));
  const manual = [...mounted.container.querySelectorAll("button")].find(
    (button) => button.textContent?.includes("选择多个")
  )!;
  await act(async () => fireEvent.click(manual));
  mounted.commits.length = 0;
  const sheet = mounted.container.querySelector<HTMLButtonElement>(
    ".workbook-sheet-list button"
  )!;
  const started = performance.now();
  await act(async () => fireEvent.click(sheet));
  const result = summarize(performance.now() - started, mounted.commits);
  await dispose(mounted.root);
  return result;
}

async function openTools(): Promise<RunSample> {
  const mounted = await mountApp();
  mounted.commits.length = 0;
  const button = mounted.container.querySelector<HTMLButtonElement>(
    'button[aria-label="我的工具"]'
  )!;
  const started = performance.now();
  await act(async () => fireEvent.click(button));
  const result = summarize(performance.now() - started, mounted.commits);
  await dispose(mounted.root);
  return result;
}

async function planEditor(): Promise<RunSample> {
  seedStorage();
  const largePlan: AnalysisPlan = {
    ...plan,
    actions: Array.from({ length: 50 }, (_, index) => ({
      type: "autofit" as const,
      sheet: demoWorkbook.activeWorksheet,
      range: `A${index + 1}:D${index + 1}`
    }))
  };
  localStorage.setItem(
    "EB_PLAN_EDITOR_PAYLOAD",
    JSON.stringify({
      type: "plan-editor-init",
      payload: {
        plan: largePlan,
        target: { kind: "message", messageId: "profile-message" },
        workbook: demoWorkbook
      }
    })
  );
  Object.assign(globalThis, {
    Office: {
      onReady: (callback: () => void) => callback(),
      context: { ui: {} },
      EventType: {
        DialogParentMessageReceived: "DialogParentMessageReceived"
      }
    }
  });
  const container = document.getElementById("profile-root")!;
  container.replaceChildren();
  const root = createRoot(container);
  const commits: CommitSample[] = [];
  const started = performance.now();
  await act(async () => {
    root.render(
      <Profiler id="PlanEditorDialog" onRender={profileCallback(commits)}>
        <PlanEditorDialog />
      </Profiler>
    );
  });
  await flush();
  const result = summarize(performance.now() - started, commits);
  await dispose(root);
  return result;
}

function escapeXml(value: string): string {
  return value.replace(/[&<>\"]/g, (character) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '\"': "&quot;"
  })[character]!);
}

function renderChart(result: ScenarioResult): string {
  const width = 900;
  const height = 430;
  const plotLeft = 72;
  const plotTop = 92;
  const plotWidth = 780;
  const plotHeight = 250;
  const maxValue = Math.max(
    20,
    ...result.runs.map((run) => run.maxCommitMs),
    result.p95.maxCommitMs
  );
  const xStep = plotWidth / Math.max(1, result.runs.length - 1);
  const y = (value: number) =>
    plotTop + plotHeight - (Math.min(value, maxValue) / maxValue) * plotHeight;
  const points = result.runs
    .map((run, index) => `${plotLeft + index * xStep},${y(run.maxCommitMs)}`)
    .join(" ");
  const thresholdY = y(16);
  return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
  <rect width="100%" height="100%" fill="#0f172a"/>
  <text x="48" y="42" fill="#f8fafc" font-size="24" font-family="Segoe UI, sans-serif" font-weight="700">${escapeXml(result.label)}</text>
  <text x="48" y="68" fill="#94a3b8" font-size="14" font-family="Segoe UI, sans-serif">React Profiler actualDuration · ${escapeXml(result.scale)} · ${SAMPLE_COUNT} 次</text>
  <line x1="${plotLeft}" y1="${thresholdY}" x2="${plotLeft + plotWidth}" y2="${thresholdY}" stroke="#f59e0b" stroke-width="2" stroke-dasharray="7 6"/>
  <text x="${plotLeft + plotWidth - 90}" y="${thresholdY - 8}" fill="#fbbf24" font-size="13" font-family="Segoe UI, sans-serif">16ms 阈值</text>
  <line x1="${plotLeft}" y1="${plotTop + plotHeight}" x2="${plotLeft + plotWidth}" y2="${plotTop + plotHeight}" stroke="#475569"/>
  <line x1="${plotLeft}" y1="${plotTop}" x2="${plotLeft}" y2="${plotTop + plotHeight}" stroke="#475569"/>
  <polyline points="${points}" fill="none" stroke="#38bdf8" stroke-width="3"/>
  ${result.runs.map((run, index) => `<circle cx="${plotLeft + index * xStep}" cy="${y(run.maxCommitMs)}" r="4" fill="#7dd3fc"/>`).join("\n  ")}
  <text x="48" y="382" fill="#e2e8f0" font-size="15" font-family="Segoe UI, sans-serif">中位最大 commit：${result.median.maxCommitMs.toFixed(2)}ms</text>
  <text x="310" y="382" fill="#e2e8f0" font-size="15" font-family="Segoe UI, sans-serif">P95 最大 commit：${result.p95.maxCommitMs.toFixed(2)}ms</text>
  <text x="560" y="382" fill="#e2e8f0" font-size="15" font-family="Segoe UI, sans-serif">P95 交互墙钟：${result.p95.wallMs.toFixed(2)}ms</text>
  <text x="48" y="410" fill="#64748b" font-size="12" font-family="Segoe UI, sans-serif">CPU/DOM 渲染采样；不包含网络、Excel COM 与真实浏览器布局绘制。</text>
</svg>`;
}

const scenarios = [
  await collectScenario(
    "initial-render",
    "打开插件：初始渲染",
    "100 条消息 + 50 个已保存工具（工具箱关闭）",
    initialRender
  ),
  await collectScenario(
    "typing-burst",
    "连续输入：重渲染频率诊断",
    "100 条历史消息，连续输入 7 个字符",
    typingBurst
  ),
  await collectScenario(
    "send-message",
    "发送消息：对话列表更新",
    "100 条历史消息，新增第 101 条后按上限裁剪",
    sendMessage
  ),
  await collectScenario(
    "switch-worksheet",
    "切换工作表：范围选择器更新",
    "100 条消息 + 演示工作簿 2 张表",
    switchWorksheet
  ),
  await collectScenario(
    "open-tools",
    "打开“我的工具”",
    "50 个工具（配置上限）",
    openTools
  ),
  await collectScenario(
    "open-plan-editor",
    "打开计划编辑器",
    "50 个白名单动作的格式化 JSON",
    planEditor
  )
];

fs.mkdirSync(OUTPUT_DIR, { recursive: true });
fs.writeFileSync(
  path.join(OUTPUT_DIR, "results.json"),
  `${JSON.stringify(
    {
      generatedAt: new Date().toISOString(),
      runtime: {
        node: process.version,
        react: React.version,
        samplesPerScenario: SAMPLE_COUNT,
        thresholdMs: 16,
        harness: "React Profiler + jsdom; source production component tree"
      },
      scenarios
    },
    null,
    2
  )}\n`,
  "utf8"
);
for (const scenario of scenarios) {
  fs.writeFileSync(
    path.join(OUTPUT_DIR, `${scenario.id}.svg`),
    renderChart(scenario),
    "utf8"
  );
}

process.stdout.write(
  `${JSON.stringify(
    scenarios.map(({ id, median, p95 }) => ({ id, median, p95 })),
    null,
    2
  )}\n`
);
