/**
 * 反向变异驱动：「回到底部浮标」——**可见性判据 + 两种字形渲染 + 两端接线**的判别力证明。
 *
 * 被测机制分四层，各自有独立实现，所以必须**分拆各证各的**：
 *   1. 判据本身 `useMessageAutoFollowRuntime.tsx` 的 `awayFromBottom`（只看位置）；
 *   2. 共用组件 `ScrollToLatestButton.tsx` 的两态渲染与 `data-mode` 出口；
 *   3. Excel 宿主 `App.tsx` 的接线（可见性条件 + `running={busy}`）；
 *   4. WPS 宿主 `WpsAgentPanel.tsx` 的接线（同上）。
 *
 * 注意这一层**不是 CSS 守卫**（那条走 `run_progress_dot_contrast_mutations.mjs`）：
 * 这里证明的是"浮标到底什么时候出现、三点 / 锚有没有真的按运行状态渲染出来"。少了
 * 判据那一条，就会出现"往上滑什么都不出现"（用户实测报过）；少了字形接线，则会出现
 * "样式对了但界面永远显示同一种字形"。
 *
 * 用法：node scripts/run_scroll_hint_mutations.mjs
 */
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { closeSync, openSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const ADDIN = resolve(REPO, "apps", "excel-addin");
const WPS = resolve(REPO, "apps", "wps-addin");
const FOLLOW_HOOK = resolve(ADDIN, "src", "hooks", "useMessageAutoFollowRuntime.tsx");
const COMPONENT = resolve(ADDIN, "src", "components", "ScrollToLatestButton.tsx");
const EXCEL_APP = resolve(ADDIN, "src", "App.tsx");
const WPS_PANEL = resolve(WPS, "src", "WpsAgentPanel.tsx");
const SCRATCH = resolve(REPO, "outputs", "scroll-hint-mutation.log");
const VITEST = resolve(REPO, "node_modules", "vitest", "vitest.mjs");

const COMPONENT_TEST = ["run", "src/components/ScrollToLatestButton.test.tsx", "-t", "ScrollToLatestButton"];
const HOOK_TEST = ["run", "src/hooks/useMessageAutoFollow.test.ts"];

const CASES = [
  {
    id: "F1",
    why: "判据：滚上去不再当场置真（退回「等下一次 signal 变化」的旧口径）⇒ 往上翻历史时浮标根本不出现",
    file: FOLLOW_HOOK,
    seed: [
      "    // 滚出去就放出来、滚回来就收回 —— **不要求**期间有新内容。",
      "    setAwayFromBottom(!following);"
    ].join("\n"),
    mutant: "    // 旧口径：滚上去不动，等下一次 signal 变化再判。",
    cwd: ADDIN,
    args: HOOK_TEST,
    minTests: 7
  },
  {
    id: "C1",
    why: "组件：把 `{running ? …}` 写成恒真 ⇒ 跑完也还是三点，永远看不到锚",
    file: COMPONENT,
    seed: "      {running ? (",
    mutant: "      {true ? (",
    cwd: ADDIN,
    args: COMPONENT_TEST,
    minTests: 3
  },
  {
    id: "C2",
    why: "组件：`data-mode` 写死成 running ⇒ 状态出口在说谎（宿主测试读的就是它）",
    file: COMPONENT,
    seed: '      data-mode={running ? "running" : "idle"}',
    mutant: '      data-mode="running"',
    cwd: ADDIN,
    args: COMPONENT_TEST,
    minTests: 3
  },
  {
    id: "C3",
    why: "组件：运行分支不再用 `.progress-dots` ⇒ 运行中不再是三点",
    file: COMPONENT,
    seed: '        <span className="progress-dots" aria-hidden="true">',
    mutant: '        <span className="progress-anchor-icon" aria-hidden="true">',
    cwd: ADDIN,
    args: COMPONENT_TEST,
    minTests: 3
  },
  {
    id: "H1",
    why: "Excel 接线：running={false} ⇒ 跑着的时候也被画成「已停下」",
    file: EXCEL_APP,
    seed: "        <ScrollToLatestButton running={busy} onClick={scrollToLatest} />",
    mutant: "        <ScrollToLatestButton running={false} onClick={scrollToLatest} />",
    cwd: ADDIN,
    args: ["run", "src/messageFirstRouting.test.tsx", "-t", "swaps to the anchor"],
    minTests: 1
  },
  {
    id: "H2",
    why: "Excel 接线：running={true} ⇒ 跑完还是三点，永远看不到锚",
    file: EXCEL_APP,
    seed: "        <ScrollToLatestButton running={busy} onClick={scrollToLatest} />",
    mutant: "        <ScrollToLatestButton running={true} onClick={scrollToLatest} />",
    cwd: ADDIN,
    args: ["run", "src/messageFirstRouting.test.tsx", "-t", "idle anchor"],
    minTests: 2
  },
  {
    id: "H3",
    why: "Excel 可见性接线：条件写成恒真 ⇒ 回到最新之后浮标还赖在屏幕上",
    file: EXCEL_APP,
    seed: "      {awayFromBottom && (",
    mutant: "      {true && (",
    cwd: ADDIN,
    args: ["run", "src/messageFirstRouting.test.tsx", "-t", "idle anchor"],
    minTests: 2
  },
  {
    id: "W1",
    why: "WPS 接线：running={false} ⇒ 跑着的时候也被画成「已停下」",
    file: WPS_PANEL,
    seed: "        <ScrollToLatestButton running={busy} onClick={scrollToLatest} />",
    mutant: "        <ScrollToLatestButton running={false} onClick={scrollToLatest} />",
    cwd: WPS,
    args: ["run", "--config", "vite.config.ts", "src/WpsAgentPanel.test.tsx", "-t", "swaps to the anchor"],
    minTests: 1
  },
  {
    id: "W2",
    why: "WPS 接线：running={true} ⇒ 跑完还是三点，永远看不到锚",
    file: WPS_PANEL,
    seed: "        <ScrollToLatestButton running={busy} onClick={scrollToLatest} />",
    mutant: "        <ScrollToLatestButton running={true} onClick={scrollToLatest} />",
    cwd: WPS,
    args: ["run", "--config", "vite.config.ts", "src/WpsAgentPanel.test.tsx", "-t", "idle anchor"],
    minTests: 2
  },
  {
    id: "W3",
    why: "WPS 可见性接线：条件写成恒真 ⇒ 回到最新之后浮标还赖在屏幕上",
    file: WPS_PANEL,
    seed: "      {awayFromBottom && (",
    mutant: "      {true && (",
    cwd: WPS,
    args: ["run", "--config", "vite.config.ts", "src/WpsAgentPanel.test.tsx", "-t", "idle anchor"],
    minTests: 2
  }
];

const LF = "\n";

/** 目标文件自带换行（CRLF/LF）都要先归一化，否则 includes/split 会误报。 */
function normalize(text) {
  return text.split(/\r\n/).join(LF);
}
/** 写回必须还原**原文件自己的**换行，否则"逐字节一致"这条判据会自己失败。 */
function denormalize(text, original) {
  return text.split(LF).join(original.includes("\r\n") ? "\r\n" : LF);
}
function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

/** 用**文件描述符**接子进程输出：流水线里创建管道会被拦，而判据必须读真实摘要。 */
function runVitest(cwd, args) {
  const fd = openSync(SCRATCH, "w");
  const result = spawnSync(process.execPath, [VITEST, ...args], {
    cwd,
    stdio: ["ignore", fd, fd]
  });
  closeSync(fd);
  const raw = readFileSync(SCRATCH, "utf8").replace(/\x1b\[[0-9;]*m/g, "");
  const summary =
    raw.split(/\r?\n/).find((line) => /Tests\s+\d+ (passed|failed)/.test(line)) ?? "";
  const passed = /Tests\s+(\d+) passed/.exec(summary);
  return {
    status: result.status,
    summary: summary.trim(),
    passed: passed ? Number(passed[1]) : 0
  };
}

/**
 * apply / restore 都必须读**磁盘上的当前内容**，不能用内存里的基线去算：
 * 用基线算会让"apply 覆盖掉上一次没还原干净的变异"和"restore 找不到变异态"两件事
 * 同时静默发生 —— 前者看起来一切正常，后者把 7 个变异全部报成失败。
 */
function applyMutation(target, newlineStyle) {
  const text = normalize(readFileSync(target.file, "utf8"));
  const occurrences = text.split(target.seed).length - 1;
  if (occurrences !== 1) return `SEED 出现 ${occurrences} 次（期望 1）`;
  writeFileSync(
    target.file,
    denormalize(text.replace(target.seed, target.mutant), newlineStyle),
    "utf8"
  );
  return null;
}

function restore(target, newlineStyle) {
  const text = normalize(readFileSync(target.file, "utf8"));
  const occurrences = text.split(target.mutant).length - 1;
  if (occurrences !== 1) return `还原时变异态出现 ${occurrences} 次（期望 1）`;
  writeFileSync(
    target.file,
    denormalize(text.replace(target.mutant, target.seed), newlineStyle),
    "utf8"
  );
  return null;
}

// ---- 前置状态闸：文件若停在上一次被中断的变异态，当场报出来，不白跑几轮 vitest ----
let bad = 0;
for (const target of CASES) {
  const original = readFileSync(target.file, "utf8");
  const text = normalize(original);
  const seedCount = text.split(target.seed).length - 1;
  const mutantCount = text.split(target.mutant).length - 1;
  if (seedCount !== 1 || mutantCount !== 0) {
    console.error(
      `CHECK FAIL ${target.id} ${target.file} SEED=${seedCount}（期望 1）变异态=${mutantCount}（期望 0）`
    );
    bad += 1;
  }
}
if (bad > 0) {
  console.error("[scroll-hint] 前置状态闸未通过。");
  process.exit(1);
}
console.log("[scroll-hint] 前置状态闸通过：所有目标文件都在 SEED 态");

const baselines = new Map();
for (const target of CASES) {
  if (!baselines.has(target.file)) baselines.set(target.file, sha256(target.file));
}
for (const [path, sha] of baselines) console.log(`[scroll-hint] 基线 ${path} sha256=${sha}`);

let failures = 0;
// 换行风格只从**基线**取一次（变异过程不动它），之后 apply/restore 都用它写回。
const newlineStyles = new Map(
  [...baselines.keys()].map((path) => [path, readFileSync(path, "utf8").includes("\r\n") ? "\r\n" : LF])
);

for (const target of CASES) {
  const newlineStyle = newlineStyles.get(target.file);
  const applied = applyMutation(target, newlineStyle);
  if (applied) {
    console.error(`[scroll-hint] ${target.id} apply 失败：${applied}`);
    failures += 1;
    continue;
  }
  const run = runVitest(target.cwd, target.args);
  const verdict = run.status !== 0 ? "RED(OK)" : "GREEN(NG!)";
  console.log(`[scroll-hint] ${target.id} ${verdict} ${run.summary || "(无摘要)"}`);
  if (run.status === 0) {
    console.error(`  ${target.id} 拆掉机制后用例仍全绿 —— 断言没有判别力：${target.why}`);
    failures += 1;
  } else if (run.passed >= target.minTests) {
    // 红也要红在**预期规模**上：通过数没掉下去，说明红的不是这条机制。
    console.error(
      `  ${target.id} 红得可疑：仍通过 ${run.passed} 条，期望少于 ${target.minTests} 条`
    );
    failures += 1;
  }
  const restored = restore(target, newlineStyle);
  if (restored) {
    console.error(`[scroll-hint] ${target.id} restore 失败：${restored}`);
    failures += 1;
    continue;
  }
  if (sha256(target.file) !== baselines.get(target.file)) {
    console.error(`[scroll-hint] ${target.id} 还原后 ${target.file} 与基线不一致（sha256 变了）`);
    failures += 1;
  }
}

for (const [path, sha] of baselines) {
  const finalSha = sha256(path);
  if (finalSha !== sha) {
    console.error(`[scroll-hint] 最终 ${path} 与基线不一致：${finalSha} != ${sha}`);
    failures += 1;
  }
}
console.log(`[scroll-hint] 目标文件最终 sha256 与基线一致：${baselines.size} 个`);

console.log("[scroll-hint] 还原后复跑基线应为 GREEN…");
const baselineRuns = [
  { cwd: ADDIN, args: HOOK_TEST, minTests: 7 },
  { cwd: ADDIN, args: COMPONENT_TEST, minTests: 3 },
  { cwd: ADDIN, args: ["run", "src/messageFirstRouting.test.tsx"], minTests: 19 },
  { cwd: WPS, args: ["run", "--config", "vite.config.ts", "src/WpsAgentPanel.test.tsx"], minTests: 126 }
];
for (const item of baselineRuns) {
  const run = runVitest(item.cwd, item.args);
  if (run.status !== 0 || run.passed < item.minTests) {
    console.error(
      `[scroll-hint] 还原后复跑未通过：${item.cwd} ${item.args.join(" ")} → ${run.summary || "无摘要"}`
    );
    failures += 1;
  } else {
    console.log(`[scroll-hint] 复跑 GREEN：${run.summary}`);
  }
}
rmSync(SCRATCH, { force: true });

if (failures > 0) {
  console.error(`反向变异未通过：${failures} 处问题。`);
  process.exit(1);
}
console.log("反向变异通过：F1 / C1–C3 / H1–H3 / W1–W3 各自变红，还原逐字节一致，复跑全绿。");
