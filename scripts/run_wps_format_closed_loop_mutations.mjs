/**
 * 反向变异驱动：证明 Task #53 第二轮「WPS 格式动作面板闭环」三个阻断点修复的
 * **新增守卫**真的有判别力（细节见 `scripts/mutate_wps_format_closed_loop.mjs`）。
 *
 * 被测机制分三处，测试文件各打各自的判别点：
 *   - 阻断点 1（副本审阅格式分派）  ← apps/wps-addin/src/wpsFormatFlow.test.ts（阻断点 1 用例）
 *   - 阻断点 2（属性级恢复 / 同区域多格式动作） ← wpsFormatFlow.test.ts（阻断点 2 + 多动作组合用例）
 *   - 阻断点 3（多格缺 Item 拒绝）  ← wpsFormatFlow.test.ts（阻断点 3 用例）
 *
 * 四条判据缺一不算过：
 * 1. 未变异时基线 GREEN（退出码 0 + 有效测试摘要 + 用例数达标）；
 * 2. 每个变异 apply 后其目标用例必须变红（退出码非 0 且**有测试摘要**）；
 * 3. restore 后**被改的那个文件**逐字节一致（sha256 比对）；
 * 4. 全部还原后复跑全绿，且三个文件最终都与基线一致。
 *
 * 起手另有前置状态闸（`check`）：文件停在上一次被中断的变异态时当场退出。
 *
 * 用法：node scripts/run_wps_format_closed_loop_mutations.mjs
 */
import { spawnSync } from "node:child_process";
import { openSync, closeSync, readFileSync, rmSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");
const wpsDir = resolve(repoRoot, "apps", "wps-addin");
const scratch = resolve(repoRoot, "outputs", "wps-format-closed-loop-mutation.log");
const vitest = resolve(repoRoot, "node_modules", "vitest", "vitest.mjs");

const FLOW_SPEC = {
  name: "vitest:wps/wpsFormatFlow",
  minTests: 10,
  command: process.execPath,
  args: [vitest, "run", "--config", "vite.config.ts", "src/wpsFormatFlow.test.ts"],
  cwd: wpsDir
};
const PANEL_SPEC = {
  name: "vitest:wps/WpsAgentPanelFormatReview",
  minTests: 2,
  command: process.execPath,
  args: [vitest, "run", "--config", "vite.config.ts", "src/WpsAgentPanelFormatReview.test.tsx"],
  cwd: wpsDir
};
const BOUNDARY_SPEC = {
  name: "vitest:wps/wpsFormatCombinationBoundary",
  minTests: 11,
  command: process.execPath,
  args: [vitest, "run", "--config", "vite.config.ts", "src/wpsFormatCombinationBoundary.test.ts"],
  cwd: wpsDir
};

/**
 * 每个变异要跑哪些用例：
 * - 模块集成证据（`wpsFormatFlow.test.ts`）覆盖六个变异；
 * - **面板点击证据**（`WpsAgentPanelFormatReview.test.tsx`，真实审阅 + 真实执行）
 *   额外盖 M1（副本格式应用）与 M5（最终状态），证明面板链路真的落到这些机制上；
 * - **组合边界证据**（`wpsFormatCombinationBoundary.test.ts`，行级行高宿主）
 *   盖 M7（冲突检测整行语义）与 M8（恢复范围含行高）。
 */
const MUTATIONS = [
  { id: "M1", specs: [FLOW_SPEC, PANEL_SPEC] },
  { id: "M2", specs: [FLOW_SPEC] },
  { id: "M3", specs: [FLOW_SPEC] },
  { id: "M4", specs: [FLOW_SPEC] },
  { id: "M5", specs: [FLOW_SPEC, PANEL_SPEC] },
  { id: "M6", specs: [FLOW_SPEC] },
  { id: "M7", specs: [BOUNDARY_SPEC] },
  { id: "M8", specs: [BOUNDARY_SPEC] }
];

const WATCHED = [
  resolve(wpsDir, "src", "wpsNativeComparison.ts"),
  resolve(wpsDir, "src", "wpsPlanExecutor.ts"),
  resolve(wpsDir, "src", "wpsFormatSnapshot.ts")
];

/**
 * 输出走**文件描述符**而不是管道：管道会被拦，而判据又必须读到真实摘要行。
 * vitest 的摘要有 ANSI 色码，解析前先剥掉。
 */
function runSpec(spec) {
  const fd = openSync(scratch, "w");
  const result = spawnSync(spec.command, spec.args, {
    cwd: spec.cwd,
    stdio: ["ignore", fd, fd]
  });
  closeSync(fd);
  const raw = readFileSync(scratch, "utf8").replace(/\x1b\[[0-9;]*m/g, "");
  const summaryLine =
    raw.split(/\r?\n/).find((line) => /Tests\s+\d+ (passed|failed)/.test(line)) ??
    raw.split(/\r?\n/).find((line) => /^\d+ (passed|failed)/.test(line)) ??
    raw
      .split(/\r?\n/)
      .reverse()
      .find((line) => /\b\d+ passed\b|\b\d+ failed\b/.test(line)) ??
    "";
  const passed = Number(/(\d+) passed/.exec(summaryLine)?.[1] ?? 0);
  const failed = Number(/(\d+) failed/.exec(summaryLine)?.[1] ?? 0);
  return {
    status: result.status,
    summary: summaryLine.trim(),
    passed,
    failed,
    error: result.error ? String(result.error.code ?? result.error) : null
  };
}

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

function invokeMutation(id, restore) {
  const args = [resolve(here, "mutate_wps_format_closed_loop.mjs"), id, ...(restore ? ["restore"] : [])];
  const fd = openSync(scratch, "w");
  const result = spawnSync(process.execPath, args, {
    cwd: repoRoot,
    stdio: ["ignore", fd, fd]
  });
  closeSync(fd);
  return {
    status: result.status,
    out: readFileSync(scratch, "utf8").trim(),
    error: result.error ? String(result.error.code) : null
  };
}

const baselineShas = new Map(WATCHED.map((path) => [path, sha256(path)]));
for (const [path, sha] of baselineShas) {
  console.log(`[mutate] 基线 ${path} sha256=${sha}`);
}

// 前置状态闸：停在上一次被中断的变异态时当场退出，不白跑四轮测试。
const preflight = invokeMutation("check");
if (preflight.status !== 0) {
  console.error(`[mutate] 前置状态闸未通过：\n${preflight.out}`);
  process.exit(1);
}
console.log(`[mutate] 前置状态闸通过：${Object.keys(MUTATIONS).length} 个 SEED 全部在位`);
rmSync(scratch, { force: true });

const BASELINE_SPECS = [FLOW_SPEC, PANEL_SPEC, BOUNDARY_SPEC];

console.log("[mutate] 基线（未变异）应为 GREEN…");
for (const spec of BASELINE_SPECS) {
  const result = runSpec(spec);
  const ok = result.status === 0 && result.summary && result.passed >= spec.minTests;
  console.log(
    `[mutate]   基线 ${spec.name}: status=${result.status} 摘要="${result.summary}" ${ok ? "OK" : "不合格"}`
  );
  if (!ok) {
    console.error(`[mutate] 基线就不绿或没跑到预期用例数（要求 >= ${spec.minTests}），中止。`);
    process.exit(1);
  }
}

const failures = [];
for (const { id, specs } of MUTATIONS) {
  const applied = invokeMutation(id, false);
  if (applied.status !== 0) {
    failures.push(`${id} apply 失败：${applied.out || applied.error}`);
    console.error(`[mutate] ${id} apply 失败：${applied.out || applied.error}`);
    continue;
  }
  for (const spec of specs) {
    const red = runSpec(spec);
    // 判据 2 要求"变红"：退出码非 0 **且**有有效测试摘要 —— 编译/收集期直接崩掉
    // 不算判别力，那只能证明我们把代码写坏了。
    const wentRed = red.status !== 0 && Boolean(red.summary);
    console.log(
      `[mutate] ${id} apply → ${applied.out.split(" sha256=")[0]} ; ${spec.name}: status=${
        red.status
      } 摘要="${red.summary}" ${wentRed ? "RED(合格)" : "未变红(不合格)"}`
    );
    if (!wentRed) {
      failures.push(`${id} 未让 ${spec.name} 变红（status=${red.status}, 摘要=${red.summary || "无"}）`);
    }
  }
  const restored = invokeMutation(id, true);
  if (restored.status !== 0) {
    failures.push(`${id} restore 失败：${restored.out || restored.error}`);
    console.error(`[mutate] ${id} restore 失败：${restored.out || restored.error}`);
  }
}

// 判据 3：每个被改过的文件都必须逐字节回到基线。
for (const [path, sha] of baselineShas) {
  const now = sha256(path);
  if (now !== sha) {
    failures.push(`${path} 还原后与基线不一致：${now} != ${sha}`);
    console.error(`[mutate] 还原不一致 ${path}`);
  }
}
if (!failures.length) console.log("[mutate] 三个文件全部逐字节还原");

console.log("[mutate] 还原后复跑全绿…");
for (const spec of BASELINE_SPECS) {
  const result = runSpec(spec);
  const ok = result.status === 0 && result.summary && result.passed >= spec.minTests;
  console.log(`[mutate]   复跑 ${spec.name}: status=${result.status} 摘要="${result.summary}" ${ok ? "OK" : "不合格"}`);
  if (!ok) failures.push(`还原后复跑 ${spec.name} 不为绿（status=${result.status}）`);
}

rmSync(scratch, { force: true });

if (failures.length) {
  console.error("\n[mutate] 未通过：");
  for (const line of failures) console.error(`  - ${line}`);
  process.exit(1);
}
console.log(
  `\n[mutate] 全部通过：${MUTATIONS.length} 个变异各自让目标用例变红，三个文件逐字节还原，复跑全绿。`
);
