/**
 * 反向变异驱动：证明「WPS 任务窗格顶部那条开发诊断导航条已整条移除」这条判据**真有判别力**。
 *
 * 四条判据缺一不算过：
 * 1. 未变异时基线 GREEN（退出码 0 + 有效测试摘要 + 用例数达标）；
 * 2. 每个变异 apply 后用例必须变红（退出码非 0）；
 * 3. restore 后文件**逐字节一致**（sha256 比对）；
 * 4. 还原后复跑全绿。
 *
 * 两个变异钉两处**互不相同**的属性：
 * M1 钉「界面上还有没有那三颗按钮 / 那条 nav」（空类名选择器不够，得连 nav 与按钮数一起判）；
 * M2 钉「诊断路由的 DEV 门控还在不在」（非 DEV 手改 hash 不能打开诊断页）。
 *
 * 用法：node scripts/run_wps_dev_nav_removed_mutations.mjs
 */
import { spawnSync } from "node:child_process";
import { openSync, closeSync, readFileSync, mkdirSync, rmSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");
const appDir = resolve(repoRoot, "apps", "wps-addin");
const scratchDir = resolve(repoRoot, "outputs");
const scratch = resolve(scratchDir, "wps-dev-nav-mutation.log");
const TARGET = resolve(appDir, "src", "AppRoutes.tsx");
const TEST_FILES = ["src/main.test.tsx", "src/main.production.test.tsx"];
const MUTATIONS = ["M1", "M2"];
// 两个文件基线共 6 条（main.test 3 + main.production 3）。
const MIN_TESTS = 6;

const vitest = resolve(repoRoot, "node_modules", "vitest", "vitest.mjs");

mkdirSync(scratchDir, { recursive: true });

/**
 * 用**文件描述符**而不是 pipe 接子进程输出：流水线里创建管道会被拦，
 * 而判据又必须读到真实的摘要行。
 */
function runVitest() {
  const fd = openSync(scratch, "w");
  const result = spawnSync(process.execPath, [vitest, "run", ...TEST_FILES], {
    cwd: appDir,
    stdio: ["ignore", fd, fd]
  });
  closeSync(fd);
  const raw = readFileSync(scratch, "utf8").replace(/\x1b\[[0-9;]*m/g, "");
  const summary =
    raw.split(/\r?\n/).find((line) => /Tests\s+\d+ (passed|failed)/.test(line)) ?? "";
  const counts = /Tests\s+(\d+) passed/.exec(summary);
  return {
    status: result.status,
    summary: summary.trim(),
    passed: counts ? Number(counts[1]) : 0
  };
}

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

function mutate(which, restore) {
  const args = [
    resolve(here, "mutate_wps_dev_nav_removed.mjs"),
    which,
    ...(restore ? ["restore"] : [])
  ];
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

const baselineSha = sha256(TARGET);
console.log(`[mutate] 基线 ${TARGET} sha256=${baselineSha}`);

console.log("[mutate] 基线（未变异）应为 GREEN…");
const baseline = runVitest();
if (baseline.status !== 0 || !baseline.summary || baseline.passed < MIN_TESTS) {
  console.error(
    `基线就不绿（status=${baseline.status}, 摘要=${baseline.summary || "无"}）`
  );
  process.exit(1);
}
console.log(`[mutate] 基线 GREEN：${baseline.summary}`);

let failures = 0;
for (const which of MUTATIONS) {
  const applied = mutate(which, false);
  if (applied.status !== 0 || applied.error) {
    console.error(
      `[mutate] ${which} apply 失败：status=${applied.status} err=${applied.error} out=${applied.out}`
    );
    failures += 1;
    continue;
  }
  const run = runVitest();
  const verdict = run.status !== 0 ? "RED(OK)" : "GREEN(NG!)";
  console.log(`[mutate] ${which} → ${verdict} ${run.summary || "(无摘要)"}`);
  if (run.status === 0) {
    console.error(
      `  ${which} 把导航条挂回来后用例仍然全绿 —— 断言没有判别力，见：${scratch}`
    );
    failures += 1;
  }
  const restored = mutate(which, true);
  if (restored.status !== 0 || restored.out === "ALREADY-CLEAN") {
    console.error(`[mutate] ${which} restore 失败：${restored.out}`);
    failures += 1;
    continue;
  }
  if (sha256(TARGET) !== baselineSha) {
    console.error(`[mutate] ${which} 还原后与基线不一致（sha256 变了）`);
    failures += 1;
  }
}

const finalSha = sha256(TARGET);
if (finalSha !== baselineSha) {
  console.error(`[mutate] 最终文件与基线不一致：${finalSha} != ${baselineSha}`);
  failures += 1;
}

console.log("[mutate] 还原后复跑应为 GREEN…");
const after = runVitest();
if (after.status !== 0 || after.passed < MIN_TESTS) {
  console.error(`[mutate] 还原后复跑未全绿：${after.summary}`);
  failures += 1;
} else {
  console.log(`[mutate] 还原后 GREEN：${after.summary}`);
}
rmSync(scratch, { force: true });

if (failures > 0) {
  console.error(`反向变异未通过：${failures} 处问题。`);
  process.exit(1);
}
console.log("反向变异通过：M1/M2 各自变红，还原逐字节一致，复跑全绿。");
