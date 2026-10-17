/**
 * 反向变异驱动：证明「Excel 点复制按钮要出现『已复制』」这条判据**真的有判别力**。
 *
 * 四条判据缺一不算过：
 * 1. 未变异时基线 GREEN（退出码 0 + 有效测试摘要 + 用例数达标）；
 * 2. 每个变异 apply 后用例必须变红（退出码非 0）；
 * 3. restore 后文件**逐字节一致**（sha256 比对）；
 * 4. 还原后复跑全绿。
 *
 * 三个变异各钉一处**互不相同**的属性：
 * M1 钉「挂载期开关在 StrictMode 重放后还能不能放行」；
 * M2 钉「自动清除窗口的长度」；
 * M3 钉「界面上到底有没有『已复制』这三个字」。
 *
 * 用法：node scripts/run_copy_feedback_mount_guard_mutations.mjs
 */
import { spawnSync } from "node:child_process";
import { openSync, closeSync, readFileSync, mkdirSync, rmSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");
const appDir = resolve(repoRoot, "apps", "excel-addin");
const scratchDir = resolve(repoRoot, "outputs");
const scratch = resolve(scratchDir, "copy-feedback-mutation.log");
// 与 mutate_copy_feedback_mount_guard.mjs 的 TARGETS 保持一致（用于 sha256 复位判据）。
const TOUCHED = [
  resolve(appDir, "src", "hooks", "useCopyFeedback.ts"),
  resolve(appDir, "src", "components", "MessageCopyButton.tsx")
];
const TEST_FILES = [
  "src/hooks/useCopyFeedback.test.ts",
  "src/components/MessageCopyButton.test.tsx"
];
const MUTATIONS = ["M1", "M2", "M3"];
const MIN_TESTS = 10;

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
    resolve(here, "mutate_copy_feedback_mount_guard.mjs"),
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

const baselineSha = new Map(TOUCHED.map((path) => [path, sha256(path)]));
for (const [path, sha] of baselineSha) {
  console.log(`[mutate] 基线 ${path} sha256=${sha}`);
}

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
      `  ${which} 拆掉机制后用例仍然全绿 —— 断言没有判别力，见：${scratch}`
    );
    failures += 1;
  }
  const restored = mutate(which, true);
  if (restored.status !== 0 || restored.out === "ALREADY-CLEAN") {
    console.error(`[mutate] ${which} restore 失败：${restored.out}`);
    failures += 1;
    continue;
  }
  for (const path of TOUCHED) {
    if (sha256(path) !== baselineSha.get(path)) {
      console.error(`[mutate] ${which} 还原后与基线不一致（sha256 变了）：${path}`);
      failures += 1;
    }
  }
}

for (const path of TOUCHED) {
  const now = sha256(path);
  if (now !== baselineSha.get(path)) {
    console.error(`[mutate] 最终文件与基线不一致：${path} ${now}`);
    failures += 1;
  }
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
console.log("反向变异通过：M1/M2/M3 各自变红，还原逐字节一致，复跑全绿。");
