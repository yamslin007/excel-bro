/**
 * 反向变异驱动：证明「发送即送达 + 停下来先核对时把请求还回输入框」这条链**真的有判别力**。
 *
 * 四条判据缺一不算过：
 * 1. 未变异时基线 GREEN（退出码 0 + 有效测试摘要 + 用例数达标）；
 * 2. 每个变异 apply 后用例必须变红（退出码非 0）；
 * 3. restore 后文件**逐字节一致**（sha256 比对）；
 * 4. 还原后复跑全绿。
 *
 * 四个变异各钉一处**互不相同**的属性：
 * M1 钉「输入框在受理那一刻是否真的空出来」（连带运行中按钮是 ■ 还是 ↑）；
 * M2 钉「用户气泡是否在受理那一刻落盘」；
 * M3 钉「目标确认后的回调重入会不会把同一句再送一遍」；
 * M4 钉「停下来让用户先核对目标时，请求有没有被还回输入框」。
 *
 * 用法：node scripts/run_message_delivery_on_send_mutations.mjs
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
const scratch = resolve(scratchDir, "message-delivery-mutation.log");
const TARGET = resolve(appDir, "src", "App.tsx");
const TEST_FILE = "src/messageFirstRouting.test.tsx";
const MUTATIONS = ["M1", "M2", "M3", "M4"];
// 该文件在基线上共 21 条（其中 2 条是本轮为这条链新增的判据）。
const MIN_TESTS = 21;

const vitest = resolve(repoRoot, "node_modules", "vitest", "vitest.mjs");

mkdirSync(scratchDir, { recursive: true });

/**
 * 用**文件描述符**而不是 pipe 接子进程输出：流水线里创建管道会被拦，
 * 而判据又必须读到真实的摘要行。
 */
function runVitest() {
  const fd = openSync(scratch, "w");
  const result = spawnSync(process.execPath, [vitest, "run", TEST_FILE], {
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
    resolve(here, "mutate_message_delivery_on_send.mjs"),
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
console.log("反向变异通过：M1/M2/M3/M4 各自变红，还原逐字节一致，复跑全绿。");
