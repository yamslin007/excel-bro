/**
 * 反向变异驱动：证明 SDK 读表桥那三处**生命周期竞态**的修复真有判别力。
 *
 * 四条判据缺一不算过（与仓库既有驱动同口径）：
 * 1. 未变异时基线 GREEN（退出码 0 + 有效摘要 + 用例数达标）；
 * 2. apply 后必须变红，且**点名那条用例**真的出现在失败列表里；
 * 3. restore 后文件**逐字节一致**（sha256 比对）；
 * 4. 还原后复跑全绿。
 *
 * 只卡"退出码非 0"不算：写崩一行 TSX 也是非 0，那是运行期错误，不是断言在起作用。
 *
 * 用法：
 *   node scripts/run_bridge_lifecycle_race_mutations.mjs        # 跑 R1/R2/R3
 *   node scripts/run_bridge_lifecycle_race_mutations.mjs R2     # 只跑一条（调试用）
 */
import { spawnSync } from "node:child_process";
import { openSync, closeSync, readFileSync, mkdirSync, existsSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");
const scratchDir = resolve(repoRoot, "outputs");
const scratch = resolve(scratchDir, "bridge-lifecycle-race-mutation.log");
const vitest = resolve(repoRoot, "node_modules", "vitest", "vitest.mjs");
const mutateScript = resolve(here, "mutate_bridge_lifecycle_races.mjs");

mkdirSync(scratchDir, { recursive: true });

const CASES = {
  R1: {
    label: "旧轮次退出时不接续新一轮排队的工作",
    cwd: "apps/excel-addin",
    file: "src/hooks/useSdkToolBridge.ts",
    tests: ["src/hooks/useSdkToolBridge.test.tsx"],
    minTests: 19,
    expectTitle: "旧轮次在途时到达的指针，旧轮次退出后必须被接上"
  },
  R2: {
    label: "迟到的初始化照旧提交（不核验轮次）",
    cwd: "apps/excel-addin",
    file: "src/hooks/useSdkToolBridge.ts",
    tests: ["src/hooks/useSdkToolBridge.test.tsx"],
    minTests: 19,
    expectTitle: "初始化期间 cancel：迟到的初始化不得把桥重新激活"
  },
  R3: {
    label: "读表返回后不核验轮次，直接发回执",
    cwd: "packages/core",
    file: "src/sdkToolBridge.ts",
    tests: ["src/sdkToolBridge.test.ts"],
    minTests: 35,
    expectTitle: "读表返回后才发现本轮已失效"
  }
};

function runVitest(kase) {
  const fd = openSync(scratch, "w");
  const result = spawnSync(process.execPath, [vitest, "run", ...kase.tests], {
    cwd: resolve(repoRoot, kase.cwd),
    stdio: ["ignore", fd, fd]
  });
  closeSync(fd);
  const raw = readFileSync(scratch, "utf8").replace(/\x1b\[[0-9;]*m/g, "");
  const summary =
    raw.split(/\r?\n/).find((line) => /Tests\s+\d+ (passed|failed)/.test(line)) ?? "";
  const passed = /Tests\s+(\d+) passed/.exec(summary);
  const failed = /(\d+) failed/.exec(summary);
  return {
    status: result.status,
    raw,
    summary: summary.trim(),
    passed: passed ? Number(passed[1]) : 0,
    failed: failed ? Number(failed[1]) : 0
  };
}

function mutate(which, restore) {
  const fd = openSync(scratch, "w");
  const result = spawnSync(
    process.execPath,
    [mutateScript, which, ...(restore ? ["restore"] : [])],
    { cwd: repoRoot, stdio: ["ignore", fd, fd] }
  );
  closeSync(fd);
  return {
    status: result.status,
    out: readFileSync(scratch, "utf8").trim(),
    error: result.error ? String(result.error.code) : null
  };
}

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

const selected = process.argv[2] ? [process.argv[2]] : Object.keys(CASES);
for (const which of selected) {
  if (!CASES[which]) {
    console.error(`未知变异：${which}`);
    process.exit(2);
  }
}

let failures = 0;
for (const which of selected) {
  const kase = CASES[which];
  const target = resolve(repoRoot, kase.cwd, kase.file);
  if (!existsSync(target)) {
    console.error(`[mutate] ${which} 目标文件不存在：${target}`);
    failures += 1;
    continue;
  }
  const baselineSha = sha256(target);
  console.log(`\n[mutate] ${which} ${kase.label}`);
  console.log(`[mutate] 目标 ${kase.cwd}/${kase.file} 基线 sha256=${baselineSha}`);

  const baseline = runVitest(kase);
  if (baseline.status !== 0 || baseline.passed < kase.minTests) {
    console.error(
      `[mutate] ${which} 基线就不绿（status=${baseline.status}, 摘要=${baseline.summary || "无"}）`
    );
    failures += 1;
    continue;
  }
  console.log(`[mutate] 基线 GREEN：${baseline.summary}`);

  const applied = mutate(which, false);
  if (applied.status !== 0 || applied.error) {
    console.error(`[mutate] ${which} apply 失败：${applied.out}`);
    failures += 1;
    continue;
  }

  const run = runVitest(kase);
  const named = run.raw.includes(kase.expectTitle);
  const red = run.status !== 0 && run.failed >= 1 && run.failed <= 8 && named;
  console.log(
    `[mutate] ${which} → ${red ? "RED(OK)" : "GREEN(NG!)"} ${run.summary || "(无摘要)"}` +
      ` named=${named}`
  );
  if (!red) {
    console.error(
      `  ${which} 未按预期变红（status=${run.status} failed=${run.failed} named=${named}）`
    );
    for (const line of run.raw.split(/\r?\n/)) {
      if (line.includes("×") || line.includes("FAIL ")) console.error(`  ${line.trim()}`);
    }
    failures += 1;
  }

  const restored = mutate(which, true);
  if (restored.status !== 0 || restored.out === "ALREADY-CLEAN") {
    console.error(`[mutate] ${which} restore 失败：${restored.out}`);
    failures += 1;
    continue;
  }
  if (sha256(target) !== baselineSha) {
    console.error(`[mutate] ${which} 还原后与基线不一致（sha256 变了）`);
    failures += 1;
  }
}

if (failures > 0) {
  console.error(`\n反向变异未通过：${failures} 处问题。`);
  process.exit(1);
}

// 第四条判据：全部还原之后再复跑一次，确认绿不是"某次 restore 恰好没生效"。
console.log("\n[mutate] 全部还原后复跑…");
for (const which of selected) {
  const kase = CASES[which];
  const run = runVitest(kase);
  if (run.status !== 0 || run.passed < kase.minTests) {
    console.error(`[mutate] ${which} 还原后复跑未全绿：${run.summary}`);
    failures += 1;
  } else {
    console.log(`[mutate] ${which} 复跑 GREEN：${run.summary}`);
  }
}
if (failures > 0) {
  console.error(`\n反向变异未通过：${failures} 处问题。`);
  process.exit(1);
}
console.log("\n反向变异通过：R1/R2/R3 各自变红并点名到对应用例，还原逐字节一致，复跑全绿。");
