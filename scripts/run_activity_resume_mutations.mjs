/**
 * 反向变异驱动：证明「暂停 / 继续的步骤身份与终态」这组判据**真有判别力**。
 *
 * 四条判据缺一不算过：
 * 1. 未变异时基线 GREEN（退出码 0 + 有效摘要 + 用例数达标）—— **core 与 hook 两侧都要**；
 * 2. 每个变异 apply 后，**它自己声明的那条用例**必须失败：
 *    - 退出码非 0（不能只卡这一条：写崩 JSX 也非 0）；
 *    - 失败摘要里点名了这条用例；
 *    - 该用例的失败原文里有它自己那处证据（T1/T4 看 `warning`，T3/T6 看 `completed`）；
 * 3. restore 后目标文件**逐字节一致**（sha256 与基线相同）；
 * 4. 还原后复跑两侧全绿。
 *
 * 六个变异各钉一处（见 scripts/mutate_activity_resume.mjs 的注释）：
 * T1/T4 钉「恢复时凭什么重开一步」，T2/T5 钉「哪些步骤算被收尾关掉的」，
 * T3/T6 钉「这一步自己的终态听谁的」。同一判别点的变异分别落在共享层与调用方，
 * 证明"两端读同一份判据"不是靠巧合。
 *
 * 用法：node scripts/run_activity_resume_mutations.mjs
 */
import { spawnSync } from "node:child_process";
import { openSync, closeSync, readFileSync, mkdirSync, rmSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");
const scratchDir = resolve(repoRoot, "outputs");
const scratch = resolve(scratchDir, "activity-resume-mutation.log");
const vitest = resolve(repoRoot, "node_modules", "vitest", "vitest.mjs");

const SUITES = {
  core: {
    dir: resolve(repoRoot, "packages", "core"),
    test: "src/activityTimeline.test.ts",
    /** 该文件 16 条：身份与去重 / 终态 / 暂停后继续 / 计时口径。 */
    minTests: 16
  },
  hook: {
    dir: resolve(repoRoot, "apps", "excel-addin"),
    test: "src/hooks/useActivityProgress.test.ts",
    /** 该文件 23 条：起点/推进/收尾/持久化/恢复 + 本批补的"收尾两档结论"2 条。 */
    minTests: 23
  }
};

/** 变异 → 它自己声明的预期失败用例名 + 失败原文里必须出现的那处证据。 */
const MUTATIONS = {
  T1: { suite: "core", expect: "暂停顺手关掉的那一步", evidence: "toBe(\"progress\")" },
  T2: { suite: "core", expect: "收尾时这一步自己已做完", evidence: "closedByFinish" },
  T3: { suite: "core", expect: "收尾时这一步自己已做完", evidence: "completed" },
  T4: { suite: "core", expect: "自己报了 warning 结束的步骤", evidence: "warning" },
  T5: { suite: "hook", expect: "暂停收尾时那一步没有结论", evidence: "closedByFinish" },
  T6: { suite: "hook", expect: "带着待确认结果收尾时", evidence: "completed" }
};

const FILES = [
  resolve(repoRoot, "packages", "core", "src", "activityTimeline.ts"),
  resolve(repoRoot, "apps", "excel-addin", "src", "hooks", "useActivityProgress.ts")
];

mkdirSync(scratchDir, { recursive: true });

/** 用文件描述符接子进程输出：判据要读到真实摘要，不能只卡退出码。 */
function runSuite(suite) {
  const fd = openSync(scratch, "w");
  const result = spawnSync(process.execPath, [vitest, "run", SUITES[suite].test], {
    cwd: SUITES[suite].dir,
    stdio: ["ignore", fd, fd]
  });
  closeSync(fd);
  const raw = readFileSync(scratch, "utf8").replace(/\x1b\[[0-9;]*m/g, "");
  const summary = raw.split(/\r?\n/).find((line) => /Tests\s+\d+ (passed|failed)/.test(line)) ?? "";
  const passed = /Tests\s+(\d+) passed/.exec(summary);
  return {
    status: result.status,
    raw,
    summary: summary.trim(),
    passed: passed ? Number(passed[1]) : 0
  };
}

function runBoth() {
  return { core: runSuite("core"), hook: runSuite("hook") };
}

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

function mutate(which, restore) {
  const args = [resolve(here, "mutate_activity_resume.mjs"), which, ...(restore ? ["restore"] : [])];
  const fd = openSync(scratch, "w");
  const result = spawnSync(process.execPath, args, { cwd: repoRoot, stdio: ["ignore", fd, fd] });
  closeSync(fd);
  return {
    status: result.status,
    out: readFileSync(scratch, "utf8").trim(),
    error: result.error ? String(result.error.code) : null
  };
}

/** 失败段落里是否点名了这条用例，并且带着它自己那处证据。 */
function inspectFailure(run, expect, evidence) {
  const lines = run.raw.split(/\r?\n/);
  const failedLine = lines.find((line) => line.includes("FAIL") && line.includes(expect));
  const body = run.raw.slice(run.raw.indexOf(failedLine ?? "\u0000"));
  return {
    named: Boolean(failedLine),
    evidence: body.includes(evidence)
  };
}

const baselineSha = new Map(FILES.map((file) => [file, sha256(file)]));
for (const [file, hash] of baselineSha) console.log(`[mutate] 基线 ${file} sha256=${hash}`);

console.log("[mutate] 先跑变异串定点闸…");
const preflight = mutate("check", false);
if (preflight.status !== 0) {
  console.error(preflight.out);
  console.error("[mutate] 变异串已失配，拒绝继续（先修 scripts/mutate_activity_resume.mjs 的种子）。");
  process.exit(1);
}
console.log(preflight.out.split(/\r?\n/).filter(Boolean).join("  |  "));

console.log("[mutate] 基线（未变异）两侧都应为 GREEN…");
const baseline = runBoth();
let failures = 0;
for (const suite of ["core", "hook"]) {
  const run = baseline[suite];
  if (run.status !== 0 || !run.summary || run.passed < SUITES[suite].minTests) {
    console.error(`基线就不绿（${suite}: status=${run.status}, 摘要=${run.summary || "无"}）`);
    failures += 1;
  } else {
    console.log(`[mutate] 基线 GREEN ${suite}：${run.summary}`);
  }
}
if (failures > 0) process.exit(1);

for (const [which, spec] of Object.entries(MUTATIONS)) {
  const applied = mutate(which, false);
  if (applied.status !== 0 || applied.error) {
    console.error(`[mutate] ${which} apply 失败：status=${applied.status} err=${applied.error} out=${applied.out}`);
    failures += 1;
    continue;
  }
  const runs = runBoth();
  const mine = runs[spec.suite];
  const other = spec.suite === "core" ? "hook" : "core";
  const seen = inspectFailure(mine, spec.expect, spec.evidence);
  const ok = mine.status !== 0 && seen.named && seen.evidence;
  console.log(
    `[mutate] ${which}（应由 ${spec.suite} 的「${spec.expect}」判红）→ ${ok ? "RED(OK)" : "GREEN(NG!)"}` +
      ` status=${mine.status} 点名=${seen.named} 原文证据=${seen.evidence} 摘要=${mine.summary || "(无)"}` +
      `；另一侧 ${other}=${runs[other].status === 0 ? "GREEN" : "RED"}`
  );
  if (!ok) {
    console.error(`  ${which} 之后没有按预期变红 —— 该断言没有判别力，见：${scratch}`);
    failures += 1;
  }
  const restored = mutate(which, true);
  if (restored.status !== 0 || restored.out === "ALREADY-CLEAN") {
    console.error(`[mutate] ${which} restore 失败：${restored.out}`);
    failures += 1;
    continue;
  }
  for (const [file, hash] of baselineSha) {
    if (sha256(file) !== hash) {
      console.error(`[mutate] ${which} 还原后 ${file} 与基线不一致（sha256 变了）`);
      failures += 1;
    }
  }
}

for (const [file, hash] of baselineSha) {
  const finalSha = sha256(file);
  if (finalSha !== hash) {
    console.error(`[mutate] 最终 ${file} 与基线不一致：${finalSha} != ${hash}`);
    failures += 1;
  }
}

console.log("[mutate] 还原后复跑两侧应为 GREEN…");
const after = runBoth();
for (const suite of ["core", "hook"]) {
  if (after[suite].status !== 0 || after[suite].passed < SUITES[suite].minTests) {
    console.error(`[mutate] 还原后 ${suite} 未全绿：${after[suite].summary}`);
    failures += 1;
  } else {
    console.log(`[mutate] 还原后 GREEN ${suite}：${after[suite].summary}`);
  }
}
rmSync(scratch, { force: true });

if (failures > 0) {
  console.error(`反向变异未通过：${failures} 处问题。`);
  process.exit(1);
}
console.log("反向变异通过：T1..T6 各自把它自己那条用例判红（含失败原文里的自证），还原逐字节一致，复跑两侧全绿。");
