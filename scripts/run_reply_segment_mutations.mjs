/**
 * 反向变异驱动：证明本批四处修复的守卫**真有判别力**。
 *
 * 四条判据缺一不算过：
 * 1. 未变异时基线 GREEN（退出码 0 + 有效摘要 + 用例数达标）；
 * 2. 每个变异 apply 后，**它自己声明的那条用例**必须失败：
 *    - 退出码非 0（不能只卡这一条：写崩语法也非 0）；
 *    - 失败摘要里点名了这条用例；
 *    - 该用例的失败原文里有它自己那处证据；
 * 3. restore 后目标文件**逐字节一致**（sha256 与基线相同）；
 * 4. 还原后复跑全绿。
 *
 * 覆盖：缺陷 1（早期恢复必须新开段，Excel/WPS 各一处）、缺陷 2（段链接最新段）、
 * 缺陷 3（"正在运行"只认执行器真实启动 + 未开始执行不带执行版本）、
 * 缺陷 4（增量按事件身份判重，不按文本）。
 *
 * 用法：node scripts/run_reply_segment_mutations.mjs
 */
import { spawnSync } from "node:child_process";
import { openSync, closeSync, readFileSync, mkdirSync, rmSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");
const outputsDir = resolve(repoRoot, "outputs", "acceptance-2026-10-22", "review-fixes");
const scratch = resolve(outputsDir, "mutation-run.log");
const vitest = resolve(repoRoot, "node_modules", "vitest", "vitest.mjs");
const python = resolve(repoRoot, ".venv", "Scripts", "python.exe");

const SUITES = {
  core: {
    kind: "vitest",
    cwd: resolve(repoRoot, "packages", "core"),
    test: "src/pythonCodeArtifact.test.ts",
    /** 该文件 31 条：状态机、增量判重、缺片、未开始执行。 */
    minTests: 31
  },
  segments: {
    kind: "vitest",
    cwd: resolve(repoRoot, "packages", "core"),
    test: "src/replySegments.test.ts",
    /** 该文件 18 条：段链 / 幂等 / 段能不能收进度。 */
    minTests: 18
  },
  excel: {
    kind: "vitest",
    cwd: resolve(repoRoot, "apps", "excel-addin"),
    test: "src/messageFirstRouting.test.tsx",
    /** 该文件 42 条：消息优先路由 + 早期暂停三阶段 × 两入口。 */
    minTests: 42
  },
  wps: {
    kind: "vitest",
    cwd: resolve(repoRoot, "apps", "wps-addin"),
    test: "src/wpsPlanningResumeClosure.test.tsx",
    /** 该文件 36 条：续跑闭包 + 段链 + 幂等身份。 */
    minTests: 36
  },
  wpsPause: {
    kind: "vitest",
    cwd: resolve(repoRoot, "apps", "wps-addin"),
    test: "src/wpsPlanningPauseCard.test.tsx",
    /** 该文件 28 条：暂停卡契约 + 早期暂停三阶段 × 两入口。 */
    minTests: 28
  },
  stage: {
    kind: "pytest",
    cwd: repoRoot,
    test: "server/tests/test_python_code_stage_evidence.py",
    /** 该文件 8 条：正常启动 / 关闭 / 无运行器 / 无绑定 / 入参非法 / 两种取消 / 异常。 */
    minTests: 8
  },
  guard: {
    kind: "vitest",
    cwd: resolve(repoRoot, "apps", "excel-addin"),
    test: "src/activityMessages.test.ts",
    /** 该文件 12 条：活动日志写回路径 + 只有活动段能接收进度。 */
    minTests: 12
  }
};

/** 变异 → 它自己声明的预期失败用例 + 失败原文里必须出现的那处证据。 */
const MUTATIONS = {
  S1: { suite: "wps", expect: "a resumed stream can be paused again", evidence: "task-progress-jump" },
  S7: { suite: "segments", expect: "新段接在该任务", evidence: 'toBe("C")' },
  S2: { suite: "core", expect: "连续右括号", evidence: "f(g(1))" },
  S3: { suite: "excel", expect: "early 'fields' pause restarts via 'button'", evidence: "task-progress-timeline" },
  S4: { suite: "wpsPause", expect: "早期暂停", evidence: "查看后续进展" },
  S5: { suite: "stage", expect: "test_cancel_before_runner_start_is_not_started", evidence: "not_started" },
  S6: { suite: "stage", expect: "test_disabled_feature_never_claims_running", evidence: "executedVersion" },
  S8: { suite: "guard", expect: "refuses to write progress into a frozen segment", evidence: "迟到的执行进度" }
};

const FILES = [
  resolve(repoRoot, "packages", "core", "src", "replySegments.ts"),
  resolve(repoRoot, "packages", "core", "src", "pythonCodeArtifact.ts"),
  resolve(repoRoot, "apps", "excel-addin", "src", "App.tsx"),
  resolve(repoRoot, "apps", "excel-addin", "src", "activityMessages.ts"),
  resolve(repoRoot, "apps", "wps-addin", "src", "WpsAgentPanel.tsx"),
  resolve(repoRoot, "server", "app", "python_coordinator.py")
];

mkdirSync(outputsDir, { recursive: true });

function stripAnsi(text) {
  return text.replace(/\x1b\[[0-9;]*m/g, "");
}

/** 用文件描述符接子进程输出：判据要读到真实摘要，不能只卡退出码。 */
function runSuite(name) {
  const suite = SUITES[name];
  const fd = openSync(scratch, "w");
  const args = suite.kind === "vitest"
    ? [vitest, "run", suite.test]
    : ["-m", "pytest", suite.test, "-q", "--basetemp=.tmp-mut-stage"];
  const result = spawnSync(
    suite.kind === "vitest" ? process.execPath : python,
    args,
    { cwd: suite.cwd, stdio: ["ignore", fd, fd] }
  );
  closeSync(fd);
  const raw = stripAnsi(readFileSync(scratch, "utf8"));
  const summary = suite.kind === "vitest"
    ? raw.split(/\r?\n/).find((line) => /Tests\s+\d+ (passed|failed)/.test(line)) ?? ""
    : raw.split(/\r?\n/).find((line) => /\d+ (passed|failed)/.test(line)) ?? "";
  const passed = /(\d+) passed/.exec(summary);
  return {
    status: result.status,
    raw,
    summary: summary.trim(),
    passed: passed ? Number(passed[1]) : 0
  };
}

function runMany(names) {
  return Object.fromEntries(names.map((name) => [name, runSuite(name)]));
}

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

function mutate(which, restore) {
  const args = [resolve(here, "mutate_reply_segments.mjs"), which, ...(restore ? ["restore"] : [])];
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
  /*
   * 两类测试的报告格式不同：
   * - vitest：`FAIL src/x.test.ts > 用例名`，失败原文紧随其后；
   * - pytest：先打印 `______ test_name ______` 的 traceback，最后才是 `FAILED ...`
   *   —— 从 `FAILED` 行往后切会把 traceback 丢掉，"失败原文里有它自己那处证据"
   *   这条判据就永远为假（会误判成"断言没有判别力"）。
   */
  const index = lines.findIndex((line) =>
    (line.includes("FAIL") || /^_+ .+ _+$/.test(line)) && line.includes(expect));
  const body = index >= 0 ? lines.slice(index, index + 80).join("\n") : "";
  return { named: index >= 0, evidence: body.includes(evidence) };
}

const baselineSha = new Map(FILES.map((file) => [file, sha256(file)]));
for (const [file, hash] of baselineSha) console.log(`[mutate] 基线 ${file} sha256=${hash}`);

console.log("[mutate] 先跑变异串定点闸…");
const preflight = mutate("check", false);
if (preflight.status !== 0) {
  console.error(preflight.out);
  console.error("[mutate] 变异串已失配，拒绝继续（先修 scripts/mutate_reply_segments.mjs 的种子）。");
  process.exit(1);
}
console.log(preflight.out.split(/\r?\n/).filter(Boolean).join("  |  "));

const suiteNames = Object.keys(SUITES);
let failures = 0;

console.log("[mutate] 基线（未变异）七个套件都应为 GREEN…");
const baseline = runMany(suiteNames);
for (const name of suiteNames) {
  const run = baseline[name];
  if (run.status !== 0 || !run.summary || run.passed < SUITES[name].minTests) {
    console.error(`基线就不绿（${name}: status=${run.status}, 摘要=${run.summary || "无"}）`);
    failures += 1;
  } else {
    console.log(`[mutate] 基线 GREEN ${name}：${run.summary}`);
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
  const run = runSuite(spec.suite);
  const seen = inspectFailure(run, spec.expect, spec.evidence);
  const ok = run.status !== 0 && seen.named && seen.evidence;
  const detail = run.summary || "(无)";
  console.log(
    `[mutate] ${which}（应由 ${spec.suite} 的「${spec.expect}」判红）→ ${ok ? "RED(OK)" : "GREEN(NG!)"}` +
      ` status=${run.status} 点名=${seen.named} 原文证据=${seen.evidence} 摘要=${detail}`
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

console.log("[mutate] 还原后复跑七个套件应为 GREEN…");
const after = runMany(suiteNames);
for (const name of suiteNames) {
  if (after[name].status !== 0 || after[name].passed < SUITES[name].minTests) {
    console.error(`[mutate] 还原后 ${name} 未全绿：${after[name].summary}`);
    failures += 1;
  } else {
    console.log(`[mutate] 还原后 GREEN ${name}：${after[name].summary}`);
  }
}
rmSync(scratch, { force: true });

if (failures > 0) {
  console.error(`反向变异未通过：${failures} 处问题。`);
  process.exit(1);
}
console.log(
  `反向变异通过：${Object.keys(MUTATIONS).join("/")} 各自把它自己那条业务用例判红（含失败原文里的自证），` +
    "还原逐字节一致，复跑七个套件全绿。"
);
