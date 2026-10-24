// 反向变异驱动：逐个应用 `mutate_worksheet_restore.mjs` 的 SEED，跑**指定**用例，
// 用**四判据**证明它真的被这条判据挡下来：
//   ① 退出码非 0（运行期错误不算，还要看后面三条）
//   ② 有有效的测试摘要（pytest: "N failed"；vitest: "Tests  N failed"）
//   ③ 预期的那条用例真的跑了并且红了（按用例名逐字匹配）
//   ④ 失败原文里出现**它自己那一处**的证据（errorContains）
// 以及还原判据：还原后文件 sha256 必须与变异前逐字节一致。
//
// 用法：node scripts/run_worksheet_restore_mutations.mjs [R1 R2 ...]
// 结论落盘到 .tmp-mutations/ ；**不要**用 `| tail` 抓全量输出（npm notice 会挤掉摘要）。
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

const REPO = "E:\\codex\\excel-bro";
const OUT = path.join(REPO, ".tmp-mutations");
const PYTHON = "D:\\AdvancedSoftware\\python\\python3.13.3\\python.exe";
/** 直接走 vitest 的 mjs 入口：不经过 npx/shell，cwd 参数才不会被吃掉。 */
const VITEST = [process.execPath, path.join(REPO, "node_modules", "vitest", "vitest.mjs"), "run"];

/**
 * 每次运行给 pytest 一个**全新**的 basetemp。
 *
 * 复用固定目录会让 pytest 在会话开始先删旧树：本机沙箱的批量删除保护会直接
 * `SystemExit(1)`，测试一条都没跑 —— 那是环境噪声，却会被误读成"变异被抓住"。
 * 这里顺带把 `-p no:cacheprovider` 打开，避免 `.pytest_cache` 的权限告警。
 */
function PYTEST_TEMP(key) {
  return `.tmp-mutations/pytest-${key}-${process.pid}-${Date.now()}`;
}

mkdirSync(OUT, { recursive: true });

/** 每个 SEED：在哪个目录跑哪条命令、预期哪条用例变红、失败原文里必须出现的证据。 */
const CASES = {
  R1: {
    seed: "R1",
    cwd: REPO,
    cmd: [PYTHON, "-m", "pytest", "server/tests/test_worksheet_restore.py",
      "-q", "--basetemp=" + PYTEST_TEMP("R1")],
    testName: "test_restore_puts_sheets_back_before_a_later_survivor",
    errorContains: "放回顺序必须与执行前一致",
  },
  R2: {
    seed: "R2",
    cwd: REPO,
    cmd: [PYTHON, "-m", "pytest", "server/tests/test_worksheet_restore.py",
      "-q", "--basetemp=" + PYTEST_TEMP("R2")],
    testName: "test_restore_repairs_ref_formulas_and_workbook_names",
    errorContains: "#REF!",
  },
  R3: {
    seed: "R3",
    cwd: REPO,
    cmd: [PYTHON, "-m", "pytest", "server/tests/test_worksheet_restore.py",
      "-q", "--basetemp=" + PYTEST_TEMP("R3")],
    testName: "test_restore_repairs_ref_formulas_and_workbook_names",
    errorContains: "RefersTo",
  },
  R4: {
    seed: "R4",
    cwd: REPO,
    cmd: [PYTHON, "-m", "pytest", "server/tests/test_worksheet_restore.py",
      "-q", "--basetemp=" + PYTEST_TEMP("R4")],
    testName: "test_restore_refuses_when_same_name_sheet_exists",
    errorContains: "conflict",
  },
  R5: {
    seed: "R5",
    cwd: path.join(REPO, "packages", "core"),
    cmd: [...VITEST, "src/worksheetRestore.test.ts"],
    testName: "refuses sheets that this restore cannot explain",
    errorContains: "unexpected_sheets",
  },
  R6: {
    seed: "R6",
    cwd: path.join(REPO, "apps", "excel-addin"),
    cmd: [...VITEST, "src/hooks/useAuthorizedTarget.test.tsx"],
    testName: "accepts only the directory growth this restore explains",
    errorContains: "expected true to be false",
  },
  R7: {
    seed: "R7",
    cwd: path.join(REPO, "apps", "excel-addin"),
    cmd: [...VITEST, "src/hooks/useWorksheetCheckpoint.test.tsx"],
    testName: "a double click cannot dispatch a second restore",
    errorContains: "上一项检查点操作尚未结束",
  },
  R8: {
    seed: "R8",
    cwd: path.join(REPO, "apps", "excel-addin"),
    cmd: [...VITEST, "src/worksheetCheckpointReceipt.test.ts"],
    testName: "rejects an untrustworthy completion",
    errorContains: "worksheet restore receipt",
  },
};

/** 变异脚本只认 SEED 名；直接问它要目标文件，避免这里再抄一份路径。 */
function mutateFileFor(seed) {
  const probe = spawnSync(process.execPath,
    [path.join(REPO, "scripts", "mutate_worksheet_restore.mjs"), seed, "restore"],
    { cwd: REPO, encoding: "utf8" });
  const match = /(?:ALREADY-CLEAN|RESTORED)：(.+)/.exec(probe.stdout.trim());
  if (!match) throw new Error(`无法确定 ${seed} 的目标文件：${probe.stdout}${probe.stderr}`);
  return match[1].trim();
}

function sha256(file) {
  return createHash("sha256").update(readFileSync(file)).digest("hex");
}

function mutate(seed, mode) {
  const args = [path.join(REPO, "scripts", "mutate_worksheet_restore.mjs"), seed];
  if (mode) args.push(mode);
  const run = spawnSync(process.execPath, args, { cwd: REPO, encoding: "utf8" });
  return { status: run.status, text: `${run.stdout}${run.stderr}`.trim() };
}

const seed = process.argv.slice(2);
const wanted = seed.length ? seed : Object.keys(CASES);
const report = [];

for (const key of wanted) {
  const spec = CASES[key];
  if (!spec) throw new Error(`未知 SEED：${key}`);
  const file = mutateFileFor(key);
  const before = sha256(file);
  // 前置闸：确认没有别的变异在位（否则"红"可能来自上一个 SEED）。
  for (const other of Object.keys(CASES)) mutate(other, "restore");

  const applied = mutate(key);
  if (applied.status !== 0 || !applied.text.startsWith("APPLIED")) {
    report.push({ seed: key, verdict: `BLOCKED: 无法应用（${applied.text}）` });
    continue;
  }

  const logFile = path.join(OUT, `${key}.log`);
  const run = spawnSync(spec.cmd[0], spec.cmd.slice(1), {
    cwd: spec.cwd, encoding: "utf8",
    env: { ...process.env, NO_COLOR: "1", FORCE_COLOR: "0" }
  });
  const raw = `${run.stdout ?? ""}${run.stderr ?? ""}`;
  // vitest 摘要带 ANSI ⇒ 解析前剥色码（同 scripts/_strip_ansi.py 的口径）。
  const text = raw.replace(/\u001b\[[0-9;]*m/g, "");
  writeFileSync(logFile, text, "utf8");

  const restored = mutate(key, "restore");
  const after = sha256(file);
  const restoredOk = restored.status === 0 && after === before;

  const exitNonZero = run.status !== 0;
  const summary = /(\d+ failed)|Tests\s+\d+ failed/.test(text);
  const ranExpected = text.includes(spec.testName) && /(×|FAILED|✗)/.test(text);
  const hasEvidence = text.includes(spec.errorContains);

  const judgements = [
    `① 退出码非 0: ${exitNonZero ? "OK" : "NG"}（${run.status}）`,
    `② 有效摘要: ${summary ? "OK" : "NG"}`,
    `③ 预期用例变红: ${ranExpected ? "OK" : "NG"}（${spec.testName}）`,
    `④ 失败原文含证据「${spec.errorContains}」: ${hasEvidence ? "OK" : "NG"}`,
    `⑤ 还原后 sha256 一致: ${restoredOk ? "OK" : "NG"}`,
  ];
  const passed = exitNonZero && summary && ranExpected && hasEvidence && restoredOk;
  report.push({ seed: key, verdict: passed ? "RED(OK)" : "GREEN(NG)", judgements, log: logFile });
}

let failed = 0;
for (const item of report) {
  console.log(`\n=== ${item.seed}: ${item.verdict}`);
  if (item.judgements) for (const line of item.judgements) console.log(`    ${line}`);
  if (item.log) console.log(`    log: ${item.log}`);
  if (!/RED\(OK\)/.test(item.verdict)) failed += 1;
}
writeFileSync(path.join(OUT, "summary.json"), JSON.stringify(report, null, 2), "utf8");
process.exit(failed === 0 ? 0 : 1);
