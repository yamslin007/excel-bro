// 反向变异驱动：逐个应用 `mutate_scope_authorization.mjs` 的 SEED，跑**指定**用例，
// 用**四判据**证明它真的被这条判据挡下来：
//   ① 退出码非 0（运行期错误不算，还要看后面三条）
//   ② 有有效的测试摘要（vitest: "Tests  N failed"）
//   ③ 预期的那条用例真的跑了并且红了（按用例名逐字匹配）
//   ④ 失败原文里出现**它自己那一处**的证据（errorContains）
// 以及还原判据：还原后文件 sha256 必须与变异前逐字节一致。
//
// 用法：node scripts/run_scope_authorization_mutations.mjs [A1 A2 ...]
// 结论落盘到 .tmp-mutations/ ；**不要**用 `| tail` 抓全量输出（npm notice 会挤掉摘要）。
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

const REPO = "E:\\codex\\excel-bro";
const OUT = path.join(REPO, ".tmp-mutations");
/** 直接走 vitest 的 mjs 入口：不经过 npx/shell，cwd 参数才不会被吃掉。 */
const VITEST = [process.execPath, path.join(REPO, "node_modules", "vitest", "vitest.mjs"), "run"];
const EXCEL = path.join(REPO, "apps", "excel-addin");

mkdirSync(OUT, { recursive: true });

/** 每个 SEED：在哪跑、预期哪条用例变红、失败原文里必须出现的证据。 */
const CASES = {
  A1: {
    cwd: EXCEL,
    file: "src/excelWriteEndToEnd.test.tsx",
    testName: "⑨切到另一套口径后直接关掉弹层：后台不误报，发送也不静默改读",
    errorContains: "只是看了看另一套口径就弹范围变化",
  },
  A2: {
    cwd: EXCEL,
    file: "src/excelWriteEndToEnd.test.tsx",
    testName: "④编辑草稿、切换模式后点「取消」：已提交授权与守卫基线不变",
    errorContains: "已固定 2 个工作表",
  },
  A3: {
    cwd: EXCEL,
    file: "src/hooks/useScopeSelection.test.ts",
    testName: "重读宿主只刷新草稿：已提交的范围与口径一字不改",
    errorContains: "to deeply equal",
  },
  A4: {
    cwd: EXCEL,
    file: "src/hooks/useAuthorizedTarget.test.tsx",
    testName: "旧读数不得判成范围变化：作废重读后一致就放行",
    errorContains: "旧读数作废后必须重读一次",
  },
  A5: {
    cwd: EXCEL,
    file: "src/excelWriteEndToEnd.test.tsx",
    testName: "⑧跟随当前口径：确认后切到另一张表要先确认，不会静默改读",
    errorContains: "活动表变了却仍按旧的活动表放行",
  },
  A6: {
    cwd: EXCEL,
    file: "src/excelWriteEndToEnd.test.tsx",
    testName: "⑨切到另一套口径后直接关掉弹层：后台不误报，发送也不静默改读",
    errorContains: "请先确认当前对话的数据范围",
  },
};

/** 变异脚本只认 SEED 名；直接问它要目标文件，避免这里再抄一份路径。 */
function mutateFileFor(seed) {
  const probe = spawnSync(process.execPath,
    [path.join(REPO, "scripts", "mutate_scope_authorization.mjs"), seed, "restore"],
    { cwd: REPO, encoding: "utf8" });
  const match = /(?:ALREADY-CLEAN|RESTORED)：(.+)/.exec(probe.stdout.trim());
  if (!match) throw new Error(`无法确定 ${seed} 的目标文件：${probe.stdout}${probe.stderr}`);
  return match[1].trim();
}

function sha256(file) {
  return createHash("sha256").update(readFileSync(file)).digest("hex");
}

function mutate(seed, mode) {
  const args = [path.join(REPO, "scripts", "mutate_scope_authorization.mjs"), seed];
  if (mode) args.push(mode);
  const run = spawnSync(process.execPath, args, { cwd: REPO, encoding: "utf8" });
  return { status: run.status, text: `${run.stdout}${run.stderr}`.trim() };
}

/** 被杀掉也要把变异还回去：否则工作区会带着一个变异进入下一次运行。 */
const filesTouched = new Set();
function restoreAll() {
  for (const key of Object.keys(CASES)) mutate(key, "restore");
}
for (const signal of ["SIGINT", "SIGTERM", "SIGHUP"]) {
  process.on(signal, () => {
    restoreAll();
    process.exit(130);
  });
}
process.on("uncaughtException", (reason) => {
  restoreAll();
  console.error("UNCAUGHT：", reason);
  process.exit(1);
});

const wanted = process.argv.slice(2);
const keys = wanted.length ? wanted : Object.keys(CASES);
const report = [];

for (const key of keys) {
  const spec = CASES[key];
  if (!spec) throw new Error(`未知 SEED：${key}`);
  const file = mutateFileFor(key);
  filesTouched.add(file);
  // 前置闸：先把所有 SEED 都还原，确认没有别的变异在位
  // （否则"红"可能来自上一个 SEED，而不是这一处）。
  restoreAll();
  const before = sha256(file);

  const applied = mutate(key);
  if (applied.status !== 0 || !applied.text.startsWith("APPLIED")) {
    report.push({ seed: key, verdict: `BLOCKED: 无法应用（${applied.text}）` });
    continue;
  }

  const logFile = path.join(OUT, `${key}.log`);
  const run = spawnSync(VITEST[0], [...VITEST.slice(1), spec.file, "-t", spec.testName, "--reporter=verbose"], {
    cwd: spec.cwd,
    encoding: "utf8",
    env: { ...process.env, NO_COLOR: "1", FORCE_COLOR: "0" }
  });
  const raw = `${run.stdout ?? ""}${run.stderr ?? ""}`;
  // vitest 摘要带 ANSI ⇒ 解析前剥色码。
  const text = raw.replace(/\u001b\[[0-9;]*m/g, "");
  writeFileSync(logFile, text, "utf8");

  const restored = mutate(key, "restore");
  const after = sha256(file);
  const restoredOk = restored.status === 0 && after === before;

  const exitNonZero = run.status !== 0;
  const summary = /Tests\s+\d+ failed/.test(text);
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
writeFileSync(path.join(OUT, "summary-scope-authorization.json"), JSON.stringify(report, null, 2), "utf8");
process.exit(failed === 0 ? 0 : 1);
