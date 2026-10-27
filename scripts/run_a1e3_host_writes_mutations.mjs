// 反向变异驱动：逐个应用 `mutate_a1e3_host_writes.mjs` 的 SEED，跑**指定**用例，
// 用**四判据**证明它真的被这条判据挡下来：
//   ① 退出码非 0（运行期错误不算，还要看后面三条）
//   ② 有有效的测试摘要（vitest: "Tests  N failed" / pytest: "N failed"）
//   ③ 预期的那条用例真的跑了并且红了（按用例名逐字匹配）
//   ④ 失败原文里出现**它自己那一处**的证据（errorContains）
// 以及还原判据：还原后文件 sha256 必须与**变异前**逐字节一致。
//
// 用法：node scripts/run_a1e3_host_writes_mutations.mjs [H1 H2 ...]
// 结论落盘到 .tmp-mutations/ ；**不要**用 `| tail` 抓全量输出（npm notice 会挤掉摘要）。
//
// 环境纪律：pytest 的 `--basetemp` 一律放 OS 临时目录 —— 放仓库里会触发环境的
// safe-delete 闸门，跑到中途整片 F/E（见 .workbuddy/memory 环境噪声一节）。
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";

const REPO = "E:\\codex\\excel-bro";
const OUT = path.join(REPO, ".tmp-mutations");
/** 直接走 vitest 的 mjs 入口：不经过 npx/shell，cwd 参数才不会被吃掉。 */
const VITEST = [process.execPath, path.join(REPO, "node_modules", "vitest", "vitest.mjs"), "run"];
const PYTHON = path.join(REPO, ".venv", "Scripts", "python.exe");
const WPS = path.join(REPO, "apps", "wps-addin");
const CORE = path.join(REPO, "packages", "core");

mkdirSync(OUT, { recursive: true });

/** 每个 SEED：在哪跑、预期哪条用例变红、失败原文里必须出现的证据。 */
const CASES = {
  H1: {
    kind: "pytest",
    file: "server/tests/test_host_write_observation.py",
    testName: "test_no_report_is_not_zero_writes",
    errorContains: "'complete' == 'not_collected'"
  },
  H2: {
    kind: "pytest",
    file: "server/tests/test_host_write_observation.py",
    testName: "test_report_outside_the_ledger_boundary_is_flagged_not_counted_as_zero",
    errorContains: "'zero' == 'unknown'"
  },
  H4: {
    kind: "vitest",
    cwd: CORE,
    file: "src/hostWriteAudit.test.ts",
    testName: "矛盾的线上数据不许被升级：zero 但覆盖不全 / 有边界外上报",
    errorContains: "expected true to be false"
  },
  H5: {
    kind: "vitest",
    cwd: WPS,
    file: "src/wpsHostWriteAudit.test.ts",
    // `-t` 的值是**正则**：用例名里的字面 `+`（"written + 回读"）会被当成量词，
    // 整文件于是全 skip、退出码 0 —— 那是"没跑"，不是"没红"。这里只用无元字符的前缀。
    testName: "一个动作只写一块时",
    errorContains: '"hostWriteCalls": 0'
  },
  H6: {
    kind: "vitest",
    cwd: WPS,
    file: "src/wpsHostWriteAudit.test.ts",
    testName: "文本保真用的 NumberFormat 只进主机调用数",
    // 变异后：2 格 @ 格式 + 2 格恢复 + 1 次计划写全部算成计划写 ⇒ 3，不是 1。
    errorContains: "expected 3 to be 1"
  },
  H7: {
    kind: "pytest",
    file: "server/tests/test_source_changed_stop.py",
    testName: "test_dev_fault_refuses_to_arm_without_a_workbook_identity",
    errorContains: "DID NOT RAISE"
  },
  H8: {
    kind: "pytest",
    file: "server/tests/test_source_changed_stop.py",
    testName: "test_dev_fault_refuses_when_the_name_carries_two_identities",
    errorContains: "DID NOT RAISE"
  },
  H9: {
    kind: "pytest",
    file: "server/tests/test_source_changed_stop.py",
    testName: "test_dev_fault_does_not_bind_a_task_that_only_carries_a_name",
    // 变异后按名字也绑上了 ⇒ 阶段从 `armed` 变成 `bound`，比后面那句断言更早暴露。
    errorContains: "'bound' == 'armed'"
  }
};

/** 变异脚本只认 SEED 名；直接问它要目标文件，避免这里再抄一份路径。 */
function mutateFileFor(seed) {
  const probe = spawnSync(process.execPath,
    [path.join(REPO, "scripts", "mutate_a1e3_host_writes.mjs"), seed, "restore"],
    { cwd: REPO, encoding: "utf8" });
  const match = /(?:ALREADY-CLEAN|RESTORED)：(.+)/.exec(probe.stdout.trim());
  if (!match) throw new Error(`无法确定 ${seed} 的目标文件：${probe.stdout}${probe.stderr}`);
  return match[1].trim();
}

function sha256(file) {
  return createHash("sha256").update(readFileSync(file)).digest("hex");
}

function mutate(seed, mode) {
  const args = [path.join(REPO, "scripts", "mutate_a1e3_host_writes.mjs"), seed];
  if (mode) args.push(mode);
  const run = spawnSync(process.execPath, args, { cwd: REPO, encoding: "utf8" });
  return { status: run.status, text: `${run.stdout}${run.stderr}`.trim() };
}

/** 被杀掉也要把变异还回去：否则工作区会带着一个变异进入下一次运行。 */
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
  // 前置闸：先把所有 SEED 都还原，确认没有别的变异在位
  // （否则"红"可能来自上一个 SEED，而不是这一处）。
  restoreAll();
  const before = sha256(file);

  const applied = mutate(key);
  if (applied.status !== 0 || !applied.text.startsWith("APPLIED")) {
    report.push({ seed: key, verdict: `BLOCKED: 无法应用（${applied.text}）` });
    continue;
  }
  // 变异必须真的改了内容：文件没变却报 APPLIED 就是"假装有证据"。
  const mutatedSha = sha256(file);

  const logFile = path.join(OUT, `${key}.log`);
  const basetemp = path.join(os.tmpdir(), `eb-mut-${key}`);
  const run = spec.kind === "pytest"
    ? spawnSync(PYTHON, [
        "-m", "pytest", `${spec.file}::${spec.testName}`,
        "-q", "--no-header", "-p", "no:cacheprovider", `--basetemp=${basetemp}`
      ], { cwd: REPO, encoding: "utf8", env: { ...process.env, PYTHONIOENCODING: "utf-8" } })
    : spawnSync(VITEST[0], [...VITEST.slice(1), spec.file, "-t", spec.testName, "--reporter=verbose"], {
        cwd: spec.cwd,
        encoding: "utf8",
        env: { ...process.env, NO_COLOR: "1", FORCE_COLOR: "0" }
      });
  const raw = `${run.stdout ?? ""}${run.stderr ?? ""}`;
  // 摘要带 ANSI ⇒ 解析前剥色码。
  const text = raw.replace(/\u001b\[[0-9;]*m/g, "");
  writeFileSync(logFile, text, "utf8");

  const restored = mutate(key, "restore");
  const after = sha256(file);
  const restoredOk = restored.status === 0 && after === before;

  const exitNonZero = run.status !== 0;
  const summary = spec.kind === "pytest"
    ? /\b\d+ failed\b/.test(text)
    : /Tests\s+\d+ failed/.test(text);
  const ranExpected = spec.kind === "pytest"
    ? text.includes(`FAILED ${spec.file}::${spec.testName}`)
    : text.includes(spec.testName) && /(×|FAILED|✗)/.test(text);
  const hasEvidence = text.includes(spec.errorContains);

  const judgements = [
    `① 退出码非 0: ${exitNonZero ? "OK" : "NG"}（${run.status}）`,
    `② 有效摘要: ${summary ? "OK" : "NG"}`,
    `③ 预期用例变红: ${ranExpected ? "OK" : "NG"}（${spec.testName}）`,
    `④ 失败原文含证据「${spec.errorContains}」: ${hasEvidence ? "OK" : "NG"}`,
    `⑤ 还原后 sha256 与变异前一致: ${restoredOk ? "OK" : "NG"}`,
    `⑥ 变异确实改了文件: ${mutatedSha !== before ? "OK" : "NG"}`
  ];
  const passed = exitNonZero && summary && ranExpected && hasEvidence && restoredOk && mutatedSha !== before;
  report.push({ seed: key, verdict: passed ? "RED(OK)" : "GREEN(NG)", judgements, log: logFile });
}

let failed = 0;
for (const item of report) {
  console.log(`\n=== ${item.seed}: ${item.verdict}`);
  if (item.judgements) for (const line of item.judgements) console.log(`    ${line}`);
  if (item.log) console.log(`    log: ${item.log}`);
  if (!/RED\(OK\)/.test(item.verdict)) failed += 1;
}
writeFileSync(path.join(OUT, "summary-a1e3-host-writes.json"), JSON.stringify(report, null, 2), "utf8");
process.exit(failed === 0 ? 0 : 1);
