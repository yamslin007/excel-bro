// 反向变异脚本：WPS 撤销/放弃后**跨重启终态识别**的判别力（2026-10-31，第三十一轮）
//
// 缺陷（本轮的根因）：WPS 撤销成功后 `undoWpsExecution` 只把宿主缓存里的
// `record.state` 改成 `restored`，账本 `status` / 游标都不动；`retireWpsRecoveryPackage`
// 删后端材料是 best-effort。跨重启后缓存没了、材料若删失败还在，水合会按账本游标把
// state 推回 `partial_write_recoverable`，续跑被重新开放。修复是把账本 disposition
// （`undoneAt` / `abandonedAt`）投影进 WPS 状态响应，并在**结论**与**执行器**两层各加
// 一道终态闸。两个判别点各钉一处：
//
//   W1  拆掉 `wpsReconcileVerdict` 的 disposition 终态判据
//       ⇒「撤销过 → undone」那条必须红：撤销后卡面重新判成"可继续/待核对"。
//
//   W2  拆掉 `wpsResumeUnavailableReason` 的 disposition 终态闸
//       ⇒「撤销后跨重启…仍挡住续跑」那条必须红：材料删失败 + 缓存清空后，执行器放行续跑。
//
//   W3  拆掉 `ensureWpsRecoveryMaterial` 水合的 disposition 终态闸
//       ⇒「撤销后跨重启：水合…撤销入口不重新开放」那条必须红：水合按游标推回
//       partial_write_recoverable，`latestWpsUndo` 重新交出撤销入口。
//
//   W4  拆掉 `prepareWpsResume` 对**已完成动作写后现场**的核对
//       ⇒「留痕失败 + 材料删失败 + 跨重启…」那条必须红。disposition 闸在留痕失败时
//       不生效，这条核对是执行入口自身的兜底：撤销后现场是写前值，账本登记的是写后
//       指纹，拆掉它就会继续往下走（由 `assertPriorWritesCurrent` 或真实写入暴露）。
//
// 四判据（不许只卡"退出码非 0"）：① 退出码非 0；② 预期用例名在 FAIL 里；
// ③ 失败原文带它自己那处证据；④ 还原后逐字节 sha256 一致（`check` 再复跑该用例必须绿）。
//
// 用法（**一次一个变异**）：
//   node scripts/mutate_wps_disposition_terminal.mjs               # 打印基线 sha256
//   node scripts/mutate_wps_disposition_terminal.mjs W1            # 应用（手工调试）
//   node scripts/mutate_wps_disposition_terminal.mjs W1 restore    # 还原
//   node scripts/mutate_wps_disposition_terminal.mjs W1 check      # 跑回归并按四判据判定
import { readFileSync, writeFileSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const RECONCILE = resolve(REPO, "apps", "wps-addin", "src", "wpsExecutionReconcile.ts");
const RESUME = resolve(REPO, "apps", "wps-addin", "src", "wpsExecutionResume.ts");
const PACKAGE = resolve(REPO, "apps", "wps-addin", "src", "wpsRecoveryPackage.ts");
const EXECUTOR = resolve(REPO, "apps", "wps-addin", "src", "wpsPlanExecutor.ts");

const TARGETS = {
  W1: {
    file: RECONCILE,
    app: "wps-addin",
    testFile: "src/wpsExecutionReconcile.test.ts",
    test: "撤销过（disposition.undoneAt）→ 终态 undone，不给任何按钮",
    seed: [
      '  if (input.state.disposition?.undoneAt) {',
      '    return { kind: "undone" };',
      '  }'
    ].join(LF),
    mutant: '  // W1（变异体，勿提交）：撤销后不再判终态，卡面重新开放续跑',
    evidence: "to be 'undone'"
  },
  W2: {
    file: RESUME,
    app: "wps-addin",
    testFile: "src/wpsExecutionResume.test.ts",
    test: "撤销后跨重启：宿主缓存清空、材料删失败仍在，disposition 终态仍挡住续跑",
    seed: '  if (state.disposition?.undoneAt) return "这次执行已经撤销过，不能再续跑剩余动作。";',
    mutant:
      '  // W2（变异体，勿提交）：撤销后跨重启不再挡续跑，材料删失败会被重新水合并放行',
    evidence: "已经撤销过"
  },
  W3: {
    file: PACKAGE,
    app: "wps-addin",
    testFile: "src/wpsExecutionResume.test.ts",
    test: "撤销后跨重启：水合读 disposition 终态，把记录推回 restored 而非 partial_write_recoverable，撤销入口不重新开放",
    seed: "state.disposition?.undoneAt || state.disposition?.abandonedAt",
    mutant:
      "false /* W3（变异体，勿提交）：水合不读 disposition，按游标推回 partial_write_recoverable，撤销入口重新开放 */",
    evidence: "expected 'partial_write_recoverable' to be 'restored'"
  },
  W4: {
    file: EXECUTOR,
    app: "wps-addin",
    testFile: "src/wpsExecutionResume.test.ts",
    test: "留痕失败 + 材料删失败 + 跨重启：旧材料水合后，撤销过的执行续跑与重复撤销都零写入",
    seed: [
      '  for (const [index, action] of prepared.actions.entries()) {',
      '    const descriptor = state.descriptor.actions[index];',
      '    if (action.previewFingerprint !== (index < state.nextActionIndex ? descriptor.expectedFingerprint : descriptor.previewFingerprint))',
      '      throw new Error("已完成或未完成目标已变化，拒绝续跑。");',
      '  }'
    ].join(LF),
    mutant:
      '  // W4（变异体，勿提交）：续跑不再核对已完成动作的写后现场，撤销过的执行会被重新写入',
    // 拆掉这一层后仍零写入（`assertPriorWritesCurrent` 兜住），但**越过了执行边界**：
    // 会真的去 claim 执行权。所以失败的断言是"claim 不该被调用"。
    evidence: "to not be called at all"
  },
  W5: {
    file: EXECUTOR,
    app: "wps-addin",
    testFile: "src/wpsExecutionResume.test.ts",
    test: "动作之间现场被改：后续动作写入前按已完成动作的写后现场复核拦住（零写入）",
    seed: [
      '  const assertPriorWritesCurrent = () => {',
      '    for (const prior of record.actions) {',
      '      if (prior.appliedRows === 0) continue;',
      '      const priorSheet = getStoredSheet(application, record, prior);',
      '      const priorRange = rangeFor(priorSheet, prior.range);',
      '      const matches = isFormatActionType(prior.type)',
      '        ? formatActionSafeToUndo(asFormatFields(prior), prior.afterFormatSnapshot, captureWpsFormatSnapshot(priorRange, prior.rowCount, prior.columnCount))',
      '        : readMatrices(priorRange, prior.rowCount, prior.columnCount).fingerprint === prior.afterFingerprint;',
      '      if (!matches) throw new Error("等待期间已执行区域发生外部变化，已停止后续动作并保留恢复点。");',
      '    }',
      '  };'
    ].join(LF),
    mutant:
      '  const assertPriorWritesCurrent = () => { /* W5（变异体，勿提交）：不再复核已完成动作的写后现场 */ };',
    // 这一层是"检查之后、写入之前"的唯一防线：拆掉后剩余动作会**真的写进原表**，
    // 所以失败原文里首先出现的是实际写入的地址。
    evidence: "Target!A3"
  }
};

function normalize(text) {
  return text.split(/\r\n/).join(LF);
}
function denormalize(text, original) {
  const native = original.includes("\r\n") ? "\r\n" : LF;
  return text.split(LF).join(native);
}
function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}
function writeVerified(path, text) {
  const expected = normalize(text);
  let last = "";
  for (let attempt = 1; attempt <= 5; attempt += 1) {
    writeFileSync(path, text, "utf8");
    last = normalize(readFileSync(path, "utf8"));
    if (last === expected) return true;
  }
  console.error(`写回校验失败（5 次）：${path}`);
  return false;
}

const which = process.argv[2];
if (!which) {
  for (const [id, target] of Object.entries(TARGETS)) {
    console.log(`BASELINE ${id} ${target.file} sha256=${sha256(target.file)}`);
  }
  process.exit(0);
}
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_wps_disposition_terminal.mjs [W1|W2|W3|W4] [restore|check]");
  process.exit(2);
}

const target = TARGETS[which];
const original = readFileSync(target.file, "utf8");

function restoreQuietly() {
  try {
    const normalized = normalize(readFileSync(target.file, "utf8"));
    if (!normalized.includes(normalize(target.mutant))) return;
    writeVerified(target.file, denormalize(normalized.split(normalize(target.mutant)).join(normalize(target.seed)), original));
    console.error(`[hook] 已还原 ${which} → ${target.file}`);
  } catch (error) {
    console.error(`[hook] 还原失败：${error instanceof Error ? error.message : String(error)}`);
  }
}
for (const signal of ["SIGINT", "SIGTERM", "SIGHUP"]) {
  process.on(signal, () => {
    restoreQuietly();
    process.exit(130);
  });
}
process.on("uncaughtException", (error) => {
  restoreQuietly();
  console.error(error);
  process.exit(1);
});

if (process.argv[3] === "restore") {
  const hits = normalize(original).split(normalize(target.mutant)).length - 1;
  if (hits === 0) {
    console.log("ALREADY-CLEAN");
    process.exit(0);
  }
  const restored = denormalize(normalize(original).split(normalize(target.mutant)).join(normalize(target.seed)), original);
  if (!writeVerified(target.file, restored)) process.exit(4);
  console.log(`restore: OK ${target.file} sha256=${sha256(target.file)}`);
  process.exit(0);
}

const occurrences = normalize(original).split(normalize(target.seed)).length - 1;
if (occurrences !== 1) {
  console.log(`SKIP: SEED 出现 ${occurrences} 次（期望 1）于 ${target.file}`);
  process.exit(3);
}

function runVitest(pattern) {
  const escaped = pattern.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return spawnSync(
    process.execPath,
    [resolve(REPO, "node_modules", "vitest", "vitest.mjs"), "run", target.testFile, "-t", escaped, "--reporter=verbose"],
    {
      cwd: resolve(REPO, "apps", target.app),
      encoding: "utf8",
      timeout: 600_000,
      env: { ...process.env, NO_COLOR: "1", FORCE_COLOR: "0" }
    }
  );
}
function clean(run) {
  return `${run.stdout ?? ""}\n${run.stderr ?? ""}`.replace(/\u001b\[[0-9;]*m/g, "");
}

if (process.argv[3] === "check") {
  const baseline = sha256(target.file);
  const mutated = denormalize(normalize(original).replace(normalize(target.seed), normalize(target.mutant)), original);
  if (!writeVerified(target.file, mutated)) process.exit(4);
  let run;
  try {
    run = runVitest(target.test);
  } finally {
    if (!writeVerified(target.file, denormalize(normalize(original), original))) process.exit(4);
  }
  const output = clean(run);
  if (run.error) console.log(`spawn 失败：${run.error.message}`);
  const failedLines = output.split("\n").filter((line) => line.includes("FAIL"));
  const criteria = [
    { id: "①退出码非 0", ok: run.status !== 0, detail: `status=${run.status}` },
    {
      id: "②预期用例红了",
      ok: failedLines.some((line) => line.includes(target.test)),
      detail: failedLines[0]?.trim() ?? "（输出里没有任何 FAIL 行）"
    },
    { id: "③失败原文带自家证据", ok: output.includes(target.evidence), detail: target.evidence },
    { id: "④还原后逐字节一致", ok: sha256(target.file) === baseline, detail: `${baseline} → ${sha256(target.file)}` }
  ];
  for (const item of criteria) console.log(`${item.ok ? "OK  " : "FAIL"} ${item.id} :: ${item.detail}`);
  const after = runVitest(target.test);
  const afterOutput = clean(after);
  const greenAfter = after.status === 0 && /passed/.test(afterOutput) && !/FAIL/.test(afterOutput);
  console.log(`${greenAfter ? "OK  " : "FAIL"} ④b还原后复跑绿 :: status=${after.status}`);
  if (!greenAfter) console.log(afterOutput.slice(-1500));
  const allOk = criteria.every((item) => item.ok) && greenAfter;
  console.log(allOk ? `\n${which}: 四判据全部成立` : `\n${which}: 判据不成立，别当成"已经验证过"`);
  process.exit(allOk ? 0 : 1);
}

const mutated = denormalize(normalize(original).replace(normalize(target.seed), normalize(target.mutant)), original);
if (!writeVerified(target.file, mutated)) process.exit(4);
console.log(`APPLIED ${which} → ${target.file} sha256=${sha256(target.file)}`);
