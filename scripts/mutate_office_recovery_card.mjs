// 反向变异脚本：**Excel 恢复卡两处状态错误**的判别力（2026-10-31，第二十六轮）
//
// 缺陷现场：来源变化被续跑正确拒绝后，重开窗格恢复卡却写"来源和目标已核对，可继续剩余动作"；
// 撤销成功后同一张卡同时出现"已撤销本次执行"与"可继续剩余动作"，且"继续剩余动作"按钮还能点。
// 本轮补的三个判别点各自钉住一件不同的事：
//
//   S1  共享层 `executionRecoveryResumeVerdict` 把 `unverified`（没复核来源）当成可继续
//       ⇒ core 用例「没复核过来源就不出 resume」必须红：没读过来源却说"可继续"。
//
//   S2  适配层 `officeRecoveryVerdictFor` 丢掉本机/账本撤销终态
//       ⇒ E2E「撤销成功后迟到的发现/重核不把终态覆盖回可继续」必须红：终态被覆盖回 resume。
//
//   S3  点击处理器 `undoOfficeRecoveredExecution` 不再同步记终态（ref 写被删）
//       ⇒ E2E 错误 #2 断言（撤销后同卡无"继续剩余动作"）必须红：卡面结论仍按旧账本算 resume。
//
// 四判据（不许只卡"退出码非 0"）：① 退出码非 0；② 预期用例红了；③ 失败原文带它自己那处证据；
// ④ 还原后逐字节 sha256 一致；④b 还原后复跑该用例绿。
//
// 用法（**一次一个变异**）：
//   node scripts/mutate_office_recovery_card.mjs             # 打印基线 sha256 + 是否干净
//   node scripts/mutate_office_recovery_card.mjs S1          # 应用（手工调试）
//   node scripts/mutate_office_recovery_card.mjs S1 restore  # 还原
//   node scripts/mutate_office_recovery_card.mjs S1 check    # 跑回归并按四判据判定
import { readFileSync, writeFileSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const CORE_VIEW = resolve(REPO, "packages", "core", "src", "executionRecoveryView.ts");
const APP = resolve(REPO, "apps", "excel-addin", "src", "App.tsx");

const CORE_TEST = "src/executionRecoveryView.test.ts";
const E2E = "src/excelWriteEndToEnd.test.tsx";

const TARGETS = {
  S1: {
    file: CORE_VIEW,
    runner: "vitest",
    app: "core",
    testFile: CORE_TEST,
    test: "没复核过来源就不出 resume，也不许说“来源已核对”",
    seed: '  if (source.state === "unverified") {',
    mutant:
      '  if (source.state === "unverified") {\n    return { kind: "resume", remaining, source: "matched" };  // S1（变异体，勿提交）：没复核也放行',
    evidence: "to be 'blocked'"
  },
  S2: {
    file: APP,
    runner: "vitest",
    app: "excel-addin",
    testFile: E2E,
    test: "撤销成功后迟到的发现/重核不把终态覆盖回可继续",
    seed: '    updateOfficeRecoveryMessage(\n      operationId,\n      "这次执行已撤销：本次写入已回到执行前，来源保持你现在的修改。",\n      officeRecoveryCardFor(state, OFFICE_RECOVERY_NO_EVIDENCE)\n    );',
    mutant:
      '    // S2（变异体，勿提交）：撤销成功后不重写卡面结论，verdict 仍是撤销前的 resume',
    evidence: 'data-recovery-action="resume"'
  },
  S3: {
    file: APP,
    runner: "vitest",
    app: "excel-addin",
    testFile: E2E,
    test: "中断 → 核对 → 改来源（续跑被拒）→ 重开且本机索引不在 → 撤销只回退已写入的目标，来源保持用户的修改",
    seed: "    officeUndoTerminalRef.current = done;",
    mutant:
      "    // S3（变异体，勿提交）：撤销成功不写同步终态，卡面结论退化成非 undone",
    // 注：S3 钉的是「撤销后同 tick 同步终态」这一件事（错误 #2 的卡面收敛），
    // **不是** ⑧ 的证据。⑧ 的守卫是点击处理器的「最新处置闸」，见 S5。
    evidence: 'data-recovery-kind="undone"'
  },
  // ⑥ 的守卫：现场已不在写后状态（`sceneChanged`）⇒ 阻断续跑，且撤销行一并收走。
  // 删掉 `sceneChanged` 标记后，verdict 仍是 blocked，但撤销行不再被收走 ⇒ 回归⑥
  // 「重开后撤销按钮为 null」必须红：撤销按钮又冒出来（点下去只会 fail-closed）。
  S4: {
    file: APP,
    runner: "vitest",
    app: "excel-addin",
    testFile: E2E,
    test: "回归⑥",
    seed: '        ? { kind: "blocked", detail: evidence.scene.detail, sceneChanged: true }',
    mutant:
      '        ? { kind: "blocked", detail: evidence.scene.detail }  // S4（变异体，勿提交）：现场已变也不标 sceneChanged',
    evidence: 'data-recovery-action="undo"'
  },
  // ⑧ 的守卫：点击「继续剩余动作」处理器里的「最新处置闸」。把 `if (terminal)`
  // 换成恒假后，旧回调不再被账本最新处置挡住 ⇒ 回归⑧必须红：claim 数增加、发生写入。
  S5: {
    file: APP,
    runner: "vitest",
    app: "excel-addin",
    testFile: E2E,
    test: "回归⑧",
    seed: '       * 已经撤销 ⇒ 一格都不写，也不重置账本游标（不重跑旧执行）。\n       */\n      const terminal = officeExecutionTerminalReason(state);\n      if (terminal) {',
    mutant:
      '       * 已经撤销 ⇒ 一格都不写，也不重置账本游标（不重跑旧执行）。\n       */\n      const terminal = officeExecutionTerminalReason(state);\n      if (false) {  // S5（变异体，勿提交）：最新处置闸被短路，旧回调不再被终态拒绝',
    evidence: "已经撤销过（账本已留处置痕）"
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
  for (let attempt = 1; attempt <= 5; attempt += 1) {
    writeFileSync(path, text, "utf8");
    if (normalize(readFileSync(path, "utf8")) === expected) return true;
  }
  console.error(`写回校验失败（5 次）：${path}`);
  return false;
}

const which = process.argv[2];
if (!which) {
  for (const [id, target] of Object.entries(TARGETS)) {
    const text = normalize(readFileSync(target.file, "utf8"));
    const dirty = text.includes(normalize(target.mutant));
    console.log(
      `BASELINE ${id} ${target.file} sha256=${sha256(target.file)} ${dirty ? "!! 变异体仍在工作区" : "clean"}`
    );
  }
  process.exit(0);
}
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_office_recovery_card.mjs [S1|S2|S3] [restore|check]");
  process.exit(2);
}

const target = TARGETS[which];
const original = readFileSync(target.file, "utf8");

function restoreQuietly() {
  try {
    const normalized = normalize(readFileSync(target.file, "utf8"));
    if (!normalized.includes(normalize(target.mutant))) return;
    writeVerified(
      target.file,
      denormalize(normalized.split(normalize(target.mutant)).join(normalize(target.seed)), original)
    );
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
  const restored = denormalize(
    normalize(original).split(normalize(target.mutant)).join(normalize(target.seed)),
    original
  );
  if (!writeVerified(target.file, restored)) process.exit(4);
  console.log(`restore: OK ${target.file} sha256=${sha256(target.file)}`);
  process.exit(0);
}

const occurrences = normalize(original).split(normalize(target.seed)).length - 1;
if (occurrences !== 1) {
  console.log(`SKIP: SEED 出现 ${occurrences} 次（期望 1）于 ${target.file}`);
  process.exit(3);
}

function runTest() {
  const escaped = target.test.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const vitest = resolve(REPO, "node_modules", "vitest", "vitest.mjs");
  const args = ["run", target.testFile, "-t", escaped, "--reporter=verbose"];
  // core 的 cwd 是 packages/core，Excel 是 apps/excel-addin。
  return spawnSync(process.execPath, [vitest, ...args], {
    cwd: resolve(REPO, target.app === "core" ? "packages/core" : "apps/excel-addin"),
    encoding: "utf8",
    timeout: 600_000,
    env: { ...process.env, NO_COLOR: "1", FORCE_COLOR: "0" }
  });
}
function clean(run) {
  return `${run.stdout ?? ""}\n${run.stderr ?? ""}`.replace(/\u001b\[[0-9;]*m/g, "");
}
function failedLines(output) {
  return output.split("\n").filter((line) => /\bFAIL(ED)?\b/.test(line));
}
function green(run, output) {
  return run.status === 0 && /passed/.test(output) && !/FAIL/.test(output);
}

if (process.argv[3] === "check") {
  const baseline = sha256(target.file);
  const mutated = denormalize(
    normalize(original).replace(normalize(target.seed), normalize(target.mutant)),
    original
  );
  if (!writeVerified(target.file, mutated)) process.exit(4);
  let run;
  try {
    run = runTest();
  } finally {
    if (!writeVerified(target.file, denormalize(normalize(original), original))) process.exit(4);
  }
  const output = clean(run);
  if (run.error) console.log(`spawn 失败：${run.error.message}`);
  const fails = failedLines(output);
  const criteria = [
    { id: "①退出码非 0", ok: run.status !== 0, detail: `status=${run.status}` },
    {
      id: "②预期用例红了",
      ok: fails.some((line) => line.includes(target.test.split("（")[0])) || output.includes(target.test),
      detail: fails[0]?.trim() ?? "（输出里没有任何失败行）"
    },
    { id: "③失败原文带自家证据", ok: output.includes(target.evidence), detail: target.evidence },
    { id: "④还原后逐字节一致", ok: sha256(target.file) === baseline, detail: `${baseline} → ${sha256(target.file)}` }
  ];
  for (const item of criteria) console.log(`${item.ok ? "OK  " : "FAIL"} ${item.id} :: ${item.detail}`);
  const after = runTest();
  const afterOutput = clean(after);
  const greenAfter = green(after, afterOutput);
  console.log(`${greenAfter ? "OK  " : "FAIL"} ④b还原后复跑绿 :: status=${after.status}`);
  if (!greenAfter) console.log(afterOutput.slice(-1500));
  const allOk = criteria.every((item) => item.ok) && greenAfter;
  console.log(allOk ? `\n${which}: 四判据全部成立` : `\n${which}: 判据不成立，别当成"已经验证过"`);
  process.exit(allOk ? 0 : 1);
}

const mutated = denormalize(
  normalize(original).replace(normalize(target.seed), normalize(target.mutant)),
  original
);
if (!writeVerified(target.file, mutated)) process.exit(4);
console.log(`APPLIED ${which} → ${target.file} sha256=${sha256(target.file)}`);
