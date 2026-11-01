// 反向变异脚本：Excel「未完成执行」发现入口的三处判据是不是真的在被测（2026-10-29）
//
// 背景：Excel 端此前只有 `localStorage` 单条索引（`pending-office-execution`）这一个
// 入口，挂载效应只比身份、不比授权。补齐的是四件事：共享判据（`pendingResumeReadiness`）、
// 后端发现、多候选显式选择、用户确认范围后的"同一条重核"。这里逐个拆掉，逐个必须红。
//
// 每个变异只打**一个**判别点，互不代替：
//
// D1 = `recordConfirmedScope` 不再写"用户确认过的那一份"（只递增版本，不记值）
//      ⇒「单条候选自动摆卡；用户确认收窄授权后同一条执行被重新核验并如实阻断」必红。
//      判别点：**卡面会不会按用户的新授权变红**（版本号照样推进，重核照样跑，
//      只有"比对哪一份名单"这一处被拆掉 —— 所以它红必定是这一处）。
// D2 = 去掉迟到响应的代数闸（`isCurrent` 恒真）
//      ⇒「迟到的发现结果不许落盘」必红。
//      判别点：**旧响应能不能落到新现场**。
// D3 = 多候选自动采用排序出来的第一条
//      ⇒「多条候选只给选择入口」必红。
//      判别点：**要不要用户选**。
// D4 = 清空授权那条提交路径不再记"用户确认过的范围"（只commit、不记账）
//      ⇒「授权被撤销（清空并确认）后：卡面立刻阻断」必红。
//      判别点：**撤权之后界面认不认这件事**（D1 管的是"确认收窄"，D4 管的是"清空"；
//      两条提交路径不同，一个修好不代表另一个修好 —— 这正是本轮实测漏掉的那条）。
// （原 D5「续跑不再复核写前快照」已于 2026-10-29 第十轮**退役**：那段代码不存在了 ——
//  "在途动作 → 只核对" 成了独立结论，复核落在 `officeRecoveryVerdictFor` 与
//  `pendingResumeReadiness` 上，续跑侧只剩"按最新读数重新判定"这道闸，
//  由 `scripts/mutate_office_recovery_split.mjs` 的 S4 钉住。保留旧 SEED 只会 SEED 未命中，
//  那等于假装有证据。）
//
// 四判据（不许只卡"退出码非 0"）：
//   ① 退出码非 0；② 预期用例名出现在 FAIL 里；③ 失败原文带它**自己那处**证据；
//   ④ 还原后逐字节 sha256 与基线一致。
//
// 用法：
//   node scripts/mutate_office_recovery_discovery.mjs                 # 打印基线 sha256
//   node scripts/mutate_office_recovery_discovery.mjs D1              # 应用变异
//   node scripts/mutate_office_recovery_discovery.mjs D1 restore       # 还原
//   node scripts/mutate_office_recovery_discovery.mjs D1 check         # 跑回归并按四判据判定
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const APP = resolve(REPO, "apps", "excel-addin", "src", "App.tsx");
const TEST_FILE = "src/excelWriteEndToEnd.test.tsx";

const TARGETS = {
  D1: {
    file: APP,
    seed: [
      "  function recordConfirmedScope(names: readonly string[] | null): void {",
      "    officeScopeConfirmedRef.current = names ? [...names] : null;",
      "    setOfficeScopeConfirmedRevision((value) => value + 1);"
    ].join(LF),
    mutant: [
      "  function recordConfirmedScope(names: readonly string[] | null): void {",
      "    void names; // D1: 用户确认过的那一份不再记账（变异体，勿提交）",
      "    setOfficeScopeConfirmedRevision((value) => value + 1);"
    ].join(LF),
    test: "单条候选自动摆卡；用户确认收窄授权后同一条执行被重新核验并如实阻断",
    evidence: "当前授权里没有工作表"
  },
  D2: {
    file: APP,
    seed: [
      "    const isCurrent = () => generation === officeDiscoveryGenerationRef.current;",
      "    void discoverOfficeRecovery(officeRecoveryIdentityKey, isCurrent);"
    ].join(LF),
    mutant: [
      "    // D2: 代数闸被拆掉，迟到响应照样落盘（变异体，勿提交）",
      "    const isCurrent = () => generation >= 0;",
      "    void discoverOfficeRecovery(officeRecoveryIdentityKey, isCurrent);"
    ].join(LF),
    test: "迟到的发现结果不许落盘：授权变过之后旧响应既不摆卡也不换对象",
    evidence: "getOfficeExecutionStatus"
  },
  D3: {
    file: APP,
    seed: [
      "    if (selection.kind === \"choose\") {",
      "      setOfficeRecoveryChoices({ candidates: selection.candidates, truncated: selection.truncated });",
      "      setOfficeDiscoveryFailure(null);",
      "      return;",
      "    }"
    ].join(LF),
    mutant: [
      "    if (selection.kind === \"choose\") {",
      "      // D3: 多候选直接采用排序出来的第一条（变异体，勿提交）",
      "      await adoptOfficeRecoveryCandidate(selection.candidates[0], isCurrent, \"auto\");",
      "      return;",
      "    }"
    ].join(LF),
    test: "多条候选只给选择入口：不自动选、不查状态、不写任何单元格",
    evidence: "office-recovery-choices"
  },
  D4: {
    file: APP,
    seed: "        recordConfirmedScope(sourceMode === \"folder\" ? null : []);",
    mutant: "        // D4: 清空授权这条提交路径不再记账（变异体，勿提交）",
    test: "授权被撤销（清空并确认）后：卡面立刻阻断，且发现不会把表加回来",
    evidence: "当前授权里没有工作表"
  },

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

const which = process.argv[2];
if (!which) {
  console.log(`BASELINE ${APP} sha256=${sha256(APP)}`);
  process.exit(0);
}
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_office_recovery_discovery.mjs [D1|D2|D3|D4] [restore|check]");
  process.exit(2);
}

const target = TARGETS[which];
const original = readFileSync(target.file, "utf8");

/*
 * 写回必须读回校验并重试：Windows 上偶发"写进去了、读回来还是旧的"。
 * 变异体留在工作区比变异失败更糟 —— 后面的回归会以"红"的面目出现，而人以为那是产品缺陷。
 */
function writeVerified(path, text) {
  const expected = normalize(text);
  let last = "";
  for (let attempt = 1; attempt <= 5; attempt += 1) {
    writeFileSync(path, text, "utf8");
    last = normalize(readFileSync(path, "utf8"));
    if (last === expected) return true;
  }
  console.error(`写回校验失败（5 次）：${path}`);
  console.error(`实际片段：${last.slice(0, 120)}`);
  return false;
}

/*
 * 进程被打断也要还原：变异体泄漏进工作区，会让下一次回归以"红"的面目出现，
 * 而那与产品无关。SIGINT/TERM/HUP 与未捕获异常都装钩子。
 */
function restoreQuietly() {
  try {
    const normalized = normalize(readFileSync(target.file, "utf8"));
    if (!normalized.includes(target.mutant)) return;
    writeVerified(target.file, denormalize(normalized.split(target.mutant).join(target.seed), original));
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
  const hits = normalize(original).split(target.mutant).length - 1;
  if (hits === 0) {
    console.log("ALREADY-CLEAN");
    process.exit(0);
  }
  const restored = denormalize(normalize(original).split(target.mutant).join(target.seed), original);
  if (!writeVerified(target.file, restored)) process.exit(4);
  console.log(`restore: OK ${target.file} sha256=${sha256(target.file)}`);
  process.exit(0);
}

const occurrences = normalize(original).split(target.seed).length - 1;
if (occurrences !== 1) {
  // SEED 未命中 = 假装有证据。宁可退出 3，也不要"跑了一条不该红的用例然后说没事"。
  console.log(`SKIP: SEED 出现 ${occurrences} 次（期望 1）于 ${target.file}`);
  process.exit(3);
}

if (process.argv[3] === "check") {
  const baseline = sha256(target.file);
  const mutated = denormalize(normalize(original).replace(target.seed, target.mutant), original);
  if (!writeVerified(target.file, mutated)) process.exit(4);
  // `-t` 是正则：用例名里的 `：`、`「」`、括号必须先转义，否则会匹配到别的用例。
  const pattern = target.test.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  let run;
  try {
    run = spawnSync(
      process.execPath,
      [resolve(REPO, "node_modules", "vitest", "vitest.mjs"), "run", TEST_FILE, "-t", pattern, "--reporter=verbose"],
      {
        cwd: resolve(REPO, "apps", "excel-addin"),
        encoding: "utf8",
        timeout: 600_000,
        env: { ...process.env, NO_COLOR: "1", FORCE_COLOR: "0" }
      }
    );
  } finally {
    const restored = denormalize(normalize(original), original);
    if (!writeVerified(target.file, restored)) process.exit(4);
  }
  // vitest 摘要带 ANSI ⇒ 解析前剥色码。
  const output = `${run.stdout ?? ""}\n${run.stderr ?? ""}`.replace(/\u001b\[[0-9;]*m/g, "");
  if (run.error) console.log(`spawn 失败：${run.error.message}`);
  const failures = [
    { id: "①退出码非 0", ok: run.status !== 0, detail: `status=${run.status}` },
    {
      id: "②预期用例红了",
      ok: output.split("\n").some((line) => line.includes("FAIL") && line.includes(target.test)),
      detail: target.test
    },
    { id: "③失败原文带自家证据", ok: output.includes(target.evidence), detail: target.evidence },
    { id: "④还原后逐字节一致", ok: sha256(target.file) === baseline, detail: `${baseline} → ${sha256(target.file)}` }
  ];
  for (const item of failures) console.log(`${item.ok ? "OK  " : "FAIL"} ${item.id} :: ${item.detail}`);
  const allOk = failures.every((item) => item.ok);
  console.log(allOk ? `\n${which}: 四判据全部成立` : `\n${which}: 判据不成立，别当成"已经验证过"`);
  process.exit(allOk ? 0 : 1);
}

const mutated = denormalize(normalize(original).replace(target.seed, target.mutant), original);
if (!writeVerified(target.file, mutated)) process.exit(4);
console.log(`APPLIED ${which} → ${target.file} sha256=${sha256(target.file)}`);
