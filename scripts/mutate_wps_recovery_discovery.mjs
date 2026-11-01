// 反向变异脚本：「重新发现未完成执行」的接线是不是真的在被测（2026-10-27）
//
// 背景：真机现场里账本进度（`needs_reconciliation / nextActionIndex=1 /
// inFlightActionIndex=1`）与原表写入现场都在，**恢复入口没了** —— 因为浏览器
// `localStorage` 的单条索引是唯一入口，而完整退出 WPS 之后它不在了。
//
// 修法是把索引降级为缓存，由后端按**当前工作簿稳定身份**发现。于是这一轮的判据必须
// 能证明"拆掉那条接线，回归一定红"；否则 test 全绿只说明测试没在测它。
//
// 每个变异只打**一个**判别点，互不代替：
//
// D1 = 拆掉后端发现调用（`discoverWpsPendingExecutions` 直接不调）
//      ⇒ 「完整退出 WPS 丢掉浏览器索引与宿主缓存后，面板经后端发现找回在途执行…」必红。
//      判别点：**有没有后端发现这个动作**（后面所有判据都以它为前提）。
// D2 = 缓存归属不按身份（有本地索引就先拿它摆卡）
//      ⇒ 「打开另一个工作簿时：属于原簿的本地索引不在这里摆卡」必红。
//      判别点：**缓存属于哪个工作簿**（D1 不改这一条）。
// D3 = 多候选自动选第一条
//      ⇒ 「工作簿有多条未完成写入时只给选择入口…」必红。
//      判别点：**要不要用户选**（D1/D2 都不碰多候选）。
// D4 = 服务不可达当成"没有未完成任务"
//      ⇒ 「服务不可达时说"暂时读不到"，绝不说成"没有未完成任务"」必红。
//      判别点：**查不到 ≠ 没有**。
// D5 = 去掉迟到响应的代数闸
//      ⇒ 「旧查询迟到不覆盖新工作簿…」必红。
//      判别点：**旧响应能不能落到新现场**。
// D6 = 身份观察只剩挂载那一次（不装定时器与焦点/可见性监听）
//      ⇒ 「先打开另一个工作簿再切回原簿…」必红。
//      判别点：**换簿/元数据晚到之后会不会重跑发现**。
//
// 四判据（不许只卡"退出码非 0"）：
//   ① 退出码非 0；② 预期用例名出现在 FAIL 里；③ 失败原文带它**自己那处**证据
//   （`until()` 的说明，而不是随便一个断言）；④ 还原后逐字节 sha256 与基线一致。
//
// 用法：
//   node scripts/mutate_wps_recovery_discovery.mjs                  # 打印基线 sha256
//   node scripts/mutate_wps_recovery_discovery.mjs D1               # 应用变异
//   node scripts/mutate_wps_recovery_discovery.mjs D1 restore        # 还原
//   node scripts/mutate_wps_recovery_discovery.mjs D1 check          # 跑回归并按四判据判定
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const PANEL = resolve(REPO, "apps", "wps-addin", "src", "WpsAgentPanel.tsx");
const TEST_FILE = "src/WpsAgentPanel.test.tsx";

const TARGETS = {
  D1: {
    file: PANEL,
    seed: "    const outcome = await discoverWpsPendingExecutions(getWpsApplication());",
    mutant: "    const outcome = { kind: \"no_identity\" } as const; // D1: 后端发现被拆掉（变异体，勿提交）",
    test: "完整退出 WPS 丢掉浏览器索引与宿主缓存后，面板经后端发现找回在途执行，且零自动动作",
    evidence: "索引丢了也必须经后端发现重新摆出待核对入口"
  },
  D2: {
    file: PANEL,
    // 2026-10-28 重新定点：上一节把"身份为 null 也采纳索引"删掉后，这一行变成下面这样。
    // SEED 未命中时脚本会退 3，不会假装跑过 —— 改这一行必须同步改这里。
    seed: "    if (pending && identity !== null && pending.workbookIdentity === identity) {",
    mutant: "    if (pending) { // D2: 缓存不再校验归属（变异体，勿提交）",
    test: "打开另一个工作簿时：属于原簿的本地索引不在这里摆卡（发现只认当前身份）",
    evidence: "应当按当前簿的身份去发现"
  },
  D3: {
    file: PANEL,
    seed: [
      "    if (selection.kind === \"choose\") {",
      "      setWpsRecovery(null);",
      "      setWpsRecoveryChoices({ candidates: selection.candidates, truncated: selection.truncated });",
      "      setWpsDiscoveryFailure(null);",
      "      return;",
      "    }"
    ].join(LF),
    mutant: [
      "    if (selection.kind === \"choose\") {",
      "      // D3: 多候选直接采用排序出来的第一条（变异体，勿提交）",
      "      await adoptWpsRecoveryCandidate(selection.candidates[0], isCurrent);",
      "      return;",
      "    }"
    ].join(LF),
    test: "工作簿有多条未完成写入时只给选择入口：不自动挑最新、也不自动查状态或写入",
    evidence: "两条候选都必须摆出来给用户选"
  },
  D4: {
    file: PANEL,
    seed: [
      "    if (outcome.kind === \"unavailable\") {",
      "      keepCurrentCard();",
      "      setWpsDiscoveryFailure(outcome.reason ? `${outcome.detail}（${outcome.reason}）` : outcome.detail);",
      "      return;",
      "    }"
    ].join(LF),
    mutant: [
      "    if (outcome.kind === \"unavailable\") {",
      "      // D4: 查不到当成没有（变异体，勿提交）",
      "      keepCurrentCard();",
      "      setWpsDiscoveryFailure(null);",
      "      return;",
      "    }"
    ].join(LF),
    test: "服务不可达时说\"暂时读不到\"，绝不说成\"没有未完成任务\"",
    evidence: "查不到必须如实说读不到"
  },
  D5: {
    file: PANEL,
    seed: [
      "    const outcome = await discoverWpsPendingExecutions(getWpsApplication());",
      "    if (!isCurrent()) return;"
    ].join(LF),
    mutant: [
      "    const outcome = await discoverWpsPendingExecutions(getWpsApplication());",
      "    // D5: 迟到响应的代数闸被拆掉（变异体，勿提交）"
    ].join(LF),
    test: "旧查询迟到不覆盖新工作簿：先前那次发现的结果不许画到当前簿上",
    evidence: "换到新簿后迟到的旧响应不许画出恢复卡"
  },
  D6: {
    file: PANEL,
    seed: [
      "    const timer = window.setInterval(",
      "      read,",
      "      checkpointCapabilities.taskRecovery.workbookIdentityWatchIntervalMs",
      "    );",
      "    window.addEventListener(\"focus\", read);",
      "    document.addEventListener(\"visibilitychange\", read);"
    ].join(LF),
    mutant: [
      "    // D6: 身份观察只剩挂载那一次（变异体，勿提交）",
      "    const timer = window.setInterval(read, 2_147_483_647);"
    ].join(LF),
    test: "先打开另一个工作簿再切回原簿：靠身份观察重跑发现，仍能找到在途执行",
    evidence: "切回原簿后必须重新发现，而不是停在\"没有未完成任务\""
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

const which = process.argv[2];
if (!which) {
  console.log(`BASELINE ${PANEL} sha256=${sha256(PANEL)}`);
  process.exit(0);
}
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_wps_recovery_discovery.mjs [D1..D6] [restore|check]");
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
  // `-t` 是正则：用例名里的 `：`、引号、括号必须先转义，否则会匹配到别的用例。
  const pattern = target.test.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  let run;
  try {
    // 直接起 node + vitest.mjs：与仓库里其它变异脚本一致（`npx.cmd` 在部分 shell 下起不来）。
    run = spawnSync(
      process.execPath,
      [resolve(REPO, "node_modules", "vitest", "vitest.mjs"), "run", TEST_FILE, "-t", pattern, "--reporter=verbose"],
      {
        cwd: resolve(REPO, "apps", "wps-addin"),
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
    { id: "②预期用例红了", ok: output.split("\n").some((line) => line.includes("FAIL") && line.includes(target.test)),
      detail: target.test },
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
