// 反向变异脚本：Excel 恢复卡「只补记的核对」与「执行剩余动作」是不是真的分开（2026-10-29）
//
// 背景：WPS 早已把恢复结论分成六种、把「核对在途动作」（只补记回执）与「继续剩余动作」
// （只跑没跑过的动作）分成两个按钮；Excel 端只有一个「核对并继续」，点一下先补记回执、
// 紧接着 executePlan 续跑。本轮把两者拆开，并用四个变异钉住"拆开"这件事本身 ——
// 只改按钮文案、不改执行路径的"假拆分"，会被 S1 直接抓出来。
//
// 每个变异只打**一个**判别点，互不代替：
//
// S1 = 核对之后顺手把剩余动作也跑掉（把「核对」和「继续」重新合回一次点击）
//      ⇒「A1 三动作中途断开：核对只补记回执（零写入）…」必红。
//      判别点：**核对这一下有没有写单元格**。文案仍是"核对在途动作"，只有执行路径被合并，
//      所以它红必定是"合并执行"而不是"按钮改名"。
// S2 = 续跑忽略服务端进度，从第 1 个动作重来
//      ⇒「在途动作先核对为「已写入」：跳过该步，只写剩余步骤，不重复写」必红。
//      判别点：**续跑的起点**（S1 管"核对要不要跑剩下的"，S2 管"续跑从哪一格开始"）。
// S3 = 卡片不再把 `backfill` 算成本端能执行的动作
//      ⇒「全部动作已写入、只缺最终回执：只给「补记完成（不重写原表）」…」必红。
//      判别点：**只剩回执时给不给补记按钮**（拆开之后最容易漏掉的那一支）。
// S4 = 续跑前不再按最新读数复核结论（卡面说能继续就直接续）
//      ⇒「续跑不信任卡面：状态里出现没有写前快照的在途动作时立刻停下」必红。
//      判别点：**点击之后有没有重新判定**（S1–S3 都不碰这道闸）。
//
// 四判据（不许只卡"退出码非 0"）：
//   ① 退出码非 0；② 预期用例名出现在 FAIL 里；③ 失败原文带它**自己那处**证据；
//   ④ 还原后逐字节 sha256 与基线一致。
//
// 用法：
//   node scripts/mutate_office_recovery_split.mjs                 # 打印基线 sha256
//   node scripts/mutate_office_recovery_split.mjs S1              # 应用变异
//   node scripts/mutate_office_recovery_split.mjs S1 restore       # 还原
//   node scripts/mutate_office_recovery_split.mjs S1 check         # 跑回归并按四判据判定
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
const A1_TEST = "A1 三动作中途断开：核对只补记回执（零写入），继续才只跑剩余动作";

const TARGETS = {
  S1: {
    file: APP,
    seed: [
      "      const fresh = await getOfficeExecutionStatus(next.operationId);",
      "      state = fresh;"
    ].join(LF),
    mutant: [
      "      const fresh = await getOfficeExecutionStatus(next.operationId);",
      "      state = fresh;",
      "      // S1: 核对之后顺手把剩余动作也跑掉（变异体，勿提交）",
      "      if (officeRecoveryVerdictFor(fresh).kind === \"resume\") {",
      "        await resumeOfficeRemainingActions({ ...message, officeResume: officeRecoveryCardFor(fresh) });",
      "      }"
    ].join(LF),
    test: A1_TEST,
    evidence: "A3:B3"
  },
  S2: {
    file: APP,
    seed: "          startActionIndex: state.nextActionIndex,",
    mutant: "          startActionIndex: 0, // S2: 续跑从第一个动作重来（变异体，勿提交）",
    test: "在途动作先核对为「已写入」：跳过该步，只写剩余步骤，不重复写",
    evidence: "A1:B1"
  },
  S3: {
    file: APP,
    seed: [
      "const EXCEL_RECOVERY_ACTIONS: readonly ExecutionRecoveryAction[] = [",
      "  \"reconcile\",",
      "  \"resume\",",
      "  \"backfill\"",
      "];"
    ].join(LF),
    mutant: [
      "const EXCEL_RECOVERY_ACTIONS: readonly ExecutionRecoveryAction[] = [",
      "  \"reconcile\",",
      "  \"resume\" // S3: 只缺回执时不再给补记按钮（变异体，勿提交）",
      "];"
    ].join(LF),
    test: "全部动作已写入、只缺最终回执：只给「补记完成（不重写原表）」",
    evidence: "补记完成（不重写原表）"
  },
  S4: {
    file: APP,
    seed: [
      "      const verdict = officeRecoveryVerdictFor(state);",
      "      if (verdict.kind !== \"resume\" || !executionRecoveryActions(verdict).includes(\"resume\")) {"
    ].join(LF),
    mutant: [
      "      const verdict = officeRecoveryVerdictFor(state);",
      "      // S4: 卡面说能继续就直接续，不再按最新读数复核（变异体，勿提交）",
      "      if (false) {"
    ].join(LF),
    test: "续跑不信任卡面：状态里出现没有写前快照的在途动作时立刻停下，零写入",
    evidence: "在途动作缺少写前快照"
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
  console.log(`BASELINE ${APP} sha256=${sha256(APP)}`);
  process.exit(0);
}
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_office_recovery_split.mjs [S1|S2|S3|S4] [restore|check]");
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
