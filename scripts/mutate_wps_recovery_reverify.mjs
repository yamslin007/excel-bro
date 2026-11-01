// 反向变异脚本：「授权变化后对同一条已选执行重新核验」是不是真的在被测（2026-10-28）
//
// 背景：真机现场 `WPS恢复验收.xlsx`，用户已选中 `wps-execution-…0f5bffdc`（在途、待核对），
// 误选了三张表 ⇒ 卡面正确报「需重新规划：工作簿、宿主版本或授权已变化」。用户回到数据范围
// 面板取消那张不该勾的表（三张 → 原两张）后，`authorizationKey` 确实变了、发现 effect 确实
// 重跑了，**卡面却一动不动** —— 因为 `discoverWpsRecovery` 在
// `current?.source === "selected"` 处直接 `return`。
//
// 那条保护的本意是"不许被别的候选改判"，被写成了"这条以后一律不重算"。修法是抽
// `reverifySelectedWpsRecovery`：同一条 operationId、同一份判定、只读、带代数闸。
//
// 于是本轮判据要能证明两件事，各打一个判别点：
//
// R1 = 回到"只保护、不重算"（selected 分支直接 return，不调重新核验）
//      ⇒ 「授权从三张表收回成原两张后，同一条已选执行重新核验并改判」必红。
//      判别点：**授权变了会不会重算这一条**（R2 不管这个）。
// R2 = 迟到响应无条件写回卡片（isCurrent 三道闸 + operationId/taskId 比对一次拆全）
//      ⇒ 「连续切换授权时迟到的核验结果不许覆盖新现场…」必红。
//      判别点：**旧响应能不能落到新现场**（R1 不动写回条件）。
//
// 为什么 R2 要一次拆三层：那三道闸防的是同一件事（"这次结论是否还属于当前现场"），
// 属于冗余层。只拆一层的话，另外两层会先把变异体挡住，用例照样绿 —— 那会得出
// "写回条件不重要"的错误结论。冗余层必须一次拆全。
//
// 四判据（不许只卡"退出码非 0"）：①退出码非 0；②预期用例名出现在 FAIL 里；
// ③失败原文带它**自己那处**证据（用例里那句 `until`/`expect` 的说明，不是随便一个断言）；
// ④还原后逐字节 sha256 与基线一致。
//
// 用法：
//   node scripts/mutate_wps_recovery_reverify.mjs                  # 打印基线 sha256
//   node scripts/mutate_wps_recovery_reverify.mjs R1               # 应用变异
//   node scripts/mutate_wps_recovery_reverify.mjs R1 restore        # 还原
//   node scripts/mutate_wps_recovery_reverify.mjs R1 check          # 跑回归并按四判据判定
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
  R1: {
    file: PANEL,
    // 只替换"重新核验"这一句：后面紧跟的 `return;` 原样保留 ⇒ 精确回到缺陷现场。
    seed: "        await reverifySelectedWpsRecovery(isCurrent);",
    mutant: "        // R1: 回到「只保护、不重算」（变异体，勿提交）",
    test: "授权从三张表收回成原两张后，同一条已选执行重新核验并改判",
    evidence: "收回第三张表后必须对同一条执行重新核验，而不是停在按旧范围算出的结论上"
  },
  R2: {
    file: PANEL,
    seed: [
      "    if (!isCurrent()) return;",
      "    if (!latest) {",
      "      setWpsDiscoveryFailure(PENDING_DISCOVERY_UNAVAILABLE_DETAIL);",
      "      return;",
      "    }",
      "    const { identity: nextIdentity, verdict } = await wpsRecoveryVerdictFor(latest);",
      "    if (!isCurrent()) return;",
      "    // 读结论期间用户可能已经点了核对/续跑，或者卡被换成另一条：两种情况都不写回。",
      "    if (recoveryActionLock.current) return;",
      "    setWpsRecovery((previous) =>",
      "      previous && previous.operationId === operationId && previous.taskId === taskId",
      "        ? { ...previous, state: latest, identity: nextIdentity, verdict, checking: false }",
      "        : previous",
      "    );"
    ].join(LF),
    mutant: [
      "    // R2: 迟到响应无条件覆盖当前卡片（变异体，勿提交）",
      "    if (!latest) {",
      "      setWpsDiscoveryFailure(PENDING_DISCOVERY_UNAVAILABLE_DETAIL);",
      "      return;",
      "    }",
      "    const { identity: nextIdentity, verdict } = await wpsRecoveryVerdictFor(latest);",
      "    setWpsRecovery((previous) => previous",
      "      ? { ...previous, state: latest, identity: nextIdentity, verdict, checking: false }",
      "      : previous);"
    ].join(LF),
    test: "连续切换授权时迟到的核验结果不许覆盖新现场；换工作簿后也不许再画回来",
    evidence: "先发那次的迟到结论不许翻案"
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
  console.error("用法: node scripts/mutate_wps_recovery_reverify.mjs [R1|R2] [restore|check]");
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
