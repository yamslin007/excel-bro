// 反向变异脚本：「写前恢复准备失败」能不能把这次执行清成「未执行」（2026-10-27 第四轮）
//
// 背景：线上旧账本缺列导致恢复包 PUT 500，写前落盘必然失败。失败后界面要能如实
// 收尾 —— 该清的清（真的一格都没写，账本也承认没动），不该清的绝不清。
// `noActionStarted` 是**正向证明**，由四个独立事实拼成：
//   ① 宿主正面证据（`restored`，且是真的读到了，不是"读不出来"）；
//   ② 账本游标还没越过第 0 步；
//   ③ 没有在途动作；
//   ④ 这次执行在账本侧已经没人还在跑（`ownerActive === false`）。
// 少任何一条都只能保留本机待办与恢复入口。四个变异各打一条，互不代替。
//
// 为什么不另加一条 `status !== "running"`：服务端 `_wps_state_payload` 用
// `status == "running" && 租约未过期` 算 `ownerActive`，二者在真实数据里同步 ——
// 写两遍只会多一层抓不到变异的装饰（M6 当时拆掉它，四条用例全绿，就是证据）。
//
// M6 = 去掉 `ownerActive === false` 判据（挂起失败后账本仍 running 也照清）
//      ⇒ 「写前恢复准备失败但挂起失败：账本仍 running，本端仍握着执行权：保留本机待办…」必须红。
//      判别点：**这次执行是否还在进行**（不是"写了没有"）。
// M7 = 去掉 `nextActionIndex === 0` 判据（账本已推进也照清）
//      ⇒ 「…账本已推进到第二步…」必须红。判别点：**账本游标**。
// M8 = 去掉 `inFlightActionIndex == null` 判据（第 0 步在途也照清）
//      ⇒ 「…第 0 步在途、回执还没回来…」必须红。判别点：**在途动作**。
// M9 = `classifyWpsWriteEvidence` 把"读失败"折成"未开始"
//      ⇒ 「宿主恢复点读不出来时不得当成「没写过」：保留本机待办」必须红。
//      判别点：**证据读不出来 ≠ 没写过**（M6–M8 都不碰这一条）。
//
// 四判据（不许只卡"退出码非 0"）：
//   ① 退出码非 0；② 红的那条用例**名字对得上**；③ 失败原文里带它自己的证据；
//   ④ 还原后文件逐字节 sha256 与基线一致。
//
// 用法：
//   node scripts/mutate_wps_recovery_clear_gate.mjs                 # 打印基线 sha256
//   node scripts/mutate_wps_recovery_clear_gate.mjs M6               # 应用变异
//   node scripts/mutate_wps_recovery_clear_gate.mjs M6 restore        # 还原
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const PANEL = resolve(REPO, "apps", "wps-addin", "src", "WpsAgentPanel.tsx");
const EVIDENCE = resolve(REPO, "apps", "wps-addin", "src", "wpsExecutionRecovery.ts");
const LEDGER = resolve(REPO, "server", "app", "task_ledger.py");

const TARGETS = {
  M6: {
    file: PANEL,
    seed: "        wpsLedgerState.inFlightActionIndex == null && wpsLedgerState.ownerActive === false;",
    mutant: "        wpsLedgerState.inFlightActionIndex == null;"
  },
  M7: {
    file: PANEL,
    seed: [
      "        wpsLedgerState?.operationId === operationId && wpsLedgerState.nextActionIndex === 0 &&",
      "        wpsLedgerState.inFlightActionIndex == null && wpsLedgerState.ownerActive === false;"
    ].join(LF),
    mutant: [
      "        wpsLedgerState?.operationId === operationId &&",
      "        wpsLedgerState.inFlightActionIndex == null && wpsLedgerState.ownerActive === false;"
    ].join(LF)
  },
  M8: {
    file: PANEL,
    seed: [
      "        wpsLedgerState?.operationId === operationId && wpsLedgerState.nextActionIndex === 0 &&",
      "        wpsLedgerState.inFlightActionIndex == null && wpsLedgerState.ownerActive === false;"
    ].join(LF),
    mutant: [
      "        wpsLedgerState?.operationId === operationId && wpsLedgerState.nextActionIndex === 0 &&",
      "        wpsLedgerState.ownerActive === false;"
    ].join(LF)
  },
  M9: {
    file: EVIDENCE,
    seed: '  if (input.readFailed) return { kind: "unknown" };',
    mutant:
      '  if (input.readFailed) return { kind: "not_started", executionId: input.executionId ?? "" };'
  },
  /*
   * M10 = 拆掉"schema 已是最新就直接返回"的早退（每次实例都重跑 journal_mode 与迁移）
   *       ⇒ 「test_restart_on_a_current_schema_never_waits_for_a_write_lock」必须红。
   *       判别点：**启动路径要不要写锁**。它和 M6–M9 的"清不清"是另一件事：
   *       那些管一次失败怎么收尾，这条管每次开进程会不会把别人的写事务卡住。
   */
  M10: {
    file: LEDGER,
    seed: ["        if _schema_is_current(path):", "            return"].join(LF),
    mutant: [
      "        # M10: 重启早退被拆掉（变异体，勿提交）",
      "        if False and _schema_is_current(path):",
      "            return"
    ].join(LF)
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
  for (const [, target] of Object.entries(TARGETS)) {
    console.log(`BASELINE ${target.file} sha256=${sha256(target.file)}`);
  }
  process.exit(0);
}
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_wps_recovery_clear_gate.mjs [M6|M7|M8|M9|M10] [restore]");
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
  console.error(`期望片段：${expected.slice(0, 80)}`);
  console.error(`实际片段：${last.slice(0, 80)}`);
  return false;
}

/*
 * 进程被打断也要还原：变异体泄漏进工作区，会让下一次回归以"红"的面目出现，
 * 而那与产品无关。SIGINT/TERM/HUP 与未捕获异常都装钩子。
 */
function restoreQuietly() {
  try {
    const current = normalize(readFileSync(target.file, "utf8"));
    if (!current.includes(target.mutant)) return;
    writeVerified(target.file, denormalize(current.split(target.mutant).join(target.seed), original));
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

const mutated = denormalize(normalize(original).replace(target.seed, target.mutant), original);
if (!writeVerified(target.file, mutated)) process.exit(4);
console.log(`APPLIED ${which} → ${target.file} sha256=${sha256(target.file)}`);
