// 反向变异脚本：SDK 读表桥的**三处生命周期竞态**（2026-10-20 第三批）
//
// 每个变异只撤回**一处**修复，各自钉一条业务断言：
//
// R1 = 旧轮次退出时不再接续新一轮排队的工作
//      把 stale 分支里的"回循环顶部继续服务"删掉，退回无条件 return。表现是
//      "继续分析之后一动不动，必须再等一个进度事件才恢复"。
// R2 = 迟到的初始化照旧提交（不再核验轮次）
//      删掉 `await liveContext()` 之后的代数核验。表现是"取消之后桥又自己活了"，
//      以及"先发起的初始化后回来，把新轮次覆盖掉"。
// R3 = 读表返回后不再核验轮次，直接发回执
//      删掉 execute 与 receipt 之间那道 isStale。表现是旧轮次的读表结果以
//      "本轮结果"的身份进到新一轮的桥上。
//
// 用法：
//   node scripts/mutate_bridge_lifecycle_races.mjs               # 打印三处基线 sha256
//   node scripts/mutate_bridge_lifecycle_races.mjs R1            # 应用
//   node scripts/mutate_bridge_lifecycle_races.mjs R1 restore    # 还原
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const HOOK = resolve(REPO, "apps", "excel-addin", "src", "hooks", "useSdkToolBridge.ts");
const SHARED = resolve(REPO, "packages", "core", "src", "sdkToolBridge.ts");

const TARGETS = {
  R1: {
    file: HOOK,
    seed: [
      "        if (isStale()) {",
      "          // 这一轮已经被收尾或被新一轮取代：**不提交**（提交会覆盖新 begin 刚建好的",
      '          // 绑定）。但"不提交"不等于"没活儿"—— 新一轮排过队就接着为它服务。',
      "          if (queuedRef.current && !stateRef.current.cancelled && liveRef.current) {",
      "            continue;",
      "          }",
      "          return;",
      "        }"
    ].join(LF),
    mutant: [
      "        if (isStale()) {",
      "          // [变异R1] 旧轮次退出时不再接续新一轮排队的工作",
      "          return;",
      "        }"
    ].join(LF)
  },
  R2: {
    file: HOOK,
    seed: [
      "    if (epochRef.current !== epoch) {",
      "      return false;",
      "    }"
    ].join(LF),
    // 保留变量引用（避免 noUnusedLocals 掩盖成编译错误），只把核验本身关掉。
    mutant: [
      "    if (epoch < 0) {",
      "      // [变异R2] 迟到的初始化照旧提交（不再核验轮次）",
      "      return false;",
      "    }"
    ].join(LF)
  },
  R3: {
    file: SHARED,
    seed: [
      "    if (isStale()) {",
      "      return { state: next, polled, executed, outcomes, stale: true };",
      "    }",
      "    const response = await ports.receipt(binding, call, outcome);"
    ].join(LF),
    mutant: [
      "    // [变异R3] 读表返回后不再核验轮次，直接发回执",
      "    const response = await ports.receipt(binding, call, outcome);"
    ].join(LF)
  }
};

/** 目标文件自带换行（CRLF/LF）都要先归一化，否则 includes/split 会误判。 */
function normalize(text) {
  return text.split(/\r\n/).join(LF);
}
/** 写回必须还原**原文件自己的**换行，否则"逐字节一致"这条判据会自己失败。 */
function denormalize(text, original) {
  const native = original.includes("\r\n") ? "\r\n" : LF;
  return text.split(LF).join(native);
}

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

const which = process.argv[2];
if (!which) {
  for (const [id, target] of Object.entries(TARGETS)) {
    console.log(`BASELINE ${id} ${target.file} sha256=${sha256(target.file)}`);
  }
  process.exit(0);
}
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_bridge_lifecycle_races.mjs [R1|R2|R3] [restore]");
  process.exit(2);
}

const target = TARGETS[which];
const original = readFileSync(target.file, "utf8");

if (process.argv[3] === "restore") {
  const hits = normalize(original).split(target.mutant).length - 1;
  if (hits === 0) {
    console.log("ALREADY-CLEAN");
    process.exit(0);
  }
  writeFileSync(
    target.file,
    denormalize(normalize(original).split(target.mutant).join(target.seed), original),
    "utf8"
  );
  console.log(`restore: OK ${target.file}`);
  process.exit(0);
}

const occurrences = normalize(original).split(target.seed).length - 1;
if (occurrences !== 1) {
  console.log(`SKIP: SEED 出现 ${occurrences} 次（期望 1）于 ${target.file}`);
  process.exit(3);
}

writeFileSync(
  target.file,
  denormalize(normalize(original).replace(target.seed, target.mutant), original),
  "utf8"
);
console.log(`APPLIED ${which} → ${target.file}`);
