// 反向变异脚本：暂停 / 继续的**步骤身份与终态**共享层（2026-10-21）
//
// 六个变异各钉一处互不相同的属性，全部挂在"这条机制真正管辖的那一行"上：
//
//   T1  恢复时不再读 `closedByFinish`（只留历史 `warning` 兜底）
//       → 被暂停强制关掉的步骤恢复后仍是终态，第一条进度被"已结束不许重激活"丢掉。
//   T2  收尾时**无条件**把没结束的步骤都记进 `closedByFinish`（不再只记 warning）
//       → 收尾判成 completed 的步骤恢复时被重新打开，迟到的"进行中"事件能拉回它。
//   T3  收尾不再接受调用方给的 `currentStatus`（退回按活动级 `outcome` 推）
//       → "带着待确认结果收尾"时这一步被写成 warning，明明做完了却显示没做完。
//   T4  `closedByFinish` 缺省时不再区分"老日志"（等价于永远按 warning 兜底）
//       → 自己报了 warning 结束的步骤（真实结论）恢复后被重开。
//   T5  hook 落盘时把 `closedByFinish` 换成空数组（假装没有步骤被收尾关掉）
//       → 落盘的那份丢了证据，恢复只能按 warning 猜。
//   T6  hook 收尾时不再把 `effectiveStatus` 交给共享层（第 7 参传 undefined）
//       → 与 T3 同一个判别点，但管的是**调用方有没有把真实终态递下去**。
//
// 用法：
//   node scripts/mutate_activity_resume.mjs                 # 打印基线 sha256
//   node scripts/mutate_activity_resume.mjs check           # 变异串定点闸
//   node scripts/mutate_activity_resume.mjs T1              # 应用
//   node scripts/mutate_activity_resume.mjs T1 restore      # 还原
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const TIMELINE = resolve(REPO, "packages", "core", "src", "activityTimeline.ts");
const HOOK = resolve(REPO, "apps", "excel-addin", "src", "hooks", "useActivityProgress.ts");

/** 应当由哪一组用例判红（驱动脚本按这个跑测试）。 */
const TARGETS = {
  T1: {
    file: TIMELINE,
    suite: "core",
    expect: "暂停顺手关掉的那一步，恢复后还能接收进度",
    seed: "  const interrupted = new Set(log.closedByFinish ?? []);",
    // [变异T1] 恢复时不再认"上一次收尾关掉的步骤"。
    mutant: "  const interrupted = new Set<string>();"
  },
  T2: {
    file: TIMELINE,
    suite: "core",
    expect: "收尾时这一步自己已做完",
    seed: '    if (status === "warning") closedByFinish.push(step.id as string);',
    // [变异T2] 无条件记录：completed/failed/cancelled 也被当成"被收尾关掉的"。
    mutant: "    closedByFinish.push(step.id as string);"
  },
  T3: {
    file: TIMELINE,
    suite: "core",
    expect: "收尾时这一步自己已做完",
    // 种子带上下一行：只写 `? terminalStatus` 会在本文件里命中两处（第 259 行也有），
    // 那样"还原"会改到无关分支。
    seed: [
      "      ? currentStatus ?? terminalStatus",
      '      : terminalStatus === "failed" || terminalStatus === "cancelled"'
    ].join(LF),
    // [变异T3] 步骤终态退回按活动级 outcome 推。
    mutant: [
      "      ? terminalStatus",
      '      : terminalStatus === "failed" || terminalStatus === "cancelled"'
    ].join(LF)
  },
  T4: {
    file: TIMELINE,
    suite: "core",
    expect: "自己报了 warning 结束的步骤，恢复后不重开",
    seed: "  const legacyLog = log.closedByFinish === undefined;",
    // [变异T4] 不区分"老日志没这个字段"，一律按 warning 兜底。
    mutant: "  const legacyLog = true;"
  },
  T5: {
    file: HOOK,
    suite: "hook",
    expect: "暂停收尾时那一步没有结论",
    seed: "          closedByFinish: finished.closedByFinish",
    // [变异T5] 落盘时把这份证据抹成空数组。
    mutant: "          closedByFinish: []"
  },
  T6: {
    file: HOOK,
    suite: "hook",
    expect: "带着待确认结果收尾时，这一步就是做完了",
    seed: [
      "        // 这一步自己的终态由调用方给（`finalStatus`），不按活动级 `outcome` 推：",
      '        // "带着待确认结果收尾"的活动中，这一步可能确实已经做完了。',
      "        effectiveStatus"
    ].join(LF),
    // [变异T6] 不再把调用方给的步骤终态递下去。
    mutant: [
      "        // [变异T6] 步骤终态退回按活动级 outcome 推",
      "        undefined"
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

/**
 * 写回必须**读回校验并重试**。
 *
 * Windows 上偶发"写进去了、读回来还是旧的"（索引器/防病毒短暂占句柄）。
 * 变异体留在工作区比变异失败更糟：后面的回归会以"红"的面目出现，
 * 而人以为那是产品缺陷。读回不等于预期就原样再写一次。
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
  return false;
}

const which = process.argv[2];
if (!which) {
  for (const file of [TIMELINE, HOOK]) {
    console.log(`BASELINE ${file} sha256=${sha256(file)}`);
  }
  process.exit(0);
}

/**
 * 前置状态闸：每个种子出现次数必须恰好 1、变异串未变异时必须出现 0 次。
 * 源码重写（重排注释、换行归一）会打散旧串，所以改完源码必跑 `check`；
 * 没跑 `check` 就等于"假装有证据"。
 */
if (which === "check") {
  let bad = 0;
  for (const [id, target] of Object.entries(TARGETS)) {
    const current = normalize(readFileSync(target.file, "utf8"));
    const seedHits = current.split(target.seed).length - 1;
    const mutantHits = current.split(target.mutant).length - 1;
    const ok = (seedHits === 1 && mutantHits === 0) || (seedHits === 0 && mutantHits === 1);
    if (!ok) {
      bad += 1;
      console.log(`MISS ${id} (${target.file}): SEED=${seedHits} MUTANT=${mutantHits}（期望 SEED=1,MUTANT=0）`);
    } else {
      console.log(`OK   ${id} SEED=${seedHits} MUTANT=${mutantHits}`);
    }
  }
  console.log(bad === 0 ? "check: OK" : `check: ${bad} 个变异串已失配`);
  process.exit(bad === 0 ? 0 : 3);
}

if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_activity_resume.mjs [check|T1..T6] [restore]");
  process.exit(2);
}

const target = TARGETS[which];
const original = readFileSync(target.file, "utf8");

if (process.argv[3] === "restore") {
  const text = normalize(original);
  const hits = text.split(target.mutant).length - 1;
  if (hits === 0) {
    console.log("ALREADY-CLEAN");
    process.exit(0);
  }
  if (hits !== 1) {
    // 不许盲替：命中多于一处说明串不唯一，"还原"会改到无关代码。
    console.error(`REFUSE: MUTANT 命中 ${hits} 处（期望 1）于 ${target.file}，拒绝还原`);
    process.exit(4);
  }
  if (!writeVerified(target.file, denormalize(text.replace(target.mutant, target.seed), original))) {
    process.exit(4);
  }
  console.log(`restore: OK ${which} → ${target.file} sha256=${sha256(target.file)}`);
  process.exit(0);
}

const occurrences = normalize(original).split(target.seed).length - 1;
if (occurrences !== 1) {
  console.log(`SKIP: SEED 出现 ${occurrences} 次（期望 1）于 ${target.file}`);
  process.exit(3);
}
if (!writeVerified(target.file, denormalize(normalize(original).replace(target.seed, target.mutant), original))) {
  process.exit(4);
}
console.log(`APPLIED ${which} → ${target.file} sha256=${sha256(target.file)}`);
