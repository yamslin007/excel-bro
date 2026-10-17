// 反向变异脚本：源变化/取消/重授权时必须连带作废**已登记结果**（2026-10-10 D 要求1）
// 用法：node scripts/mutate_python_result_invalidation.mjs M1|M2 [restore]
// M1 = drop_python_state 不再作废已登记结果（回到"只清登记+绑定"的旧行为）
// M2 = close 也连带作废结果（把"结束会话"与"作废授权态"混为一谈）
//
// 铁律：脚本只做 apply/restore；测试由 bash 直调。
import { readFileSync, writeFileSync } from "node:fs";

const REPO = "E:\\codex\\excel-bro";
const FILE = `${REPO}\\server\\app\\task_bridge.py`;
const CRLF = "\r\n";
const NL = CRLF;

const TARGETS = {
  // M1：拆掉结果作废（唯一判别点：源变化后旧结果仍可解析）
  M1: {
    seed: [
      '    coordinator.release(task_id)',
      '    if not keep_results:',
      '        local_results.release(task_id)',
    ].join(NL),
    mutant: [
      '    coordinator.release(task_id)',
      '    # [变异M1] 不作废已登记结果',
    ].join(NL),
  },
  // M2：把 keep_results 开关拆掉（close 也作废结果 → 终端结果卡读不出来）
  M2: {
    seed: '        drop_python_state(task_id, keep_results=True)',
    mutant: '        drop_python_state(task_id)  # [变异M2]',
  },
};

const which = process.argv[2];
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_python_result_invalidation.mjs M1|M2 [restore]");
  process.exit(2);
}
const target = TARGETS[which];
const original = readFileSync(FILE, "utf8");

function restore() {
  const hit = original.split(target.mutant).length - 1;
  if (hit === 0) {
    console.log("ALREADY-CLEAN：" + FILE);
    return;
  }
  writeFileSync(FILE, original.split(target.mutant).join(target.seed), "utf8");
  console.log("RESTORED：" + FILE);
}

if (process.argv[3] === "restore") {
  restore();
  process.exit(0);
}

const normalized = original.replaceAll("\r\n", "\n");
const occurrences = normalized.split(target.seed.replaceAll("\r\n", "\n")).length - 1;
if (occurrences !== 1) {
  console.log(`SKIP: SEED 出现 ${occurrences} 次（期望 1）。`);
  process.exit(3);
}
writeFileSync(FILE, original.replace(target.seed, target.mutant), "utf8");
console.log(`APPLIED ${which} → ${FILE}`);
