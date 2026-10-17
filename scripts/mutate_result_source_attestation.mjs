// 反向变异脚本：结果计划来源依赖「服务端声明 ∩ 本端记账」（2026-10-10 D 要求2）
// 用法：node scripts/mutate_result_source_attestation.mjs M1|M2|M3 [restore]
// M1 = 退回单侧记账（等于 attestSourceDependency）→ 虚报"读过但没用到的表"
// M2 = 直接采信服务端声明（本端无凭据也放行）
// M3 = 交集为空时留空数组（把"依赖不确定"写成"依赖为空"）
import { readFileSync, writeFileSync } from "node:fs";

const REPO = "E:\\codex\\excel-bro";
const FILE = `${REPO}\\packages\\core\\src\\planSourceDependency.ts`;
const LF = "\n";

const TARGETS = {
  M1: {
    seed: [
      '  const trusted =',
      '    declared.length === 0',
      '      ? read',
      '      : read.filter((name) =>',
      '          declared.some((item) => item.toLocaleLowerCase() === name.toLocaleLowerCase())',
      '        );',
    ].join(LF),
    mutant: '  const trusted = read; // [变异M1] 只认本端记账并集',
  },
  M2: {
    seed: '  const trusted =\n    declared.length === 0\n      ? read\n      : read.filter((name) =>\n          declared.some((item) => item.toLocaleLowerCase() === name.toLocaleLowerCase())\n        );',
    mutant: '  const trusted = declared.length === 0 ? read : declared; // [变异M2] 采信服务端声明',
  },
  M3: {
    seed: [
      '  const next = { ...plan };',
      '  if (trusted.length === 0) {',
      '    delete next.sourceDependencySheets;',
      '    return next;',
      '  }',
      '  next.sourceDependencySheets = trusted;',
      '  return next;',
      '}',
    ].join(LF),
    mutant: [
      '  const next = { ...plan };',
      '  // [变异M3] 交集为空也照写空数组',
      '  next.sourceDependencySheets = trusted;',
      '  return next;',
      '}',
    ].join(LF),
  },
};

const which = process.argv[2];
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_result_source_attestation.mjs M1|M2|M3 [restore]");
  process.exit(2);
}
const target = TARGETS[which];
const original = readFileSync(FILE, "utf8");

if (process.argv[3] === "restore") {
  const hit = original.split(target.mutant).length - 1;
  if (hit === 0) { console.log("ALREADY-CLEAN"); process.exit(0); }
  writeFileSync(FILE, original.split(target.mutant).join(target.seed), "utf8");
  console.log("RESTORED：" + FILE);
  process.exit(0);
}

const occurrences = original.split(target.seed).length - 1;
if (occurrences !== 1) {
  console.log(`SKIP: SEED 出现 ${occurrences} 次（期望 1）。`);
  process.exit(3);
}

writeFileSync(FILE, original.replace(target.seed, target.mutant), "utf8");
console.log(`APPLIED ${which} → ${FILE}`);
