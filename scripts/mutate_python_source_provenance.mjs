// 反向变异脚本：「宿主成功读过的范围」记账（2026-10-10 D 阶段 要求1）
// 用法：node scripts/mutate_python_source_provenance.mjs M1|M2|M3|M4
// M1 = Excel prepare_python_data 分支不记账
// M2 = WPS prepare_python_data 分支不记账
// M3 = WPS query_table 分支不记账（拆掉新增的那一层）
// M4 = Excel query_table 分支不记账（回归：既有断言仍须有判别力）
//
// 铁律：脚本只做 apply/restore；测试由 bash 直调（沙箱里受管 node 再 spawn 受管 node 会 EBUSY）。
import { readFileSync, writeFileSync } from "node:fs";

const REPO = "E:\\codex\\excel-bro";
const EXCEL = `${REPO}\\apps\\excel-addin\\src\\sdkBridgeHost.ts`;
const WPS = `${REPO}\\apps\\wps-addin\\src\\wpsSdkBridgeHost.ts`;
const LF = "\n";
const CRLF = "\r\n";

const TARGETS = {
  M1: {
    file: EXCEL,
    nl: LF,
    seed: [
      '          recordTrustedToolReads(',
      '            binding.taskId,',
      '            blocks.map((block) => sources[block.sourceIndex]?.sheet ?? "").filter(Boolean)',
      '          );',
    ].join(LF),
    mutant: '          // [变异M1] prepare_python_data 不记账',
  },
  M2: {
    file: WPS,
    nl: LF,
    seed: [
      '          recordTrustedToolReads(',
      '            binding.taskId,',
      '            blocks.map((block) => sources[block.sourceIndex]?.sheet ?? "").filter(Boolean)',
      '          );',
    ].join(LF),
    mutant: '          // [变异M2] WPS prepare_python_data 不记账',
  },
  M3: {
    file: WPS,
    nl: LF,
    seed: [
      '        recordTrustedToolReads(',
      '          binding.taskId,',
      '          sheets.map((sheet) => sheet.name)',
      '        );',
    ].join(LF),
    mutant: '        // [变异M3] WPS query_table 不记账',
  },
  M4: {
    file: EXCEL,
    nl: LF,
    seed: '        recordTrustedToolReads(binding.taskId, result.sourceSheets ?? []);',
    mutant: '        // [变异M4] Excel query_table 不记账',
  },
};

const which = process.argv[2];
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_python_source_provenance.mjs M1|M2|M3|M4");
  process.exit(2);
}
const target = TARGETS[which];
const original = readFileSync(target.file, "utf8");

const touched = new Set();
function restoreAll() {
  if (touched.has(target.file)) {
    writeFileSync(target.file, original, "utf8");
    touched.clear();
  }
}
for (const sig of ["SIGINT", "SIGTERM", "SIGHUP"]) {
  process.on(sig, () => { restoreAll(); process.exit(130); });
}
process.on("uncaughtException", (err) => { restoreAll(); console.error("uncaughtException（已还原）:", err); process.exit(1); });

const mode = process.argv[3] === "restore" ? "restore" : "apply";
if (mode === "restore") {
  /*
   * 反向替换还原：不能用"本次进程读到的文件"当原稿 —— restore 是独立进程，
   * 那时读到的已经是变异后的内容，写回等于把变异固化（曾导致 M2 残留叠加 M3）。
   * 只认 mutant 文本本身，把它换回 seed；找不到说明本来就干净。
   */
  const mutantText = target.nl === "\r\n" ? target.mutant.replaceAll("\n", "\r\n") : target.mutant;
  const seedText = target.nl === "\r\n" ? target.seed.replaceAll("\n", "\r\n") : target.seed;
  const hit = original.split(mutantText).length - 1;
  if (hit === 0) {
    console.log("ALREADY-CLEAN：" + target.file);
    process.exit(0);
  }
  writeFileSync(target.file, original.split(mutantText).join(seedText), "utf8");
  console.log("RESTORED：" + target.file);
  process.exit(0);
}

const normalized = original.replaceAll("\r\n", "\n");
// 文本按目标文件的换行符归一化后再比对（本仓混合 CRLF/LF）。
const seed = target.nl === "\r\n" ? target.seed.replaceAll("\n", "\r\n") : target.seed;
const mutant = target.nl === "\r\n" ? target.mutant.replaceAll("\n", "\r\n") : target.mutant;

const occurrences = normalized.split(target.seed).length - 1;
if (occurrences !== 1) {
  console.log(`SKIP: SEED 出现 ${occurrences} 次（期望 1）。`);
  process.exit(3);
}

touched.add(target.file);
writeFileSync(target.file, original.replace(seed, mutant), "utf8");
console.log(`APPLIED ${which} → ${target.file}`);
