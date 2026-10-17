// 反向变异脚本：WPS 并排审阅副本「只承载**写入目标**那一张表」（2026-10-10）
// 用法：node scripts/mutate_wps_write_sheet.mjs M1|M2|M3 [restore]
// M1 = 副本承载表退回"授权表第一张"（多表授权时会复制错表）
// M2 = 去掉跨表写入拒绝（副本预览无法自洽的跨表计划也被放行）
// M3 = 退回旧的"授权张数必须为 1"判据（两表关联汇总的计划被整片挡住）
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const FILE = resolve(REPO, "apps", "wps-addin", "src", "wpsNativeComparison.ts");
const LF = "\n";

const TARGETS = {
  M1: {
    seed: [
      "  const name = prepared.actions[0].sheetName;",
      "  const authorized = prepared.authorization.sheets.find((sheet) => sheet.name === name);"
    ].join(LF),
    mutant: [
      "  const name = prepared.actions[0].sheetName;",
      "  const authorized = prepared.authorization.sheets[0]; // [变异M1] 复制授权表第一张，而不是写入目标表"
    ].join(LF)
  },
  M2: {
    seed: [
      "  if (written.size !== 1) {",
      '    throw new Error("原生并排审阅当前只支持单张工作表上的写入计划；跨表写入不会复制整本工作簿。");',
      "  }"
    ].join(LF),
    mutant: "  // [变异M2] 不再拒绝跨表写入"
  },
  M3: {
    seed: "  const writeSheet = writeSheetOf(prepared);",
    mutant: [
      '  if (prepared.authorization.sheets.length !== 1) throw new Error("原生并排审阅当前只支持一张已授权工作表；跨表计划不会复制整本工作簿。");',
      "  const writeSheet = writeSheetOf(prepared); // [变异M3] 退回“授权张数必须为 1”的判据"
    ].join(LF)
  }
};

/** 目标文件自带换行（CRLF/LF）都要先归一化，否则 includes/split 会误报 ALREADY-CLEAN。 */
function normalize(text) {
  return text.split(/\r\n/).join(LF);
}
/** 写回必须还原**原文件自己的**换行，否则"逐字节一致"这条判据会自己失败。 */
function denormalize(text, original) {
  const native = original.includes("\r\n") ? "\r\n" : LF;
  return text.split(LF).join(native);
}

const which = process.argv[2];
const original = readFileSync(FILE, "utf8");
if (!which) {
  console.log(`BASELINE ${FILE} sha256=${createHash("sha256").update(original).digest("hex")}`);
  process.exit(0);
}
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_wps_write_sheet.mjs M1|M2|M3 [restore]");
  process.exit(2);
}
const target = TARGETS[which];

if (process.argv[3] === "restore") {
  const hits = normalize(original).split(target.mutant).length - 1;
  if (hits === 0) {
    console.log("ALREADY-CLEAN");
    process.exit(0);
  }
  writeFileSync(FILE, denormalize(normalize(original).split(target.mutant).join(target.seed), original), "utf8");
  console.log(`S1 restore: OK ${FILE}`);
  process.exit(0);
}

const occurrences = normalize(original).split(target.seed).length - 1;
if (occurrences !== 1) {
  console.log(`SKIP: SEED 出现 ${occurrences} 次（期望 1）。`);
  process.exit(3);
}

writeFileSync(
  FILE,
  denormalize(normalize(original).replace(target.seed, target.mutant), original),
  "utf8"
);
console.log(`APPLIED M${which.slice(1)} → ${FILE}`);
