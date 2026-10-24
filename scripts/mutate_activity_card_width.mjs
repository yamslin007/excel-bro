// 反向变异脚本：执行过程卡与结果展示卡**同宽**的契约（2026-10-21）
//
// W1 = 把运行态卡的宽度上限放回来（`width: 100%` → `min(420px, calc(100% - 28px))`）
//      「运行态卡不设宽度上限」那条必须红：它同时钉着"显式铺满"和"没有 px 上限/内缩"。
// W2 = 把已完成态卡的水平内边距改掉（12px → 20px）
//      「水平内边距一致」那条必须红。这条钉的是**另一个**属性（左右文字边线），
//      所以 W1 不该动它、W2 也不该动 W1 —— 一个变异只允许一个判别点。
//
// 用法：
//   node scripts/mutate_activity_card_width.mjs              # 打印基线 sha256
//   node scripts/mutate_activity_card_width.mjs W1           # 应用
//   node scripts/mutate_activity_card_width.mjs W1 restore   # 还原
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const STYLES = resolve(REPO, "apps", "excel-addin", "src", "styles.css");

const TARGETS = {
  W1: {
    file: STYLES,
    seed: [
      ".activity-card {",
      "  box-sizing: border-box;",
      "  width: 100%;",
      "  min-width: 0;"
    ].join(LF),
    mutant: [
      ".activity-card {",
      "  box-sizing: border-box;",
      "  width: min(420px, calc(100% - 28px));",
      "  min-width: 0;"
    ].join(LF)
  },
  W2: {
    file: STYLES,
    seed: [".activity-log-card {", "  margin-top: 8px;", "  padding: 8px 12px;"].join(LF),
    mutant: [".activity-log-card {", "  margin-top: 8px;", "  padding: 8px 20px;"].join(LF)
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
  for (const [id, target] of Object.entries(TARGETS)) {
    console.log(`BASELINE ${id} ${target.file} sha256=${sha256(target.file)}`);
  }
  process.exit(0);
}
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_activity_card_width.mjs [W1|W2] [restore]");
  process.exit(2);
}

const target = TARGETS[which];
const original = readFileSync(target.file, "utf8");

/*
 * 写回必须读回校验并重试。
 *
 * Windows 上偶发"写进去了、读回来还是旧的"（杀毒/索引器短暂占句柄），
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
  console.error(`期望片段：${expected.slice(0, 80)}`);
  console.error(`实际片段：${last.slice(0, 80)}`);
  return false;
}

if (process.argv[3] === "restore") {
  const hits = normalize(original).split(target.mutant).length - 1;
  if (hits === 0) {
    console.log("ALREADY-CLEAN");
    process.exit(0);
  }
  const restored = denormalize(
    normalize(original).split(target.mutant).join(target.seed),
    original
  );
  if (!writeVerified(target.file, restored)) process.exit(4);
  console.log(`restore: OK ${target.file} sha256=${sha256(target.file)}`);
  process.exit(0);
}

const occurrences = normalize(original).split(target.seed).length - 1;
if (occurrences !== 1) {
  console.log(`SKIP: SEED 出现 ${occurrences} 次（期望 1）于 ${target.file}`);
  process.exit(3);
}

const mutated = denormalize(
  normalize(original).replace(target.seed, target.mutant),
  original
);
if (!writeVerified(target.file, mutated)) process.exit(4);
console.log(`APPLIED ${which} → ${target.file} sha256=${sha256(target.file)}`);
