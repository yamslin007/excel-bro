// 反向变异脚本：「发送即送达」与「停下来先核对时把请求还回输入框」（2026-10-16）
//
// 背景：Excel 端原先要把「理解消息」这个模型往返（外加操作目标核对、授权表结构读取）
// 全部跑完，才在 `sendMessage` 里 append 用户气泡并 `setPrompt("")`。期间输入框里
// 还留着那句话，于是运行中的发送键停在「打断并补充」（↑）——用户再按一次 Enter 就
// 拿着**同一句**去转向重跑；对话里也看不到自己刚发的话。
// 修法：受理那一刻（第一个 await 之前）就送达 + 清空；并补上「停下来让用户先调整」
// 的几条出口把请求放回输入框；用 `targetConfirmedRetry` 标记目标确认后的回调重入，
// 避免同一句送两遍（与 WPS 的 `targetConfirmedRetry` 同义）。
//
// 用法：node scripts/mutate_message_delivery_on_send.mjs [M1|M2|M3|M4] [restore]
// 不带参数：只打印目标文件的基线 sha256。
// M1 = 受理时不再清空输入框
// M2 = 受理时不再落用户气泡（只清输入框）
// M3 = 去掉 `targetConfirmedRetry` 判据（回调重入会再送一遍）
// M4 = 目标核对出口不再把请求放回输入框
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";
const APP = resolve(REPO, "apps", "excel-addin", "src", "App.tsx");

const TARGETS = {
  M1: {
    file: APP,
    seed: [
      "      });",
      '      setPrompt("");',
      "    }",
      "    // 转向重跑：把被打断那句的意图并进这句，让 checkIntent 收到完整承接式需求。"
    ].join(LF),
    mutant: [
      "      });",
      "      // [变异M1] 不再在受理时清空输入框",
      "    }",
      "    // 转向重跑：把被打断那句的意图并进这句，让 checkIntent 收到完整承接式需求。"
    ].join(LF)
  },
  M2: {
    file: APP,
    seed: [
      "    if (!answeringClarification && options?.targetConfirmedRetry !== true) {",
      "      appendMessage({",
      '        role: "user",',
      '        text: enteredText || "请结合附件图片分析当前工作簿。",',
      "        attachmentNames: pendingImages.map((image) => image.name)",
      "      });",
      '      setPrompt("");',
      "    }"
    ].join(LF),
    mutant: [
      "    if (!answeringClarification && options?.targetConfirmedRetry !== true) {",
      "      // [变异M2] 气泡不再在受理时落盘（只清输入框）",
      '      setPrompt("");',
      "    }"
    ].join(LF)
  },
  M3: {
    file: APP,
    seed: "    if (!answeringClarification && options?.targetConfirmedRetry !== true) {",
    mutant:
      "    if (!answeringClarification) { // [变异M3] 去掉“这句已经送过”的判据"
  },
  M4: {
    file: APP,
    seed: [
      "    if (!(await targetGuard.checkForAnalysis(() => retrySendRef.current(intentText, understanding)))) {",
      "      keepRequestInComposer();"
    ].join(LF),
    mutant: [
      "    if (!(await targetGuard.checkForAnalysis(() => retrySendRef.current(intentText, understanding)))) {",
      "      // [变异M4] 停下来时不再把请求放回输入框"
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
  console.error(
    "用法: node scripts/mutate_message_delivery_on_send.mjs [M1|M2|M3|M4] [restore]"
  );
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
