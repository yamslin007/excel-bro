// 反向变异脚本：A0 补充修复的**客户端三处**（2026-10-20）
//
// 三处对应三条真实缺口，各自只有一个判别点：
//
// C1 = WPS 不采集必要来源指纹
//      把 `sourceRequired: status.continuation?.sourceVerificationRequired !== false`
//      写成 `false`。共享层于是不读工作簿，交上去的是 `not_collected`；服务端只能
//      回 SOURCE_UNVERIFIED，而界面上只表现为"点了继续没反应"。
// C2 = 继续时不重新初始化桥
//      把 `ensureLive` 的 `liveRef.current &&` 去掉 —— 也就是"上一轮留下过绑定就
//      当桥还活着"。`end()` 之后绑定仍在（结果卡还要按引用取值），于是续跑把读表
//      指针交给一条已收尾的桥，谁也不再轮询。
// C3 = 不转发 bridgeTool
//      进度映射只留文案字段，把指针整个丢掉。续跑时模型发出的读表请求永远到不了
//      本端，服务端一路等到超时。
//
// 用法：
//   node scripts/mutate_continuation_bridge.mjs                # 打印三处基线 sha256
//   node scripts/mutate_continuation_bridge.mjs C1             # 应用
//   node scripts/mutate_continuation_bridge.mjs C1 restore     # 还原
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const WPS_PANEL = resolve(REPO, "apps", "wps-addin", "src", "WpsAgentPanel.tsx");
const WPS_AGENT = resolve(REPO, "apps", "wps-addin", "src", "wpsReadonlyAgent.ts");
const EXCEL_BRIDGE = resolve(
  REPO,
  "apps",
  "excel-addin",
  "src",
  "hooks",
  "useSdkToolBridge.ts"
);

const TARGETS = {
  C1: {
    file: WPS_PANEL,
    seed: [
      '        // 服务端从检查点判定；缺省按"需要核实"处理。',
      "        sourceRequired: status.continuation?.sourceVerificationRequired !== false,"
    ].join(LF),
    mutant: [
      "        // [变异C1] 不再追问服务端要不要核实来源，直接当成无需采集",
      "        sourceRequired: false,"
    ].join(LF)
  },
  C2: {
    file: EXCEL_BRIDGE,
    seed: [
      "      if (",
      "        liveRef.current &&",
      "        existing !== null &&",
      "        existing.taskId === options.taskId",
      "      ) {"
    ].join(LF),
    mutant: [
      "      if (",
      "        // [变异C2] 只要上一轮留下过绑定就当桥还活着",
      "        existing !== null &&",
      "        existing.taskId === options.taskId",
      "      ) {"
    ].join(LF)
  },
  C3: {
    file: WPS_AGENT,
    seed: "    bridgeTool: event.bridgeTool ?? null,",
    mutant: "    bridgeTool: null, // [变异C3] 不再转发读表指针"
  }
};

/** 目标文件自带换行（CRLF/LF）都要先归一化，否则 includes/split 会误报。 */
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
  console.error("用法: node scripts/mutate_continuation_bridge.mjs [C1|C2|C3] [restore]");
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
