// 反向变异脚本：本批四处修复的**关键守卫**（2026-10-22）
//
// 一个变异只允许一个判别点，且必须挂在那个机制**真正管辖**的属性上：
//
// S1 = WPS 续跑父段退回"从前往后找第一条匹配 turnId 的消息"
//      ⇒ 第二次继续仍然改写第一段，链断在中间（WPS 的「查看后续进展」数量必须红）。
// S7 = 共享层 `latestReplySegmentForTask` 退回"从前往后找第一段"
//      ⇒ 纯判据自己也失效（段链接哪一段的单元用例必须红）。
// S2 = 增量按文本去重（`code.endsWith(delta)`）
//      ⇒ `f(g(1` + `)` + `)` 丢掉第二个右括号（core 的连续括号用例必须红）。
// S3 = 早期恢复复用旧消息 id（`persistenceKey = parentId ?? append(...)`）
//      ⇒ Excel 的早期暂停恢复退化成"就地改旧卡"（段数必须回到 1）。
// S4 = WPS 早期恢复不接父段（`parentId = null`）
//      ⇒ 旧段没有后继入口（「查看后续进展」必须红）。
// S5 = "正在运行"的启动证据被提前置真（`started = {"yes": True}`）
//      ⇒ 启动前取消被说成"已暂停"（not_started 用例必须红）。
// S6 = 未开始执行也带 `executedVersion`
//      ⇒ 一段从没跑过的代码被显示成"已运行"（executedVersion 用例必须红）。
// S8 = 活动日志写回时不再问"目标段还是不是活动段"（共享谓词被短路）
//      ⇒ 已冻结的历史段被迟到进度重新激活（Excel 写回路径用例必须红）。
//
// 用法：
//   node scripts/mutate_reply_segments.mjs              # 打印基线 sha256
//   node scripts/mutate_reply_segments.mjs check        # 只校验 SEED 命中
//   node scripts/mutate_reply_segments.mjs S1           # 应用
//   node scripts/mutate_reply_segments.mjs S1 restore   # 还原
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const FILES = {
  segments: resolve(REPO, "packages", "core", "src", "replySegments.ts"),
  code: resolve(REPO, "packages", "core", "src", "pythonCodeArtifact.ts"),
  excel: resolve(REPO, "apps", "excel-addin", "src", "App.tsx"),
  excelActivity: resolve(REPO, "apps", "excel-addin", "src", "activityMessages.ts"),
  wps: resolve(REPO, "apps", "wps-addin", "src", "WpsAgentPanel.tsx"),
  coordinator: resolve(REPO, "server", "app", "python_coordinator.py")
};

const TARGETS = {
  S1: {
    file: FILES.wps,
    seed: [
      "    const original =",
      "      (continuation.previousSegmentId",
      "        ? messages.find((item) => item.id === continuation.previousSegmentId)",
      "        : undefined) ??",
      "      [...messages]",
      "        .reverse()",
      "        .find((item) => item.activityLog?.operationId === turnId) ??",
      "      (legacy?.activityLog?.operationId ? undefined : legacy);"
    ].join(LF),
    // 旧实现原样：从前往后找第一条匹配 turnId 的消息当父段。
    mutant: [
      "    const original =",
      "      messages.find((item) => item.activityLog?.operationId === turnId) ??",
      "      (legacy?.activityLog?.operationId ? undefined : legacy);"
    ].join(LF)
  },
  S7: {
    file: FILES.segments,
    seed: [
      "  for (let index = segments.length - 1; index >= 0; index -= 1) {",
      "    if (segments[index].taskId === taskId) return segments[index];"
    ].join(LF),
    mutant: [
      "  for (let index = 0; index < segments.length; index += 1) {",
      "    if (segments[index].taskId === taskId) return segments[index];"
    ].join(LF)
  },
  S2: {
    file: FILES.code,
    seed: "      code = code + event.delta!;",
    mutant: "      code = code.endsWith(event.delta!) ? code : code + event.delta!;"
  },
  S3: {
    file: FILES.excel,
    seed: [
      "    const parentId = params.previousSegmentId ?? previous?.id ?? null;",
      "    const persistenceKey = appendMessage({",
      "      role: \"assistant\",",
      "      activityPending: true",
      "    });"
    ].join(LF),
    mutant: [
      "    const parentId = params.previousSegmentId ?? previous?.id ?? null;",
      "    const persistenceKey = parentId ?? appendMessage({",
      "      role: \"assistant\",",
      "      activityPending: true",
      "    });"
    ].join(LF)
  },
  S4: {
    file: FILES.wps,
    seed: "    const parentId = previous?.id ?? null;",
    mutant: "    const parentId: string | null = null;"
  },
  S5: {
    file: FILES.coordinator,
    seed: "    started = {\"yes\": False}",
    mutant: "    started = {\"yes\": True}"
  },
  S6: {
    file: FILES.coordinator,
    seed: [
      "                \"pythonCode\": code_stage(",
      "                    refusal_stage or STAGE_NOT_STARTED,",
      "                    version=1,",
      "                    chars=len(code_text or \"\"),"
    ].join(LF),
    mutant: [
      "                \"pythonCode\": code_stage(",
      "                    refusal_stage or STAGE_NOT_STARTED,",
      "                    version=1,",
      "                    executedVersion=1,",
      "                    chars=len(code_text or \"\"),"
    ].join(LF)
  },
  S8: {
    file: FILES.excelActivity,
    seed: [
      "  if (",
      "    !segmentAcceptsProgress(",
      "      replySegmentsOf(messages),",
      "      resolvedTargetId,",
      "      log.operationId ?? null",
      "    )",
      "  ) {",
      "    return messages;",
      "  }"
    ].join(LF),
    mutant: [
      "  if (",
      "    false && !segmentAcceptsProgress(",
      "      replySegmentsOf(messages),",
      "      resolvedTargetId,",
      "      log.operationId ?? null",
      "    )",
      "  ) {",
      "    return messages;",
      "  }"
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

/** 写回必须读回校验并重试：变异体留在工作区比变异失败更糟。 */
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
  for (const [, target] of Object.entries(TARGETS)) {
    console.log(`BASELINE ${target.file} sha256=${sha256(target.file)}`);
  }
  process.exit(0);
}

if (which === "check") {
  let bad = 0;
  for (const [id, target] of Object.entries(TARGETS)) {
    const occurrences = normalize(readFileSync(target.file, "utf8")).split(target.seed).length - 1;
    console.log(`${id} SEED ${occurrences === 1 ? "OK" : `MISS(${occurrences})`} ${target.file}`);
    if (occurrences !== 1) bad += 1;
  }
  process.exit(bad === 0 ? 0 : 3);
}

if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_reply_segments.mjs [check|S1..S8] [restore]");
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
  const restored = denormalize(
    normalize(original).split(target.mutant).join(target.seed),
    original
  );
  if (!writeVerified(target.file, restored)) process.exit(4);
  console.log(`restore: OK ${which} ${target.file} sha256=${sha256(target.file)}`);
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
