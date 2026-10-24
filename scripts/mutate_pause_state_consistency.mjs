// 反向变异脚本：「暂停之后界面到底停没停」（2026-10-23）
//
// 背景（用户实测）：继续之后按暂停，请求确实 abort 了，但过程卡仍显示"进行中"、
// 耗时继续增长；再过一会儿按钮还可能重新变成"可暂停"，点下去却没有可中止对象。
//
// 三处机制各自独立，因此每个变异只拆一处、只应打红它自己那条用例：
//  - P1（共享判据）：只要"有任务在跑"就声称可暂停 → 判据本身不再要求可中止对象；
//  - P2（桥归属）：桥只在"仍持有 controller"时收 → 核对接管 controller 后桥无人收；
//  - P3（有界收尾）：去掉"按下暂停就先按归属收干净"这一步 → 收尾被迫等核对回来；
//  - P5（收尾归属，两处同拆）：P3 与"finally 回到单一 ownsTurn"必须**一起**拆——
//       两道守卫各自兜住了同一件事，只拆一处仍然全绿（冗余层一次拆全）。
//  - P4（Excel 判据）：回到 `busy` 一个变量 + 无条件声称可暂停。
//
// 用法：
//   node scripts/mutate_pause_state_consistency.mjs            → 打印基线 sha256
//   node scripts/mutate_pause_state_consistency.mjs P1         → 应用变异
//   node scripts/mutate_pause_state_consistency.mjs P1 restore → 还原
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";
const CORE = resolve(REPO, "packages", "core", "src", "pauseControl.ts");
const REPLY_SEGMENTS = resolve(REPO, "packages", "core", "src", "replySegments.ts");
const WPS = resolve(REPO, "apps", "wps-addin", "src", "WpsAgentPanel.tsx");
const EXCEL = resolve(REPO, "apps", "excel-addin", "src", "App.tsx");

const TARGETS = {
  P1: {
    file: CORE,
    seed: [
      "  if (input.taskRunning && input.controllerAlive) {",
      "    return {",
      '      kind: "pause",'
    ].join(LF),
    mutant: [
      "  // [变异P1] 只要“有任务在跑”就声称可暂停（不再要求存在可中止对象）",
      "  if (input.taskRunning) {",
      "    return {",
      '      kind: "pause",'
    ].join(LF)
  },
  P2: {
    file: WPS,
    seed: [
      "    endBridgeOwnedBy(params.owner);",
      "    return plan;"
    ].join(LF),
    mutant: [
      "    // [变异P2] 桥只在“仍持有 controller”时收（三样资源共用一个布尔）",
      "    if (plan.clearController) endBridgeOwnedBy(params.owner);",
      "    return plan;"
    ].join(LF)
  },
  P3: {
    file: WPS,
    // 只拆"按下暂停就先收尾"这一步：收尾于是退回到 finally，必须等核对回来。
    seed: [
      "          settleExecution({",
      "            owner: executionOwner,",
      "            activity,",
      "            activitySettled,",
      '            outcome: "awaiting_input"',
      "          });",
      "          activitySettled = true;"
    ].join(LF),
    mutant: [
      "          // [变异P3] 不在核对之前收尾（收尾退回到 finally，必须等核对回来）",
      "          activitySettled = activitySettled;"
    ].join(LF)
  },
  P5: {
    file: WPS,
    edits: [
      {
        seed: [
          "          settleExecution({",
          "            owner: executionOwner,",
          "            activity,",
          "            activitySettled,",
          '            outcome: "awaiting_input"',
          "          });",
          "          activitySettled = true;"
        ].join(LF),
        mutant: [
          "          // [变异P5a] 不在核对之前收尾"
        ].join(LF)
      },
      {
        seed: [
          "      settleExecution({",
          "        owner: executionOwner,",
          "        activity,",
          "        activitySettled,",
          '        outcome: "awaiting_input"',
          "      });",
          "      resumingPauseRef.current = false;"
        ].join(LF),
        mutant: [
          "      // [变异P5b] 回到单一 ownsTurn 判断：核对接管 controller 之后三样全放弃",
          "      const ownsTurn = controllerRef.current === controller;",
          "      if (ownsTurn) {",
          '        if (activity && !activitySettled) endActivity(activity, "awaiting_input");',
          "        controllerRef.current = null;",
          "        setProgress(null);",
          "        sdkBridge.end();",
          "      }",
          "      resumingPauseRef.current = false;"
        ].join(LF)
      }
    ]
  },
  P4: {
    file: EXCEL,
    seed: [
      "  const composerPauseButton = pauseButtonState({",
      "    hasInput: prompt.trim().length > 0 || pendingImages.length > 0,",
      "    paused: composerCanContinue,",
      "    taskRunning: running,",
      "    controllerAlive: Boolean(",
      "      (turnAbortRef.current && !turnAbortRef.current.signal.aborted) ||",
      "      (queryAbortRef.current && !queryAbortRef.current.signal.aborted)",
      "    )",
      "  });"
    ].join(LF),
    mutant: [
      "  // [变异P4] 回到旧判据：只看 busy，并一律声称可暂停",
      "  const composerPauseButton = pauseButtonState({",
      "    hasInput: prompt.trim().length > 0 || pendingImages.length > 0,",
      "    paused: composerCanContinue,",
      "    taskRunning: busy,",
      "    controllerAlive: true",
      "  });"
    ].join(LF)
  },
  /*
   * P6（Excel 迟到进度）：首次流的 `onProgress` 不再看"这一次是不是已经取消"。
   * 拆掉之后，暂停后到达的读表回执会把活动卡重新变回"进行中"。
   */
  P6: {
    file: EXCEL,
    seed: [
      "          onProgress: (event) => {",
      "            if (planningController?.signal.aborted) return;"
    ].join(LF),
    mutant: [
      "          onProgress: (event) => {",
      "            // [变异P6] 不再丢弃已取消那一次的迟到进度"
    ].join(LF)
  },
  /*
   * P7（暂停回执三态）：不再看 `pauseRequested`，一律当成"服务端已接受"。
   * 拆掉之后，"服务端此刻没有在跑这条任务"会被谎报成"已接受暂停请求"。
   */
  P7: {
    file: EXCEL,
    seed: "        if (ack.pauseRequested) {",
    mutant: [
      "        // [变异P7] 不看回执，一律当成服务端已接受",
      "        if (true) {"
    ].join(LF)
  },
  /*
   * P9（暂停必须先关掉这一轮的活动）：暂停分支不再收尾活动。
   * 拆掉之后，按下暂停的卡仍留在"进行中"（活动卡数 != 0）。
   */
  P9: {
    file: EXCEL,
    seed: [
      "        rememberEarlyPause(intentText);",
      '        finishActivity("warning", "需求分析已暂停，原需求已保留", "awaiting_input");',
      '      } else if (isAbortError(reason)) finishActivity("cancelled", "用户已取消");'
    ].join(LF),
    mutant: [
      "        rememberEarlyPause(intentText);",
      "        // [变异P9] 暂停不再收尾这一轮的活动（卡留在进行中）",
      '      } else if (isAbortError(reason)) finishActivity("cancelled", "用户已取消");'
    ].join(LF)
  }  ,
  /*
   * P10（段不接受迟到进度）：把"已冻结的段不再接受进度"这一层拆掉。
   * 它与 P6 的 onProgress 守卫是**两道**：只拆 P6 全绿（P6 是冗余层），
   * 拆这里才会让迟到进度真的改写冻结段。
   */
  P10: {
    file: REPLY_SEGMENTS,
    seed: [
      "  if (segment.status !== \"live\") return false;"
    ].join(LF),
    mutant: [
      "  // [变异P10] 已冻结的段也接受进度"
    ].join(LF)
  }  ,
  /*
   * P8（按下暂停必须真的取消请求）：只把活动标成已暂停，不再 abort。
   * 界面会显示停了下来，而请求还在飞 —— 这就是按了没用的那个假象。
   */
  P8: {
    file: EXCEL,
    seed: [
      '    markActivityInterrupted("已暂停接收进度，等待核对检查点。");',
      "    controller.abort();"
    ].join(LF),
    mutant: [
      '    markActivityInterrupted("已暂停接收进度，等待核对检查点。");',
      "    // [变异P8] 不再取消请求：界面说停了，请求还在飞"
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

const which = process.argv[2];
if (!which) {
  for (const [id, target] of Object.entries(TARGETS)) {
    console.log(`BASELINE ${id} ${target.file} sha256=${sha256(target.file)}`);
  }
  process.exit(0);
}
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_pause_state_consistency.mjs [P1|P2|P3|P4] [restore]");
  process.exit(2);
}

const target = TARGETS[which];
const original = readFileSync(target.file, "utf8");
/** 单个 `seed/mutant` 或一组 `edits`（冗余守卫要一次拆全，分两次拆会互相兜住）。 */
const edits = target.edits ?? [{ seed: target.seed, mutant: target.mutant }];

if (process.argv[3] === "restore") {
  let text = normalize(original);
  let restored = 0;
  for (const edit of edits) {
    if (text.split(edit.mutant).length - 1 === 0) continue;
    text = text.split(edit.mutant).join(edit.seed);
    restored += 1;
  }
  if (restored === 0) {
    console.log("ALREADY-CLEAN");
    process.exit(0);
  }
  writeFileSync(target.file, denormalize(text, original), "utf8");
  console.log(`restore: OK ${which} → ${target.file} sha256=${sha256(target.file)}`);
  process.exit(0);
}

let text = normalize(original);
for (const [index, edit] of edits.entries()) {
  const occurrences = text.split(edit.seed).length - 1;
  if (occurrences !== 1) {
    console.log(`SKIP: SEED 出现 ${occurrences} 次（期望 1）于 ${target.file} #${index + 1}`);
    process.exit(3);
  }
  text = text.replace(edit.seed, edit.mutant);
}
writeFileSync(target.file, denormalize(text, original), "utf8");
console.log(`APPLIED ${which} → ${target.file} sha256=${sha256(target.file)}`);
