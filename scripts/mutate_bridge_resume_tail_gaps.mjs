// 反向变异脚本：SDK 续跑收尾的**两处调用链缺口**（2026-10-20 第四批）
//
// E1 = 旧轮次**异常**退出时不再接续新一轮排队的工作
//      把 drain 里「异常也要走同一套出口」撤掉：让 catch 只记账、不再判断 queued，
//      退化成"异常时这一轮直接结束"。表现是旧轮次 AbortError/断网后，
//      新桥三项计数（轮询/读取/回执）都是 0，必须再等一个事件。
// E2 = 调用层不看桥返回值，桥没立起来照样发请求
//      把 App.tsx 里续跑路径的 `if (!bridgeReady) return;` 撤掉，退回无条件发
//      `continuePlanningTask`。表现是 ensureLive 返回 false 时请求仍然发出去，
//      服务端指针无人接、模型等超时。
//
// 用法：
//   node scripts/mutate_bridge_resume_tail_gaps.mjs               # 打印基线 sha256
//   node scripts/mutate_bridge_resume_tail_gaps.mjs E1            # 应用
//   node scripts/mutate_bridge_resume_tail_gaps.mjs E1 restore    # 还原
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const HOOK = resolve(REPO, "apps", "excel-addin", "src", "hooks", "useSdkToolBridge.ts");
const APP = resolve(REPO, "apps", "excel-addin", "src", "App.tsx");
const WPS = resolve(REPO, "apps", "wps-addin", "src", "WpsAgentPanel.tsx");

const TARGETS = {
  E1: {
    file: HOOK,
    seed: [
      "        } catch {",
      "          /*",
      "           * 这一轮**抛出来**了：控制流从 `poll` 或 `receipt` 一路到这里（取消 →",
      "           * AbortError、断网 → fetch reject）。",
      "           *",
      "           * 以前这里只是「记账、然后 finally 放开 runningRef」，于是排过队的那条",
      "           * 新指针永远没人兑现：桥活着、状态里有待办、一个请求都不发，表现和",
      '           * "继续分析之后一动不动"完全一样。正常的 stale 分支已经会接续队列，',
      "           * 异常分支也必须走**同一套**出口 —— 收尾不能分两种。",
      "           *",
      '           * 但"接续"不等于"重试"：见下面的 return 条件。',
      "           */",
      "          roundFailed = true;",
      "        }"
    ].join(LF),
    // 变异：把"异常也要接续"撤掉，回到旧的"异常直接结束"——roundFailed 标记丢掉，
    // 但保持代码形状不报编译错误（roundFailed 仍赋值，只是后面的分支被短路掉）。
    mutant: [
      "        } catch {",
      "          // [变异E1] 旧轮次异常退出时直接结束，不再接续排队的工作",
      "          return;",
      "        }"
    ].join(LF)
  },
  E2: {
    file: APP,
    seed: [
      "      if (!bridgeReady) {",
      "        /*",
      "         * **桥没立起来就不发继续请求。**"
    ].join(LF),
    mutant: [
      "      if (false) {",
      "        /*",
      "         * **桥没立起来就不发继续请求。**"
    ].join(LF)
  },
  E3: {
    file: WPS,
    seed: [
      "      const bridgeReady = await ensureWpsBridge(pause.turnId, evidenceIndexes, controller.signal);",
      "      if (!bridgeReady) {"
    ].join(LF),
    mutant: [
      "      const bridgeReady = await ensureWpsBridge(pause.turnId, evidenceIndexes, controller.signal);",
      "      if (false) {"
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
  console.error("用法: node scripts/mutate_bridge_resume_tail_gaps.mjs [E1|E2|E3] [restore]");
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
