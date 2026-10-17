// 反向变异脚本：Task #53 第二轮「WPS 格式动作面板闭环」三个阻断点修复的**新增守卫**。
//
// 用法：node scripts/mutate_wps_format_closed_loop.mjs M1|M2|M3|M4|check [restore]
//
// 本轮三个阻断点各引入一组守卫，被测机制分四处（详见驱动
// `run_wps_format_closed_loop_mutations.mjs` 头部注释）：
//
//   ── 阻断点 1：格式动作接入副本审阅路径 ─────────────────────────────
//   M1 apps/wps-addin/src/wpsNativeComparison.ts 的 `stageWpsNativeComparison`
//      抽掉 `applyFormatAction`（副本不再应用格式）→ 副本审阅时格式根本没被改，
//      回读判定必然失败、副本审阅当场抛错。打「格式动作必须真的在副本上改格式，
//      值/公式不变」这条闭口。
//
//   ── 阻断点 2：属性级恢复（撤销不覆盖用户对其他属性的后续修改）──────
//   M2 apps/wps-addin/src/wpsPlanExecutor.ts 的撤销分支把 `declaredFormatFields`
//      换成「恢复全部字段」→ 撤销会连用户后续改的字色/行高一起覆盖。打
//      「撤销只恢复本次动作声明过的属性」这条闭口。
//   M4 同一文件 `overlaps` 抽掉「两个格式动作之间允许重叠」→ 同区域两个格式动作
//      被判成重叠写入而整份拒绝。打「同区域多格式动作组合不得误报冲突」。
//
//   ── 阻断点 3：多格缺 Item 拒绝（不回退整块伪造逐格快照）────────────
//   M3 apps/wps-addin/src/wpsFormatSnapshot.ts 的 `captureWpsFormatSnapshot`
//      把缺 `Item` 时返回的 `captured:false` 改成 `captured:true` → 多格缺逐格
//      定位能力时不再拒绝，而是伪造"已采集"。打「多格/多行缺 Item 必须在写入前
//      拒绝、不能伪造逐格快照」。
//
//   ── 收口：组合状态（中间态 vs 最终状态）──────────────────────────────
//   M5 apps/wps-addin/src/wpsNativeComparison.ts 的 finalize 循环把
//      `if (!isFormatActionType(...)) continue;` 改成 `if (isFormatActionType(...)) continue;`
//      → 不再把副本侧快照覆盖成**最终状态**，退回逐动作中间态。打「确认与提交后
//      对照必须按整份计划最终状态，同区域『加粗→居中』不得被中间态误报」。
//   M6 apps/wps-addin/src/wpsPlanExecutor.ts 的 `formatActionSafeToUndo` 丢掉最终
//      快照、退回"声明目标值"旧口径 → 同属性后续覆盖（字号 12/13→14→16）时第一步
//      声明的 14 被误判成冲突。打「撤销冲突检测按最终恢复影响范围、逆序恢复」。
//
//   ── 第 18 轮：组合边界（部分重叠 / 行高整行语义）──────────────────────
//   M7 apps/wps-addin/src/wpsPlanExecutor.ts 的 `formatActionUnchangedSinceFinal`
//      把行高比对分支关掉（`if (declared.has("rowHeight"))` → 恒 false）→ 撤销
//      冲突检测不再覆盖行高：外部改行高（哪怕是同行无关列）也不再被判成冲突。
//      打「撤销冲突检测遵循整行语义」。
//   M8 apps/wps-addin/src/wpsFormatSnapshot.ts 的 `restoreWpsFormatSnapshot`
//      把行高恢复分支关掉（`if (fieldSet.has("rowHeight"))` → 恒 false）→ 撤销
//      不再恢复行高、逐行回读核对随即失败。打「撤销恢复范围包含整行行高」。
//
// 八个变异落在三个文件上（wpsNativeComparison.ts / wpsPlanExecutor.ts /
// wpsFormatSnapshot.ts），换行符不统一（前两者 CRLF、后者 LF），所以 seed/mutant
// 的归一化与写回都要**按文件**取目标换行。
//
// 脚本只做 apply / restore 两件事，判据由 run_wps_format_closed_loop_mutations.mjs 负责。
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const WPS_COMPARISON = resolve(REPO, "apps", "wps-addin", "src", "wpsNativeComparison.ts");
const WPS_EXECUTOR = resolve(REPO, "apps", "wps-addin", "src", "wpsPlanExecutor.ts");
const WPS_SNAPSHOT = resolve(REPO, "apps", "wps-addin", "src", "wpsFormatSnapshot.ts");

const MUTATIONS = {
  M1: {
    file: WPS_COMPARISON,
    seed: `        applyFormatAction(range, asFormatFields(action));`,
    mutant: `        void action;`
  },
  M2: {
    file: WPS_EXECUTOR,
    seed: `      restoreWpsFormatSnapshot(formatTarget, action.beforeFormatSnapshot!, declaredFormatFields(formatFields));`,
    mutant: `      restoreWpsFormatSnapshot(formatTarget, action.beforeFormatSnapshot!);`
  },
  M3: {
    file: WPS_SNAPSHOT,
    seed: `    return { cells: [], rowHeights: [], captured: false };`,
    mutant: `    return { cells: [], rowHeights: [], captured: true };`
  },
  M4: {
    file: WPS_EXECUTOR,
    seed: `  if (isFormatActionType(left.type) && isFormatActionType(right.type)) return false;`,
    mutant: `  void left; void right;`
  },
  M5: {
    // finalize 循环里把"跳过非格式动作"反转成"跳过格式动作" ⇒ 不再覆盖成最终状态，
    // 副本侧快照退回逐动作中间态。
    file: WPS_COMPARISON,
    seed: `    for (const [index, action] of prepared.actions.entries()) {\n      if (!isFormatActionType(action.type)) continue;\n      const finalFormat = captureWpsFormatSnapshot(target(sheet, action.range), action.rowCount, action.columnCount);`,
    mutant: `    for (const [index, action] of prepared.actions.entries()) {\n      if (isFormatActionType(action.type)) continue;\n      const finalFormat = captureWpsFormatSnapshot(target(sheet, action.range), action.rowCount, action.columnCount);`
  },
  M6: {
    // 撤销冲突检测丢掉最终快照、退回"声明目标值"旧口径。
    file: WPS_EXECUTOR,
    seed: `  return afterFinal\n    ? formatActionUnchangedSinceFinal(action, afterFinal, current)\n    : formatActionStillApplied(action, current);`,
    mutant: `  return formatActionStillApplied(action, current);`
  },
  M7: {
    // 撤销冲突检测不再比对行高：外部改行高（同行任意列）不再被判成冲突。
    file: WPS_EXECUTOR,
    seed: `  if (declared.has("rowHeight")) {\n    for (let r = 0; r < current.rowHeights.length; r += 1) {\n      if (current.rowHeights[r] !== finalSnapshot.rowHeights[r]) return false;\n    }\n  }`,
    mutant: `  if (declared.has("bold") && false) {\n    for (let r = 0; r < current.rowHeights.length; r += 1) {\n      if (current.rowHeights[r] !== finalSnapshot.rowHeights[r]) return false;\n    }\n  }`
  },
  M8: {
    // 恢复快照不再恢复行高：撤销后行高保持执行后的值，逐行回读核对随即失败。
    file: WPS_SNAPSHOT,
    seed: `  if (fieldSet.has("rowHeight")) {\n    for (let r = 0; r < snapshot.rowHeights.length; r += 1) {\n      const height = snapshot.rowHeights[r];\n      if (height === null) continue;\n      cellAt(range, r, 0, singleCell).RowHeight = height;\n    }\n  }`,
    mutant: `  if (fieldSet.has("bold") && false) {\n    for (let r = 0; r < snapshot.rowHeights.length; r += 1) {\n      const height = snapshot.rowHeights[r];\n      if (height === null) continue;\n      cellAt(range, r, 0, singleCell).RowHeight = height;\n    }\n  }`
  }
};

function newlineOf(text) {
  return text.includes("\r\n") ? "\r\n" : LF;
}

function normalize(text) {
  return text.replace(/\r\n/g, LF);
}

function readTarget(mutation) {
  return readFileSync(mutation.file, "utf8");
}

function apply(name) {
  const mutation = MUTATIONS[name];
  if (!mutation) throw new Error(`未知变异：${name}`);
  const raw = readTarget(mutation);
  const eol = newlineOf(raw);
  const source = normalize(raw);
  const seed = normalize(mutation.seed);
  const mutant = normalize(mutation.mutant);
  const occurrences = source.split(seed).length - 1;
  if (occurrences !== 1) {
    throw new Error(`${name} 的 SEED 在 ${mutation.file} 里命中 ${occurrences} 次（要求恰好 1 次）`);
  }
  const next = source.replace(seed, mutant);
  writeFileSync(mutation.file, eol === LF ? next : next.replace(/\n/g, eol), "utf8");
  return { file: mutation.file, eol };
}

function restore(name) {
  const mutation = MUTATIONS[name];
  if (!mutation) throw new Error(`未知变异：${name}`);
  const raw = readTarget(mutation);
  const eol = newlineOf(raw);
  const source = normalize(raw);
  const seed = normalize(mutation.seed);
  const mutant = normalize(mutation.mutant);
  const occurrences = source.split(mutant).length - 1;
  if (occurrences !== 1) {
    throw new Error(`${name} 的 MUTANT 在 ${mutation.file} 里命中 ${occurrences} 次（要求恰好 1 次）`);
  }
  const next = source.replace(mutant, seed);
  writeFileSync(mutation.file, eol === LF ? next : next.replace(/\n/g, eol), "utf8");
  return { file: mutation.file, eol };
}

/** 前置状态闸：文件停在 SEED 态（未变异）才允许继续。 */
function check() {
  for (const [name, mutation] of Object.entries(MUTATIONS)) {
    const raw = readTarget(mutation);
    const source = normalize(raw);
    const seed = normalize(mutation.seed);
    const mutant = normalize(mutation.mutant);
    if (!source.includes(seed)) {
      return { ok: false, name, file: mutation.file, reason: "找不到 SEED（可能停在上一次被中断的变异态）" };
    }
    if (source.includes(mutant)) {
      return { ok: false, name, file: mutation.file, reason: "已经处在 MUTANT 态（上一次 restore 没跑完）" };
    }
  }
  return { ok: true };
}

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

const mode = process.argv[2];
const argument = process.argv[3];

// 中断也要还原：把已 apply 的变异原路退回。
let applied = null;
const guard = (signal) => {
  if (applied) {
    try {
      restore(applied);
      process.stderr.write(`[restore:${signal}] ${applied} 已还原\n`);
    } catch (error) {
      process.stderr.write(`[restore:${signal}] 还原失败：${error.message}\n`);
    }
  }
  process.exit(1);
};
for (const signal of ["SIGINT", "SIGTERM", "SIGHUP"]) {
  process.on(signal, () => guard(signal));
}
process.on("uncaughtException", (error) => {
  guard(`uncaught:${error.message}`);
});

try {
  if (mode === "check") {
    const result = check();
    if (!result.ok) {
      process.stdout.write(`CHECK FAILED ${result.name} ${result.file}: ${result.reason}\n`);
      process.exit(1);
    }
    process.stdout.write(`CHECK OK 全部 ${Object.keys(MUTATIONS).length} 个 SEED 都在位\n`);
    process.exit(0);
  }
  if (!MUTATIONS[mode]) throw new Error(`未知用法：${mode}`);
  if (argument === "restore") {
    const result = restore(mode);
    process.stdout.write(`RESTORED ${mode} → ${result.file} sha256=${sha256(result.file)}\n`);
    process.exit(0);
  }
  const result = apply(mode);
  applied = mode;
  process.stdout.write(`APPLIED ${mode} → ${result.file} sha256=${sha256(result.file)}\n`);
  process.exit(0);
} catch (error) {
  process.stderr.write(`${error.message}\n`);
  process.exit(1);
}
