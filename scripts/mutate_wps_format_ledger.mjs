// 反向变异脚本：Task #62「格式动作恢复账本链路」本轮新增守卫的判别力。
//
// 用法：node scripts/mutate_wps_format_ledger.mjs M1|M2|M3|M4|M5|check [restore]
//
// 背景：真实 WPS 验收在「确认写回」时失败，界面原文
// 「无法持久登记本次写入现场，已取消写入；原表未被修改。WPS 恢复描述的动作定位或形状无效」。
// 根因是"格式动作没有可用的现场证据"这条链在四层里各断一段。下面五个变异各打一段，
// 判别点互相不同（详见驱动脚本头部）。
//
//   M1 apps/excel-addin/src/contracts.ts 的 `WPS_RECOVERABLE_ACTION_TYPES` 退回
//      "只认三个值动作" ⇒ 格式动作在 fetch 之前就被出站契约拒。打「出站闸必须认识
//      已开放的格式动作」这条闭口 —— 就是实机故障的那一段。
//   M2 apps/wps-addin/src/wpsPlanExecutor.ts 的 `composedFormatSnapshot` 关掉前序
//      叠加（`earlier < index` → `earlier < 0`）⇒ 同区域同属性的后一个动作拿"预览
//      原始快照"当写前基准，写前复核与账本登记都对不上。打「写前现场按属性级叠加」。
//   M3 apps/wps-addin/src/wpsRecoveryInspection.ts 的格式分支关掉 ⇒ 现场核对退回
//      值矩阵指纹；格式动作不改值，核对永远判不出"已完成/未开始"。打「对账/续跑
//      必须按格式证据指纹核对现场」。
//   M4 server/app/main.py 的 `_validate_wps_format_evidence` 关掉属性白名单 ⇒ 服务端
//      放行未验证属性（underline / vertical / columnWidth）与"不属于该动作的属性"。
//      打「服务端第三道闸的属性级校验」。
//   M5 apps/wps-addin/src/wpsPlanExecutor.ts 的 `composedFormatSnapshot` 关掉**行高**
//      叠加分支 ⇒ 行高（整行属性）在列不相交时不再被叠加。打「行高按整行范围叠加」。
//
// 五个变异落在四个文件上，换行符不统一（contracts.ts 与 main.py 是 CRLF，
// wpsPlanExecutor.ts 与 wpsRecoveryInspection.ts 是 LF），所以归一化与写回都
// **按文件**取目标换行。
//
// 本脚本只做 apply / restore；判据由 `run_wps_format_ledger_mutations.mjs` 负责。
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const CONTRACTS = resolve(REPO, "apps", "excel-addin", "src", "contracts.ts");
const EXECUTOR = resolve(REPO, "apps", "wps-addin", "src", "wpsPlanExecutor.ts");
const INSPECTION = resolve(REPO, "apps", "wps-addin", "src", "wpsRecoveryInspection.ts");
const MAIN_PY = resolve(REPO, "server", "app", "main.py");

const MUTATIONS = {
  M1: {
    file: CONTRACTS,
    seed: `export const WPS_RECOVERABLE_ACTION_TYPES = [
  "writeValues",
  "writeFormulas",
  "clearRange",
  "setFont",
  "setAlignment",
  "resizeRange"
] as const;`,
    mutant: `export const WPS_RECOVERABLE_ACTION_TYPES = [
  "writeValues",
  "writeFormulas",
  "clearRange"
] as const;`
  },
  M2: {
    file: EXECUTOR,
    seed: `  for (let earlier = 0; earlier < index; earlier += 1) {`,
    mutant: `  for (let earlier = 0; earlier < 0; earlier += 1) {`
  },
  M3: {
    file: INSPECTION,
    seed: `      if (isFormatActionType(action.type)) {`,
    mutant: `      if (isFormatActionType(action.type) && false) {`
  },
  M4: {
    file: MAIN_PY,
    seed: `    unsupported = [name for name in declared if name not in allowed]
    if unsupported:`,
    mutant: `    unsupported = [name for name in declared if False]
    if unsupported:`
  },
  M5: {
    file: EXECUTOR,
    seed: `      if (field === "rowHeight") {`,
    mutant: `      if (field === "rowHeight" && false) {`
  }
};

function newlineOf(text) {
  return text.includes("\r\n") ? "\r\n" : LF;
}

function normalize(text) {
  return text.replace(/\r\n/g, LF);
}

/** 读盘上的**当前**内容（不能拿内存基线去算）。 */
function readTarget(mutation) {
  return readFileSync(mutation.file);
}

function writeTarget(mutation, sourceWithLf) {
  const eol = newlineOf(readFileSync(mutation.file, "utf8"));
  writeFileSync(mutation.file, eol === LF ? sourceWithLf : sourceWithLf.replace(/\n/g, eol));
}

function apply(name) {
  const mutation = MUTATIONS[name];
  if (!mutation) throw new Error(`未知变异：${name}`);
  const source = normalize(readTarget(mutation).toString("utf8"));
  const seed = normalize(mutation.seed);
  const mutant = normalize(mutation.mutant);
  const occurrences = source.split(seed).length - 1;
  if (occurrences !== 1) {
    throw new Error(`${name} 的 SEED 在 ${mutation.file} 里命中 ${occurrences} 次（要求恰好 1 次）`);
  }
  writeTarget(mutation, source.replace(seed, mutant));
  return { file: mutation.file };
}

function restore(name) {
  const mutation = MUTATIONS[name];
  if (!mutation) throw new Error(`未知变异：${name}`);
  const source = normalize(readTarget(mutation).toString("utf8"));
  const seed = normalize(mutation.seed);
  const mutant = normalize(mutation.mutant);
  const occurrences = source.split(mutant).length - 1;
  if (occurrences !== 1) {
    throw new Error(`${name} 的 MUTANT 在 ${mutation.file} 里命中 ${occurrences} 次（要求恰好 1 次）`);
  }
  writeTarget(mutation, source.replace(mutant, seed));
  return { file: mutation.file };
}

/** 前置状态闸：所有文件都停在 SEED 态（未变异）才允许继续。 */
function check() {
  for (const [name, mutation] of Object.entries(MUTATIONS)) {
    const source = normalize(readTarget(mutation).toString("utf8"));
    if (!source.includes(normalize(mutation.seed))) {
      return { ok: false, name, file: mutation.file, reason: "找不到 SEED（可能停在上一次被中断的变异态）" };
    }
    if (source.includes(normalize(mutation.mutant))) {
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
      const result = restore(applied);
      process.stderr.write(`[restore:${signal}] ${applied} → ${result.file} 已还原\n`);
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
