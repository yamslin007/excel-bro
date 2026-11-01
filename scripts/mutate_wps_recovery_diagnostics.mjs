// 反向变异脚本：恢复核对**诊断层**（2026-10-29）
//
// 这一层只观察、不判定，所以变异要打在"诊断能不能分开"的判据上，而不是"结论对不对"上：
//
// M1 = 把"已选表为空"并回 `authorize_threw`
//      「已选表为空 → authorize_no_selection」必须红。判别点：两支原因码不再可辨。
// M2 = 去掉 DEV 门控（非 DEV 也记录）
//      「非 DEV 一条都不留」必须红。判别点：生产里会留痕。
// M3 = 去掉异常 name 的白名单正则
//      「异常只留白名单 name」必须红。判别点：带路径的名字会被记下来。
// M4 = 把"只读重试"记成 `discovery`
//      「首次发现与只读重试各留一条诊断」必须红。判别点：两条路径不再能对照。
//
// 一个变异只允许一个判别点：M1 只动原因码选择，M2 只动门控，M3 只动名字过滤，
// M4 只动 stage 传参。它们互不遮蔽 —— 各自挂在不同的断言上。
//
// 用法：
//   node scripts/mutate_wps_recovery_diagnostics.mjs              # 打印基线 sha256
//   node scripts/mutate_wps_recovery_diagnostics.mjs M1           # 应用
//   node scripts/mutate_wps_recovery_diagnostics.mjs M1 restore   # 还原
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const DIAGNOSTICS = resolve(REPO, "apps", "wps-addin", "src", "wpsRecoveryDiagnostics.ts");
const INSPECTION = resolve(REPO, "apps", "wps-addin", "src", "wpsRecoveryInspection.ts");
const PANEL = resolve(REPO, "apps", "wps-addin", "src", "WpsAgentPanel.tsx");

const TARGETS = {
  M1: {
    file: INSPECTION,
    seed: 'if (!selectedSheetCount) return stop("unknown", "authorize_no_selection");',
    mutant: 'if (!selectedSheetCount) return stop("unknown", "authorize_threw");'
  },
  M2: {
    file: DIAGNOSTICS,
    seed: ["  if (!isDev()) return;", "  sequence += 1;"].join(LF),
    mutant: "  sequence += 1;"
  },
  M3: {
    file: DIAGNOSTICS,
    seed: 'return typeof name === "string" && /^[A-Za-z_$][A-Za-z0-9_$]{0,40}$/.test(name) ? name : undefined;',
    mutant: 'return typeof name === "string" ? name : undefined;'
  },
  M4: {
    file: PANEL,
    seed: 'wpsRecoveryVerdictFor(state, "recheck")',
    mutant: 'wpsRecoveryVerdictFor(state, "discovery")'
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
  console.error("用法: node scripts/mutate_wps_recovery_diagnostics.mjs [M1|M2|M3|M4] [restore]");
  process.exit(2);
}

const target = TARGETS[which];
const original = readFileSync(target.file, "utf8");

/*
 * 写回必须读回校验并重试。
 *
 * Windows 上偶发"写进去了、读回来还是旧的"（索引器/防病毒短暂占句柄），
 * 变异体留在工作区比变异失败更糟：后面的回归会以"红"的面目出现，而人以为那是产品缺陷。
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
