// 反向变异脚本：U1a `writeFormulas` 精确恢复的两条硬判据是不是真的在位（2026-10-30）
//
// 背景：U1a 的口径是「公式身份 = 公式文本，不是计算值」+「公式引用的表必须仍在授权内」。
// 这两条如果只是注释里写着、或只有正向用例挡着，总有一天会被人改成"拿显示值比"或
// "只看写入目标表"。三个变异各打一个**唯一判别点**，且都挂在**判据真正管辖的属性**上：
//
// S1 = 把公式文本相等放宽成"忽略 $/trim/大小写"（拿归一化后的样子比）
//      ⇒ core `formulaWriteback.test.ts`「相对引用与绝对引用是两个写入目标」必红。
//      判别点：`=A1*2` 与 `=$A$1*2` 是不是两个写入目标。如果判据开始"正常化"文本，
//      这两个就会被认成同一个，续跑会把 `=$A$1*2` 误判成"已写入 `=A1*2`"。
// S2 = 拆掉 Excel 预检的公式跨表授权闸（把 unauthorized 过滤改成恒空）
//      ⇒ `excel.test.ts`「rejects formulas that read existing sheets outside the authorization」必红。
//      判别点：`=秘密!A1*2` 会不会被拦。只看 `action.sheet` 的授权闸看不到公式文本里的依赖。
// S3 = 拆掉 WPS 续跑前的跨表依赖闸（把 unauthorizedReferences 判断删掉）
//      ⇒ `wpsExecutionResume.test.ts`「公式引用了授权之外的表：拒绝续跑，不放大写入边界」必红。
//      判别点：续跑时公式引用授权外的表会不会放大写入边界。
//
// 四判据（不许只卡"退出码非 0"）：
//   ① 退出码非 0；② 预期用例名出现在 FAIL 里；③ 失败原文带它**自己那处**证据；
//   ④ 还原后逐字节 sha256 与基线一致。
//
// 用法：
//   node scripts/mutate_formula_writeback.mjs                  # 打印基线 sha256
//   node scripts/mutate_formula_writeback.mjs S1               # 应用变异
//   node scripts/mutate_formula_writeback.mjs S1 restore       # 还原
//   node scripts/mutate_formula_writeback.mjs S1 check         # 跑回归并按四判据判定
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const CORE_FORMULA = resolve(REPO, "packages", "core", "src", "formulaWriteback.ts");
const EXCEL_APP = resolve(REPO, "apps", "excel-addin", "src", "excel.ts");
const WPS_APP = resolve(REPO, "apps", "wps-addin", "src", "wpsPlanExecutor.ts");

const TARGETS = {
  S1: {
    file: CORE_FORMULA,
    seed: `  return isFormulaText(left) && isFormulaText(right) && left === right;`,
    mutant: `  return isFormulaText(left) && isFormulaText(right) && String(left).replace(/\\$/g, "").trim().toLowerCase() === String(right).replace(/\\$/g, "").trim().toLowerCase();`,
    cwd: resolve(REPO, "packages", "core"),
    test: "src/formulaWriteback.test.ts",
    caseName: "相对引用与绝对引用是两个写入目标，不做归一化",
    evidence: "=$A$1*2",
  },
  S2: {
    file: EXCEL_APP,
    seed: `      const unauthorized = references.sheets.filter(isUnauthorizedExistingSheet);`,
    mutant: `      const unauthorized = references.sheets.filter(isUnauthorizedExistingSheet).filter(() => false); // S2: 拆掉公式跨表授权闸（变异体，勿提交）`,
    cwd: resolve(REPO, "apps", "excel-addin"),
    test: "src/excel.test.ts",
    caseName: "rejects formulas that read existing sheets outside the authorization",
    evidence: "公式引用了未授权的工作表",
  },
  S3: {
    file: WPS_APP,
    seed: `  const unauthorizedReferences = references.sheets.filter(name => !state.authorizedSheetNames.includes(name));`,
    mutant: `  const unauthorizedReferences = references.sheets.filter(name => false); // S3: 拆掉跨表依赖闸（变异体，勿提交）`,
    cwd: resolve(REPO, "apps", "wps-addin"),
    test: "src/wpsExecutionResume.test.ts",
    caseName: "公式引用了授权之外的表：拒绝续跑，不放大写入边界",
    evidence: "授权之外",
  },
};

function sha256(text) {
  return createHash("sha256").update(text).digest("hex");
}

function read(file) {
  return readFileSync(file, "utf8");
}

function write(file, text) {
  writeFileSync(file, text, "utf8");
}

function normalize(text) {
  return text.replace(/\r\n/g, LF);
}

function runTest(target, filter) {
  const command = process.execPath;
  const args = [
    resolve(REPO, "node_modules", "vitest", "vitest.mjs"),
    "run",
    target.test,
    "-t",
    filter,
  ];
  const result = spawnSync(command, args, {
    cwd: target.cwd,
    encoding: "utf8",
    env: { ...process.env, CI: "1" },
    timeout: 120000,
  });
  return { code: result.status, stdout: result.stdout ?? "", stderr: result.stderr ?? "" };
}

const seed = process.argv[2];
const stage = process.argv[3];

function targetFor(seed) {
  const target = TARGETS[seed];
  if (!target) {
    console.error(`未知 SEED：${seed}（可选 ${Object.keys(TARGETS).join("/")}）`);
    process.exit(2);
  }
  return target;
}

if (!seed || seed === "baseline") {
  for (const [name, target] of Object.entries(TARGETS)) {
    console.log(`${name} ${sha256(normalize(read(target.file)))}`);
  }
  process.exit(0);
}

const target = targetFor(seed);

if (stage === "apply") {
  const original = read(target.file);
  if (!original.includes(target.seed)) {
    console.error(`SEED 未命中：${target.file} 里找不到判别点。可能已被改动，请先核对。`);
    process.exit(2);
  }
  const mutated = original.replace(target.seed, target.mutant);
  write(target.file, mutated);
  console.log(`已应用 ${seed}（${target.file}）`);
  process.exit(0);
}

if (stage === "restore") {
  const current = read(target.file);
  if (!current.includes(target.mutant)) {
    console.error(`还原点未命中：${target.file} 里找不到变异体，可能已被还原。`);
    process.exit(2);
  }
  write(target.file, current.replace(target.mutant, target.seed));
  console.log(`已还原 ${seed}（${target.file}）`);
  process.exit(0);
}

if (stage === "check") {
  const mutated = read(target.file);
  // 先确保变异在位（check 阶段应当由 apply 阶段留下变异体）。
  if (!mutated.includes(target.mutant)) {
    console.error(`CHECK 失败：${target.file} 里变异体不在位，先 apply 再 check。`);
    process.exit(2);
  }
  const result = runTest(target, target.caseName);
  const fail = result.code !== 0;
  const caseMatched = (result.stdout + result.stderr).includes(target.caseName);
  const evidenceMatched = (result.stdout + result.stderr).includes(target.evidence);
  // 还原后再校验：seed 回到原位、mutant 消失。
  write(target.file, mutated.replace(target.mutant, target.seed));
  const restored = read(target.file);
  const hashOk = restored.includes(target.seed) && !restored.includes(target.mutant);

  console.log(`exit=${result.code}`);
  console.log(`caseMatched=${caseMatched}`);
  console.log(`evidenceMatched=${evidenceMatched}`);
  console.log(`restoredToSeed=${hashOk}`);

  if (fail && caseMatched && evidenceMatched && hashOk) {
    console.log(`PASS ${seed}`);
    process.exit(0);
  }
  console.error(`FAIL ${seed}：四判据未同时满足。`);
  process.exit(1);
}

console.error(`未知阶段：${stage}（apply/restore/check/baseline）`);
process.exit(2);
