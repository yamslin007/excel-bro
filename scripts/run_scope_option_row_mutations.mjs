/**
 * 反向变异驱动：证明「数据范围面板 · 选项行」的四条判据**真有判别力**。
 *
 * 四条判据缺一不算过：
 * 1. 未变异时基线 GREEN（退出码 0 + 有效测试摘要 + 用例数达标）——**两端都要**；
 * 2. 每个变异 apply 后，**它声明的那一端**必须变红（退出码非 0）；
 * 3. restore 后四个目标文件**逐字节一致**（sha256 比对）；
 * 4. 还原后复跑两端全绿。
 *
 * 七个变异各钉一处互不相同的属性（见 scripts/mutate_scope_option_rows.mjs 的注释）：
 * M1/M2 钉「行内还敢不敢显示授权文案」，M3/M4 钉「勾选框是共享一份还是被覆盖」，
 * M5/M6 钉「工作表行还跟不跟模式卡同高」，M7 钉「内层 fieldset 还读不读间距 token」。
 *
 * 真实像素由 `scripts/check_scope_option_rows.mjs` 在无头浏览器里量（自带 `--legacy`
 * 反向对照：修复前实测行高 38px、勾选框 25px 灰方块、块间距 0px）。
 *
 * 用法：node scripts/run_scope_option_row_mutations.mjs
 */
import { spawnSync } from "node:child_process";
import { openSync, closeSync, readFileSync, mkdirSync, rmSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");
const scratchDir = resolve(repoRoot, "outputs");
const scratch = resolve(scratchDir, "scope-option-row-mutation.log");
const vitest = resolve(repoRoot, "node_modules", "vitest", "vitest.mjs");

const APPS = {
  excel: {
    dir: resolve(repoRoot, "apps", "excel-addin"),
    test: "src/scopeOptionRow.test.tsx",
    /** 该文件共 6 条：行内文案 2 + 勾选框 2 + 行高 1 + 块间距 1。 */
    minTests: 6
  },
  wps: {
    dir: resolve(repoRoot, "apps", "wps-addin"),
    test: "src/wpsScopeOptionRow.test.tsx",
    /** 该文件共 3 条：行内文案/结构 1 + 勾选框 1 + 行高 1。 */
    minTests: 3
  }
};
const TARGETS = [
  resolve(repoRoot, "apps", "wps-addin", "src", "WpsProbePanel.tsx"),
  resolve(repoRoot, "apps", "wps-addin", "src", "styles.css"),
  resolve(repoRoot, "apps", "excel-addin", "src", "styles.css"),
  resolve(repoRoot, "apps", "excel-addin", "src", "App.tsx")
];
/** 变异 → 应当在哪一端变红（每一条都必须由**它自己的**那一端判红）。 */
const MUTATIONS = {
  M1: "wps",
  M2: "excel",
  M3: "excel",
  M4: "wps",
  M5: "wps",
  M6: "excel",
  M7: "excel"
};

mkdirSync(scratchDir, { recursive: true });

/**
 * 用**文件描述符**而不是 pipe 接子进程输出：流水线里创建管道会被拦，
 * 而判据又必须读到真实的摘要行。vitest 摘要带 ANSI 色码，解析前先剥掉。
 */
function runTests(app) {
  const fd = openSync(scratch, "w");
  const result = spawnSync(process.execPath, [vitest, "run", APPS[app].test], {
    cwd: APPS[app].dir,
    stdio: ["ignore", fd, fd]
  });
  closeSync(fd);
  const raw = readFileSync(scratch, "utf8").replace(/\x1b\[[0-9;]*m/g, "");
  const summary =
    raw.split(/\r?\n/).find((line) => /Tests\s+\d+ (passed|failed)/.test(line)) ?? "";
  const passed = /Tests\s+(\d+) passed/.exec(summary);
  return {
    status: result.status,
    summary: summary.trim(),
    passed: passed ? Number(passed[1]) : 0
  };
}

function runBoth() {
  return { excel: runTests("excel"), wps: runTests("wps") };
}

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

function mutate(which, restore) {
  const args = [
    resolve(here, "mutate_scope_option_rows.mjs"),
    which,
    ...(restore ? ["restore"] : [])
  ];
  const fd = openSync(scratch, "w");
  const result = spawnSync(process.execPath, args, {
    cwd: repoRoot,
    stdio: ["ignore", fd, fd]
  });
  closeSync(fd);
  return {
    status: result.status,
    out: readFileSync(scratch, "utf8").trim(),
    error: result.error ? String(result.error.code) : null
  };
}

const baselineSha = new Map(TARGETS.map((file) => [file, sha256(file)]));
for (const [file, hash] of baselineSha) console.log(`[mutate] 基线 ${file} sha256=${hash}`);

/*
 * 前置状态闸：七个变异串都必须**定点**（改动一处源码就可能打散旧串）。
 * 这条闸不是形式主义 —— 旧 M5 串只写 `min-height: 38px;`，与本文件里
 * `.connection-shell > button` 的同名声明撞车，"还原"顺手把连接页按钮
 * 撑成了 59px。串一失配就必须先停，别让后面的判据在错误的基线上跑。
 */
console.log("[mutate] 先跑变异串定点闸…");
const preflight = mutate("check", false);
if (preflight.status !== 0) {
  console.error(preflight.out);
  console.error("[mutate] 变异串已失配，拒绝继续（先修 scripts/mutate_scope_option_rows.mjs 的种子）。");
  process.exit(1);
}
console.log(preflight.out.split(/\r?\n/).filter(Boolean).join("  |  "));

console.log("[mutate] 基线（未变异）两端都应为 GREEN…");
const baseline = runBoth();
let failures = 0;
for (const app of ["excel", "wps"]) {
  const run = baseline[app];
  if (run.status !== 0 || !run.summary || run.passed < APPS[app].minTests) {
    console.error(
      `基线就不绿（${app}: status=${run.status}, 摘要=${run.summary || "无"}）`
    );
    failures += 1;
  } else {
    console.log(`[mutate] 基线 GREEN ${app}：${run.summary}`);
  }
}
if (failures > 0) process.exit(1);

for (const [which, app] of Object.entries(MUTATIONS)) {
  const applied = mutate(which, false);
  if (applied.status !== 0 || applied.error) {
    console.error(
      `[mutate] ${which} apply 失败：status=${applied.status} err=${applied.error} out=${applied.out}`
    );
    failures += 1;
    continue;
  }
  const runs = runBoth();
  const other = app === "excel" ? "wps" : "excel";
  const verdict = runs[app].status !== 0 ? "RED(OK)" : "GREEN(NG!)";
  console.log(
    `[mutate] ${which}（应由 ${app} 判红）→ ${verdict} ${runs[app].summary || "(无摘要)"}` +
      `；另一端 ${other}=${runs[other].status === 0 ? "GREEN" : "RED"}`
  );
  if (runs[app].status === 0) {
    console.error(`  ${which} 之后 ${app} 仍然全绿 —— 该断言没有判别力，见：${scratch}`);
    failures += 1;
  }
  const restored = mutate(which, true);
  if (restored.status !== 0 || restored.out === "ALREADY-CLEAN") {
    console.error(`[mutate] ${which} restore 失败：${restored.out}`);
    failures += 1;
    continue;
  }
  for (const [file, hash] of baselineSha) {
    if (sha256(file) !== hash) {
      console.error(`[mutate] ${which} 还原后 ${file} 与基线不一致（sha256 变了）`);
      failures += 1;
    }
  }
}

for (const [file, hash] of baselineSha) {
  const finalSha = sha256(file);
  if (finalSha !== hash) {
    console.error(`[mutate] 最终 ${file} 与基线不一致：${finalSha} != ${hash}`);
    failures += 1;
  }
}

console.log("[mutate] 还原后复跑两端应为 GREEN…");
const after = runBoth();
for (const app of ["excel", "wps"]) {
  if (after[app].status !== 0 || after[app].passed < APPS[app].minTests) {
    console.error(`[mutate] 还原后 ${app} 未全绿：${after[app].summary}`);
    failures += 1;
  } else {
    console.log(`[mutate] 还原后 GREEN ${app}：${after[app].summary}`);
  }
}
rmSync(scratch, { force: true });

if (failures > 0) {
  console.error(`反向变异未通过：${failures} 处问题。`);
  process.exit(1);
}
console.log("反向变异通过：M1..M7 各自把它自己那一端判红，还原逐字节一致，复跑两端全绿。");
