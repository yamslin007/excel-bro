/**
 * 反向变异驱动：证明「进度浮标外圈卡片浅色渐变 + 主题色前景图形（三点 / 锚）」的
 * **对比度与令牌守卫**真的有判别力。
 *
 * 被测机制 = styles.test.ts 里那条 `keeps the progress-dot contrast above the
 * graphics threshold in every preset`。它读**四个**真实来源：
 *   - 外圈渐变引用了哪些**令牌** ← styles.css 的 `.new-progress-button` 规则本体
 *   - 圆点 / 锚图标**各自**引用了哪个令牌 ← styles.css 的两条字形规则本体
 *   - 那些令牌在四个预设里各是什么**色值** ← useTheme.ts 的色板
 *   - `:root` 里那份与 default 预设是否同值
 * 所以七个变异分别打这些地方，各自都必须让用例变红：
 *
 *   M1 vivid-green 的前景色换回纯品牌色 → 未压深只有 2.93:1，掉到门槛以下
 *   M2 外圈退回写死的深色药丸       → 上色声明里出现字面色值，负向锚点触发
 *   M3 default 的 brandHover 改深    → 若守卫只看渐变起点而不取最暗端点，此变异不会红
 *   M4 外圈那条"关掉边框"的声明换回"画一条 1px 实线" → 打「外圈不许画边」这条
 *      负向断言。它同时钉着两个方向：真会画线的声明（`solid` + 非零宽度）必须红，
 *      而 `border: none` 这种"一个像素都不画"的声明必须放行 —— 后者是用户裁决后
 *      用来按掉浏览器 UA 默认按钮边框的（不写它浅绿圆外会冒黑环）。
 *   M5 圆点退回「降级行 + color-mix」 → 上次事故本体；守卫的「不许 color-mix」必须拦下
 *   M6 锚图标（实心剪影）换成另一个主题令牌 → 若守卫只查"是不是令牌"、不查"是不是同一个令牌"，不会红
 *   M7 锚图标取色写成同值字面量      → 若守卫不强制 `var(--color-*)` 引用，不会红
 *
 * 四条判据缺一不算过：
 * 1. 未变异时基线 GREEN（退出码 0 + 有效测试摘要 + 用例数达标）；
 * 2. 每个变异 apply 后用例必须变红（退出码非 0）；
 * 3. restore 后**被改的那个文件**逐字节一致（sha256 比对）；
 * 4. 还原后复跑全绿，且两个文件最终都与基线一致。
 *
 * 起手另有**前置状态闸**（`... check`）：文件若停在上一次被中断的变异态，
 * 当场报出来并退出，不去跑那几轮 vitest。
 *
 * 用法：node scripts/run_progress_dot_contrast_mutations.mjs
 */
import { spawnSync } from "node:child_process";
import { openSync, closeSync, readFileSync, rmSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");
const addinDir = resolve(repoRoot, "apps", "excel-addin");
const scratch = resolve(repoRoot, "outputs", "progress-dot-contrast-mutation.log");
const MUTATIONS = ["M1", "M2", "M3", "M4", "M5", "M6", "M7"];
// M5/M6/M7 的字形守卫、M4 的边框守卫都是加在**同一个 `it` 里的断言**，不新增用例，
// 所以这个下限不变。它的作用是钉住"没人把这个文件里的用例整条删掉"。
const MIN_TESTS = 25;
const TEST_FILE = "src/styles.test.ts";

const WATCHED = [
  resolve(addinDir, "src", "styles.css"),
  resolve(addinDir, "src", "hooks", "useTheme.ts")
];

const vitest = resolve(repoRoot, "node_modules", "vitest", "vitest.mjs");

/**
 * 用**文件描述符**而不是 pipe 接子进程输出：流水线里创建管道会被拦，
 * 而判据又必须读到真实的摘要行。
 */
function runVitest() {
  const fd = openSync(scratch, "w");
  const result = spawnSync(process.execPath, [vitest, "run", TEST_FILE], {
    cwd: addinDir,
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
    passed: passed ? Number(passed[1]) : 0,
    raw
  };
}

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

function mutate(which, restore) {
  const args = [
    resolve(here, "mutate_progress_dot_contrast.mjs"),
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

const baselineShas = new Map(WATCHED.map((path) => [path, sha256(path)]));
for (const [path, sha] of baselineShas) {
  console.log(`[mutate] 基线 ${path} sha256=${sha}`);
}

// 前置状态闸：文件停在上一次被中断的变异态时当场退出，不白跑几轮 vitest。
const preflight = mutate("check", false);
if (preflight.status !== 0) {
  console.error(`[mutate] 前置状态闸未通过：\n${preflight.out}`);
  process.exit(1);
}
console.log("[mutate] 前置状态闸通过：两个文件都在 SEED 态");

console.log("[mutate] 基线（未变异）应为 GREEN…");
const baseline = runVitest();
if (baseline.status !== 0 || !baseline.summary || baseline.passed < MIN_TESTS) {
  console.error(`基线就不绿（status=${baseline.status}, 摘要=${baseline.summary || "无"}）`);
  process.exit(1);
}
console.log(`[mutate] 基线 GREEN：${baseline.summary}`);

let failures = 0;
for (const which of MUTATIONS) {
  const applied = mutate(which, false);
  if (applied.status !== 0 || applied.error) {
    console.error(
      `[mutate] ${which} apply 失败：status=${applied.status} err=${applied.error} out=${applied.out}`
    );
    failures += 1;
    continue;
  }
  const run = runVitest();
  const verdict = run.status !== 0 ? "RED(OK)" : "GREEN(NG!)";
  console.log(`[mutate] ${which} ${applied.out} → ${verdict} ${run.summary || "(无摘要)"}`);
  if (run.status === 0) {
    console.error(`  ${which} 拆掉守卫后用例仍然全绿 —— 断言没有判别力，见：${scratch}`);
    failures += 1;
  }
  const restored = mutate(which, true);
  if (restored.status !== 0 || restored.out === "ALREADY-CLEAN") {
    console.error(`[mutate] ${which} restore 失败：${restored.out}`);
    failures += 1;
    continue;
  }
  for (const [path, sha] of baselineShas) {
    if (sha256(path) !== sha) {
      console.error(`[mutate] ${which} 还原后 ${path} 与基线不一致（sha256 变了）`);
      failures += 1;
    }
  }
}

for (const [path, sha] of baselineShas) {
  const finalSha = sha256(path);
  if (finalSha !== sha) {
    console.error(`[mutate] 最终 ${path} 与基线不一致：${finalSha} != ${sha}`);
    failures += 1;
  }
}
console.log(`[mutate] 两个文件最终 sha256 与基线一致：${[...baselineShas.values()].length} 个`);

console.log("[mutate] 还原后复跑应为 GREEN…");
const after = runVitest();
if (after.status !== 0 || after.passed < MIN_TESTS) {
  console.error(`[mutate] 还原后复跑未全绿：${after.summary}`);
  failures += 1;
} else {
  console.log(`[mutate] 还原后 GREEN：${after.summary}`);
}
rmSync(scratch, { force: true });

if (failures > 0) {
  console.error(`反向变异未通过：${failures} 处问题。`);
  process.exit(1);
}
console.log("反向变异通过：M1–M7 各自变红，还原逐字节一致，复跑全绿。");
