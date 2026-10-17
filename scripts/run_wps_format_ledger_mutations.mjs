/**
 * 反向变异驱动：证明 Task #62「格式动作恢复账本链路」本轮新增守卫**真的有判别力**。
 *
 * 被测机制 → 判别用例（每个变异一个**不同**的判别点）：
 *   M1 出站契约的动作白名单（contracts.ts）
 *      ← scripts/run_wps_format_ledger_link.mjs（真实面板 + 真实出站校验 + 真实 HTTP
 *        + 真实服务端账本）。这正是实机失败的那一段。
 *   M2 写前现场按属性级叠加（composedFormatSnapshot 的前序叠加）
 *      ← apps/wps-addin/src/wpsFormatExecutor.test.ts
 *   M3 现场核对按格式证据指纹（inspectWpsRecoveryState 的格式分支）
 *      ← apps/wps-addin/src/wpsRecoveryInspection.test.ts
 *   M4 服务端属性级白名单（main._validate_wps_format_evidence）
 *      ← server/tests/test_wps_execution_api.py
 *   M5 行高按**整行**范围叠加（composedFormatSnapshot 的 rowHeight 分支）
 *      ← apps/wps-addin/src/wpsFormatCombinationBoundary.test.ts
 *
 * 四条判据缺一不算过：
 * 1. 未变异时基线 GREEN（退出码 0 + 有效测试摘要 + 用例数达标）；
 * 2. 每个变异 apply 后其目标用例必须变红（退出码非 0 **且**有测试摘要 ——
 *    编译/收集期崩掉不算判别力）；
 * 3. restore 后每个被改的文件逐字节回到基线（sha256）；
 * 4. 全部还原后复跑全绿。
 *
 * 起手另有前置状态闸（`check`）：文件停在上一次被中断的变异态时当场退出。
 *
 * 用法：node scripts/run_wps_format_ledger_mutations.mjs
 */
import { spawnSync } from "node:child_process";
import { closeSync, existsSync, openSync, readFileSync, rmSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");
const wpsDir = resolve(repoRoot, "apps", "wps-addin");
const scratch = resolve(repoRoot, "outputs", "wps-format-ledger-mutation.log");
const vitest = resolve(repoRoot, "node_modules", "vitest", "vitest.mjs");
const python = (() => {
  const venv = resolve(repoRoot, ".venv", "Scripts", "python.exe");
  const venvPosix = resolve(repoRoot, ".venv", "bin", "python");
  if (existsSync(venv)) return venv;
  if (existsSync(venvPosix)) return venvPosix;
  return process.platform === "win32" ? "python" : "python3";
})();

const vitestSpec = (name, file, minTests) => ({
  name,
  minTests,
  command: process.execPath,
  args: [vitest, "run", "--config", "vite.config.ts", file],
  cwd: wpsDir
});

const EXECUTOR_SPEC = vitestSpec("vitest:wps/wpsFormatExecutor", "src/wpsFormatExecutor.test.ts", 14);
const INSPECTION_SPEC = vitestSpec("vitest:wps/wpsRecoveryInspection", "src/wpsRecoveryInspection.test.ts", 20);
const BOUNDARY_SPEC = vitestSpec("vitest:wps/wpsFormatCombinationBoundary", "src/wpsFormatCombinationBoundary.test.ts", 12);

/** 跨进程链路：自带临时服务端、真实 HTTP、真实 SQLite 账本；模型规划固定、宿主为替身。 */
const LEDGER_CHAIN_SPEC = {
  name: "node/run_wps_format_ledger_link",
  minTests: 2,
  command: process.execPath,
  args: [resolve(repoRoot, "scripts", "run_wps_format_ledger_link.mjs")],
  cwd: repoRoot
};

const PYTEST_SPEC = {
  name: "pytest/test_wps_execution_api",
  minTests: 47,
  command: python,
  args: ["-m", "pytest", "server/tests/test_wps_execution_api.py", "-q", "--basetemp=.tmp-mut-wps-api"],
  cwd: repoRoot
};

const MUTATIONS = [
  { id: "M1", specs: [LEDGER_CHAIN_SPEC] },
  { id: "M2", specs: [EXECUTOR_SPEC] },
  { id: "M3", specs: [INSPECTION_SPEC] },
  { id: "M4", specs: [PYTEST_SPEC] },
  { id: "M5", specs: [BOUNDARY_SPEC] }
];

const WATCHED = [
  resolve(repoRoot, "apps", "excel-addin", "src", "contracts.ts"),
  resolve(repoRoot, "apps", "wps-addin", "src", "wpsPlanExecutor.ts"),
  resolve(repoRoot, "apps", "wps-addin", "src", "wpsRecoveryInspection.ts"),
  resolve(repoRoot, "server", "app", "main.py")
];

/**
 * 输出走**文件描述符**而不是管道：管道会被拦，而判据又必须读到真实摘要行。
 * vitest 的摘要带 ANSI 色码，解析前先剥掉。
 */
function runSpec(spec) {
  const fd = openSync(scratch, "w");
  const result = spawnSync(spec.command, spec.args, {
    cwd: spec.cwd,
    stdio: ["ignore", fd, fd]
  });
  closeSync(fd);
  const raw = readFileSync(scratch, "utf8").replace(/\x1b\[[0-9;]*m/g, "");
  const lines = raw.split(/\r?\n/);
  // vitest: `Tests  14 passed (14)`；pytest: `47 passed, 1 warning in 7.7s`；
  // 跨进程脚本里嵌的 vitest 摘要同样是 `Tests  2 passed (2)`。
  const summaryLine =
    lines.find((line) => /Tests\s+\d+ (passed|failed)/.test(line)) ??
    lines.find((line) => /^\d+ (passed|failed)/.test(line)) ??
    lines.reverse().find((line) => /\b\d+ passed\b|\b\d+ failed\b/.test(line)) ??
    "";
  const passed = Number(/(\d+) passed/.exec(summaryLine)?.[1] ?? 0);
  const failed = Number(/(\d+) failed/.exec(summaryLine)?.[1] ?? 0);
  return {
    status: result.status,
    summary: summaryLine.trim(),
    passed,
    failed,
    error: result.error ? String(result.error.code ?? result.error) : null
  };
}

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

function invokeMutation(id, restore) {
  const args = [resolve(here, "mutate_wps_format_ledger.mjs"), id, ...(restore ? ["restore"] : [])];
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

// 前置状态闸：停在上一次被中断的变异态时当场退出，不白跑五轮测试。
const preflight = invokeMutation("check");
if (preflight.status !== 0) {
  console.error(`[mutate] 前置状态闸未通过：\n${preflight.out}`);
  process.exit(1);
}
console.log(`[mutate] 前置状态闸通过：${MUTATIONS.length} 个 SEED 全部在位`);

const BASELINE_SPECS = [EXECUTOR_SPEC, INSPECTION_SPEC, BOUNDARY_SPEC, PYTEST_SPEC, LEDGER_CHAIN_SPEC];

console.log("[mutate] 基线（未变异）应为 GREEN…");
for (const spec of BASELINE_SPECS) {
  const result = runSpec(spec);
  const ok = result.status === 0 && result.summary && result.passed >= spec.minTests;
  console.log(
    `[mutate]   基线 ${spec.name}: status=${result.status} 摘要="${result.summary}" ${ok ? "OK" : "不合格"}`
  );
  if (!ok) {
    console.error(`[mutate] 基线就不绿或没跑到预期用例数（要求 >= ${spec.minTests}），中止。`);
    process.exit(1);
  }
}

const failures = [];
for (const { id, specs } of MUTATIONS) {
  const applied = invokeMutation(id, false);
  if (applied.status !== 0) {
    failures.push(`${id} apply 失败：${applied.out || applied.error}`);
    console.error(`[mutate] ${id} apply 失败：${applied.out || applied.error}`);
    continue;
  }
  for (const spec of specs) {
    const red = runSpec(spec);
    // 判据 2 要求"变红"：退出码非 0 **且**有有效测试摘要 —— 编译/收集期直接崩掉
    // 不算判别力，那只能证明我们把代码写坏了。
    const wentRed = red.status !== 0 && Boolean(red.summary);
    console.log(
      `[mutate] ${id} apply → ${applied.out.split(" sha256=")[0]} ; ${spec.name}: status=${
        red.status
      } 摘要="${red.summary}" ${wentRed ? "RED(合格)" : "未变红(不合格)"}`
    );
    if (!wentRed) {
      failures.push(`${id} 未让 ${spec.name} 变红（status=${red.status}, 摘要=${red.summary || "无"}）`);
    }
  }
  const restored = invokeMutation(id, true);
  if (restored.status !== 0) {
    failures.push(`${id} restore 失败：${restored.out || restored.error}`);
    console.error(`[mutate] ${id} restore 失败：${restored.out || restored.error}`);
  }
}

// 判据 3：每个被改过的文件都必须逐字节回到基线。
for (const [path, sha] of baselineShas) {
  const now = sha256(path);
  if (now !== sha) {
    failures.push(`${path} 还原后与基线不一致：${now} != ${sha}`);
    console.error(`[mutate] 还原不一致 ${path}`);
  }
}
if (!failures.length) console.log(`[mutate] ${WATCHED.length} 个文件全部逐字节还原`);

console.log("[mutate] 还原后复跑全绿…");
for (const spec of BASELINE_SPECS) {
  const result = runSpec(spec);
  const ok = result.status === 0 && result.summary && result.passed >= spec.minTests;
  console.log(`[mutate]   复跑 ${spec.name}: status=${result.status} 摘要="${result.summary}" ${ok ? "OK" : "不合格"}`);
  if (!ok) failures.push(`还原后复跑 ${spec.name} 不为绿（status=${result.status}）`);
}

rmSync(scratch, { force: true });

if (failures.length) {
  console.error("\n[mutate] 未通过：");
  for (const line of failures) console.error(`  - ${line}`);
  process.exit(1);
}
console.log(
  `\n[mutate] 全部通过：${MUTATIONS.length} 个变异各自让目标用例变红，${WATCHED.length} 个文件逐字节还原，复跑全绿。`
);
