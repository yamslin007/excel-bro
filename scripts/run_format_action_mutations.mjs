/**
 * 反向变异驱动：证明第 3 阶段「常用格式操作」两端闭环的新增守卫**真的有判别力**。
 *
 * 被测机制分四处（细节见 `scripts/mutate_format_actions.mjs` 头部注释）：
 *   - 共享层影响范围/对照/预览  ← packages/core/src/formatActions.test.ts
 *   - Excel 端字段级校验        ← apps/excel-addin/src/formatActionFields.test.ts
 *   - WPS 端 fail-closed + 配置漂移 ← apps/wps-addin/src/wpsPlanExecutor.test.ts
 *   - Python 协议空属性/重复边 + 跨语言漂移 ← server/tests/test_protocol_consistency.py
 *
 * 四条判据缺一不算过：
 * 1. 未变异时基线 GREEN（退出码 0 + 有效测试摘要 + 用例数达标）；
 * 2. 每个变异 apply 后其目标用例必须变红（退出码非 0 且**有测试摘要**）；
 * 3. restore 后**被改的那个文件**逐字节一致（sha256 比对）；
 * 4. 全部还原后复跑全绿，且五个文件最终都与基线一致。
 *
 * 起手另有前置状态闸（`check`）：文件停在上一次被中断的变异态时当场退出。
 *
 * 用法：node scripts/run_format_action_mutations.mjs
 */
import { spawnSync } from "node:child_process";
import { openSync, closeSync, readFileSync, rmSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");
const addinDir = resolve(repoRoot, "apps", "excel-addin");
const wpsDir = resolve(repoRoot, "apps", "wps-addin");
const coreDir = resolve(repoRoot, "packages", "core");
const scratch = resolve(repoRoot, "outputs", "format-action-mutation.log");
const PYTHON = resolve(repoRoot, ".venv", "Scripts", "python.exe");
const vitest = resolve(repoRoot, "node_modules", "vitest", "vitest.mjs");

const PYTEST_SPEC = {
  name: "pytest:test_protocol_consistency",
  minTests: 55,
  command: PYTHON,
  args: [
    "-m",
    "pytest",
    "server/tests/test_protocol_consistency.py",
    "-q",
    "--basetemp=.tmp-pytest/btm",
    "-p",
    "no:cacheprovider"
  ],
  cwd: repoRoot
};

function vitestSpec(name, cwd, config, file, minTests) {
  return {
    name,
    minTests,
    command: process.execPath,
    args: [vitest, "run", "--config", config, file],
    cwd
  };
}

const CORE_SPEC = vitestSpec(
  "vitest:core/formatActions",
  coreDir,
  "vitest.config.ts",
  "src/formatActions.test.ts",
  25
);
const ADDIN_SPEC = vitestSpec(
  "vitest:addin/formatActionFields",
  addinDir,
  "vite.config.ts",
  "src/formatActionFields.test.ts",
  13
);
const WPS_SPEC = vitestSpec(
  "vitest:wps/wpsPlanExecutor",
  wpsDir,
  "vite.config.ts",
  "src/wpsPlanExecutor.test.ts",
  30
);

const MUTATIONS = [
  { id: "V1", spec: CORE_SPEC },
  { id: "V2", spec: ADDIN_SPEC },
  { id: "V3", spec: WPS_SPEC },
  { id: "V4", spec: WPS_SPEC },
  { id: "V5", spec: CORE_SPEC },
  { id: "V6", spec: CORE_SPEC },
  { id: "V7", spec: CORE_SPEC },
  { id: "V8", spec: CORE_SPEC },
  { id: "P1", spec: PYTEST_SPEC },
  { id: "P2", spec: PYTEST_SPEC },
  { id: "P3", spec: PYTEST_SPEC },
  { id: "P4", spec: PYTEST_SPEC },
  { id: "P5", spec: PYTEST_SPEC },
  { id: "P6", spec: PYTEST_SPEC },
  { id: "P7", spec: PYTEST_SPEC },
  { id: "P8", spec: PYTEST_SPEC }
];

const WATCHED = [
  resolve(coreDir, "src", "formatActions.ts"),
  resolve(addinDir, "src", "contracts.ts"),
  resolve(wpsDir, "src", "wpsPlanExecutor.ts"),
  resolve(repoRoot, "config", "capabilities.json"),
  resolve(repoRoot, "server", "app", "models.py")
];

/**
 * 输出走**文件描述符**而不是管道：管道会被拦，而判据又必须读到真实摘要行。
 * vitest 的摘要有 ANSI 色码，解析前先剥掉。
 */
function runSpec(spec) {
  const fd = openSync(scratch, "w");
  const result = spawnSync(spec.command, spec.args, {
    cwd: spec.cwd,
    stdio: ["ignore", fd, fd]
  });
  closeSync(fd);
  const raw = readFileSync(scratch, "utf8").replace(/\x1b\[[0-9;]*m/g, "");
  const summaryLine =
    raw.split(/\r?\n/).find((line) => /Tests\s+\d+ (passed|failed)/.test(line)) ??
    raw.split(/\r?\n/).find((line) => /^\d+ (passed|failed)/.test(line)) ??
    raw
      .split(/\r?\n/)
      .reverse()
      .find((line) => /\b\d+ passed\b|\b\d+ failed\b/.test(line)) ??
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
  const args = [resolve(here, "mutate_format_actions.mjs"), id, ...(restore ? ["restore"] : [])];
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

// 前置状态闸：停在上一次被中断的变异态时当场退出，不白跑十几轮测试。
const preflight = invokeMutation("check");
if (preflight.status !== 0) {
  console.error(`[mutate] 前置状态闸未通过：\n${preflight.out}`);
  process.exit(1);
}
console.log("[mutate] 前置状态闸通过：16 个 SEED 全部在位");
rmSync(scratch, { force: true });

const SPECS = [PYTEST_SPEC, CORE_SPEC, ADDIN_SPEC, WPS_SPEC];
console.log("[mutate] 基线（未变异）应为 GREEN…");
for (const spec of SPECS) {
  const result = runSpec(spec);
  const ok = result.status === 0 && result.summary && result.passed >= spec.minTests;
  console.log(
    `[mutate]   基线 ${spec.name}: status=${result.status} 摘要="${result.summary}" ${
      ok ? "OK" : "不合格"
    }`
  );
  if (!ok) {
    console.error(`[mutate] 基线就不绿或没跑到预期用例数（要求 >= ${spec.minTests}），中止。`);
    process.exit(1);
  }
}

const failures = [];
for (const { id, spec } of MUTATIONS) {
  const applied = invokeMutation(id, false);
  if (applied.status !== 0) {
    failures.push(`${id} apply 失败：${applied.out || applied.error}`);
    console.error(`[mutate] ${id} apply 失败：${applied.out || applied.error}`);
    continue;
  }
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
if (!failures.length) console.log("[mutate] 五个文件全部逐字节还原");

console.log("[mutate] 还原后复跑全绿…");
for (const spec of SPECS) {
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
  `\n[mutate] 全部通过：${MUTATIONS.length} 个变异各自让目标用例变红，五个文件逐字节还原，复跑全绿。`
);
