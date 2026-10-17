// 反向变异驱动：通用合成案例（要求5）关键守卫。
//
// 每个变异：基线 GREEN → apply → 跑指定用例（须变红）→ restore → 复跑（须全绿）
//          → **逐字节**比对源码与变异前一致（不是"搜不到残留标记"就算干净）。
//
// 用法：node scripts/run_synthetic_case_mutations.mjs
//
// ── 判定纪律（四道必须同时满足，缺一不算通过）────────────────────────────
// 1. **退出码为 0**：`execFileSync` 拿到非零退出码就是失败，不能只看 stdout。
// 2. **有有效测试摘要**：必须解析出 `Tests <n> passed ...` 这样的真实摘要行；
//    解析不出来（进程崩了、没跑用例）一律判失败 —— 不能把"没摘要"当"没失败"。
// 3. **实际执行了预期用例**：`passed + failed >= minTests`。用例被整文件跳过
//    （例如 describe 条件不成立）时摘要会显示 0 个用例，那不能算"RED 成功"。
// 4. **逐字节还原**：变异前先记 sha256，还原后再算一次，必须相等。
//
// apply / restore 命令自身失败（非零退出码、SEED 未命中）**立即报错终止**，
// 绝不继续跑后面的用例 —— 否则会拿着半变异/未变异的源码得出假结论。
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";

const REPO = "E:\\codex\\excel-bro";
const CORE = `${REPO}\\packages\\core`;
const EXCEL = `${REPO}\\apps\\excel-addin`;
const VITEST = `${REPO}\\node_modules\\vitest\\vitest.mjs`;
const MUTATOR = `${REPO}\\scripts\\mutate_synthetic_case_guards.mjs`;
const NO_COLOR = /\x1b\[[0-9;]*m/g;

const CASES = [
  // S1：目标表不在声明来源里 —— 合成案例第 3 段的负向用例必须变红。
  {
    id: "S1",
    cwd: EXCEL,
    files: ["src/syntheticResultEndToEnd.test.tsx"],
    minFailed: 1,
    minTests: 8
  },
  // S2：来源快照与 scope 不一致 —— 由 core 结果流用例判别（两端各一）。
  {
    id: "S2",
    cwd: CORE,
    files: ["src/pythonResultFlow.test.ts"],
    minFailed: 1,
    minTests: 4
  },
  // S3 单独拆没判别力（captureSources 自己也会抛）；必须与 S3b 同时拆。
  {
    id: "S3+S3b",
    cwd: EXCEL,
    files: ["src/syntheticResultEndToEnd.test.tsx"],
    minFailed: 1,
    minTests: 8,
    combined: ["S3", "S3b"]
  },
  // S4：缺字段不跳过 —— 合成案例第 0 段"缺字段整张跳过"用例必须变红。
  {
    id: "S4",
    cwd: EXCEL,
    files: ["src/syntheticResultEndToEnd.test.tsx"],
    minFailed: 1,
    minTests: 8
  }
];

function run(command, args, cwd) {
  try {
    const output = execFileSync(command, args, {
      cwd,
      encoding: "utf8",
      stdio: ["ignore", "pipe", "pipe"]
    });
    return { output, status: 0 };
  } catch (error) {
    return {
      output: `${error.stdout ?? ""}${error.stderr ?? ""}`,
      status: typeof error.status === "number" ? error.status : 1
    };
  }
}

/**
 * 跑用例并**严格解析**摘要。
 *
 * 返回 `{ ok, passed, failed, total, summary, reason }`：
 * - `ok` 为 false 时 `reason` 说明为什么这次结果不可采信（退出码异常 / 无摘要 / 零用例）。
 * - 摘要行带 ANSI 色码，解析前先剥掉（否则全部误判"无摘要"）。
 */
function vitest(cwd, files, minTests) {
  const result = run(process.execPath, [VITEST, "run", ...files], cwd);
  const plain = result.output.replace(NO_COLOR, "");
  const summaryLine = plain
    .split("\n")
    .filter((line) => /^\s*Tests\s+/.test(line))
    .at(-1);
  if (!summaryLine) {
    return {
      ok: false,
      passed: 0,
      failed: 0,
      total: 0,
      summary: "(无摘要)",
      reason: `未解析到测试摘要行（退出码 ${result.status}）`
    };
  }
  const passed = Number(summaryLine.match(/(\d+)\s+passed/)?.[1] ?? "0");
  const failed = Number(summaryLine.match(/(\d+)\s+failed/)?.[1] ?? "0");
  const total = passed + failed;
  const summary = summaryLine.trim().replace(/^Tests\s+/, "");
  if (total < minTests) {
    return {
      ok: false,
      passed,
      failed,
      total,
      summary,
      reason: `实际只执行了 ${total} 个用例（要求至少 ${minTests}）`
    };
  }
  if (result.status !== 0 && failed === 0) {
    // 退出码非零却一个 failed 都没有：进程层面出了问题，结果不可采信。
    return {
      ok: false,
      passed,
      failed,
      total,
      summary,
      reason: `退出码 ${result.status} 但无失败用例（进程异常）`
    };
  }
  return { ok: true, passed, failed, total, summary, reason: "" };
}

/**
 * apply / restore 变异。**失败立即抛错**：非零退出码或 SEED 未命中都不许继续，
 * 否则后续用例是拿半变异源码跑出来的假结论。
 */
function mutate(id, ...rest) {
  const label = rest.includes("restore") ? "restore" : "apply";
  const result = run(process.execPath, [MUTATOR, id, ...rest], REPO);
  const output = result.output.trim();
  if (result.status !== 0) {
    throw new Error(
      `变异 ${id} ${label} 失败（退出码 ${result.status}）：\n${output}`
    );
  }
  const first = output.split("\n")[0] ?? "";
  // 应用必须真的命中并改写；还原必须真的回到 SEED（ALREADY-CLEAN 是合法幂等结果）。
  if (label === "apply" && !/apply:\s*OK/.test(first)) {
    throw new Error(`变异 ${id} apply 未生效：${first}`);
  }
  if (label === "restore" && !/restore:\s*(OK|ALREADY-CLEAN)/.test(first)) {
    throw new Error(`变异 ${id} restore 未生效：${first}`);
  }
  return first;
}

function cleanAll() {
  const result = run(process.execPath, [MUTATOR, "restore-all"], REPO);
  if (result.status !== 0) {
    throw new Error(`restore-all 失败（退出码 ${result.status}）：${result.output.trim()}`);
  }
}

const SOURCES = [
  `${REPO}\\packages\\core\\src\\pythonResult.ts`,
  `${REPO}\\packages\\core\\src\\pythonData.ts`,
  `${EXCEL}\\src\\officePythonResult.ts`,
  `${EXCEL}\\src\\dataTools.ts`
];

function sha256(file) {
  return createHash("sha256").update(readFileSync(file)).digest("hex");
}

/** 变异前的源码指纹（每个文件一份）。 */
function snapshotSources() {
  return Object.fromEntries(SOURCES.map((file) => [file, sha256(file)]));
}

/**
 * 还原后**逐字节**比对：源码必须与变异前完全一致。
 * 只"搜不到 `变异S` 标记"是不够的 —— 那会漏掉被截断、被重复替换、换行符被改写的残体。
 */
function assertByteIdentical(before) {
  const drifted = SOURCES.filter((file) => sha256(file) !== before[file]);
  if (drifted.length) {
    throw new Error(
      `还原后源码与变异前不逐字节一致：\n${drifted
        .map((file) => `  ${file}\n    before=${before[file]}\n    after =${sha256(file)}`)
        .join("\n")}`
    );
  }
}

const rows = [];
let failedOverall = false;
cleanAll();
const pristine = snapshotSources();
try {
  for (const entry of CASES) {
    const baseline = vitest(entry.cwd, entry.files, entry.minTests);
    if (!baseline.ok || baseline.failed !== 0) {
      rows.push([
        entry.id,
        "基线",
        baseline.summary,
        `NG：基线不是 GREEN${baseline.reason ? `（${baseline.reason}）` : ""}`
      ]);
      failedOverall = true;
      continue;
    }
    const appliedIds = entry.combined ?? [entry.id];
    const applied = appliedIds.map((id) => mutate(id));
    const mutated = vitest(entry.cwd, entry.files, entry.minTests);
    appliedIds.forEach((id) => mutate(id, "restore"));
    const restored = vitest(entry.cwd, entry.files, entry.minTests);
    // 逐字节还原校验：这是"还原到位"的唯一证据。
    assertByteIdentical(pristine);

    const redOk = mutated.ok && mutated.failed >= entry.minFailed;
    const greenOk = restored.ok && restored.failed === 0 && restored.passed > 0;
    if (!redOk || !greenOk) failedOverall = true;
    const detail = [
      redOk ? "" : `RED 不可采信（${mutated.reason || `failed=${mutated.failed}<${entry.minFailed}`}）`,
      greenOk ? "" : `GREEN 不可采信（${restored.reason || "还原后非全绿"}）`
    ]
      .filter(Boolean)
      .join("；");
    rows.push([
      entry.id,
      applied.join(" / "),
      mutated.summary,
      redOk && greenOk
        ? `RED(OK) → 还原 ${restored.summary} → 源码 sha256 一致`
        : `判别力不足：${detail || "RED/GREEN 未达标"}`
    ]);
  }
} finally {
  cleanAll();
  // 收尾再验一次：即使中途 throw，也要确认源码回到变异前。
  try {
    assertByteIdentical(pristine);
  } catch (error) {
    console.error(String(error instanceof Error ? error.message : error));
    failedOverall = true;
  }
}

for (const [id, applied, red, verdict] of rows) {
  console.log(`${id}\t${applied}\t${red}\t${verdict}`);
}
console.log(failedOverall ? "RESULT: FAIL" : "RESULT: OK");
process.exit(failedOverall ? 1 : 0);
