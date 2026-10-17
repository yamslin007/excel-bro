// 反向变异驱动：逐个应用变异 → 跑指定用例 → 还原 → 复核 GREEN。
//
// 为什么要有这个驱动：变异必须**单独施加**（同一文件上多个 SEED 会互相吞掉），
// 而且任何中途失败都可能把变异留在源码里。这里每步都做 apply/cleanup 兜底，
// 并在结束时**逐字节**比对源码与变异前一致。
//
// 用法：node scripts/run_office_undo_mutations.mjs
//
// ── 判定纪律（四道必须同时满足，缺一不算通过）────────────────────────────
// 1. **退出码为 0**：`execFileSync` 拿到非零退出码就是失败，不能只看 stdout。
// 2. **有有效测试摘要**：必须解析出 `Tests <n> passed ...`；解析不出来一律判失败。
// 3. **实际执行了预期用例**：`passed + failed >= minTests`（防整文件跳过）。
// 4. **逐字节还原**：变异前记 sha256，还原后再算一次，必须相等（不是搜残留标记）。
//
// apply / restore 自身失败（非零退出码、SEED 未命中）**立即报错终止**。
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";

const REPO = "E:\\codex\\excel-bro";
const EXCEL = `${REPO}\\apps\\excel-addin`;
const WPS = `${REPO}\\apps\\wps-addin`;
const VITEST = `${REPO}\\node_modules\\vitest\\vitest.mjs`;
const MUTATOR = `${REPO}\\scripts\\mutate_office_undo_and_same_sheet.mjs`;
const NO_COLOR = /\x1b\[[0-9;]*m/g;

/** 每个变异：在哪个目录跑、跑哪些文件、期望"至少 N 个用例失败"、至少跑了多少用例。 */
const CASES = [
  { id: "M1", cwd: EXCEL, files: ["src/excel.test.ts", "src/excelWriteEndToEnd.test.tsx"], minFailed: 1, minTests: 20 },
  { id: "M2", cwd: EXCEL, files: ["src/excel.test.ts", "src/excelWriteEndToEnd.test.tsx"], minFailed: 1, minTests: 20 },
  { id: "M3", cwd: EXCEL, files: ["src/excelWriteEndToEnd.test.tsx"], minFailed: 1, minTests: 5 },
  { id: "M4", cwd: WPS, files: ["src/wpsPythonResult.test.ts", "src/wpsPlanExecutor.test.ts"], minFailed: 1, minTests: 10 },
  { id: "M5", cwd: WPS, files: ["src/wpsPlanExecutor.test.ts"], minFailed: 1, minTests: 5 },
  // M6b 单独拆没判别力（M4 那道闸仍会抛）；必须与 M4 同时拆才算一次完整变异。
  {
    id: "M4+M6b",
    cwd: WPS,
    files: ["src/wpsPythonResult.test.ts"],
    minFailed: 1,
    minTests: 5,
    combined: ["M4", "M6b"]
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
 * 跑用例并**严格解析**摘要（`ok:false` 时 `reason` 说明为何不可采信）。
 * 摘要行带 ANSI 色码，解析前先剥掉。
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
 * 变异器的输出词形（**按各自实现写死**，不猜）。
 *
 * `mutate_office_undo_and_same_sheet.mjs` 说的是 `APPLIED M1 → <path>` /
 * `RESTORED：<path>`；另一个变异器说 `S1 apply: OK`。把词形与退出码**合并判定**，
 * 缺一不算成功：
 *
 * - 退出码非 0 ⇒ 立刻抛。该变异器有三条硬退出路径（`exit 2` 用法错、
 *   `exit 3` SEED 次数≠1、`exit 3` 替换为空操作），正好兜住"`replace` 静默命中
 *   0 次却报告成功"这类假成功。
 * - 退出码 0 但**没报告这次操作** ⇒ 也抛。防的是变异器将来改成"打了日志就退 0"。
 *
 * 只放松词形、不放松语义：`ALREADY-CLEAN` 只对 **restore** 合法（apply 时
 * 若报 ALREADY-CLEAN，说明源码本来就是变异的，RED 就成了伪造的）。
 */
const MUTATOR_LANGUAGE = {
  apply: /^(APPLIED\s+\S+|.*apply:\s*OK)\s/,
  restore: /^(RESTORED：|.*restore:\s*(OK|ALREADY-CLEAN))/
};

/** apply / restore 变异；失败**立即抛错**，绝不拿半变异源码继续跑。 */
function mutate(id, ...rest) {
  const label = rest.includes("restore") ? "restore" : "apply";
  const result = run(process.execPath, [MUTATOR, id, ...rest], REPO);
  const output = result.output.trim();
  if (result.status !== 0) {
    throw new Error(`变异 ${id} ${label} 失败（退出码 ${result.status}）：\n${output}`);
  }
  const first = output.split("\n")[0] ?? "";
  if (!MUTATOR_LANGUAGE[label].test(first)) {
    throw new Error(
      `变异 ${id} ${label} 未生效（退出码 0 但输出未报告 ${label}）：${first || "(空输出)"}`
    );
  }
  if (label === "apply" && /ALREADY-CLEAN/.test(first)) {
    throw new Error(`变异 ${id} apply 报告 ALREADY-CLEAN（源码已是变异态，RED 会是伪造的）：${first}`);
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
  `${EXCEL}\\src\\excel.ts`,
  `${WPS}\\src\\wpsPlanExecutor.ts`,
  `${WPS}\\src\\wpsPythonResult.ts`
];

function sha256(file) {
  return createHash("sha256").update(readFileSync(file)).digest("hex");
}

function snapshotSources() {
  return Object.fromEntries(SOURCES.map((file) => [file, sha256(file)]));
}

/** 还原后**逐字节**比对：必须与变异前完全一致。 */
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
