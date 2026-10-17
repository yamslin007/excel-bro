/**
 * 反向变异驱动：证明「当前文件卡 / 触发器不再显示日期」的判据**真有判别力**。
 *
 * 四条判据缺一不算过：
 * 1. 未变异时基线 GREEN（退出码 0 + 有效测试摘要 + 用例数达标）；
 * 2. 每个变异 apply 后必须**按声明的样子**变红 —— 不只"退出码非 0"，还得
 *    **恰好**红在它该红的用例上（`failing`），并且失败原文里出现**它自己那一处**的证据
 *    （`errorContains`）。这一条是这一轮的关键：H1 与 H4 红的都是同一条用例
 *    （"文件名里有日期也不显示"），只靠用例名根本分不出是**卡片**被塞了日期还是
 *    **触发器**被塞了日期 —— 那就等于两个调用点里永远有一个没被测到。
 *    H1 的失败原文里有 `报表日期 2026-03-15 · 当前工作表 Sheet2`（卡片那一行），
 *    H4 的是 `2026-03-15 · 当前表 Sheet2`（触发器那一行）。
 * 3. restore 后 `App.tsx` **逐字节一致**（sha256 比对）；
 * 4. 还原后复跑全绿。
 *
 * 为什么还要卡"红在哪几条"：这一轮的改动落在 JSX 里，写崩了 jsdom 一样会让 vitest 非 0 退出 ——
 * 那是运行期错误，不是判据在起作用。变异自身也不许靠"引用一个已删除的标识符"来变红：
 * H1/H4 都是**就地内联**日期识别，不依赖已被删掉的 `extractWorkbookDataPeriod`。
 *
 * 用法：node scripts/run_workbook_identity_card_mutations.mjs
 */
import { spawnSync } from "node:child_process";
import { openSync, closeSync, readFileSync, mkdirSync, rmSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");
const scratchDir = resolve(repoRoot, "outputs");
const scratch = resolve(scratchDir, "workbook-identity-card-mutation.log");
const vitest = resolve(repoRoot, "node_modules", "vitest", "vitest.mjs");
const mutateScript = resolve(here, "mutate_workbook_identity_card.mjs");

const APP_DIR = resolve(repoRoot, "apps", "excel-addin");
const APP_FILE = resolve(APP_DIR, "src", "App.tsx");
const TEST_FILE = "src/workbookIdentityCard.test.tsx";
/** 该文件共 2 条：没有日期 1 + 有日期 1。少一条就说明判据被删。 */
const MIN_TESTS = 2;

/** 用例短名 —— 与 `workbookIdentityCard.test.tsx` 里的 it(...) 标题逐字一致。 */
const NO_DATE_CASE = "文件名里没有日期";
const DATED_CASE = "文件名里有日期也不显示";

/** 变异 → 它应当（且仅应当）让哪几条用例变红，以及失败原文里必须出现的证据。 */
const MUTATIONS = {
  H1: {
    failing: [DATED_CASE],
    errorContains: "报表日期 2026-03-15 · 当前工作表 Sheet2",
    note: "把日期请回当前文件卡"
  },
  H2: {
    failing: [NO_DATE_CASE, DATED_CASE],
    errorContains: "文件名中未识别到报表日期 · 当前工作表 Sheet2",
    note: "退回否定式说明"
  },
  H3: {
    failing: [NO_DATE_CASE, DATED_CASE],
    errorContains: "not to be null",
    note: "整行被删掉"
  },
  H4: {
    failing: [DATED_CASE],
    errorContains: "2026-03-15 · 当前表 Sheet2",
    note: "把日期加回触发器"
  }
};

mkdirSync(scratchDir, { recursive: true });

/**
 * 用**文件描述符**而不是 pipe 接子进程输出：流水线里创建管道会被拦，
 * 而判据又必须读到真实的摘要行与失败原文。vitest 输出带 ANSI 色码，先剥掉。
 */
function runTests() {
  const fd = openSync(scratch, "w");
  const result = spawnSync(process.execPath, [vitest, "run", TEST_FILE], {
    cwd: APP_DIR,
    stdio: ["ignore", fd, fd]
  });
  closeSync(fd);
  const raw = readFileSync(scratch, "utf8").replace(/\x1b\[[0-9;]*m/g, "");
  const summary =
    raw.split(/\r?\n/).find((line) => /Tests\s+\d+ (passed|failed)/.test(line)) ?? "";
  const passed = /Tests\s+(?:(\d+) failed \| )?(\d+) passed/.exec(summary);
  const failed = /Tests\s+(\d+) failed/.exec(summary);
  // 文件级清单里失败用例的标记是 `×`（ANSI 已剥离），拿它的**短名**做比对。
  const failingTitles = raw
    .split(/\r?\n/)
    .filter((line) => /^\s*×\s+/.test(line))
    .map((line) => line.trim());
  return {
    status: result.status,
    summary: summary.trim(),
    passed: passed ? Number(passed[2]) : 0,
    failed: failed ? Number(failed[1]) : 0,
    failingTitles,
    raw
  };
}

/** 某条短名是不是**出现在失败清单里**（而不是只出现在别的什么提示文本里）。 */
function isFailing(run, shortName) {
  return run.failingTitles.some((line) => line.includes(shortName));
}

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

function mutate(which, restore) {
  const args = [mutateScript, which, ...(restore ? ["restore"] : [])];
  const fd = openSync(scratch, "w");
  const result = spawnSync(process.execPath, args, {
    cwd: repoRoot,
    stdio: ["ignore", fd, fd]
  });
  closeSync(fd);
  return { status: result.status, out: readFileSync(scratch, "utf8").trim() };
}

const baselineSha = sha256(APP_FILE);
console.log(`[mutate] 基线 ${APP_FILE} sha256=${baselineSha}`);

/* 前置状态闸：四个变异都必须完整在位或完整缺失。 */
console.log("[mutate] 先跑变异串定点闸…");
const preflight = mutate("check", false);
if (preflight.status !== 0) {
  console.error(preflight.out);
  console.error("[mutate] 变异串已失配，拒绝继续（先修 scripts/mutate_workbook_identity_card.mjs 的种子）。");
  process.exit(1);
}
console.log(preflight.out.split(/\r?\n/).filter(Boolean).join("  |  "));

console.log("[mutate] 基线（未变异）应为 GREEN…");
const baseline = runTests();
if (baseline.status !== 0 || !baseline.summary || baseline.passed < MIN_TESTS || baseline.failed !== 0) {
  console.error(`基线就不绿：status=${baseline.status}, 摘要=${baseline.summary || "无"}`);
  process.exit(1);
}
console.log(`[mutate] 基线 GREEN：${baseline.summary}`);

let failures = 0;
for (const [which, spec] of Object.entries(MUTATIONS)) {
  const applied = mutate(which, false);
  if (applied.status !== 0) {
    console.error(`[mutate] ${which} apply 失败：status=${applied.status} out=${applied.out}`);
    failures += 1;
    continue;
  }
  const run = runTests();
  // 「红了几条」+「红的正是声明的那几条」：条数相等 + 每条都在场 ⇒ 失败集合恰好等于声明集合，
  // 不会出现"该红的没红、不该红的替它红了"。（两条用例的短名互不为子串，比对上不会串。）
  const titlesOk =
    run.status !== 0 &&
    run.failed === spec.failing.length &&
    spec.failing.every((name) => isFailing(run, name));
  const evidenceOk = run.raw.includes(spec.errorContains);
  console.log(
    `[mutate] ${which}（${spec.note}）→ ${titlesOk && evidenceOk ? "RED(OK)" : "NOT-AS-DECLARED(NG!)"} ` +
      `${run.summary || "(无摘要)"}｜失败清单=${run.failingTitles.length}`
  );
  if (!titlesOk || !evidenceOk) {
    console.error(
      `  ${which} 期望恰好红在 ${spec.failing.length} 条 [${spec.failing.join(" / ")}]` +
        `、失败原文含「${spec.errorContains}」；实际 status=${run.status} failed=${run.failed}` +
        ` evidence=${evidenceOk}；见 ${scratch}`
    );
    failures += 1;
  }
  const restored = mutate(which, true);
  if (restored.status !== 0 || restored.out === "ALREADY-CLEAN") {
    console.error(`[mutate] ${which} restore 失败：${restored.out}`);
    failures += 1;
    continue;
  }
  if (sha256(APP_FILE) !== baselineSha) {
    console.error(`[mutate] ${which} 还原后 ${APP_FILE} 与基线不一致（sha256 变了）`);
    failures += 1;
  }
}

const finalSha = sha256(APP_FILE);
if (finalSha !== baselineSha) {
  console.error(`[mutate] 最终 ${APP_FILE} 与基线不一致：${finalSha} != ${baselineSha}`);
  failures += 1;
}

console.log("[mutate] 还原后复跑应为 GREEN…");
const after = runTests();
if (after.status !== 0 || after.passed < MIN_TESTS || after.failed !== 0) {
  console.error(`[mutate] 还原后未全绿：${after.summary}`);
  failures += 1;
} else {
  console.log(`[mutate] 还原后 GREEN：${after.summary}`);
}
rmSync(scratch, { force: true });

if (failures > 0) {
  console.error(`反向变异未通过：${failures} 处问题。`);
  process.exit(1);
}
console.log(
  "反向变异通过：H2/H3 两条都红、H1 只红「有日期」那条（失败原文是卡片那一行）、" +
    "H4 只红「有日期」那条（失败原文是触发器那一行），还原逐字节一致，复跑全绿。"
);
