/**
 * 反向变异驱动：证明 A0 补充修复里**客户端三处接线**真的有判别力。
 *
 * 四条判据缺一不算过（与仓库既有驱动同口径）：
 * 1. 未变异时基线 GREEN（退出码 0 + 有效摘要 + 用例数达标）；
 * 2. apply 后必须变红，且**点名那条用例**真的在失败列表里；
 * 3. restore 后文件**逐字节一致**（sha256 比对）；
 * 4. 还原后复跑全绿。
 *
 * 只卡"退出码非 0"是不行的：一行写崩了 TSX 也是非 0，那是运行期错误，
 * 不是断言在起作用。所以还要读到失败原文里它**自己那一处**。
 *
 * 用法：
 *   node scripts/run_continuation_bridge_mutations.mjs            # 跑 C1/C2/C3
 *   node scripts/run_continuation_bridge_mutations.mjs C1         # 只跑一条（调试用）
 */
import { spawnSync } from "node:child_process";
import { openSync, closeSync, readFileSync, mkdirSync, existsSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");
const scratchDir = resolve(repoRoot, "outputs");
const scratch = resolve(scratchDir, "continuation-bridge-mutation.log");
const vitest = resolve(repoRoot, "node_modules", "vitest", "vitest.mjs");
const mutateScript = resolve(here, "mutate_continuation_bridge.mjs");

mkdirSync(scratchDir, { recursive: true });

const CASES = {
  C1: {
    label: "WPS 不采集必要来源指纹",
    app: "wps-addin",
    file: "src/WpsAgentPanel.tsx",
    tests: ["src/wpsPlanningPauseCard.test.tsx"],
    // 该文件基线 13 条（其中最后一条是本轮新增的来源指纹判据）。
    minTests: 13,
    // 失败原文里必须出现的用例标题片段。
    expectTitle: "服务端要求核实来源时，继续请求必须带上此刻重读的指纹"
  },
  C2: {
    label: "继续时不重新初始化桥（旧 binding 当成活桥）",
    app: "excel-addin",
    file: "src/hooks/useSdkToolBridge.ts",
    tests: ["src/hooks/useSdkToolBridge.test.tsx", "src/messageFirstRouting.test.tsx"],
    // 14（钩子，本轮 +3）+ 22（路由，本轮 +1）= 36。
    minTests: 36,
    expectTitle: "end() 之后绑定还在但桥不再服务；ensureLive 用同一任务重建后才继续读表"
  },
  C3: {
    label: "不转发 bridgeTool 指针",
    app: "wps-addin",
    file: "src/wpsReadonlyAgent.ts",
    tests: ["src/wpsReadonlyAgent.test.ts"],
    minTests: 47,
    expectTitle: "续跑期间必须把 bridgeTool 指针透传出去"
  }
};

function runVitest(kase) {
  const appDir = resolve(repoRoot, "apps", kase.app);
  const fd = openSync(scratch, "w");
  const result = spawnSync(
    process.execPath,
    [vitest, "run", ...kase.tests],
    { cwd: appDir, stdio: ["ignore", fd, fd] }
  );
  closeSync(fd);
  const raw = readFileSync(scratch, "utf8").replace(/\x1b\[[0-9;]*m/g, "");
  const summary =
    raw.split(/\r?\n/).find((line) => /Tests\s+\d+ (passed|failed)/.test(line)) ?? "";
  const passed = /Tests\s+(\d+) passed/.exec(summary);
  const failed = /Tests\s+\d+ failed/.exec(summary);
  return {
    status: result.status,
    raw,
    summary: summary.trim(),
    passed: passed ? Number(passed[1]) : 0,
    failed: failed ? Number(/(\d+) failed/.exec(summary)[1]) : 0
  };
}

function mutate(which, restore) {
  const fd = openSync(scratch, "w");
  const result = spawnSync(
    process.execPath,
    [mutateScript, which, ...(restore ? ["restore"] : [])],
    { cwd: repoRoot, stdio: ["ignore", fd, fd] }
  );
  closeSync(fd);
  return {
    status: result.status,
    out: readFileSync(scratch, "utf8").trim(),
    error: result.error ? String(result.error.code) : null
  };
}

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

const selected = process.argv[2] ? [process.argv[2]] : Object.keys(CASES);
for (const which of selected) {
  if (!CASES[which]) {
    console.error(`未知变异：${which}`);
    process.exit(2);
  }
}

let failures = 0;
for (const which of selected) {
  const kase = CASES[which];
  const target = resolve(repoRoot, "apps", kase.app, kase.file);
  if (!existsSync(target)) {
    console.error(`[mutate] ${which} 目标文件不存在：${target}`);
    failures += 1;
    continue;
  }
  const baselineSha = sha256(target);
  console.log(`\n[mutate] ${which} ${kase.label}`);
  console.log(`[mutate] 目标 ${kase.file} 基线 sha256=${baselineSha}`);

  const baseline = runVitest(kase);
  if (baseline.status !== 0 || baseline.passed < kase.minTests) {
    console.error(
      `[mutate] ${which} 基线就不绿（status=${baseline.status}, 摘要=${baseline.summary || "无"}）`
    );
    failures += 1;
    continue;
  }
  console.log(`[mutate] 基线 GREEN：${baseline.summary}`);

  const applied = mutate(which, false);
  if (applied.status !== 0 || applied.error) {
    console.error(`[mutate] ${which} apply 失败：${applied.out}`);
    failures += 1;
    continue;
  }

  const run = runVitest(kase);
  const named = run.raw.includes(kase.expectTitle);
  const red =
    run.status !== 0 && run.failed >= 1 && run.failed <= 8 && named;
  console.log(
    `[mutate] ${which} → ${red ? "RED(OK)" : "GREEN(NG!)"} ${run.summary || "(无摘要)"}` +
      ` named=${named}`
  );
  if (!red) {
    console.error(
      `  ${which} 未按预期变红（status=${run.status} failed=${run.failed} named=${named}）`
    );
    // 把失败清单摆出来，便于定点：不点名说明断言抓的不是这一处。
    for (const line of run.raw.split(/\r?\n/)) {
      if (line.includes("×") || line.includes("FAIL ")) console.error(`  ${line.trim()}`);
    }
    failures += 1;
  }

  const restored = mutate(which, true);
  if (restored.status !== 0 || restored.out === "ALREADY-CLEAN") {
    console.error(`[mutate] ${which} restore 失败：${restored.out}`);
    failures += 1;
    continue;
  }
  if (sha256(target) !== baselineSha) {
    console.error(`[mutate] ${which} 还原后与基线不一致（sha256 变了）`);
    failures += 1;
  }
}

if (failures > 0) {
  console.error(`\n反向变异未通过：${failures} 处问题。`);
  process.exit(1);
}

// 第四条判据：全部还原之后再复跑一次，确认绿不是"某次 restore 恰好没生效"造成的。
console.log("\n[mutate] 全部还原后复跑…");
for (const which of selected) {
  const kase = CASES[which];
  const run = runVitest(kase);
  if (run.status !== 0 || run.passed < kase.minTests) {
    console.error(`[mutate] ${which} 还原后复跑未全绿：${run.summary}`);
    failures += 1;
  } else {
    console.log(`[mutate] ${which} 复跑 GREEN：${run.summary}`);
  }
}
if (failures > 0) {
  console.error(`\n反向变异未通过：${failures} 处问题。`);
  process.exit(1);
}
console.log("\n反向变异通过：C1/C2/C3 各自变红并点名到对应用例，还原逐字节一致，复跑全绿。");
