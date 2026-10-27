/**
 * A1-E2「结果引用过期恢复全链」的**一条可复跑命令**（`npm run test:a1-e2`）。
 *
 * 它按真实顺序把两半串起来：
 *
 * 1. **生产侧**（`scripts/drive_real_sdk_result_expiry.py`）：真实模型（Agent SDK）+
 *    真实服务端 + 真实结果登记器 + 真实 Docker 隔离执行，跑
 *    `读表 → 重算 → 新引用 → 计划编译`，途中由 A1-E1 的入口把**旧引用**精确失效。
 *    本脚本**不**自己跑它（要真实模型预算与已启动的服务端），只消费它写出的工件。
 * 2. **消费侧**：Excel 端真实 `<App/>`、WPS 端真实 `WpsAgentPanel`，各自消费
 *    **自己那一轮**的工件，完成"执行一次 → 回读 → 审阅（确认/放弃）"。
 *
 * 为什么生产侧不在这里跑：它要真模型与真容器，跑一次几十秒到几分钟，且需要
 * `EXCEL_BRO_RESULT_REF_FAULT=1` 的服务端实例。把它硬塞进这条命令会让"消费侧
 * 用例"变成一个不可控的长过程。所以本命令的诚实口径是：
 *
 * - 工件不存在 ⇒ 打印「未完成：生产侧没有产出工件」并以非零退出（**不**用旧工件
 *   或 mock 顶替）；
 * - 工件存在但自洽性检查不过（引用不是 2 个、旧引用没 404、计划没引用新引用）
 *   ⇒ 同样非零退出；
 * - 消费侧任一端用例被跳过或失败 ⇒ 非零退出（"跳过"不等于通过）。
 */
import { spawnSync } from "node:child_process";
import { existsSync, readFileSync, statSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");
const isWin = process.platform === "win32";

/** 两端各自的工件：Excel 端不许拿 WPS 端那一轮替自己作证。 */
const ARTIFACTS = {
  wps: resolve(repoRoot, "outputs", "a1-e2-recovery.json"),
  excel: resolve(repoRoot, "outputs", "a1-e2-recovery-excel.json")
};

/**
 * 「来源在分析期间变化」这一分支的工件（两端各一份）。
 *
 * 它与主场景的工件**分开校验**：主场景证明"过期 → 重算 → 新引用 → 只写一次"，
 * 这一份证明"模型没有再次读取时，来源变更只能靠写前复核拦住"。两份工件缺一份，
 * 这一轮就没结论 —— 不做"有没有都行"的软判。
 */
const SOURCE_CHANGE_ARTIFACTS = {
  wps: resolve(repoRoot, "outputs", "a1-e2-source-change.json"),
  excel: resolve(repoRoot, "outputs", "a1-e2-source-change-excel.json")
};

/**
 * 「模型**确实**再读了一次」的工件（B 路，两端各一份）。
 *
 * 与上面两份**分开校验**：A 路证明"模型没再读时只能靠写前复核"；这一份证明
 * "模型再读时被生产桥的授权/来源复核以 `source_changed` 拒交值"。它刻意**不启用**
 * 引用过期 —— 旧引用仍然有效，计划仍可能被编译出来，于是"迟到旧计划不得写入"
 * 这件事才有东西可验。缺一份即未完成，不使用旧工件或替身。
 */
const REREAD_ARTIFACTS = {
  wps: resolve(repoRoot, "outputs", "a1-e2-source-change-reread.json"),
  excel: resolve(repoRoot, "outputs", "a1-e2-source-change-reread-excel.json")
};

/** 旧来源的汇总值 / 改后来源的汇总值：只作人可读的旁证，不当硬判据。 */
const OLD_TOTALS = [400, 1000];
const NEW_TOTALS = [550, 1150];

function fail(message, code = 1) {
  console.error(`[a1-e2] ${message}`);
  process.exit(code);
}

function checkSourceChangeArtifact(path, expectedHost) {
  if (!existsSync(path)) {
    fail(
      `未完成：来源变化分支没有产出 ${path}。用\n` +
        `  python scripts/drive_real_sdk_result_expiry.py --scenario source-change --host ${expectedHost} --out ${path}\n` +
        `再跑一轮；本命令不使用旧工件或替身。`
    );
  }
  const artifact = JSON.parse(readFileSync(path, "utf8"));
  const problems = [];
  const change = artifact.sourceChange;
  if (artifact.scenario !== "source-change") problems.push("工件场景不是 source-change");
  if (!change) problems.push("没有来源变更记录：确定性屏障没有生效");
  else {
    if (change.fingerprintBefore === change.fingerprintAfter) {
      problems.push("改动前后来源指纹相同：改的那一格不在指纹覆盖范围内，等于没改");
    }
    if (change.before === change.after) problems.push("改动前后单元格值相同");
  }
  const prepareCount = artifact.bridge?.prepareCount ?? 0;
  if (prepareCount < 1) problems.push("一次成功交值都没有：模型没读到来源，这条分支什么也证明不了");
  if (change && prepareCount >= 1) {
    // 变更只能发生在"首次交值之后"：早于它的话，模型读到的就是新值。
    const served = artifact.bridge?.served ?? [];
    const firstRead = served.find((item) => item.tool === "prepare_python_data" && !item.refused);
    if (!firstRead || (change.afterServedCall ?? 0) < firstRead.index) {
      problems.push("来源变更早于首次交值：模型读到的其实是新值，前提不成立");
    }
  }
  const basis = artifact.planSourceBasis;
  if (basis !== "old" && basis !== "new") problems.push(`planSourceBasis=${basis}（无法判定计划依据哪一版来源）`);
  if (artifact.hostWritesDuringPlanning !== 0) problems.push("规划期发生了宿主写入");
  if (artifact.workbook?.hostId !== expectedHost) {
    problems.push(`工件宿主是 ${artifact.workbook?.hostId}，不是 ${expectedHost}`);
  }
  if (problems.length > 0) {
    fail(`来源变化工件不自洽（${path}）：\n  - ${problems.join("\n  - ")}`);
  }
  const numbers = artifact.planNumbers ?? [];
  const has = (value) => numbers.some((item) => Math.abs(item - value) < 1e-9);
  const refusals = (artifact.bridge?.refusals ?? []).length;
  const planActions = artifact.planActionCount ?? 0;
  console.log(
    `[a1-e2] 来源变化 ${path}\n` +
      `[a1-e2]   turnId=${artifact.turnId} 宿主=${artifact.workbook.hostId}\n` +
      `[a1-e2]   改 ${change.sheet}!${change.cell} ${change.before} → ${change.after}` +
      `（在第 ${change.afterServedCall} 次桥调用之后）\n` +
      `[a1-e2]   桥交值 ${prepareCount} 次；被拒 ${refusals} 次；计划动作 ${planActions} 个\n` +
      `[a1-e2]   计划依据=${basis}（old = 只拿到过变更前的值）；` +
      `计划数值含旧汇总 ${OLD_TOTALS.filter(has).join("/") || "无"}、` +
      `含新汇总 ${NEW_TOTALS.filter(has).join("/") || "无"}`
  );
  return artifact;
}

/**
 * B 路工件（模型确实再读一次）的自洽校验。任何一条不过都判未完成。
 *
 * 判据只看**可复核的事实**：桥调用序列（序号/操作身份/裁决码）、账本里那次调用的
 * 终态、改动与读取的先后、计划依据、规划期写入次数。不根据模型说了什么推断。
 */
function checkReReadArtifact(path, expectedHost) {
  if (!existsSync(path)) {
    fail(
      `未完成：B 路（模型再次读取被 source_changed 拒绝）没有产出 ${path}。用\n` +
        `  python scripts/drive_real_sdk_result_expiry.py --scenario source-change-reread --host ${expectedHost} --out ${path}\n` +
        `再跑一轮；未覆盖该分支的轮次不能用别的分支顶替。`
    );
  }
  const artifact = JSON.parse(readFileSync(path, "utf8"));
  const problems = [];
  const change = artifact.sourceChange;
  const second = artifact.secondRead ?? {};
  const attempts = second.attempts ?? [];
  const served = artifact.bridge?.served ?? [];
  const events = artifact.events ?? [];

  if (artifact.scenario !== "source-change-reread") problems.push("工件场景不是 source-change-reread");
  if (artifact.fault?.armed !== false) problems.push("该场景不应启用引用过期（要单独隔离来源变化）");
  if (!change) problems.push("没有来源变更记录：确定性屏障没有生效");
  else {
    if (change.before === change.after) problems.push("改动前后单元格值相同");
    if (change.fingerprintBefore === change.fingerprintAfter) {
      problems.push("改动前后来源指纹相同：改的那一格不在指纹覆盖范围内，等于没改");
    }
  }

  const firstServed = served.find((item) => item.tool === "prepare_python_data" && !item.refused);
  if (!firstServed) problems.push("一次成功交值都没有：模型没读到来源，这条分支什么也证明不了");
  if (firstServed && change && (change.afterServedCall ?? 0) < firstServed.index) {
    problems.push("来源变更早于首次交值：模型读到的其实是新值，前提不成立");
  }

  // 分支判定：必须真的是 B 路（再次调用 + 被拒），不能是"A 路 + 一句结论"。
  if (artifact.branch !== "B") problems.push(`branch=${artifact.branch}（不是 B：模型没有再次读取，或再次读取没被拒）`);
  if (!second.attempted) problems.push("变更之后没有新的 prepare 调用");
  if (!second.refused) problems.push("再次读取没有被拒绝");
  if (second.code !== "source_changed") problems.push(`拒绝码是 ${second.code}，不是 source_changed`);
  if (second.sameScope !== true) {
    problems.push("第二次读取的来源范围与第一次不同：两次之间混进了别的变量");
  }
  if (attempts.length === 0) problems.push("secondRead.attempts 为空");
  for (const attempt of attempts) {
    if (attempt.operationId !== second.operationId) continue;
    if (!/^op_[0-9a-f]{32}$/.test(String(attempt.operationId ?? ""))) {
      problems.push(`操作身份不像生产桥下发的 id：${attempt.operationId}`);
    }
    if (!String(attempt.stepId ?? "").startsWith("bridge.prepare_python_data:")) {
      problems.push(`步骤身份不是桥调用步骤：${attempt.stepId}`);
    }
    const sequence = events.filter((event) => event.stepId === attempt.stepId);
    if (sequence.length === 0) problems.push("账本里没有这次被拒的读取：服务端没记下它");
    // source_changed 在服务端被记成 warning（"没做成"，不是产品崩溃）。
    else if (!sequence.some((event) => event.status === "warning")) {
      problems.push(`账本里这次读取的终态不是 warning：${sequence.map((e) => e.status).join("/")}`);
    }
  }

  if (artifact.planSourceBasis !== "old") {
    problems.push(`planSourceBasis=${artifact.planSourceBasis}（变了之后不该再有成功交值）`);
  }
  if ((artifact.hostWritesDuringPlanning ?? 0) !== 0) problems.push("规划期发生了宿主写入");
  if (artifact.workbook?.hostId !== expectedHost) {
    problems.push(`工件宿主是 ${artifact.workbook?.hostId}，不是 ${expectedHost}`);
  }
  if (problems.length > 0) {
    fail(`B 路工件不自洽（${path}）：\n  - ${problems.join("\n  - ")}`);
  }

  const planActions = artifact.planActionCount ?? 0;
  console.log(
    `[a1-e2] B 路 ${path}\n` +
      `[a1-e2]   turnId=${artifact.turnId} 宿主=${artifact.workbook.hostId}\n` +
      `[a1-e2]   改 ${change.sheet}!${change.cell} ${change.before} → ${change.after}` +
      `（在第 ${change.afterServedCall} 次桥调用之后）\n` +
      `[a1-e2]   首次交值 #${firstServed.index}；第二次读取 ${second.stepId} → ${second.code}\n` +
      `[a1-e2]   桥交值 ${artifact.bridge?.prepareCount} 次；被拒 ${attempts.length} 次；` +
      `计划动作 ${planActions} 个；规划期写入 ${artifact.hostWritesDuringPlanning} 次\n` +
      (planActions > 0
        ? `[a1-e2]   本轮模型仍提交了一份按旧来源算出的计划：写前复核由两端消费侧用例验（应 0 写入）\n`
        : `[a1-e2]   本轮模型没有提交计划（plan=null）：写前复核仍由 A 路工件 + 两端用例覆盖\n`)
  );
  return artifact;
}

function checkArtifact(path, expectedHost) {
  if (!existsSync(path)) {
    fail(
      `未完成：生产侧没有产出 ${path}。先按 docs/SDK_A1_RESULT_EXPIRY_2026-10-24.md 的\n` +
        `「两端最短原生操作步骤」跑一次驱动（--host ${expectedHost}），本命令不使用旧工件或替身。`
    );
  }
  const artifact = JSON.parse(readFileSync(path, "utf8"));
  const refs = artifact.registeredResultRefs ?? [];
  const problems = [];
  if (refs.length !== 2) problems.push(`登记引用应为 2 个，实际 ${refs.length}`);
  if (artifact.plan?.localResultRef !== refs[1]) {
    problems.push("计划引用的不是新引用（服务端没用重算后的结果编译）");
  }
  const oldProbe = (artifact.probes ?? []).find((probe) => probe.resultRef === refs[0]);
  if (!oldProbe || oldProbe.ok !== false || oldProbe.status !== 404) {
    problems.push("旧引用没有被拒绝（应 404）");
  }
  if (artifact.repeatOldRef && artifact.repeatOldRef.ok !== false) {
    problems.push("重复取旧引用居然可读（缓存/重试把它复活了）");
  }
  if ((artifact.planActionCount ?? 0) < 1) problems.push("计划里没有写入动作");
  if (artifact.workbook?.hostId !== expectedHost) {
    problems.push(`工件宿主是 ${artifact.workbook?.hostId}，不是 ${expectedHost}`);
  }
  const phases = (artifact.fault?.timeline ?? []).map((item) => item.phase);
  if (!phases.includes("bound") || phases[phases.length - 1] !== "consumed") {
    problems.push("故障时序不是 armed → bound → consumed");
  }
  if (problems.length > 0) {
    fail(`工件不自洽（${path}）：\n  - ${problems.join("\n  - ")}`);
  }
  console.log(
    `[a1-e2] 工件 ${path}\n` +
      `[a1-e2]   turnId=${artifact.turnId} 宿主=${artifact.workbook.hostId}\n` +
      `[a1-e2]   旧引用 ${refs[0]} → 404；新引用 ${refs[1]} → 可读\n` +
      `[a1-e2]   计划动作 ${artifact.planActionCount} 个；落盘 ${statSync(path).mtime.toISOString()}`
  );
  return artifact;
}

checkArtifact(ARTIFACTS.wps, "wps-et");
checkArtifact(ARTIFACTS.excel, "excel-office");
checkSourceChangeArtifact(SOURCE_CHANGE_ARTIFACTS.wps, "wps-et");
checkSourceChangeArtifact(SOURCE_CHANGE_ARTIFACTS.excel, "excel-office");
checkReReadArtifact(REREAD_ARTIFACTS.wps, "wps-et");
checkReReadArtifact(REREAD_ARTIFACTS.excel, "excel-office");

function runVitest(appDir, file, extraEnv) {
  const vitest = resolve(repoRoot, "node_modules", ".bin", isWin ? "vitest.cmd" : "vitest");
  const result = spawnSync(vitest, ["run", file], {
    cwd: resolve(repoRoot, "apps", appDir),
    env: { ...process.env, EB_A1_E2: "1", ...extraEnv },
    encoding: "utf8",
    // Windows 上 `node_modules/.bin/vitest.cmd` 是批处理，必须经 shell 起。
    shell: isWin
  });
  const output = `${result.stdout ?? ""}${result.stderr ?? ""}`;
  process.stdout.write(output);
  return { status: result.status ?? 1, output };
}

console.log("[a1-e2] 消费侧（Excel）：真实 <App/> 执行一次 → 回读 → 确认/撤销…");
const excel = runVitest("excel-addin", "src/resultExpiryRecoveryChain.test.tsx", {
  EB_A1_E2_ARTIFACT: "outputs/a1-e2-recovery-excel.json"
});
if (excel.status !== 0) fail("Excel 端消费侧用例未通过（见上）。");
if (/\b(\d+)\s+skipped\b/.test(excel.output) && !/\b0\s+skipped\b/.test(excel.output)) {
  fail("Excel 端有用例被**跳过** —— 跳过不等于通过。");
}

console.log("[a1-e2] 消费侧（WPS）：真实 WpsAgentPanel 执行一次 → 回读 → 确认/放弃…");
const wps = runVitest("wps-addin", "src/WpsAgentPanel.test.tsx", {});
if (wps.status !== 0) fail("WPS 端消费侧用例未通过（见上）。");

console.log(
  "[a1-e2] 连续链路完成：生产侧（真实模型 + 真实服务端 + 真实容器）与两端消费侧已接通。\n" +
    "[a1-e2] 注意：A1-E2 完成 ≠ A1 整体验收；其他缺口仍按计划表单独记录。"
);
