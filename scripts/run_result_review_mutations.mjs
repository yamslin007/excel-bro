/**
 * 结果审阅新逻辑的**反向变异**（证明这些断言真的有判别力）。
 *
 * 每条变异只做一件事：把**被测机制**拆掉，然后跑那条本该抓住它的用例。
 * 判红的标准不是"退出码非 0"——那连语法错误、采集失败都算 —— 而是：
 *   - 失败条数 > 0；
 *   - 失败用例名里**包含预期的业务断言**（`expectContains`）；
 *   - 失败原文里出现该业务断言的证据（`errorContains`）。
 *
 * 用法：
 *   node scripts/run_result_review_mutations.mjs check
 *   node scripts/run_result_review_mutations.mjs apply R1
 *   node scripts/run_result_review_mutations.mjs restore R1
 *   node scripts/run_result_review_mutations.mjs run R1
 */
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const STATE = join(root, ".tmp-result-review-mutations");
const isWin = process.platform === "win32";

/** 每条：拆掉哪个机制、跑哪个用例、必须看到哪句业务失败。 */
const MUTATIONS = {
  // R1：prepare 再次依赖"建副本"接口 —— 可选双视图又变回写入前置条件。
  R1: {
    file: "apps/wps-addin/src/WpsAgentPanel.tsx",
    seed: [
      "        return captureWpsBeforeResult(getWpsApplication(), prepared);"
    ].join("\n"),
    mutant: [
      "        const before = captureWpsBeforeResult(getWpsApplication(), prepared);",
      "        stageWpsNativeComparison(getWpsApplication(), prepared, true); // [变异R1] 执行前就建副本",
      "        return before;"
    ].join("\n"),
    cwd: "apps/wps-addin",
    test: "src/WpsAgentPanel.test.tsx",
    filter: "执行阶段不建副本；只有点",
    expectContains: "执行阶段不建副本",
    errorContains: "stageWpsNativeComparison"
  },
  // R2：双视图直接沿用"复制到的当前表"（= 执行后状态），不写回执行前快照。
  R2: {
    file: "apps/wps-addin/src/wpsNativeComparison.ts",
    seed: [
      "          if (typeof formula === \"string\" && formula.startsWith(\"=\")) cell.Formula = formula;",
      "          else cell.Value2 = value ?? null;"
    ].join("\n"),
    mutant: [
      "          void formula; void value; void cell; // [变异R2] 不写回执行前快照"
    ].join("\n"),
    cwd: "apps/wps-addin",
    test: "src/wpsBeforeComparison.test.ts",
    filter: "值是执行前的：副本按执行前快照重建",
    expectContains: "值是执行前的",
    errorContains: "旧值A"
  },
  // R3：确认变更时**再执行一次**计划。
  R3: {
    file: "packages/core/src/resultReview.ts",
    seed: "      await this.operations!.validateResult?.(this.receipt);",
    mutant: [
      "      await this.operations!.execute(this.receipt as never); // [变异R3] 确认时又写一遍",
      "      await this.operations!.validateResult?.(this.receipt);"
    ].join("\n"),
    cwd: "apps/wps-addin",
    test: "src/WpsAgentPanel.test.tsx",
    filter: "执行 → 单视图结果 → 确认：只写一次，确认不重放计划",
    expectContains: "只写一次",
    errorContains: "toBe"
  },
  // R5：遗留标记不再按"本次随机公式"认领，改成"整块区域全删" ⇒ 会误删别人的规则。
  R5: {
    file: "apps/wps-addin/src/wpsResultReviewMarkers.ts",
    seed: "        if (rule?.Formula1 === stored.formula) matches.push(rule);",
    mutant: "        if (rule) matches.push(rule); // [变异R5] 按区域全删，不核对归属",
    cwd: "apps/wps-addin",
    test: "src/wpsResultReviewMarkers.test.ts",
    filter: "同区域其它规则不碰",
    expectContains: "同区域其它规则不碰",
    errorContains: "=A1>0"
  },
  // R6：标记没落盘 ⇒ 创建到一半失败后，已创建的规则变成无主标记。
  R6: {
    file: "apps/wps-addin/src/wpsResultReviewMarkers.ts",
    seed: "      writeStored(marks, tracked);",
    mutant: "      void tracked; // [变异R6] 不逐条落盘",
    cwd: "apps/wps-addin",
    test: "src/wpsResultReviewMarkers.test.ts",
    filter: "创建失败且清理也失败",
    expectContains: "创建失败且清理也失败",
    errorContains: "ranges"
  },
  // R7：放弃时先清标记、失败就整体失败 ⇒ 恢复成功被"标记没清掉"改写。
  R7: {
    file: "packages/core/src/resultReview.ts",
    seed: "      await this.operations!.restore(this.receipt);",
    mutant: "      await this.clearMarker();\n      await this.operations!.restore(this.receipt); // [变异R7] 标记失败拦住恢复",
    cwd: "packages/core",
    test: "src/resultReview.test.ts",
    filter: "放弃时标记清理失败",
    expectContains: "放弃时标记清理失败",
    // 变异后 discard 先清标记、失败即整体返回 false：业务断言直接变红。
    errorContains: "expected false to be true"
  },
  // R8：确认清理失败后丢掉标记句柄 ⇒ 不可重试（界面只能报"不能重试"）。
  R8: {
    file: "packages/core/src/resultReview.ts",
    seed: [
      "        markerPending: notCurrent ? false : this.marker !== undefined",
      "      });"
    ].join("\n"),
    mutant: [
      "        markerPending: false",
      "      }); this.marker = undefined; // [变异R8] 标记失败后丢弃句柄"
    ].join("\n"),
    cwd: "packages/core",
    test: "src/resultReview.test.ts",
    filter: "确认时清理失败",
    expectContains: "确认时清理失败",
    errorContains: "markerPending"
  },
  // R9：当前性判据读回活动簿 ⇒ 双视图激活副本后被判成"换了工作簿"（真机症状）。
  R9: {
    file: "apps/wps-addin/src/WpsAgentPanel.tsx",
    seed: [
      "        String(wpsReadWorkbook(getWpsApplication())?.FullName ?? '') === prepared.authorization.workbookFullName;"
    ].join("\n"),
    mutant: [
      "        String(getWpsApplication().ActiveWorkbook?.FullName ?? '') === prepared.authorization.workbookFullName; // [变异R9] 又读活动簿"
    ].join("\n"),
    cwd: "apps/wps-addin",
    test: "src/WpsAgentPanel.test.tsx",
    filter: "并排后副本获得焦点",
    expectContains: "并排后副本获得焦点",
    // 判据失效后确认走不下去：until 超时，原文就是这条业务等待的说明。
    errorContains: "确认应成功"
  },
  // R10：身份失败也算"可重试清理标记" ⇒ 把"目标不对"说成"收尾没做完"。
  R10: {
    file: "packages/core/src/resultReview.ts",
    seed: "      const notCurrent = reason instanceof ReviewNotCurrentError;",
    mutant: "      const notCurrent = false; // [变异R10] 身份失败也按标记失败报",
    cwd: "packages/core",
    test: "src/resultReview.test.ts",
    filter: "当前性不成立",
    expectContains: "确认时当前性不成立",
    errorContains: "可重试清理标记"
  },
  // R11：成功终态不收起固定条 ⇒ 确认/放弃之后它继续占着输入框上方。
  R11: {
    file: "apps/excel-addin/src/components/ResultReviewBar.tsx",
    seed: "  if (state.phase === 'idle' || (settled && !needsAttention)) return null;",
    mutant: "  if (state.phase === 'idle') return null; // [变异R11] 成功终态不收起",
    cwd: "apps/wps-addin",
    test: "src/WpsAgentPanel.test.tsx",
    filter: "并排后副本获得焦点",
    expectContains: "并排后副本获得焦点",
    errorContains: "result-review-bar"
  },
  // R12：数据范围面板掉回输入框那一层（`.composer` 的后代）⇒ 又被挤/被裁/叠在输入区上。
  R12: {
    file: "apps/wps-addin/src/WpsAgentPanel.tsx",
    seed: [
      "      {scopeOverlay ? <div className=\"scope-layer\">{scopeOverlay}</div> : null}",
      "",
      "      <footer className=\"composer\">"
    ].join("\n"),
    mutant: [
      "      <footer className=\"composer\">",
      "        <div className=\"scope-composer-inline\">{scopeOverlay}</div> // [变异R12] 面板回到输入框那一层"
    ].join("\n"),
    cwd: "apps/wps-addin",
    // 必须跑在**真实**面板上：替身自己在 JSX 里写了 `.scope-layer`，抓不到这条变异。
    test: "src/WpsAgentPanel.test.tsx",
    filter: "真实面板：数据范围面板本体住在独立呈现层",
    expectContains: "面板本体不是 .composer 的后代",
    errorContains: "composer"
  },
  // R13：标题回到滚动容器里 ⇒ 列表一滚，标题与关闭按钮一起滚走。
  R13: {
    file: "apps/wps-addin/src/WpsProbePanel.tsx",
    seed: [
      "          <div className=\"scope-popover-header\">",
      "            <div>",
      "              <strong>这次对话使用哪些数据？</strong>",
      "              <span>",
      "                {draftSource === \"folder\"",
      "                  ? \"批量处理本地工作簿；数据由本机服务端读取\"",
      "                  : \"默认跟随你当前正在查看的工作表\"}",
      "              </span>",
      "            </div>",
      "            <button type=\"button\" aria-label=\"关闭数据范围\" onClick={discardScopeEdit}>×</button>",
      "          </div>",
      "          <div className=\"scope-popover-scroll\">"
    ].join("\n"),
    mutant: [
      "          <div className=\"scope-popover-scroll\">",
      "          <div className=\"scope-popover-header\">",
      "            <div>",
      "              <strong>这次对话使用哪些数据？</strong>",
      "              <span>",
      "                {draftSource === \"folder\"",
      "                  ? \"批量处理本地工作簿；数据由本机服务端读取\"",
      "                  : \"默认跟随你当前正在查看的工作表\"}",
      "              </span>",
      "            </div>",
      "            <button type=\"button\" aria-label=\"关闭数据范围\" onClick={discardScopeEdit}>×</button>",
      "          </div> // [变异R13] 标题回到滚动容器里"
    ].join("\n"),
    cwd: "apps/wps-addin",
    test: "src/wpsScopeDraft.test.tsx",
    filter: "面板住在独立呈现层里",
    expectContains: "面板住在独立呈现层里",
    errorContains: "scope-popover-header"
  },
  // R14：列工作表回到活动簿 ⇒ 副本获得焦点时范围核对去读副本（真机症状：面板开始读副本）。
  R14: {
    file: "apps/wps-addin/src/wpsWorkbookReader.ts",
    // 种子带上只属于这条规则的相邻注释：`wpsReadWorkbook(application)` 在本文件里
    // 还有别处，单行种子会让"还原"顺手改掉无关的那一行。
    seed: [
      "  const workbook = wpsReadWorkbook(application);",
      "  if (!workbook) throw new WpsAuthorizationError(\"WPS 表格中当前没有打开的工作簿。\");",
      "",
      "  const activeSheet = application.ActiveSheet ?? workbook.ActiveSheet;"
    ].join("\n"),
    mutant: [
      "  const workbook = application.ActiveWorkbook; // [变异R14] 又读活动簿",
      "  if (!workbook) throw new WpsAuthorizationError(\"WPS 表格中当前没有打开的工作簿。\");",
      "",
      "  const activeSheet = application.ActiveSheet ?? workbook.ActiveSheet;"
    ].join("\n"),
    cwd: "apps/wps-addin",
    test: "src/WpsAgentPanel.test.tsx",
    filter: "副本获得焦点时读数据范围",
    expectContains: "副本获得焦点时读数据范围",
    // 判据读的是"原表的名字"：换成活动簿后拿到的是副本名。
    errorContains: "工作簿1"
  },
  // R15：把"按计划重叠关系算样式"拆成"只看单个动作" ⇒ 同区域/部分重叠的写值或 setFont
  // 也会被叠上审阅底色，黄色盖住 setFill 的真实填充结果。
  R15: {
    file: "packages/core/src/reviewMarkStyle.ts",
    seed: '    const fillOwned = !rect || overlapsType("setFill");',
    mutant: '    const fillOwned = !rect || actionTouchesFill(action); // [变异R15] 退回单动作判据',
    cwd: "packages/core",
    test: "src/reviewMarkStyle.test.ts",
    filter: "写值区域与 setFill 区域部分重叠",
    expectContains: "写值区域与 setFill 区域部分重叠",
    errorContains: "border-only"
  },
  // R16：副本清理从"本次完整公式"退回到"前缀匹配" ⇒ 另一轮审阅的标记（同前缀、不同 nonce）被误删。
  R16: {
    file: "apps/wps-addin/src/wpsResultReviewMarkers.ts",
    seed: "        if (rule?.Formula1 === formula && typeof rule?.Delete === 'function') rule.Delete();",
    mutant: [
      "        const f = String(rule?.Formula1 ?? '');",
      "        if (f.startsWith('=N(\"EB_REVIEW_') && typeof rule?.Delete === 'function') rule.Delete(); // [变异R16] 退回前缀匹配"
    ].join("\n"),
    cwd: "apps/wps-addin",
    test: "src/wpsResultReviewMarkers.test.ts",
    filter: "副本清理只删本次完整公式",
    expectContains: "副本清理只删本次完整公式",
    errorContains: "EB_REVIEW_another_run_123"
  },
  // R17：Excel 标记也退回"只看单个动作" ⇒ 与 WPS 口径漂移，同区域重叠时 Excel 会叠底色。
  R17: {
    file: "apps/excel-addin/src/resultReviewMarkers.ts",
    seed: "        const style = styleByRange.get(`${action.sheet}\\u0000${action.range}`) ?? 'fill+border';",
    mutant: "        const style = action.type === 'setFill' ? 'border-only' : 'fill+border'; // [变异R17] Excel 退回单动作判据",
    cwd: "apps/excel-addin",
    test: "src/resultReviewMarkers.test.ts",
    filter: "写值区域与 setFill 区域部分重叠",
    expectContains: "写值区域与 setFill 区域部分重叠",
    errorContains: "null"
  },
  // R4：执行摘要卡绑定到错误的执行版本（消息 id 写死）⇒ 用户看到的"本次计划"不是本次的。
  R4: {
    file: "apps/wps-addin/src/WpsAgentPanel.tsx",
    seed: [
      "      if (pendingPlanMessageId) setExecutedPlan({ messageId: pendingPlanMessageId, plan, revised: planRevised });"
    ].join("\n"),
    mutant: [
      "      if (pendingPlanMessageId) setExecutedPlan({ messageId: \"stale-plan\", plan, revised: planRevised }); // [变异R4] 摘要绑错版本"
    ].join("\n"),
    cwd: "apps/wps-addin",
    test: "src/WpsAgentPanel.test.tsx",
    filter: "mirrors the Excel plan card hierarchy",
    expectContains: "mirrors the Excel plan card hierarchy",
    // 摘要卡绑错版本后，那条用例读到的是"没有摘要卡"，失败原文就是"空数组 vs 两个备注块"。
    errorContains: "expected [] to deeply equal"
  }
};

const sha256 = (buffer) => createHash("sha256").update(buffer).digest("hex");
const eolOf = (text) => (text.includes("\r\n") ? "\r\n" : "\n");
const toLf = (text) => text.split("\r\n").join("\n");
const fromLf = (text, eol) => (eol === "\r\n" ? text.split("\n").join("\r\n") : text);
const backup = (id) => join(STATE, `${id}.bak`);
const marker = (id) => `[变异${id}]`;

function target(id) {
  const spec = MUTATIONS[id];
  if (!spec) throw new Error(`未知变异：${id}`);
  return { spec, path: join(root, spec.file) };
}

function checkOne(id) {
  const { spec, path } = target(id);
  const text = toLf(readFileSync(path, "utf8"));
  const hits = text.split(spec.seed).length - 1;
  if (hits !== 1) throw new Error(`${id} SEED 命中 ${hits} 次（期望 1）：${spec.file}`);
  if (text.includes(marker(id)) && existsSync(backup(id)) && sha256(readFileSync(path)) !== sha256(readFileSync(backup(id)))) {
    throw new Error(`${id} 仍有未还原的变异（文件与备份不一致）`);
  }
  return true;
}

function apply(id) {
  const { spec, path } = target(id);
  const raw = readFileSync(path, "utf8");
  const eol = eolOf(raw);
  const text = toLf(raw);
  if (text.includes(marker(id))) throw new Error(`${id} 已经处于变异状态`);
  const hits = text.split(spec.seed).length - 1;
  if (hits !== 1) throw new Error(`${id} SEED 命中 ${hits} 次（期望 1）：${spec.file}`);
  mkdirSync(STATE, { recursive: true });
  writeFileSync(backup(id), raw);
  writeFileSync(path, fromLf(text.replace(spec.seed, spec.mutant), eol));
  console.log(`APPLIED ${id} → ${spec.file}`);
}

function restore(id) {
  const { spec, path } = target(id);
  if (!existsSync(backup(id))) throw new Error(`${id} 没有备份，无法还原`);
  const expected = readFileSync(backup(id));
  const current = readFileSync(path);
  // 还原按**磁盘当前内容**做替换，只在确实处于变异态时动手。
  if (sha256(current) === sha256(expected)) {
    console.log(`RESTORED ${id} → 已是基线（sha256 ${sha256(current).slice(0, 12)}）`);
    return;
  }
  const eol = eolOf(current.toString("utf8"));
  const text = toLf(current.toString("utf8"));
  if (!text.includes(spec.mutant)) throw new Error(`${id} 当前文件里找不到变异串，拒绝盲还原：${spec.file}`);
  writeFileSync(path, fromLf(text.replace(spec.mutant, spec.seed), eol));
  const after = readFileSync(path);
  if (sha256(after) !== sha256(expected)) throw new Error(`${id} 还原后与备份不一致：${spec.file}`);
  console.log(`RESTORED ${id} → ${spec.file}（sha256 ${sha256(after).slice(0, 12)} 一致）`);
}

function run(id) {
  const { spec } = target(id);
  const command = isWin ? "npx.cmd" : "npx";
  const result = spawnSync(command, ["vitest", "run", spec.test, "-t", spec.filter], {
    cwd: join(root, spec.cwd), encoding: "utf8", shell: isWin
  });
  // vitest 的摘要带 ANSI 颜色；不剥色码会把 "Tests 1 failed" 解析成"没失败"。
  const output = `${result.stdout ?? ""}\n${result.stderr ?? ""}`.replace(/\u001b\[[0-9;]*m/g, "");
  const summaryMatch = /Tests\s+(\d+) failed/.exec(output);
  const failedCount = summaryMatch ? Number(summaryMatch[1]) : 0;
  /*
   * 只认"预期那条用例失败"：别的用例变红不算这条机制的判别力。
   * 失败行形如 `   × <describe 路径> > <用例名>`，用例名通常在行尾，
   * 所以按"以 × 开头且包含该片段的行"判定，而不是要求片段紧跟在 × 之后。
   */
  const failedLines = output.split("\n").filter((line) => line.trimStart().startsWith("×"));
  const failed = failedCount > 0 && failedLines.some((line) => line.includes(spec.expectContains));
  const named = output.includes(spec.expectContains);
  const evidenced = output.includes(spec.errorContains);
  const summary = (output.match(/Tests\s+.*/) ?? ["(no summary)"])[0] + `；失败条数=${failedCount}`;
  if (!failed) {
    console.error(`❌ ${id} 未变红（${summary}）；断言对该机制没有判别力。`);
    process.exit(1);
  }
  if (!named) {
    console.error(`❌ ${id} 变红但不是预期用例（缺少「${spec.expectContains}」）。`);
    process.exit(1);
  }
  if (!evidenced) {
    console.error(`❌ ${id} 失败原文里没有业务证据「${spec.errorContains}」，可能是语法/采集错误。`);
    process.exit(1);
  }
  console.log(`✅ ${id} 变红且是预期用例：${summary}`);
}

const [command, id] = process.argv.slice(2);
try {
  if (command === "check") {
    Object.keys(MUTATIONS).forEach(checkOne);
    console.log(`CHECK OK 全部 ${Object.keys(MUTATIONS).length} 个 SEED 都在位（各命中 1 次）`);
  } else if (command === "apply") apply(id);
  else if (command === "restore") restore(id);
  else if (command === "run") run(id);
  else {
    console.error("用法：check | apply <ID> | restore <ID> | run <ID>");
    process.exit(2);
  }
} catch (error) {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
}