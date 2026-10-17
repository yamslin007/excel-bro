// 反向变异脚本：审阅副本生命周期（业务结果 / 关闭结果 / 磁盘清理三者不得互相冒充）。
//
// 用法：
//   node scripts/mutate_review_copy_lifecycle.mjs plan               # 打印每个 SEED 的 cwd/TEST/EXPECT（制表符分隔）
//   node scripts/mutate_review_copy_lifecycle.mjs check              # 只读自检：每个 SEED 是否仍各命中 1 次（重构后必跑）
//   node scripts/mutate_review_copy_lifecycle.mjs apply   <ID>       # 只施加一个 SEED
//   node scripts/mutate_review_copy_lifecycle.mjs restore <ID>       # 只还原一个 SEED（逐字节 sha256 校验）
//
// 铁律（与仓库其它 mutate_*.mjs 一致）：
//   - 一次只 apply/restore 一个 SEED；先归一化 LF 匹配，写回按**原文件**的 EOL。
//   - 起手前置状态闸：文件已含该 SEED 的变异标记就拒绝 apply。
//   - 装 SIGINT/SIGTERM/SIGHUP + uncaughtException/unhandledRejection 还原钩子，异常退出也要还原。
//   - 还原按 apply 时保存的**原始字节**回写，并核对 sha256 一致（不通过即非零退出）。
//   - 报错必须印文件名。
//   - **测试由 bash 直调**，本脚本不 spawn：本沙箱里 node 起子进程会 EBUSY。
//     每个 SEED 必须复跑 `plan` 给出的 TEST，确认 EXPECT 用例**实跑且变红**（四判据）。

import { createHash } from "node:crypto";
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import path from "node:path";

const REPO = path.resolve(import.meta.dirname, "..");
const STATE = path.join(REPO, ".tmp-review-copy-life-mutation");

const SEEDS = {
  // S1：拿去"写回失败默认保留副本"这道闸 ⇒ 排查材料会被自动销毁。
  S1: {
    file: "packages/core/src/reviewCopyLifecycle.ts",
    seed: "  if (input.business === 'write_failed' || input.business === 'partial_or_unknown') {",
    mutant: "  if (false) { // [变异S1] 写回失败不再强制保留副本",
    cwd: "packages/core",
    test: "src/reviewCopyLifecycle.test.ts",
    expect: "写回失败或结果不确定时保留副本不自动丢弃排查材料",
  },
  // S2：未知/未取到的关闭结果回落成 closed ⇒ 适配器出错被报成"已关闭"。
  S2: {
    file: "packages/core/src/reviewCopyLifecycle.ts",
    seed: "  const outcome: ReviewCopyCloseOutcome = template ? input.copy : 'failed';",
    mutant: "  const outcome: ReviewCopyCloseOutcome = template ? input.copy : 'closed'; // [变异S2] 未知结果冒充已关闭",
    cwd: "packages/core",
    test: "src/reviewCopyLifecycle.test.ts",
    expect: "未知的关闭结果按失败处理绝不报已关闭",
  },
  // S3：拿掉"必须位于专用临时目录之内"这一条 ⇒ 磁盘清理会越界。
  S3: {
    file: "packages/core/src/reviewCopyLifecycle.ts",
    seed: "  const within = path.toLowerCase().startsWith(`${base.toLowerCase()}/`);",
    mutant: "  const within = true; // [变异S3] 去掉越界校验",
    cwd: "packages/core",
    test: "src/reviewCopyLifecycle.test.ts",
    expect: "磁盘清理拒绝越出专用临时目录",
  },
  // S4：没有登记文件时也回 deleted ⇒ 未保存工作簿被当成磁盘文件"已删除"。
  S4: {
    file: "packages/core/src/reviewCopyLifecycle.ts",
    seed: "  if (input.tempFilePath === null) return 'none';",
    mutant: "  if (input.tempFilePath === null) return 'deleted'; // [变异S4] 未保存工作簿冒充磁盘文件",
    cwd: "packages/core",
    test: "src/reviewCopyLifecycle.test.ts",
    expect: "没有登记的临时文件时不得声称已删除文件",
  },
  // S5：不再要求"唯一已保存原表"在场就动手 ⇒ 可能关掉用户别的窗口（甚至整个 Excel 实例）。
  S5: {
    file: "server/app/excel_window_layout.py",
    seed: [
      "        originals = [info for info in observations if self._original(info, request)]",
      "        if len(originals) != 1:",
      '            return ExcelReviewCopyCloseResult(status="identity_lost", layoutVersion=4, copyClosed=False)',
    ].join("\n"),
    mutant: [
      "        originals = [info for info in observations if self._original(info, request)] or [copies[0]]  # [变异S5] 不核对审阅上下文",
      "        if len(originals) != 1:",
      '            return ExcelReviewCopyCloseResult(status="identity_lost", layoutVersion=4, copyClosed=False)',
    ].join("\n"),
    cwd: ".",
    test: "server/tests/test_excel_window_layout.py",
    expect: "test_close_never_acts_when_the_review_context_cannot_be_confirmed",
  },
  // S6：WM_CLOSE 投递失败也报 closed ⇒ 关不掉被说成关掉了。
  S6: {
    file: "server/app/excel_window_layout.py",
    seed: '        if not receipt.accepted:',
    mutant: '        if False:  # [变异S6] 忽略工作簿关闭失败',
    cwd: ".",
    test: "server/tests/test_excel_window_layout.py",
    expect: "test_close_reports_unconfirmed_close_instead_of_forcing_or_killing",
  },
  // S7：兼容模式忽略自相矛盾的 copyClosed ⇒ "说关掉了但其实没关"会被照单全收。
  S7: {
    file: "apps/excel-addin/src/excelNativeComparison.ts",
    seed: '    if (result.status === "closed" || result.status === "already_closed") return result.copyClosed ? result.status : "failed";',
    mutant: '    if (result.status === "closed" || result.status === "already_closed") return result.status; // [变异S7] 忽略 copyClosed',
    cwd: "apps/excel-addin",
    test: "src/excelNativeComparison.test.ts",
    expect: "maps the local bridge's close receipt onto the shared outcome without ever faking a close",
  },
  // S8：WPS 关闭前不再复核身份 ⇒ 用户另存过的副本会被当成我们的副本丢掉。
  S8: {
    file: "apps/wps-addin/src/wpsNativeComparison.ts",
    seed: '  return sameWorkbookIdentity(current, expected) ? null : "user_saved";',
    mutant: '  return null; // [变异S8] 不再复核副本身份',
    cwd: "apps/wps-addin",
    test: "src/wpsNativeComparison.test.ts",
    expect: "副本已被用户关闭 ⇒ already_closed（幂等）；被另存/替换 ⇒ user_saved 且不动它",
  },
  // S9：宿主拒绝关闭也返回 closed ⇒ 谎报已关闭。
  S9: {
    file: "apps/wps-addin/src/wpsNativeComparison.ts",
    seed: '  catch { return "failed"; }',
    mutant: '  catch { return "closed"; } // [变异S9] 宿主拒绝关闭也报已关闭',
    cwd: "apps/wps-addin",
    test: "src/wpsNativeComparison.test.ts",
    expect: "宿主拒绝关闭 / 返回 false ⇒ 既不谎报已关闭，也不把它说成业务失败",
  },
  // S10：写回失败/不确定时不再"保留副本并说明原因" ⇒ 排查材料被丢在现场之外。
  S10: {
    file: "apps/excel-addin/src/hooks/useExcelNativeComparison.ts",
    seed: [
      '      if (current.outcome === "write_failed" || current.outcome === "partial_or_unknown") {',
      "        publish(null);",
      '        setFeedback({ messageId, text: cleanupText(reviewCopyCleanupState({ business: "partial_or_unknown", copy: "not_attempted", tempFilePath: null })) });',
      "        return;",
      "      }",
    ].join("\n"),
    mutant: [
      "      if (false) { // [变异S10] 写回失败不再保留副本、也不说明原因",
      "        publish(null);",
      '        setFeedback({ messageId, text: cleanupText(reviewCopyCleanupState({ business: "partial_or_unknown", copy: "not_attempted", tempFilePath: null })) });',
      "        return;",
      "      }",
    ].join("\n"),
    cwd: "apps/excel-addin",
    test: "src/components/ExcelNativeReviewActions.test.tsx",
    expect: "a partial/failed execution cannot be blindly retried, and original execution evidence remains authoritative",
  },
  // S11：放弃路径直接关副本、不再核对"副本里有没有额外人工修改" ⇒ 用户的改动被静默丢弃。
  S11: {
    file: "apps/wps-addin/src/wpsNativeComparison.ts",
    seed: [
      "  const blocked = describeCopyDiscardSafety(wpsCopyContentVerdict(staged));",
      "  if (blocked) return blocked;",
    ].join("\n"),
    mutant: "  const blocked = null; // [变异S11] 不再核对副本是否被额外修改",
    cwd: "apps/wps-addin",
    test: "src/wpsNativeComparison.test.ts",
    expect: "副本里被人工改过 ⇒ 不关副本、报 extra_edits，改动不被静默丢弃",
  },
  // S12：Excel 原生关闭前不再核对副本证据 ⇒ 人工修改被静默丢弃。
  S12: {
    file: "apps/excel-addin/src/excelNativeComparison.ts",
    seed: [
      "  const blocked = describeCopyDiscardSafety(content);",
      "  if (blocked) return blocked;",
    ].join("\n"),
    mutant: "  const blocked = null; // [变异S12] 不再核对副本是否被额外修改",
    cwd: "apps/excel-addin",
    test: "src/excelNativeComparison.test.ts",
    expect: "副本被人工修改后结束审阅 ⇒ 不关副本、报 extra_edits，修改不被静默丢弃",
  },
  // S13：把"窗口还在"直接解释成"宿主正在询问是否保存" ⇒ 没有对话框证据也敢下结论。
  S13: {
    file: "packages/core/src/reviewCopyLifecycle.ts",
    seed: "    detail: '关闭请求已经发出，但窗口仍然存在。加载项没有对话框证据，所以不推断原因：可能是宿主在询问是否保存，也可能是别的东西挡住了关闭、或关闭本身较慢。请切到该窗口确认后再决定；副本内容不会写回原表。',",
    mutant: "    detail: '加载项没有对话框证据。宿主正在询问是否保存，请选「不保存」。', // [变异S13] 推断保存提示",
    cwd: "packages/core",
    test: "src/reviewCopyLifecycle.test.ts",
    expect: "close_unconfirmed 只说\"未确认\"，不声称宿主在问是否保存",
  },
  // S14：跳过恢复文件直接关闭，必须由真实适配器控制流测试抓住。
  S14: {
    file: "server/app/excel_window_layout.py",
    seed: '                recovery_path = self._archive_review_book(book)',
    mutant: '                recovery_path = None  # [变异S14] 跳过归档直接关闭',
    cwd: ".",
    test: "server/tests/test_review_copy_archive_close.py",
    expect: "test_retained_workbook_is_archived_with_extra_edits_then_closed_without_original_access",
  },
  // S15：确认不再绑在**这一份**副本上 ⇒ 放弃旧副本后，旧请求能借新副本的现场写原表。
  S15: {
    file: "apps/excel-addin/src/hooks/useExcelNativeComparison.ts",
    seed: '      if (current.session.mode === "manual" && acknowledgedCopy !== current.session.copySheetName) {',
    mutant: "      if (false) { // [变异S15] 确认不再绑在这一份副本上",
    cwd: "apps/excel-addin",
    test: "src/components/ExcelNativeReviewActions.test.tsx",
    expect: "确认勾选项绑在**这一份**副本上：带着上一份副本的确认也不能写回新副本",
  },
  // S16：不再比对现场身份 ⇒ "当前有审阅"就能放行，旧请求借新副本的确认写原表。
  S16: {
    file: "packages/core/src/reviewWriteBackGate.ts",
    seed: [
      "  return issued.review !== current.review ||",
      "    issued.plan !== current.plan ||",
      "    issued.workbookIdentity !== current.workbookIdentity;",
    ].join("\n"),
    mutant: "  return false; // [变异S16] 不再比对现场身份",
    cwd: "packages/core",
    test: "src/reviewWriteBackGate.test.ts",
    expect: "只有副本换了（放弃后又生成新副本）⇒ 已变化",
  },
  // S17：授权检查回来后不再复核现场 ⇒ "当前有审阅"就能放行，旧请求借新副本的确认写原表。
  //      判别点是那条**直接判据**用例：复合路径（在途期间放弃并生成新副本）目前被
  //      `busy` 挡在界面之外（`busy` 含 `targetGuard.checking`），端到端造不出来。
  S17: {
    file: "apps/wps-addin/src/WpsAgentPanel.tsx",
    seed: "  return !reviewWriteBackTargetChanged(issued, now.target);",
    mutant: "  return true; // [变异S17] 授权检查回来后不再复核现场",
    cwd: "apps/wps-addin",
    test: "src/WpsAgentPanel.test.tsx",
    expect: "判据逐项分开：换了副本 / 换了计划 / 换了原工作簿，各自都能单独拦下",
  },
  // S18：写回请求不再核对属于哪张卡 ⇒ 授权检查在途期间换了副本后，旧请求借新现场写原表。
  S18: {
    file: "apps/excel-addin/src/hooks/useExcelNativeComparison.ts",
    seed: '      if (!current || current.messageId !== messageId || current.outcome !== "reviewing" || mode !== "copy-compare") {',
    mutant: '      if (!current || current.outcome !== "reviewing") { // [变异S18] 不再核对请求属于哪张卡，也不再核对执行方式',
    cwd: "apps/excel-addin",
    test: "src/components/ExcelNativeReviewActions.test.tsx",
    expect: "授权检查在途期间现场换掉了（放弃并生成新副本）⇒ 旧请求回来后不写回",
  },
};

const sha256 = (buffer) => createHash("sha256").update(buffer).digest("hex");
const eolOf = (text) => (text.includes("\r\n") ? "\r\n" : "\n");
const toLf = (text) => text.split("\r\n").join("\n");
const fromLf = (text, eol) => (eol === "\r\n" ? text.split("\n").join("\r\n") : text);
const statePath = (id, suffix) => path.join(STATE, `${id}.${suffix}`);
const markerOf = (id) => `[变异${id}]`;
const applied = new Set();

function requireSeed(id) {
  const target = id && SEEDS[id];
  if (!target) {
    console.error(`未知 SEED：${id}（可选：${Object.keys(SEEDS).join(", ")}）`);
    process.exit(2);
  }
  return target;
}

function apply(id) {
  const target = requireSeed(id);
  const absolute = path.join(REPO, target.file);
  mkdirSync(STATE, { recursive: true });
  const raw = readFileSync(absolute);
  const text = raw.toString("utf8");
  const eol = eolOf(text);
  const lf = toLf(text);
  if (lf.includes(markerOf(id))) {
    console.error(`前置状态闸：${target.file} 已含 ${markerOf(id)}，请先 restore ${id}。`);
    process.exit(3);
  }
  const hits = lf.split(target.seed).length - 1;
  if (hits !== 1) {
    console.error(`SEED ${id} 在 ${target.file} 命中 ${hits} 次（期望 1）；源码重写会打散旧 SEED。`);
    process.exit(3);
  }
  writeFileSync(statePath(id, "orig"), raw);
  writeFileSync(statePath(id, "sha256"), sha256(raw) + "\n", "utf8");
  writeFileSync(absolute, Buffer.from(fromLf(lf.replace(target.seed, target.mutant), eol), "utf8"));
  applied.add(id);
  console.log(`APPLIED ${id} → ${target.file}`);
}

function restore(id, { quiet = false } = {}) {
  const target = requireSeed(id);
  const absolute = path.join(REPO, target.file);
  const backup = statePath(id, "orig");
  if (!existsSync(backup)) {
    if (!quiet) console.log(`ALREADY-CLEAN：${target.file}（未施加过 ${id}）`);
    return;
  }
  const original = readFileSync(backup);
  const expected = readFileSync(statePath(id, "sha256"), "utf8").trim();
  writeFileSync(absolute, original);
  const actual = sha256(readFileSync(absolute));
  if (actual !== expected) {
    console.error(`还原后 sha256 不一致：${target.file}（期望 ${expected}，实际 ${actual}）`);
    process.exit(4);
  }
  rmSync(backup);
  rmSync(statePath(id, "sha256"));
  applied.delete(id);
  if (!quiet) console.log(`RESTORED ${id} → ${target.file}（sha256 ${actual.slice(0, 12)} 一致）`);
}

function restoreAll() {
  for (const id of [...applied]) {
    try { restore(id, { quiet: true }); console.log(`RESTORED ${id}（异常钩子）`); }
    catch (error) { console.error(`异常钩子还原 ${id} 失败：${error.message} —— 文件：${SEEDS[id].file}`); }
  }
}

for (const signal of ["SIGINT", "SIGTERM", "SIGHUP"]) {
  process.on(signal, () => { restoreAll(); process.exit(130); });
}
process.on("uncaughtException", (error) => {
  console.error(`未捕获异常：${error?.stack ?? error}`);
  restoreAll();
  process.exit(5);
});
process.on("unhandledRejection", (reason) => {
  console.error(`未处理的 rejection：${reason?.stack ?? reason}`);
  restoreAll();
  process.exit(5);
});

/**
 * SEED 在位自检（只读，不改任何文件）。
 *
 * 源码一旦重写（抽函数、改缩进、剥 Markdown 记号），SEED 就会静默落空 ——
 * `apply` 的前置闸会如实报 `命中 0 次`，但那时你已经在跑变异了。这个 check 让
 * "重构完先看一眼种子还在不在"变成一条命令，与 `mutate_excel_format_compat.mjs check` 对齐。
 */
function check() {
  let broken = 0;
  for (const [id, target] of Object.entries(SEEDS)) {
    const absolute = path.join(REPO, target.file);
    if (!existsSync(absolute)) {
      console.error(`MISSING ${id}  ${target.file}（文件不存在）`);
      broken += 1;
      continue;
    }
    const lf = toLf(readFileSync(absolute, "utf8"));
    const hits = lf.split(target.seed).length - 1;
    if (hits !== 1) {
      console.error(`BROKEN  ${id}  ${target.file}（命中 ${hits} 次，期望 1）`);
      console.error(`   seed=${JSON.stringify(target.seed).slice(0, 160)}`);
      broken += 1;
    }
  }
  if (broken > 0) {
    console.error(`CHECK FAIL：${broken} / ${Object.keys(SEEDS).length} 个 SEED 已落空；源码重写会打散旧 SEED，请修回。`);
    process.exit(6);
  }
  console.log(`CHECK OK 全部 ${Object.keys(SEEDS).length} 个 SEED 都在位（各命中 1 次）`);
}

const [mode, id] = process.argv.slice(2);
if (mode === "apply") apply(id);
else if (mode === "restore") restore(id);
else if (mode === "check") check();
else if (mode === "plan") {
  for (const [seedId, target] of Object.entries(SEEDS)) {
    console.log([seedId, target.cwd, target.test, target.expect].join("\t"));
  }
} else {
  console.error("用法: node scripts/mutate_review_copy_lifecycle.mjs plan | apply <ID> | restore <ID> | check");
  process.exit(2);
}
