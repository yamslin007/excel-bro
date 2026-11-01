// 反向变异脚本：处置留痕（放弃 / 撤销）**接线层**的判别力（2026-10-29，第十三轮）
//
// 上一份 `mutate_execution_disposition.py` 只打服务端账本（D1 首次写赢 / D2 不碰进度 /
// D3 两值不同列）。这里补三个**端上**判别点 —— 它们分别钉着"什么时候写"和"写成什么"：
//
//   E1  Excel：把 `onRestored` 里的判定写死成 `abandoned`
//       ⇒「确认写入 → 结果卡撤销」那条必须红：确认之后撤销是 `undone`，不许被记成放弃。
//
//   E2  Excel：把 `onRestored` 挪到 `await undoExecution(snapshot)` **之前**
//       ⇒「执行后目标被人工修改，确认与放弃都不覆盖该修改」必须红：恢复失败时宿主一个
//       单元格都没回退，这时写留痕就是造假。留痕的语义是"用户确实撤回了"。
//
//   W3  WPS：把 `origin === "discard" ? "abandoned" : "undone"` 写死成 `abandoned`
//       ⇒「撤销本次执行 → 服务端留痕记「撤销」」必须红。双端读同一份判据，两端各有一个
//       判别点，不许"一端绿就算两端都对"。
//
// 四判据（不许只卡"退出码非 0"）：
//   ① 退出码非 0；② 预期用例名出现在 FAIL 里；③ 失败原文带它**自己那处**证据；
//   ④ 还原后逐字节 sha256 与基线一致（且 `check` 会再跑一次还原后的那条用例，必须绿）。
//
// 用法（**一次一个变异**，不要并发）：
//   node scripts/mutate_excel_disposition_wiring.mjs                # 打印基线 sha256
//   node scripts/mutate_excel_disposition_wiring.mjs E1             # 应用（手工调试用）
//   node scripts/mutate_excel_disposition_wiring.mjs E1 restore     # 还原
//   node scripts/mutate_excel_disposition_wiring.mjs E1 check       # 跑回归并按四判据判定
import { readFileSync, writeFileSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const EXCEL_APP = resolve(REPO, "apps", "excel-addin", "src", "App.tsx");
const EXCEL_UNDO = resolve(REPO, "apps", "excel-addin", "src", "hooks", "useUndoSnapshot.ts");
const WPS_PANEL = resolve(REPO, "apps", "wps-addin", "src", "WpsAgentPanel.tsx");

const TARGETS = {
  E1: {
    file: EXCEL_APP,
    app: "excel-addin",
    testFile: "src/excelWriteEndToEnd.test.tsx",
    test: "确认写入 → 界面给出显式撤销范围 → 撤销后逐单元格回到执行前状态",
    seed: [
      "      const disposition: ExecutionDisposition =",
      '        revertedByDiscard.current === snapshot.executionId ? "abandoned" : "undone";'
    ].join(LF),
    mutant: [
      '      const disposition: ExecutionDisposition = "abandoned";  // E1（变异体，勿提交）'
    ].join(LF),
    // vitest 的断言原文（**单引号**）就是这一句：取值被写死成 abandoned，而用例要求 undone。
    evidence: "'abandoned' to be 'undone'"
  },
  E2: {
    file: EXCEL_UNDO,
    app: "excel-addin",
    testFile: "src/excelWriteEndToEnd.test.tsx",
    test: "执行后目标被人工修改，确认与放弃都不覆盖该修改",
    seed: [
      "        for (const snapshot of targets) {",
      "          await undoExecution(snapshot);",
      "          undone += 1;",
      "          onRestored?.(snapshot);"
    ].join(LF),
    mutant: [
      "        for (const snapshot of targets) {",
      "          onRestored?.(snapshot);  // E2（变异体，勿提交）：恢复失败也回调",
      "          await undoExecution(snapshot);",
      "          undone += 1;"
    ].join(LF),
    // 负向断言 `.not.toHaveBeenCalled()` 的失败原文。
    evidence: "to not be called"
  },
  W3: {
    file: WPS_PANEL,
    app: "wps-addin",
    testFile: "src/WpsAgentPanel.test.tsx",
    test: "撤销本次执行 → 服务端留痕记「撤销」，不是「放弃」",
    seed: '      const disposition: ExecutionDisposition = origin === "discard" ? "abandoned" : "undone";',
    mutant: '      const disposition: ExecutionDisposition = "abandoned";  // W3（变异体，勿提交）',
    // 与 Excel 同一条断言的同一种失败原文（两端读同一份判据，就该这么一致）。
    evidence: "'abandoned' to be 'undone'"
  }
};

function normalize(text) {
  return text.split(/\r\n/).join(LF);
}
function denormalize(text, original) {
  const native = original.includes("\r\n") ? "\r\n" : LF;
  return text.split(LF).join(native);
}
function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

/*
 * 写回必须读回校验并重试（`Errno 22` / 索引器占句柄会把变异留在工作区）。
 * 变异体留在工作区比变异失败更糟：后续回归会以"红"的面目出现，而人以为那是产品缺陷。
 */
function writeVerified(path, text) {
  const expected = normalize(text);
  let last = "";
  for (let attempt = 1; attempt <= 5; attempt += 1) {
    writeFileSync(path, text, "utf8");
    last = normalize(readFileSync(path, "utf8"));
    if (last === expected) return true;
  }
  console.error(`写回校验失败（5 次）：${path}`);
  console.error(`期望片段：${expected.slice(0, 120)}`);
  console.error(`实际片段：${last.slice(0, 120)}`);
  return false;
}

const which = process.argv[2];
if (!which) {
  for (const [id, target] of Object.entries(TARGETS)) {
    console.log(`BASELINE ${id} ${target.file} sha256=${sha256(target.file)}`);
  }
  process.exit(0);
}
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_excel_disposition_wiring.mjs [E1|E2|W3] [restore|check]");
  process.exit(2);
}

const target = TARGETS[which];
const original = readFileSync(target.file, "utf8");

/** 失控保护：信号/未捕获异常一律先把变异还原掉。 */
function restoreQuietly() {
  try {
    const normalized = normalize(readFileSync(target.file, "utf8"));
    if (!normalized.includes(normalize(target.mutant))) return;
    writeVerified(target.file, denormalize(normalized.split(normalize(target.mutant)).join(normalize(target.seed)), original));
    console.error(`[hook] 已还原 ${which} → ${target.file}`);
  } catch (error) {
    console.error(`[hook] 还原失败：${error instanceof Error ? error.message : String(error)}`);
  }
}
for (const signal of ["SIGINT", "SIGTERM", "SIGHUP"]) {
  process.on(signal, () => {
    restoreQuietly();
    process.exit(130);
  });
}
process.on("uncaughtException", (error) => {
  restoreQuietly();
  console.error(error);
  process.exit(1);
});

if (process.argv[3] === "restore") {
  const hits = normalize(original).split(normalize(target.mutant)).length - 1;
  if (hits === 0) {
    console.log("ALREADY-CLEAN");
    process.exit(0);
  }
  const restored = denormalize(normalize(original).split(normalize(target.mutant)).join(normalize(target.seed)), original);
  if (!writeVerified(target.file, restored)) process.exit(4);
  console.log(`restore: OK ${target.file} sha256=${sha256(target.file)}`);
  process.exit(0);
}

const occurrences = normalize(original).split(normalize(target.seed)).length - 1;
if (occurrences !== 1) {
  // SEED 未命中 = 假装有证据。结构改写后必须回来重新定点。
  console.log(`SKIP: SEED 出现 ${occurrences} 次（期望 1）于 ${target.file}`);
  process.exit(3);
}

const CODE = 0;

function runVitest(pattern) {
  // `-t` 是正则：用例名里的 `：`、`→`、`「」` 必须先转义，否则会顺手匹配到别的用例。
  const escaped = pattern.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return spawnSync(
    process.execPath,
    [resolve(REPO, "node_modules", "vitest", "vitest.mjs"), "run", target.testFile, "-t", escaped, "--reporter=verbose"],
    {
      cwd: resolve(REPO, "apps", target.app),
      encoding: "utf8",
      timeout: 600_000,
      env: { ...process.env, NO_COLOR: "1", FORCE_COLOR: "0" }
    }
  );
}
function clean(run) {
  return `${run.stdout ?? ""}\n${run.stderr ?? ""}`.replace(/\u001b\[[0-9;]*m/g, "");
}

if (process.argv[3] === "check") {
  const baseline = sha256(target.file);
  const mutated = denormalize(normalize(original).replace(normalize(target.seed), normalize(target.mutant)), original);
  if (!writeVerified(target.file, mutated)) process.exit(4);
  let run;
  try {
    run = runVitest(target.test);
  } finally {
    if (!writeVerified(target.file, denormalize(normalize(original), original))) process.exit(4);
  }
  const output = clean(run);
  if (run.error) console.log(`spawn 失败：${run.error.message}`);
  const failedLines = output.split("\n").filter((line) => line.includes("FAIL"));
  const criteria = [
    { id: "①退出码非 0", ok: run.status !== 0, detail: `status=${run.status}` },
    {
      id: "②预期用例红了",
      ok: failedLines.some((line) => line.includes(target.test)),
      detail: failedLines[0]?.trim() ?? "（输出里没有任何 FAIL 行）"
    },
    { id: "③失败原文带自家证据", ok: output.includes(target.evidence), detail: target.evidence },
    { id: "④还原后逐字节一致", ok: sha256(target.file) === baseline, detail: `${baseline} → ${sha256(target.file)}` }
  ];
  for (const item of criteria) console.log(`${item.ok ? "OK  " : "FAIL"} ${item.id} :: ${item.detail}`);
  // 还原后复跑：证明"红的来源是变异本身"，而不是这段用例本来就红。
  const after = runVitest(target.test);
  const afterOutput = clean(after);
  const greenAfter = after.status === 0 && /passed/.test(afterOutput) && !/FAIL/.test(afterOutput);
  console.log(`${greenAfter ? "OK  " : "FAIL"} ④b还原后复跑绿 :: status=${after.status}`);
  if (!greenAfter) console.log(afterOutput.slice(-1500));
  const allOk = criteria.every((item) => item.ok) && greenAfter;
  console.log(allOk ? `\n${which}: 四判据全部成立` : `\n${which}: 判据不成立，别当成"已经验证过"`);
  process.exit(allOk ? CODE : 1);
}

const mutated = denormalize(normalize(original).replace(normalize(target.seed), normalize(target.mutant)), original);
if (!writeVerified(target.file, mutated)) process.exit(4);
console.log(`APPLIED ${which} → ${target.file} sha256=${sha256(target.file)}`);
