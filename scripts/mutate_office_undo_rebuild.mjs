// 反向变异脚本：**重开之后撤销未完成执行**这条链路的判别力（2026-10-31，第二十五轮）
//
// 缺陷现场：三步公式计划在第二步回读后中断 → 核对在途动作 → 改 `来源!B5` 5→500 →
// 续跑被拒（对的）→ 重开 Excel，恢复卡还在、两个公式还在表里，**没有**「撤销本次执行」。
// 本轮补的四个判别点各自钉住一件不同的事，四者缺一都不能说"修好了"：
//
//   U1  账本材料挂不进撤销栈（`ensureUndoSnapshot` 变成空操作）
//       ⇒「重开后…撤销只回退已写入的目标」必须红：材料在账本里，但栈里没有，
//       按钮点下去静默不执行 —— 这正是"有快照 ≠ 能撤销"的分界。
//
//   U2  把"未完成/待核对"当成不可撤销（判据里把 `needs_reconciliation` 并进 `completed`）
//       ⇒ 同一条用例必须红，且红在**说明**上：判据阻断时正文根本不渲染那句
//       "来源工作表「来源」保持你现在的修改"。
//
//   U3  服务端状态响应不再投影处置留痕（`disposition=None`）
//       ⇒ 服务端用例必须红：重开的窗格无法知道"这个入口已经被用掉"。
//
//   U4  同一执行的两份快照都留在栈里（断点前缀份 + 续跑完成份，或按 planId 兜底去重失效）
//       ⇒「公式续跑完成后进入审阅收尾（discard）」必须红：LIFO 恢复两遍，
//       第二次必然与写后指纹冲突 ⇒ 用户点「放弃变更」得到"恢复未成功"。
//
// 四判据（不许只卡"退出码非 0"）：
//   ① 退出码非 0；② 预期用例名出现在失败清单里；③ 失败原文带它**自己那处**证据；
//   ④ 还原后逐字节 sha256 与基线一致；④b 还原后复跑该用例必须绿（红的来源是变异本身）。
//
// 用法（**一次一个变异**，不要并发；跑前确认没有别的变异留在工作区）：
//   node scripts/mutate_office_undo_rebuild.mjs                # 打印基线 sha256 + 当前是否干净
//   node scripts/mutate_office_undo_rebuild.mjs U1             # 应用（手工调试用）
//   node scripts/mutate_office_undo_rebuild.mjs U1 restore     # 还原
//   node scripts/mutate_office_undo_rebuild.mjs U1 check       # 跑回归并按四判据判定
import { readFileSync, writeFileSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";
import { tmpdir } from "node:os";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const EXCEL_UNDO_HOOK = resolve(REPO, "apps", "excel-addin", "src", "hooks", "useUndoSnapshot.ts");
const CORE_UNDO = resolve(REPO, "packages", "core", "src", "executionUndoRebuild.ts");
const SERVER_MAIN = resolve(REPO, "server", "app", "main.py");

const E2E = "src/excelWriteEndToEnd.test.tsx";
const REOPEN_TEST =
  "中断 → 核对 → 改来源（续跑被拒）→ 重开且本机索引不在 → 撤销只回退已写入的目标，来源保持用户的修改";

const TARGETS = {
  U1: {
    file: EXCEL_UNDO_HOOK,
    runner: "vitest",
    app: "excel-addin",
    testFile: E2E,
    test: REOPEN_TEST,
    seed: "        return exists ? stack : trimStack([...stack, snapshot]);",
    mutant:
      "        return stack;  // U1（变异体，勿提交）：账本材料挂不进撤销栈",
    // 撤销根本没跑 ⇒ 那句结论永远不会出现。
    evidence: "to contain '已撤销本次执行'"
  },
  U2: {
    file: CORE_UNDO,
    runner: "vitest",
    app: "excel-addin",
    testFile: E2E,
    test: REOPEN_TEST,
    seed: '  if (input.status === "completed") {',
    mutant:
      '  if (input.status === "completed" || input.status === "needs_reconciliation") {  // U2（变异体，勿提交）',
    // 判据阻断时正文不渲染撤销说明（那句"来源保持你现在的修改"只在可撤销时出现）。
    evidence: "to contain '来源工作表「来源」保持你现在的修改'"
  },
  U3: {
    file: SERVER_MAIN,
    runner: "pytest",
    testFile: "server/tests/test_execution_disposition.py",
    test: "test_state_response_projects_disposition_so_a_reopened_pane_does_not_offer_it_again",
    seed: [
      "        disposition=OfficeExecutionDispositionState(",
      "            undoneAt=record.undone_at, abandonedAt=record.abandoned_at",
      "        ),"
    ].join(LF),
    mutant: "        disposition=None,  # U3（变异体，勿提交）：状态响应不回处置留痕",
    evidence: "AssertionError: assert None =="
  },
  U4: {
    file: EXCEL_UNDO_HOOK,
    runner: "vitest",
    app: "excel-addin",
    testFile: E2E,
    test: "公式续跑完成后进入审阅收尾（discard）",
    seed: "          ...stack.filter((entry) => !snapshot.executionId || entry.executionId !== snapshot.executionId),",
    mutant: "          ...stack,  // U4（变异体，勿提交）：同一执行的两份快照都留在栈里",
    evidence: "to contain '已放弃并恢复'"
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
  console.error(`期望片段：${expected.slice(0, 160)}`);
  console.error(`实际片段：${last.slice(0, 160)}`);
  return false;
}

const which = process.argv[2];
if (!which) {
  for (const [id, target] of Object.entries(TARGETS)) {
    const text = normalize(readFileSync(target.file, "utf8"));
    const dirty = text.includes(normalize(target.mutant));
    console.log(
      `BASELINE ${id} ${target.file} sha256=${sha256(target.file)} ${dirty ? "!! 变异体仍在工作区" : "clean"}`
    );
  }
  process.exit(0);
}
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_office_undo_rebuild.mjs [U1|U2|U3|U4] [restore|check]");
  process.exit(2);
}

const target = TARGETS[which];
const original = readFileSync(target.file, "utf8");

/** 失控保护：信号/未捕获异常一律先把变异还原掉。 */
function restoreQuietly() {
  try {
    const normalized = normalize(readFileSync(target.file, "utf8"));
    if (!normalized.includes(normalize(target.mutant))) return;
    writeVerified(
      target.file,
      denormalize(normalized.split(normalize(target.mutant)).join(normalize(target.seed)), original)
    );
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
  const restored = denormalize(
    normalize(original).split(normalize(target.mutant)).join(normalize(target.seed)),
    original
  );
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

function pythonExe() {
  return process.platform === "win32"
    ? resolve(REPO, ".venv", "Scripts", "python.exe")
    : resolve(REPO, ".venv", "bin", "python");
}

function runTest() {
  if (target.runner === "pytest") {
    return spawnSync(
      pythonExe(),
      [
        "-m",
        "pytest",
        target.testFile,
        "-q",
        "-k",
        target.test,
        // basetemp 必须落在 **OS 临时目录**：仓库内的临时目录会被环境里的
        // `sitecustomize` 安全删除钩子牵连（`Errno 22` / 计数跨进程）。
        `--basetemp=${resolve(tmpdir(), "eb-mutation-pytest")}`
      ],
      { cwd: REPO, encoding: "utf8", timeout: 600_000, env: { ...process.env, NO_COLOR: "1" } }
    );
  }
  // `-t` 是正则：用例名里的 `：`、`→`、`「」` 必须先转义，否则会顺手匹配到别的用例。
  const escaped = target.test.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return spawnSync(
    process.execPath,
    [
      resolve(REPO, "node_modules", "vitest", "vitest.mjs"),
      "run",
      target.testFile,
      "-t",
      escaped,
      "--reporter=verbose"
    ],
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
/** 失败清单：vitest 打 `FAIL`，pytest 打 `FAILED`。 */
function failedLines(output) {
  return output.split("\n").filter((line) => /\bFAIL(ED)?\b/.test(line));
}
function green(run, output) {
  if (run.status !== 0) return false;
  return target.runner === "pytest" ? !/\bfailed\b/.test(output) : /passed/.test(output) && !/FAIL/.test(output);
}

if (process.argv[3] === "check") {
  const baseline = sha256(target.file);
  const mutated = denormalize(
    normalize(original).replace(normalize(target.seed), normalize(target.mutant)),
    original
  );
  if (!writeVerified(target.file, mutated)) process.exit(4);
  let run;
  try {
    run = runTest();
  } finally {
    if (!writeVerified(target.file, denormalize(normalize(original), original))) process.exit(4);
  }
  const output = clean(run);
  if (run.error) console.log(`spawn 失败：${run.error.message}`);
  const fails = failedLines(output);
  const criteria = [
    { id: "①退出码非 0", ok: run.status !== 0, detail: `status=${run.status}` },
    {
      id: "②预期用例红了",
      ok: fails.some((line) => line.includes(target.test.split("（")[0])) || output.includes(target.test),
      detail: fails[0]?.trim() ?? "（输出里没有任何失败行）"
    },
    { id: "③失败原文带自家证据", ok: output.includes(target.evidence), detail: target.evidence },
    { id: "④还原后逐字节一致", ok: sha256(target.file) === baseline, detail: `${baseline} → ${sha256(target.file)}` }
  ];
  for (const item of criteria) console.log(`${item.ok ? "OK  " : "FAIL"} ${item.id} :: ${item.detail}`);
  const after = runTest();
  const afterOutput = clean(after);
  const greenAfter = green(after, afterOutput);
  console.log(`${greenAfter ? "OK  " : "FAIL"} ④b还原后复跑绿 :: status=${after.status}`);
  if (!greenAfter) console.log(afterOutput.slice(-1500));
  const allOk = criteria.every((item) => item.ok) && greenAfter;
  console.log(allOk ? `\n${which}: 四判据全部成立` : `\n${which}: 判据不成立，别当成"已经验证过"`);
  process.exit(allOk ? 0 : 1);
}

const mutated = denormalize(
  normalize(original).replace(normalize(target.seed), normalize(target.mutant)),
  original
);
if (!writeVerified(target.file, mutated)) process.exit(4);
console.log(`APPLIED ${which} → ${target.file} sha256=${sha256(target.file)}`);
