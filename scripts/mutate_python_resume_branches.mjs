// 反向变异脚本：隔离 Python 两个恢复分支补修（2026-10-06 第五批）
// 用法：node scripts/mutate_python_resume_branches.mjs M1|M2|M3
// M1 = Python 状态判定忽略 pending（拆 pending 纳入关联/判定）
// M2 = 名称非空就放行孤立回执（拆孤立 ID 拒绝）
// M3 = 跳过调用 ID 与工具名一致性检查（拆名称冲突拒绝）
// 铁律：基线先 GREEN；信号钩子还原；写前登记 touched；SEED 唯一性校验。
import { readFileSync, writeFileSync, mkdtempSync } from "node:fs";
import { spawnSync as spawn } from "node:child_process";
import { tmpdir } from "node:os";
import { join } from "node:path";

const REPO = "E:\\codex\\excel-bro";
const PY = join(REPO, ".venv", "Scripts", "python.exe");
const FILE = join(REPO, "server", "app", "excel_agent.py");
const NL = "\r\n";

const TARGETS = {
  M1: {
    expectFail: ["test_pending_only_python_loop_state_never_persisted"],
  },
  M2: {
    expectFail: [
      "test_named_orphan_receipt_refuses_resume",
      "test_named_orphan_receipt_never_persisted_or_sent",
    ],
  },
  M3: {
    expectFail: ["test_named_receipt_name_mismatch_refuses_resume"],
  },
};

const SEEDS = {
  // M1：让 _checkpoint_involves_python 忽略 pending 参数（判定回到只看 messages）
  M1: [
    '    candidates: list[dict[str, Any]] = list(messages)',
    '    if isinstance(pending_assistant_message, dict):',
    '        candidates.append(pending_assistant_message)',
    '    for item in candidates:',
    '        if not isinstance(item, dict):',
    '            continue',
    '        if item.get("role") == "assistant":',
  ].join(NL),
  // M2：孤立 ID 检查短路（有 name 时跳过拒绝）
  M2: [
    '            tool_call_id = copied.get("tool_call_id")',
    '            name = copied.get("name")',
    '            if not isinstance(tool_call_id, str) or tool_call_id not in association:',
    '                # 孤立 ID：无法关联到任何 assistant 调用，内容归属不明，拒绝。',
    '                raise ValueError(PYTHON_ANALYSIS_UNATTRIBUTABLE_RESUME_MESSAGE)',
  ].join(NL),
  // M3：名称冲突检查短路
  M3: [
    '            expected_name = association[tool_call_id]',
    '            if isinstance(name, str) and name != expected_name:',
    '                # 名称冲突：回执声称的工具名与调用名不一致，拒绝。',
    '                raise ValueError(PYTHON_ANALYSIS_UNATTRIBUTABLE_RESUME_MESSAGE)',
  ].join(NL),
};

const MUTANTS = {
  M1: [
    '    # [变异M1] Python 状态判定忽略 pending',
    '    candidates: list[dict[str, Any]] = list(messages)',
    '    for item in candidates:',
    '        if not isinstance(item, dict):',
    '            continue',
    '        if item.get("role") == "assistant":',
  ].join(NL),
  M2: [
    '            tool_call_id = copied.get("tool_call_id")',
    '            name = copied.get("name")',
    '            if not name and (not isinstance(tool_call_id, str) or tool_call_id not in association):',
    '                # [变异M2] 名称非空就放行孤立回执',
    '                raise ValueError(PYTHON_ANALYSIS_UNATTRIBUTABLE_RESUME_MESSAGE)',
  ].join(NL),
  M3: [
    '            expected_name = association[tool_call_id]',
    '            # [变异M3] 跳过调用 ID 与工具名一致性检查',
  ].join(NL),
};

const which = process.argv[2];
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_python_resume_branches.mjs M1|M2|M3");
  process.exit(2);
}
const target = TARGETS[which];
const seed = SEEDS[which];
const mutant = MUTANTS[which];
const original = readFileSync(FILE, "utf8");
const occurrences = original.split(seed).length - 1;
if (occurrences !== 1) {
  console.log(`SKIP: SEED 在目标文件中出现 ${occurrences} 次（期望 1）。`);
  process.exit(3);
}

const touched = new Map();
function restoreAll() {
  for (const [f, c] of touched) writeFileSync(f, c, "utf8");
  touched.clear();
}
for (const sig of ["SIGINT", "SIGTERM", "SIGHUP"]) {
  process.on(sig, () => { restoreAll(); process.exit(130); });
}
process.on("uncaughtException", (err) => { restoreAll(); console.error("uncaughtException（已还原）:", err); process.exit(1); });

function runPytest(name) {
  const baseDir = mkdtempSync(join(tmpdir(), "mutresume_"));
  const r = spawn(PY, ["-m", "pytest", "server/tests/test_python_analysis.py", "-q",
    "-p", "no:cacheprovider", "--basetemp", join(baseDir, name)], {
    cwd: REPO, encoding: "utf8",
    env: { ...process.env, TMP: baseDir, TEMP: baseDir, PYTEST_ADDOPTS: "" },
    timeout: 300000,
  });
  return `${r.stdout || ""}\n${r.stderr || ""}`;
}

const baseline = runPytest("baseline");
const bp = /(\d+) passed/.exec(baseline);
const bf = /(\d+) failed/.exec(baseline);
console.log(`[基线] ${bp ? bp[1] + " passed" : "?"}${bf ? ", " + bf[1] + " failed" : ""}`);
if (!bp || bf) {
  console.log("BASELINE-NOT-GREEN：基线未全绿，终止（源码未改动）。");
  console.log(baseline.slice(-1500));
  process.exit(4);
}

try {
  touched.set(FILE, original);
  writeFileSync(FILE, original.replace(seed, mutant), "utf8");
  const mutated = runPytest("mutated");
  const failed = [...mutated.matchAll(/FAILED (server\/tests\/\w+\.py::\w+)/g)].map((m) => m[1].split("::")[1]);
  const hits = target.expectFail.filter((n) => failed.includes(n));
  console.log(`[变异后] 失败 ${failed.length} 条：${failed.join(", ") || "（无）"}`);
  console.log(`[命中判别点] ${hits.length ? hits.join(", ") : "无"}`);
  const red = failed.length > 0 && hits.length > 0;
  console.log(red ? "RED-OK：变异使判别测试独立变红。" : "RED-MISS：变异未按预期变红。");
  if (!red) console.log(mutated.slice(-2000));

  restoreAll();
  const restored = runPytest("restored");
  const rp = /(\d+) passed/.exec(restored);
  const rf = /(\d+) failed/.exec(restored);
  console.log(`[还原后] ${rp ? rp[1] + " passed" : "?"}${rf ? ", " + rf[1] + " failed" : ""}`);
  const green = rp && !rf;
  console.log(green ? "RESTORED-GREEN：还原后复跑全绿。" : "RESTORED-NOT-GREEN：还原后仍有失败！");
  process.exit(red && green ? 0 : 5);
} catch (err) {
  restoreAll();
  throw err;
}
