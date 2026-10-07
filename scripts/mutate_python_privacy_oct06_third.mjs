// 反向变异脚本：隔离 Python 三缺陷修复（2026-10-06 第三批）
// 用法：node scripts/mutate_python_privacy_oct06_third.mjs M1|M2|M3
// M1 = 恢复"被拒绝参数原样历史保留"（拆最小白名单投影）
// M2 = 绕过"旧检查点安全恢复检查"（拆恢复现场投影）
// M3 = 恢复"执行入口异常原样回显"（拆执行错误安全映射）
// 铁律：基线先 GREEN；信号钩子保证还原；写前登记 touched；SEED 唯一性校验。
import { readFileSync, writeFileSync, mkdtempSync } from "node:fs";
import { spawnSync as spawn } from "node:child_process";
import { tmpdir } from "node:os";
import { join } from "node:path";

const REPO = "E:\\codex\\excel-bro";
const PY = join(REPO, ".venv", "Scripts", "python.exe");

const TARGETS = {
  M1: {
    file: join(REPO, "server", "app", "python_analysis.py"),
    nl: "\n",
    expectFail: [
      "test_rejected_extra_fields_never_enter_model_requests",
      "test_rejected_nested_handles_never_enter_model_requests",
      "test_rejected_arguments_never_enter_sqlite",
    ],
  },
  M2: {
    file: join(REPO, "server", "app", "excel_agent.py"),
    nl: "\r\n",
    expectFail: [
      "test_old_checkpoint_python_code_not_sent_or_executed",
      "test_old_checkpoint_pending_message_code_not_preserved",
    ],
  },  M3: {
    file: join(REPO, "server", "app", "excel_agent.py"),
    nl: "\r\n",
    expectFail: [
      "test_execution_entry_exception_fails_closed_in_loop",
    ],
  },
  M4: {
    file: join(REPO, "server", "app", "claude_sdk_agent.py"),
    nl: "\n",
    expectFail: [
      "test_execution_entry_exception_fails_closed_in_sdk",
    ],
  },
};

const SEEDS = {
  M1: [
    '    projected: dict[str, Any] = {}',
    '    handles = parsed.get("datasetHandles")',
    '    safe_handles = _project_safe_dataset_handles(handles)',
    '    if safe_handles is not None:',
    '        projected["datasetHandles"] = safe_handles',
    '    projected["code"] = PYTHON_CODE_WITHHELD_MARKER',
    '    return json.dumps(projected, ensure_ascii=False)',
  ].join("\n"),
  M2: [
    '    projected_messages: list[dict[str, Any]] = []',
    '    for item in messages:',
    '        if not isinstance(item, dict):',
    '            projected_messages.append(item)',
    '            continue',
    '        copied = dict(item)',
    '        _withhold_python_code_in_assistant_message(copied)',
    '        projected_messages.append(copied)',
    '    return projected_messages',
  ].join("\r\n"),
  M3: [
    '                    if name == "run_python_analysis":',
    '                        state.last_error_message = python_error_response_content()',
    '                        tool_content = python_error_response_content()',
    '                        correction_summary = (',
    '                            "隔离 Python 分析执行环节出错，本轮未运行任何代码。"',
    '                        )',
    '                        trace_basis = "隔离 Python 分析执行环节出错，本轮未运行任何代码。"',
    '                    else:',
  ].join("\r\n"),
  M4: [
    '                if tool_name == "run_python_analysis":',
    '                    return {',
    '                        "content": [{"type": "text", "text": python_error_response_content()}],',
    '                        "is_error": True,',
    '                    }',
  ].join("\n"),
};

const MUTANTS = {
  M1: [
    '    # [变异M1] 恢复"被拒绝参数原样历史保留"：退回只删 code 键。',
    '    projected = {key: value for key, value in parsed.items() if key != "code"}',
    '    projected["code"] = PYTHON_CODE_WITHHELD_MARKER',
    '    return json.dumps(projected, ensure_ascii=False)',
  ].join("\n"),
  M2: [
    '    # [变异M2] 绕过恢复现场投影：恒等返回原列表（不投影）。',
    '    return [dict(item) if isinstance(item, dict) else item for item in messages]',
  ].join("\r\n"),
  M3: [
    '                    if False:  # [变异M3] 恢复执行入口异常原样回显',
    '                        state.last_error_message = python_error_response_content()',
    '                        tool_content = python_error_response_content()',
    '                        correction_summary = (',
    '                            "隔离 Python 分析执行环节出错，本轮未运行任何代码。"',
    '                        )',
    '                        trace_basis = "隔离 Python 分析执行环节出错，本轮未运行任何代码。"',
    '                    else:',
  ].join("\r\n"),
  M4: [
    '                if False:  # [变异M4] SDK 恢复执行入口异常原样回显',
    '                    return {',
    '                        "content": [{"type": "text", "text": python_error_response_content()}],',
    '                        "is_error": True,',
    '                    }',
  ].join("\n"),
};

const which = process.argv[2];
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_python_privacy_oct06_third.mjs M1|M2|M3");
  process.exit(2);
}
const target = TARGETS[which];
const seed = SEEDS[which];
const mutant = MUTANTS[which];
const original = readFileSync(target.file, "utf8");
const occurrences = original.split(seed).length - 1;
if (occurrences !== 1) {
  console.log(`SKIP: SEED 在目标文件中出现 ${occurrences} 次（期望 1）。`);
  process.exit(3);
}

const touched = new Map();
function restoreAll() {
  for (const [file, content] of touched) writeFileSync(file, content, "utf8");
  touched.clear();
}
for (const sig of ["SIGINT", "SIGTERM", "SIGHUP"]) {
  process.on(sig, () => { restoreAll(); process.exit(130); });
}
process.on("uncaughtException", (err) => { restoreAll(); console.error("uncaughtException（已还原）:", err); process.exit(1); });

function runPytest(name) {
  const baseDir = mkdtempSync(join(tmpdir(), "mut3_"));
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
  touched.set(target.file, original);
  writeFileSync(target.file, original.replace(seed, mutant), "utf8");
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
