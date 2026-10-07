// 反向变异脚本：隔离 Python 隐私/入参骨架（2026-10-06）
// 用法：node scripts/mutate_python_privacy_1006.mjs M1|M2|M3|M4
// 每项流程：备份 → 下刀 → 跑 server/tests/test_python_analysis.py → 还原 → 复跑。
// 铁律：信号钩子保证还原；写前登记 touched；SEED 唯一性校验（0 或 >1 次即 SKIP）。
import { readFileSync, writeFileSync, mkdtempSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { tmpdir } from "node:os";
import { join } from "node:path";

const REPO = "E:\\codex\\excel-bro";
const PY = join(REPO, ".venv", "Scripts", "python.exe");
const TARGETS = {
  M1: {
    file: join(REPO, "server", "app", "excel_agent.py"),
    expectFail: [
      "test_python_code_never_reenters_model_requests",
      "test_python_code_never_enters_sqlite_persistence",
      "test_resume_python_call_reports_code_not_preserved",
    ],
    mustStayGreen: [],
  },
  M2: {
    file: join(REPO, "server", "app", "python_analysis.py"),
    expectFail: [
      "test_receipt_message_regenerated_from_trusted_state",
      "test_receipt_rejects_unknown_status_and_combinations",
      "test_receipt_rejects_forged_ids_handles_and_unknown_fields",
    ],
    mustStayGreen: [],
  },
  M3: {
    file: join(REPO, "server", "app", "python_analysis.py"),
    expectFail: ["test_extra_fields_rejected_at_actual_entry"],
    mustStayGreen: ["test_boundary_lengths_follow_capabilities"],
  },
  M4: {
    file: join(REPO, "server", "app", "excel_agent.py"),
    expectFail: ["test_projection_failure_fails_closed_at_both_outlets"],
    mustStayGreen: ["test_outbound_gate_passes_python_receipt_and_read_range"],
  },
};

const which = process.argv[2];
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_python_privacy_1006.mjs M1|M2|M3|M4");
  process.exit(2);
}
const target = TARGETS[which];

// ---- SEED：按目标文件换行符归一化构造（excel_agent.py=CRLF，python_analysis.py=LF）----
const NL = target.file.includes("excel_agent") ? "\r\n" : "\n";
const SEEDS = {
  M1: [
    'def _withhold_python_code_in_assistant_message(message: dict[str, Any]) -> None:',
    '    """缺陷 1 修复：``run_python_analysis`` 的原始代码不进入任何现场。',
    '',
    '    在 assistant 消息进入消息历史 / 检查点 / 恢复现场**之前**调用：',
    '    ``code`` 字段确定性替换为固定占位（不做内容扫描、不做全局字符串替换），',
    '    ``tool_call_id``、工具名与其余参数原样保留，协议配对不变。幂等。',
    '    """',
    '    tool_calls = message.get("tool_calls")',
    '    if not isinstance(tool_calls, list):',
    '        return',
    '    safe_calls: list[Any] = []',
    '    changed = False',
    '    for call in tool_calls:',
    '        if not isinstance(call, dict):',
    '            safe_calls.append(call)',
    '            continue',
    '        function = call.get("function")',
    '        if (',
    '            isinstance(function, dict)',
    '            and function.get("name") == "run_python_analysis"',
    '        ):',
    '            safe_calls.append({',
    '                **call,',
    '                "function": {',
    '                    **function,',
    '                    "arguments": withheld_python_arguments(',
    '                        function.get("arguments")',
    '                        if isinstance(function.get("arguments"), str) else None',
    '                    ),',
    '                },',
    '            })',
    '            changed = True',
    '        else:',
    '            safe_calls.append(call)',
    '    if changed:',
    '        message["tool_calls"] = safe_calls',
    '',
    '',
  ].join(NL),
  M2: [
    'def project_python_analysis_receipt(receipt: dict[str, Any]) -> dict[str, Any]:',
    '    """有类型回执校验与投影：模型可见内容的唯一出口。',
    '',
    '    规则（全部失败关闭，非法即抛 ValueError，调用出口必须转固定安全文案）：',
    '    - 回执必须是对象；字段集合 ⊆ :data:`RECEIPT_KEYS`，白名单外字段拒绝；',
    '    - ``status``/``reasonCategory`` 必须是封闭枚举，且组合在登记表内；',
    '    - 当前没有运行器与结果登记：``runId``/``resultHandle`` 出现即伪造，拒绝；',
    '      将来支持句柄时必须由可信协调器生成并核对任务归属，不能只验证',
    '      字符串前缀或正则；',
    '    - ``message`` 不读取输入——按可信状态组合重新生成固定公开文案。',
    '    """',
    '    if not isinstance(receipt, dict):',
    '        raise ValueError("Python 分析回执必须是对象")',
    '    if set(receipt.keys()) - set(RECEIPT_KEYS):',
    '        raise ValueError("回执包含未登记字段")',
    '    status = receipt.get("status")',
    '    reason_category = receipt.get("reasonCategory")',
    '    if not isinstance(status, str) or status not in RECEIPT_STATUSES:',
    '        raise ValueError("回执状态非法")',
    '    if not isinstance(reason_category, str):',
    '        raise ValueError("回执原因类别非法")',
    '    if (status, reason_category) not in _ALLOWED_RECEIPT_COMBINATIONS:',
    '        raise ValueError("回执状态组合非法")',
    '    if receipt.get("runId") is not None or receipt.get("resultHandle") is not None:',
    '        raise ValueError("回执携带未登记的运行标识或结果句柄")',
    '    return {',
    '        "status": status,',
    '        "reasonCategory": reason_category,',
    '        "message": _fixed_receipt_message(status, reason_category),',
    '    }',
    '',
    '',
  ].join(NL),
  M3: [
    '    if set(arguments.keys()) - set(_ARGUMENT_KEYS):',
    '        return False',
  ].join(NL),
  M4: [
    '    except ValueError:',
    '        if name == "run_python_analysis":',
    '            return PYTHON_PROJECTION_FAILED_CONTENT',
    '        raise',
  ].join(NL),
};
const MUTANTS = {
  M1: [
    'def _withhold_python_code_in_assistant_message(message: dict[str, Any]) -> None:',
    '    """[变异M1] 拆除持久化/历史投影保护：no-op。"""',
    '    return',
    '',
    '',
  ].join(NL),
  M2: [
    'def project_python_analysis_receipt(receipt: dict[str, Any]) -> dict[str, Any]:',
    '    """[变异M2] 拆除回执值校验：原样放行。"""',
    '    return dict(receipt)',
    '',
    '',
  ].join(NL),
  M3: [
    '    if False:  # [变异M3] 拆除未知字段拒绝',
    '        return False',
  ].join(NL),
  M4: [
    '    except ValueError:',
    '        raise  # [变异M4] 拆除异常出口保护',
  ].join(NL),
};

const seed = SEEDS[which];
const mutant = MUTANTS[which];
const original = readFileSync(target.file, "utf8");
const occurrences = original.split(seed).length - 1;
if (occurrences !== 1) {
  console.log(`SKIP: SEED 在目标文件中出现 ${occurrences} 次（期望 1）。`);
  process.exit(3);
}

// 写前登记 touched 表；信号钩子保证还原。
const touched = new Map();
function restoreAll() {
  for (const [file, content] of touched) writeFileSync(file, content, "utf8");
  touched.clear();
}
for (const sig of ["SIGINT", "SIGTERM", "SIGHUP"]) {
  process.on(sig, () => { restoreAll(); process.exit(130); });
}
process.on("uncaughtException", (err) => {
  restoreAll();
  console.error("uncaughtException（已还原）:", err);
  process.exit(1);
});

function runPytest(basetempName) {
  const baseDir = mkdtempSync(join(tmpdir(), "mut_pytest_"));
  const r = spawnSync(PY, [
    "-m", "pytest", "server/tests/test_python_analysis.py", "-q",
    "-p", "no:cacheprovider",
    "--basetemp", join(baseDir, basetempName),
  ], {
    cwd: REPO,
    encoding: "utf8",
    env: {
      ...process.env,
      TMP: baseDir,
      TEMP: baseDir,
      PYTEST_ADDOPTS: "",
    },
    timeout: 300000,
  });
  return `${r.stdout || ""}\n${r.stderr || ""}`;
}

// 1) 基线：必须 GREEN
const baselineOut = runPytest("baseline");
const baselinePass = /(\d+) passed/.exec(baselineOut);
const baselineFailed = /(\d+) failed/.exec(baselineOut);
console.log(`[基线] ${baselinePass ? baselinePass[1] + " passed" : "?"}${baselineFailed ? ", " + baselineFailed[1] + " failed" : ""}`);
if (!baselinePass || baselineFailed) {
  console.log("BASELINE-NOT-GREEN：基线未全绿，终止变异（源码未改动）。");
  console.log(baselineOut.slice(-1500));
  process.exit(4);
}

// 2) 下刀
try {
  touched.set(target.file, original);
  writeFileSync(target.file, original.replace(seed, mutant), "utf8");

  // 3) 变异后跑测试：必须 RED，且失败名单命中 expectFail
  const mutatedOut = runPytest("mutated");
  const failedNames = [...mutatedOut.matchAll(/FAILED (server\/tests\/\w+\.py::\w+)/g)]
    .map((m) => m[1].split("::")[1]);
  const hits = target.expectFail.filter((n) => failedNames.includes(n));
  const stayGreenBroken = target.mustStayGreen.filter((n) => failedNames.includes(n));
  console.log(`[变异后] 失败 ${failedNames.length} 条；命中判别点: ${hits.length ? hits.join(", ") : "无"}`);
  console.log(`[失败名单] ${failedNames.join(", ") || "（无）"}`);
  if (stayGreenBroken.length) {
    console.log(`警告：应保持绿的用例也红了（判别点不唯一）: ${stayGreenBroken.join(", ")}`);
  }
  const red = failedNames.length > 0 && hits.length > 0 && stayGreenBroken.length === 0;
  console.log(red ? "RED-OK：变异使判别测试独立变红。" : "RED-MISS：变异未按预期变红（或判别点不唯一）。");
  if (!red) console.log(mutatedOut.slice(-2000));

  // 4) 还原
  restoreAll();
  const restoredOut = runPytest("restored");
  const restoredPass = /(\d+) passed/.exec(restoredOut);
  const restoredFailed = /(\d+) failed/.exec(restoredOut);
  console.log(`[还原后] ${restoredPass ? restoredPass[1] + " passed" : "?"}${restoredFailed ? ", " + restoredFailed[1] + " failed" : ""}`);
  const green = restoredPass && !restoredFailed;
  console.log(green ? "RESTORED-GREEN：还原后复跑全绿。" : "RESTORED-NOT-GREEN：还原后仍有失败！");
  process.exit(red && green ? 0 : 5);
} catch (err) {
  restoreAll();
  throw err;
}
