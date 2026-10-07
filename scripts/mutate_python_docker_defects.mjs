// 反向变异：Docker 隔离执行器七类缺陷修复的判别力（2026-10-07, Request H）
// 用法：node scripts/mutate_python_docker_defects.mjs M1|M2|M3|M4|M5|M6|M7
//
// 每个变异**只拆掉一类修复机制**，要求对应的判别测试独立变红。
// 铁律：① 基线先 GREEN；② 信号钩子还原；③ touched 登记；④ SEED 唯一。
//
// 覆盖：
//  M1 = 拆「构建上下文/COPY 源」检查（Dockerfile 改回找不到源的形式）
//  M2 = 拆「stdin 传 bytes」（改回传 str）
//  M3 = 拆「超时清理闭环」（超时不调用 _force_cleanup）
//  M4 = 拆「有界输出读取」（改回 capture_output=True）
//  M5 = 拆「隔离配置校验」（validate_isolation_config 恒返回 True）
//  M6 = 拆「回包严格校验」（status=error 也当 completed）
//  M7 = 拆「真实测试防假通过」（网络测试读未定义变量）
import { readFileSync, writeFileSync, mkdtempSync } from "node:fs";
import { spawnSync as spawn } from "node:child_process";
import { tmpdir } from "node:os";
import { join } from "node:path";

const REPO = "E:\\codex\\excel-bro";
const PY = join(REPO, ".venv", "Scripts", "python.exe");
const RUNNER = join(REPO, "server", "app", "python_docker_runner.py");
const DOCKERFILE = join(REPO, "server", "docker", "Dockerfile");
const REALTESTS = join(REPO, "server", "tests", "test_python_docker_real.py");
const DEFECT_TESTS = "server/tests/test_python_docker_defects.py";

const NL = "\n";

const TARGETS = {
  M1: {
    file: DOCKERFILE,
    expectFail: ["test_dockerfile_copy_source_exists_in_build_context"],
    tests: DEFECT_TESTS,
  },
  M2: {
    file: RUNNER,
    expectFail: ["test_run_python_isolated_passes_bytes_to_stdin"],
    tests: DEFECT_TESTS,
  },
  M3: {
    file: RUNNER,
    expectFail: ["test_timeout_runs_named_container_and_cleans_up"],
    tests: DEFECT_TESTS,
  },
  M4: {
    file: RUNNER,
    expectFail: ["test_stderr_output_is_bounded_not_unbounded", "test_output_is_read_with_bound_not_unbounded_capture"],
    tests: DEFECT_TESTS,
  },
  M5: {
    file: RUNNER,
    expectFail: ["test_isolation_config_validation_rejects_incomplete_config"],
    tests: DEFECT_TESTS,
  },
  M6: {
    file: RUNNER,
    expectFail: ["test_worker_status_error_with_zero_exit_is_not_completed"],
    tests: DEFECT_TESTS,
  },
  M7: {
    file: REALTESTS,
    expectFail: ["test_real_test_network_probe_reads_defined_variable"],
    tests: DEFECT_TESTS,
  },
};

// SEED 必须从函数声明行开始、覆盖到函数末尾 }（或足够唯一的片断）。
const SEEDS = {
  M1: [
    'COPY app/docker_worker.py /app/worker.py',
  ].join(NL),
  M2: [
    '                    input=input_bytes,      # bytes：二进制/stdin 默认模式',
  ].join(NL),
  M3: [
    '            except subprocess.TimeoutExpired:',
    '                # 超时：容器可能仍在运行 → 强制清理（闭环）',
    '                _force_cleanup(container_name)',
    '                return DockerRunResult("timeout", None, None, "timeout")',
  ].join(NL),
  M4: [
    '                proc = subprocess.run(',
    '                    args,',
    '                    input=input_bytes,      # bytes：二进制/stdin 默认模式',
    '                    stdout=out_f,           # 有界：写文件，不缓冲进内存',
    '                    stderr=err_f,',
    '                    timeout=wall_timeout_seconds,',
    '                    shell=False,',
    '                    env=env,',
    '                )',
  ].join(NL),
  M5: [
    '    problems: list[str] = []',
  ].join(NL),
  M6: [
    '    if status == _WORKER_ERROR_STATUS:',
    '        return False, "execution_error", None',
  ].join(NL),
  M7: [
    "            \"result = '|'.join(out)\\n\"",
  ].join(NL),
};

const MUTANTS = {
  M1: [
    'COPY docker_worker.py /app/worker.py',
  ].join(NL),
  M2: [
    '                    input=json.loads(input_bytes.decode("utf-8")),  # [变异M2] str 而非 bytes',
  ].join(NL),
  M3: [
    '            except subprocess.TimeoutExpired:',
    '                # [变异M3] 拆超时清理：只返回，不清理容器',
    '                return DockerRunResult("timeout", None, None, "timeout")',
  ].join(NL),
  M4: [
    '                proc = subprocess.run(',
    '                    args,',
    '                    input=input_bytes,',
    '                    capture_output=True,  # [变异M4] 无界缓冲',
    '                    timeout=wall_timeout_seconds,',
    '                    shell=False,',
    '                    env=env,',
    '                )',
  ].join(NL),
  M5: [
    '    problems: list[str] = []',
    '    return True, None  # [变异M5] 拆隔离配置校验',
  ].join(NL),
  M6: [
    '    if status == _WORKER_ERROR_STATUS:',
    '        return True, None, payload.get("result")  # [变异M6] 把 error 当成功',
  ].join(NL),
  M7: [
    "            \"result = '|'.join(out)\\n\"  # noqa",
  ].join(NL) + NL + NL + "@@M7_PATCH@@",
};

// M7 特殊：直接把 result 赋值行删掉（模拟读未定义变量）
const M7_ORIG = "            \"result = '|'.join(out)\\n\"\n";
const M7_NEW = "            # [变异M7] 删掉 result 赋值（模拟读未定义变量）\n";

const which = process.argv[2];
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_python_docker_defects.mjs M1|M2|M3|M4|M5|M6|M7");
  process.exit(2);
}
const target = TARGETS[which];
const touched = new Map();

function restoreAll() {
  for (const [f, c] of touched) writeFileSync(f, c, "utf8");
  touched.clear();
}
for (const sig of ["SIGINT", "SIGTERM", "SIGHUP"]) {
  process.on(sig, () => { restoreAll(); process.exit(130); });
}
process.on("uncaughtException", (err) => { restoreAll(); console.error("uncaughtException（已还原）:", err); process.exit(1); });

function applyMutation() {
  if (which === "M7") {
    const original = readFileSync(REALTESTS, "utf8");
    if ((original.split(M7_ORIG).length - 1) !== 1) {
      console.log("SKIP: M7 SEED 不唯一");
      process.exit(3);
    }
    touched.set(REALTESTS, original);
    writeFileSync(REALTESTS, original.replace(M7_ORIG, M7_NEW), "utf8");
    return;
  }
  const file = target.file;
  const original = readFileSync(file, "utf8");
  const seed = SEEDS[which];
  const mutant = MUTANTS[which];
  const occurrences = original.split(seed).length - 1;
  if (occurrences !== 1) {
    console.log(`SKIP: ${which} SEED 出现 ${occurrences} 次（期望 1）。`);
    console.log("SEED 片段：\n" + seed);
    process.exit(3);
  }
  touched.set(file, original);
  writeFileSync(file, original.replace(seed, mutant), "utf8");
}

function runPytest(testPath, name) {
  const baseDir = mkdtempSync(join(tmpdir(), "mutdef_"));
  const r = spawn(PY, ["-m", "pytest", testPath, "-q",
    "-p", "no:cacheprovider", "--basetemp", join(baseDir, name)], {
    cwd: REPO, encoding: "utf8",
    env: { ...process.env, TMP: baseDir, TEMP: baseDir, PYTEST_ADDOPTS: "" },
    timeout: 180000,
  });
  return `${r.stdout || ""}\n${r.stderr || ""}`;
}

const baseline = runPytest(target.tests, "baseline");
const bp = /(\d+) passed/.exec(baseline);
const bf = /(\d+) failed/.exec(baseline);
console.log(`[基线] ${bp ? bp[1] + " passed" : "?"}${bf ? ", " + bf[1] + " failed" : ""}`);
if (!bp || bf) {
  console.log("BASELINE-NOT-GREEN：基线未全绿，终止（源码未改动）。");
  console.log(baseline.slice(-1500));
  process.exit(4);
}

try {
  applyMutation();
  const mutated = runPytest(target.tests, "mutated");
  const failed = [...mutated.matchAll(/FAILED (server\/tests\/\w+\.py::[\w\[\]\-]+)/g)].map((m) => m[1].split("::")[1].split("[")[0]);
  const uniqFailed = [...new Set(failed)];
  const hits = target.expectFail.filter((n) => uniqFailed.includes(n));
  console.log(`[变异后] 失败 ${uniqFailed.length} 条：${uniqFailed.join(", ") || "（无）"}`);
  console.log(`[命中判别点] ${hits.length ? hits.join(", ") : "无"}`);
  const red = uniqFailed.length > 0 && hits.length > 0;
  console.log(red ? "RED-OK：变异使判别测试独立变红。" : "RED-MISS：变异未按预期变红。");
  if (!red) console.log(mutated.slice(-2500));

  restoreAll();
  const restored = runPytest(target.tests, "restored");
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
