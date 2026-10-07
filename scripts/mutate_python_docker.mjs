// 反向变异：Docker 隔离运行时关键守卫（2026-10-07）
// 用法：node scripts/mutate_python_docker.mjs M1|M2|M3
// M1 = 拆引擎类型验证（Windows/不可用也放行）
// M2 = 拆网络隔离参数（--network none）
// M3 = 拆输出字节限额
// 铁律：基线 GREEN；信号钩子还原；touched 登记；SEED 唯一性。
import { readFileSync, writeFileSync, mkdtempSync } from "node:fs";
import { spawnSync as spawn } from "node:child_process";
import { tmpdir } from "node:os";
import { join } from "node:path";

const REPO = "E:\\codex\\excel-bro";
const PY = join(REPO, ".venv", "Scripts", "python.exe");
const FILE = join(REPO, "server", "app", "python_docker_runner.py");
const NL = "\n";

const TARGETS = {
  M1: { expectFail: ["test_verify_engine_rejects_non_linux"] },
  M2: { expectFail: ["test_build_run_arguments_is_fixed_and_isolated"] },
  M3: { expectFail: ["test_run_python_isolated_enforces_output_limit"] },
};

const SEEDS = {
  M1: [
    '    ostype = proc.stdout.strip()',
    '    # 只接受本机 Linux 容器引擎；拒绝 Windows 容器（模型代码是 Python，需 Linux 运行时）',
    '    return ostype == "linux"',
  ].join(NL),
  M2: [
    '    args: list[str] = [',
    '        "docker", "run", "--rm",',
    '        "--network", config.network,',
  ].join(NL),
  M3: [
    '    # 4. 输出总字节限制（宿主协调器二次把关，不依赖工作程序）',
    '    if len(proc.stdout) > max_output_bytes:',
    '        return DockerRunResult("limit_exceeded", proc.returncode, None, "output_limit_exceeded")',
  ].join(NL),
};

const MUTANTS = {
  M1: [
    '    ostype = proc.stdout.strip()',
    '    # [变异M1] 拆引擎类型验证：任何返回都当可用',
    '    return True',
  ].join(NL),
  M2: [
    '    args: list[str] = [',
    '        "docker", "run", "--rm",',
    '        # [变异M2] 拆网络隔离',
  ].join(NL),
  M3: [
    '    # [变异M3] 拆输出字节限额（宿主二次把关被移除）',
  ].join(NL),
};

const which = process.argv[2];
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_python_docker.mjs M1|M2|M3");
  process.exit(2);
}
const target = TARGETS[which];
const seed = SEEDS[which];
const mutant = MUTANTS[which];
const original = readFileSync(FILE, "utf8");
const occurrences = original.split(seed).length - 1;
if (occurrences !== 1) {
  console.log(`SKIP: SEED 出现 ${occurrences} 次（期望 1）。`);
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
  const baseDir = mkdtempSync(join(tmpdir(), "mutdocker_"));
  const r = spawn(PY, ["-m", "pytest", "server/tests/test_python_docker_runner.py", "-q",
    "-p", "no:cacheprovider", "--basetemp", join(baseDir, name)], {
    cwd: REPO, encoding: "utf8",
    env: { ...process.env, TMP: baseDir, TEMP: baseDir, PYTEST_ADDOPTS: "" },
    timeout: 120000,
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
