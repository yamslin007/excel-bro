/**
 * 跨前后端「同一次任务数据流」的**一条可复跑命令**（`npm run test:cross-process`）。
 *
 * 它把两半按真实顺序串起来：
 *
 * 1. **生产侧**（`server/tests/test_cross_process_chain_real.py`）：真实 ASGI 服务端
 *    + 真实 Docker 隔离容器，跑 桥读表 → prepare_python_data → 句柄登记 →
 *    run_python_analysis → 结果登记 → 按引用取回 → `compile_result_plan`，
 *    并把那次执行的产物导出到 `config/cross_process_chain.json`。
 *    *本机没有 Docker 时它会被 pytest 判为环境阻断（skip），**不产出工件***。
 *
 * 2. **消费侧**（`apps/excel-addin/src/crossProcessChain.test.tsx`）：读那份工件，
 *    在真实 `<App/>` 上完成 预览 → 确认 → 写回 → 回读 → 撤销。
 *    消费侧默认不跑（缺工件时会崩收集/类型检查），只在**本条命令确证工件已产出**
 *    之后，带上 `EB_CROSS_PROCESS=1` 显式门控执行。
 *
 * 诚实口径：
 * - 生产侧 Docker 不可用 ⇒ 无工件 ⇒ 本脚本打印「连续链路未完成：生产侧环境阻断」并以
 *   非零退出，**不**拿旧工件或 mock 顶替通过。
 * - 生产侧通过但未产出/未刷新工件 ⇒ 同样判「连续链路未完成」。
 */
import { spawnSync } from "node:child_process";
import { existsSync, readFileSync, rmSync, statSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");
const fixture = resolve(repoRoot, "config", "cross_process_chain.json");
const isWin = process.platform === "win32";

/** 优先用仓库 venv（既有测试约定），回退到 PATH 上的 python。 */
function pythonBin() {
  const venv = resolve(repoRoot, ".venv", "Scripts", "python.exe");
  const venvPosix = resolve(repoRoot, ".venv", "bin", "python");
  if (existsSync(venv)) return venv;
  if (existsSync(venvPosix)) return venvPosix;
  return isWin ? "python" : "python3";
}

function run(command, args, options = {}) {
  return spawnSync(command, args, {
    cwd: repoRoot,
    stdio: "inherit",
    shell: isWin,
    ...options
  });
}

// 先删旧工件：只有本次生产侧真的跑出结果，才算"接通"。
if (existsSync(fixture)) rmSync(fixture);

console.log("[cross-process] 生产侧：真实服务端 + 真实 Docker 隔离容器…");
const producer = run(pythonBin(), [
  "-m",
  "pytest",
  "server/tests/test_cross_process_chain_real.py",
  "-p",
  "no:cacheprovider"
]);
if (producer.status !== 0) {
  console.error("连续链路未完成：生产侧用例未通过（见上）。");
  process.exit(producer.status ?? 1);
}
if (!existsSync(fixture)) {
  console.error(
    "连续链路未完成：生产侧环境阻断（多半是 Docker 隔离引擎不可用），未产出 " +
      "config/cross_process_chain.json。本命令不会用旧工件或 mock 顶替通过。"
  );
  process.exit(2);
}

const artifact = JSON.parse(readFileSync(fixture, "utf8"));
console.log(
  `[cross-process] 工件已产出: ${fixture}\n` +
    `[cross-process]   结果引用: ${artifact.resultRef}\n` +
    `[cross-process]   计划: ${artifact.plan?.id} → ` +
    `${artifact.plan?.actions?.[0]?.sheet}!${artifact.plan?.actions?.[0]?.range}\n` +
    `[cross-process]   落盘时间: ${statSync(fixture).mtime.toISOString()}`
);

console.log("[cross-process] 消费侧：真实 <App/> 预览 → 确认 → 写回 → 回读 → 撤销…");
// 在**工作区目录**里跑消费侧用例：`@excel-bro/core` 别名与 jsdom 环境都在
// `apps/excel-addin/vite.config.ts` 里，从仓库根跑会解析不到。`EB_CROSS_PROCESS=1`
// 是消费侧唯一的门控 —— 只有本条命令（在确证工件已产出之后）会带上它。
const vitest = resolve(repoRoot, "node_modules", ".bin", isWin ? "vitest.cmd" : "vitest");
const consumer = run(vitest, ["run", "src/crossProcessChain.test.tsx"], {
  cwd: resolve(repoRoot, "apps", "excel-addin"),
  env: { ...process.env, EB_CROSS_PROCESS: "1" }
});
if (consumer.status !== 0) {
  console.error("连续链路未完成：消费侧用例未通过（见上）。");
  process.exit(consumer.status ?? 1);
}
console.log(
  "连续链路完成：生产侧（真实服务端 + 真实容器）与消费侧（真实面板写回）已接通。"
);
