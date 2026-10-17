/**
 * 格式动作**恢复账本链路**的一条可复跑命令（`npm run test:wps-format-ledger`）。
 *
 * 它把这件事的两半接在**同一根链**上，而不是各自用替身：
 *
 *   真实 `<WpsAgentPanel/>` 的入口按钮 → 真实副本审阅 → 真实执行器
 *   → **真实出站契约校验** → **真实 HTTP** → **真实 ASGI 服务端 + 真实 SQLite 账本**
 *   → 真实登记 / 写前检查点 / 写后回执 / 完成回执 → 回读 → 撤销。
 *
 * 命令负责四件事（都只针对**本次**运行的进程与临时文件）：
 *
 * 1. 起**本次专用服务端**（真实 `server.app.main:app`，只绑回环、用临时端口），
 *    就绪检查通过后才跑用例；
 * 2. 把 `EXCEL_BRO_CONFIG_DIR` 指向临时目录 —— 不碰开发者本机账本；
 * 3. 写一份**临时能力表**（只给服务端用），把 `wpsWrite.maxCells` 压到用例的规模，
 *    让"服务端拒绝登记 ⇒ 原表零写入"这条用例靠**真实的服务端拒绝**触发，
 *    而不是靠 fetch 桩。客户端仍用仓库内置能力表（它读打包进来的
 *    `config/capabilities.json`，不受此变量影响）——两条线本来就可能不一致，
 *    这正是该用例要覆盖的形态；
 * 4. 只关闭本次启动的服务端进程，打印日志路径与退出码。
 *
 * 诚实口径：
 * - 任一步骤失败 ⇒ 非零退出，不报成功；
 * - 模型规划在用例里是**固定**的（规划由替身给出），本命令**不**证明"真实模型自主规划"；
 * - WPS 原生对象是**有状态宿主替身**，本命令**不**证明"真实 WPS 客户端"通过。
 *   这两项照实记为未完成（见 docs）。
 */
import { spawn, spawnSync } from "node:child_process";
import { createServer } from "node:net";
import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");
const isWin = process.platform === "win32";

/** 用例计划合计 12 格：服务端上限压到 12，再用一份 60 格的计划真实触发服务端拒绝。 */
const SERVER_MAX_CELLS = 12;

/** 优先用仓库 venv（既有测试约定），回退到 PATH 上的 python。 */
function pythonBin() {
  const venv = resolve(repoRoot, ".venv", "Scripts", "python.exe");
  const venvPosix = resolve(repoRoot, ".venv", "bin", "python");
  if (existsSync(venv)) return venv;
  if (existsSync(venvPosix)) return venvPosix;
  return isWin ? "python" : "python3";
}

function run(command, args, options = {}) {
  return spawnSync(command, args, { cwd: repoRoot, stdio: "inherit", shell: isWin, ...options });
}

/** 起一次用完即弃的监听，拿到一个当前空闲的回环端口。 */
function freePort() {
  return new Promise((resolvePort, reject) => {
    const probe = createServer();
    probe.on("error", reject);
    probe.listen(0, "127.0.0.1", () => {
      const { port } = probe.address();
      probe.close(() => resolvePort(port));
    });
  });
}

async function waitForServer(baseUrl, timeoutMs = 60_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const response = await fetch(`${baseUrl}/health`);
      if (response.ok) return true;
    } catch {
      // 还没起来，继续等。
    }
    await new Promise((r) => setTimeout(r, 300));
  }
  return false;
}

// ---------------------------------------------------------------------------
// 1) 临时能力表 + 临时账本目录（只影响本次进程）
// ---------------------------------------------------------------------------
//
// 每次运行用**独立目录**，绝不复用上一轮的 SQLite 账本：复用会让"账本里零记录"
// 这类断言读到旧账。历史目录尽力清理，清理失败**不影响本次结论** —— 被杀掉的
// uvicorn 若成了孤儿进程，会占住上一轮目录里的 sqlite 文件，删不掉只说明它在
// 清理时机上还没退出，不代表本次链路有问题。
const outputsDir = resolve(repoRoot, "outputs");
mkdirSync(outputsDir, { recursive: true });
for (const entry of readdirSync(outputsDir, { withFileTypes: true })) {
  if (!entry.isDirectory() || !entry.name.startsWith("wps-format-ledger")) continue;
  try {
    rmSync(resolve(outputsDir, entry.name), { recursive: true, force: true });
  } catch {
    // 上一轮目录被占用：留着，本次换独立目录。
  }
}
const runDir = resolve(
  outputsDir,
  `wps-format-ledger-${process.pid.toString(36)}-${Date.now().toString(36)}`
);
mkdirSync(runDir, { recursive: true });

const capabilities = JSON.parse(
  readFileSync(resolve(repoRoot, "config", "capabilities.json"), "utf8")
);
capabilities.wpsWrite.maxCells = SERVER_MAX_CELLS;
const capabilitiesPath = resolve(runDir, "capabilities.json");
writeFileSync(capabilitiesPath, JSON.stringify(capabilities, null, 2), "utf8");

const configDir = resolve(runDir, "config");
mkdirSync(configDir, { recursive: true });

const logPath = resolve(runDir, "server.log");
writeFileSync(logPath, "", "utf8");

// ---------------------------------------------------------------------------
// 2) 起本次的服务端（真实 ASGI 应用、真实路由、真实 SQLite 账本）
// ---------------------------------------------------------------------------
const port = await freePort();
const baseUrl = `http://127.0.0.1:${port}`;
console.log(`[wps-format-ledger] 启动测试服务端（只绑回环）：${baseUrl}`);

const server = spawn(
  pythonBin(),
  ["-m", "uvicorn", "server.app.main:app", "--host", "127.0.0.1", "--port", String(port), "--log-level", "warning"],
  {
    cwd: repoRoot,
    stdio: ["ignore", "pipe", "pipe"],
    shell: isWin,
    env: {
      ...process.env,
      EXCEL_BRO_CAPABILITIES_PATH: capabilitiesPath,
      EXCEL_BRO_CONFIG_DIR: configDir,
      AI_BASE_URL: "http://127.0.0.1:1",
      AI_MODEL: "format-ledger-test-model"
    }
  }
);

const logStream = [];
server.stdout?.on("data", (chunk) => logStream.push(String(chunk)));
server.stderr?.on("data", (chunk) => logStream.push(String(chunk)));

/**
 * 结束服务端。Windows 上 `spawn(..., {shell: true})` 起的是 cmd.exe 包装，
 * `server.kill()` 只杀包装进程，**带不走 python 子进程**，留下占住 sqlite 的孤儿。
 * 这里按进程树 kill，并把回收结果写进日志（回收失败要看得见，不静默）。
 */
function killServer() {
  if (server.exitCode !== null) return;
  if (isWin && server.pid) {
    const killed = spawnSync("taskkill", ["/pid", String(server.pid), "/T", "/F"], {
      stdio: "ignore"
    });
    logStream.push(`\n[runner] taskkill /T /F pid=${server.pid} status=${killed.status}\n`);
    return;
  }
  try {
    server.kill("SIGTERM");
  } catch (error) {
    logStream.push(`\n[runner] 发送 SIGTERM 失败：${String(error)}\n`);
  }
}

const bye = (code) => {
  try {
    writeFileSync(logPath, logStream.join(""), "utf8");
  } catch {
    // 日志只是诊断材料，写不进去不改变结论。
  }
  killServer();
  process.exit(code);
};

const ready = await waitForServer(baseUrl);
if (!ready) {
  console.error(
    `格式动作恢复账本链路未完成：测试服务端未就绪（${baseUrl}）。\n` +
      `服务端输出：\n${logStream.join("")}`
  );
  bye(2);
}
// 被外部信号打断也要把服务端进程树带走，不留孤儿占住临时账本。
for (const signal of ["SIGINT", "SIGTERM", "SIGHUP"]) {
  process.on(signal, () => bye(1));
}
console.log(`[wps-format-ledger] 服务端就绪：${baseUrl}/health（服务端 wpsWrite.maxCells=${SERVER_MAX_CELLS}）`);

// ---------------------------------------------------------------------------
// 3) 跨进程用例（真实面板 + 真实出站校验 + 真实 HTTP + 真实账本）
// ---------------------------------------------------------------------------
const vitest = resolve(repoRoot, "node_modules", ".bin", isWin ? "vitest.cmd" : "vitest");
console.log("[wps-format-ledger] 真实面板：副本审阅 → 登记 → 确认写回 → 回读 → 撤销…");
const result = run(
  vitest,
  ["run", "--config", "vite.config.ts", "src/wpsFormatRecoveryLedger.test.tsx"],
  {
    cwd: resolve(repoRoot, "apps", "wps-addin"),
    env: {
      ...process.env,
      EB_WPS_LEDGER: "1",
      VITE_API_BASE_URL: baseUrl,
      EXCEL_BRO_CONFIG_DIR: configDir
    }
  }
);

if (result.status !== 0) {
  console.error(
    `格式动作恢复账本链路未完成：用例未通过（退出码 ${result.status}）。\n` +
      `  用例：apps/wps-addin/src/wpsFormatRecoveryLedger.test.tsx\n` +
      `  服务端日志：${logPath}`
  );
  bye(result.status ?? 1);
}

console.log(
  "格式动作恢复账本链路完成：真实面板在真实出站校验 + 真实 HTTP + 真实服务端账本上" +
    "跑通了 副本审阅 → 登记 → 确认写回 → 回读 → 撤销。\n" +
    `[wps-format-ledger] 服务端日志：${logPath}\n` +
    "[wps-format-ledger] 说明：模型规划固定、WPS 原生对象为有状态替身，" +
    "因此「真实模型自主规划」与「真实 WPS 客户端」两项**未**通过本命令证明。"
);
bye(0);
