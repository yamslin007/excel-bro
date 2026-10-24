/**
 * **双端在线连续集成**的一条可复跑命令（`npm run test:online`）。
 *
 * 它跑的是**同一次任务从头到尾不换手**的在线链路：真实面板 → 真实 HTTP → 真实
 * 规划流（SSE）→ `bridgeTool` 指针 → 面板自己的桥读宿主 → 真实回执 → 服务端登记
 * 数据集 → `run_python_analysis` → **真实 Docker 隔离执行器** → 结果登记与引用
 * → 白名单计划编译 → 面板预览 → 用户确认 → 产品执行器写回 → 回读 → 撤销。
 *
 * 命令负责的四件事（都只针对**本次**运行的进程与临时文件）：
 *
 * 1. 起**测试专用服务端**（`server/tests/online_entry.py`：显式测试入口，只绑
 *    回环地址、用专用端口、注入固定模型替身），就绪检查通过后才跑用例；
 * 2. 临时打开隔离 Python 总开关（写临时能力表，不动生产 `config/capabilities.json`），
 *    并把 `EXCEL_BRO_CONFIG_DIR` 指向临时目录 —— 否则开发者**真实的模型连接**
 *    会混进来，目录里不该出现它；
 * 3. 依次跑 Excel 与 WPS 两端的在线用例，两端共用 `config/online_link_scenario.json`
 *    的同一份数据 / 规则 / 期望；
 * 4. 只关闭本次启动的服务端进程，打印日志路径与退出码。
 *
 * 诚实口径：
 * - Docker 隔离引擎不可用 ⇒ 打印「环境阻断」并**非零退出**，不报成功；
 * - 任一端用例未通过 ⇒ 非零退出，并指明断在哪一端；
 * - 模型替身只在 `server/tests/online_entry.py` 注入，生产没有任何旁路开关。
 *
 * 注意：本命令**不能**证明"真实模型自主规划"与"真实 Excel / WPS 客户端"通过 ——
 * 本轮允许固定模型与有状态宿主替身，这两项照实记为未完成（见 docs）。
 */
import { spawn, spawnSync } from "node:child_process";
import { createServer } from "node:net";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";
import { tmpdir } from "node:os";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");
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

/**
 * 只看**退出码**：`docker version` 连得上守护进程才返回 0。
 * 这里不捕获 stdout（流水线里捕获子进程管道会被拦），输出直接进控制台。
 */
function dockerAvailable() {
  const probe = run("docker", ["version"], { stdio: "inherit" });
  return !probe.error && probe.status === 0;
}

// ---------------------------------------------------------------------------
// 1) 环境阻断检查：没有真实隔离引擎就不算在线链路接通
// ---------------------------------------------------------------------------
console.log("[online] 检查隔离 Python 引擎（Docker）…");
if (!dockerAvailable()) {
  console.error(
    "在线连续链路未完成：环境阻断 —— 本机 Docker 隔离引擎不可用，" +
      "`run_python_analysis` 无法真实执行。本命令不会用 mock 或旧工件顶替通过。"
  );
  process.exit(2);
}

// ---------------------------------------------------------------------------
// 2) 临时能力与配置（只影响本次进程）
// ---------------------------------------------------------------------------
// Preserve prior evidence and avoid deleting a directory still held by Windows.
mkdirSync(resolve(repoRoot, "outputs"), { recursive: true });
const runDir = mkdtempSync(resolve(repoRoot, "outputs", "online-link-"));
console.log(`[online] 本次证据目录：${runDir}`);

const capabilitiesPath = resolve(runDir, "capabilities.json");
const capabilities = JSON.parse(
  readFileSync(resolve(repoRoot, "config", "capabilities.json"), "utf8")
);
if (!capabilities.pythonExecution) capabilities.pythonExecution = {};
capabilities.pythonExecution.enabled = true;
writeFileSync(capabilitiesPath, JSON.stringify(capabilities, null, 2), "utf8");

// 隔离模型目录：不让开发者真实的已保存连接混进"固定模型"这次运行。
const configDir = resolve(runDir, "config");
mkdirSync(configDir, { recursive: true });

const sdkStateDir = mkdtempSync(resolve(tmpdir(), "excel-bro-online-sdk-"));
const logPath = resolve(runDir, "server.log");
if (existsSync(logPath)) rmSync(logPath);

// ---------------------------------------------------------------------------
// 3) 起本次的服务端
// ---------------------------------------------------------------------------
const port = await freePort();
const baseUrl = `http://127.0.0.1:${port}`;
console.log(`[online] 启动测试服务端（只绑回环）：${baseUrl}`);

const server = spawn(
  pythonBin(),
  [resolve(repoRoot, "server", "tests", "online_entry.py")],
  {
    cwd: repoRoot,
    stdio: "inherit",
    shell: false,
    env: {
      ...process.env,
      EB_ONLINE_PORT: String(port),
      EB_ONLINE_LOG: logPath,
      EB_ONLINE_SCENARIO: resolve(repoRoot, "config", "online_link_scenario.json"),
      EXCEL_BRO_CAPABILITIES_PATH: capabilitiesPath,
      EXCEL_BRO_CONFIG_DIR: configDir,
      AI_BASE_URL: "http://127.0.0.1:1",
      AI_MODEL: "online-test-model",
      AI_API_KEY: "online-test-key",
      ANTHROPIC_API_KEY: "",
      CLAUDE_AGENT_SDK_MODEL: "",
      CLAUDE_AGENT_SDK_STATE_DIR: sdkStateDir
    }
  }
);

const bye = (code) => {
  if (!server.killed && server.exitCode === null) {
    try {
      if (isWin && server.pid) {
        // The venv launcher can own another Python process. Stop only the tree
        // rooted at this run's child, never all Python/Excel processes.
        const stopped = spawnSync('taskkill', ['/PID', String(server.pid), '/T', '/F'],
          { shell: false, windowsHide: true, stdio: 'ignore' });
        if (stopped.status !== 0) {
          console.error('[online] 本次测试服务端清理未确认。');
          code = code || 2;
        }
      } else {
        server.kill();
      }
    } catch {
      // 进程已退出。
    }
  }
  if (dirname(resolve(sdkStateDir)) === resolve(tmpdir())) {
    rmSync(sdkStateDir, { recursive: true, force: true });
  }
  process.exit(code);
};

const ready = await waitForServer(baseUrl);
if (!ready) {
  console.error(`在线连续链路未完成：测试服务端未就绪（${baseUrl}）。日志：${logPath}`);
  bye(2);
}
console.log(`[online] 服务端就绪：${baseUrl}/health`);

// ---------------------------------------------------------------------------
// 4) 两端在线用例（同一份场景 / 同一套期望）
// ---------------------------------------------------------------------------
const vitest = resolve(repoRoot, "node_modules", ".bin", isWin ? "vitest.cmd" : "vitest");
const onlineEnv = {
  ...process.env,
  EB_ONLINE: "1",
  VITE_API_BASE_URL: baseUrl,
  EXCEL_BRO_CAPABILITIES_PATH: capabilitiesPath,
  EXCEL_BRO_CONFIG_DIR: configDir
};

function runSuite(label, workspace, file) {
  console.log(`[online] ${label}：真实面板发起任务 → … → 撤销核对…`);
  const result = run(vitest, ["run", file], {
    cwd: resolve(repoRoot, "apps", workspace),
    env: onlineEnv
  });
  if (result.status !== 0) {
    console.error(
      `在线连续链路未完成：${label}用例未通过（退出码 ${result.status}）。\n` +
        `  用例：apps/${workspace}/${file}\n` +
        `  服务端日志：${logPath}`
    );
    return false;
  }
  return true;
}

const excelOk = runSuite("Excel 端", "excel-addin", "src/onlineLink.test.tsx");
const wpsOk = runSuite("WPS 端", "wps-addin", "src/wpsOnlineLink.test.tsx");

if (!excelOk || !wpsOk) {
  const broken = [!excelOk ? "Excel" : null, !wpsOk ? "WPS" : null].filter(Boolean).join(" / ");
  console.error(`在线连续链路未完成：${broken} 端未通过。服务端日志：${logPath}`);
  bye(1);
}

console.log(
  "双端在线连续集成完成：Excel 与 WPS 两端都在真实服务端 + 真实 Docker 隔离执行器上" +
    `跑通了同一份场景（config/online_link_scenario.json）。\n` +
    `[online] 服务端日志（含 taskId / operationId / 结果引用）：${logPath}\n` +
    "[online] 说明：本轮允许固定模型工具调用序列与有状态宿主替身，因此" +
    "「真实模型自主规划」与「真实 Excel / WPS 客户端」两项**未**通过本命令证明。"
);
bye(0);
