#!/usr/bin/env node
/**
 * WPS 开发加载项的**唯一**启动入口（`npm run dev:wps`）。
 *
 * 它替代 `wpsjs debug -p 3889`，因为那条路径有三个不可控点：
 *   1. `portfinder` 会把 Windows 保留端口（EACCES）当成"不可用"静默跳过，
 *      于是端口漂到 3987/3988 —— 而后端来源白名单里没有它们；
 *   2. `debug_publish.debugVite` 会把结果端口写回 `package.json` 的 `scripts.dev`
 *      （`vite --port <扫出来的端口>`），让 package.json 变成运行时配置；
 *   3. 端口是"扫"出来的，重启后保留范围变化就会再漂一次。
 *
 * 这里改成确定性的四步：
 *   读配置候选 → 实测 bind 挑第一个可用（挑不到就报错，不漂移）
 *   → 按同一端口写 publish.xml 登记（读回校验）
 *   → 以 `EB_WPS_DEV_PORT` 起 Vite（strictPort，绝不换端口）
 *   → 实测"登记地址 / 监听地址 / 后端允许来源"三者一致。
 *
 * 不做的事：不改 node_modules、不改 Windows 保留端口、不关系统服务、不要求管理员。
 * 也不再写 `package.json`。
 */
import fs from "node:fs";
import path from "node:path";
import http from "node:http";
import { spawn, execFile } from "node:child_process";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

import {
  DevConfigError,
  DevPortUnavailableError,
  buildPluginEntry,
  defaultTryBind,
  describePortFailure,
  originForPort,
  readDevServerConfig,
  readRegisteredUrl,
  registrationUrl,
  resolveRequestedPort,
  selectDevPort,
  upsertPluginRegistration
} from "./wpsDevPorts.mjs";

const SCRIPT_DIR = path.dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = path.resolve(SCRIPT_DIR, "..");
const APP_DIR = path.join(REPO_ROOT, "apps", "wps-addin");
const CAPABILITIES_PATH = path.join(REPO_ROOT, "config", "capabilities.json");
const BACKEND_URL = (process.env.EB_BACKEND_URL || "http://127.0.0.1:8765").replace(/\/+$/, "");
const READY_TIMEOUT_MS = 60000;

const USAGE = `用法：npm run dev:wps [-- 选项]

  --port <端口>     指定端口，必须已在 config/capabilities.json 的
                    devServers.wpsPortCandidates 里（也可以用 EB_WPS_DEV_PORT）
  --launch          启动成功后顺带拉起 WPS 表格（默认只打印下一步，不动宿主）
  --no-backend-check
                    跳过"后端是否允许该来源"的实测（后端没起时可用）
  --help            显示本帮助`;

function parseArgs(argv) {
  const options = { port: null, launch: false, backendCheck: true, help: false };
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === "--help" || arg === "-h") {
      options.help = true;
    } else if (arg === "--launch") {
      options.launch = true;
    } else if (arg === "--no-backend-check") {
      options.backendCheck = false;
    } else if (arg === "--port") {
      options.port = argv[index + 1];
      index += 1;
    } else if (arg.startsWith("--port=")) {
      options.port = arg.slice("--port=".length);
    } else {
      throw new DevConfigError(`无法识别的参数：${arg}\n\n${USAGE}`);
    }
  }
  return options;
}

function publishXmlPath() {
  if (process.platform === "win32") {
    return path.resolve(process.env.APPDATA, "kingsoft/wps/jsaddons/publish.xml");
  }
  if (process.platform === "darwin") {
    return path.resolve(
      process.env.HOME,
      "Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml"
    );
  }
  return path.resolve(process.env.HOME, ".local/share/Kingsoft/wps/jsaddons/publish.xml");
}

/** 登记身份取自加载项包本身（name / addonType），与 `wpsjs` 的口径一致。 */
function readAddonIdentity() {
  const manifest = JSON.parse(fs.readFileSync(path.join(APP_DIR, "package.json"), "utf8"));
  if (!manifest.name || !manifest.addonType) {
    throw new DevConfigError("apps/wps-addin/package.json 缺少 name/addonType，无法登记加载项");
  }
  return { name: manifest.name, type: manifest.addonType };
}

function request(url, { method = "GET", headers = {}, timeoutMs = 3000 } = {}) {
  return new Promise((resolve, reject) => {
    const client = http.request(url, { method, headers, timeout: timeoutMs }, (response) => {
      let body = "";
      response.setEncoding("utf8");
      response.on("data", (chunk) => {
        body += chunk;
      });
      response.on("end", () => {
        resolve({ status: response.statusCode ?? 0, headers: response.headers, body });
      });
    });
    client.on("timeout", () => {
      client.destroy(new Error("请求超时"));
    });
    client.on("error", reject);
    client.end();
  });
}

async function waitForDevServer(port, child) {
  const deadline = Date.now() + READY_TIMEOUT_MS;
  let lastError = "";
  while (Date.now() < deadline) {
    if (child.exitCode !== null) {
      throw new Error(
        `Vite 进程已退出（退出码 ${child.exitCode}）。端口 ${port} 在检查后被别的进程占用，` +
          "或 vite 配置拒绝了该端口；请重跑 npm run dev:wps 查看上面的原始报错。"
      );
    }
    try {
      const { status } = await request(`http://127.0.0.1:${port}/index.html`, { timeoutMs: 1000 });
      if (status === 200) return;
      lastError = `HTTP ${status}`;
    } catch (error) {
      lastError = error.message;
    }
    await new Promise((resolve) => setTimeout(resolve, 250));
  }
  throw new Error(
    `等待 http://127.0.0.1:${port}/index.html 就绪超时（${READY_TIMEOUT_MS / 1000}s，最后错误：${lastError}）`
  );
}

/**
 * 实测后端是否真的接受这个来源。
 *
 * 两个**无副作用**探测，正好对应任务窗格失效的两条真实路径：
 *   - `GET /health`：浏览器能否**读到**响应，取决于有没有回
 *     `access-control-allow-origin`；缺失就是 `Failed to fetch` 的直接原因。
 *   - `POST /api/__wps_dev_origin_probe`（不存在的路由）：来源守卫在路由之前，
 *     所以被允许的来源得到 404、未被允许的来源得到 403 `FORBIDDEN_ORIGIN`。
 *     这正是任务窗格发送 POST 时踩的那条闸门，且不会创建任务、不写账本。
 */
async function probeBackend(origin) {
  const health = await request(`${BACKEND_URL}/health`, { headers: { Origin: origin } });
  const allowOrigin = health.headers["access-control-allow-origin"];
  const guarded = await request(`${BACKEND_URL}/api/__wps_dev_origin_probe`, {
    method: "POST",
    headers: { Origin: origin }
  });
  let guardCode = null;
  try {
    guardCode = JSON.parse(guarded.body)?.detail?.code ?? null;
  } catch {
    guardCode = null;
  }
  return {
    healthStatus: health.status,
    allowOrigin,
    corsOk: allowOrigin === origin,
    guardStatus: guarded.status,
    guardOk: guarded.status === 404 && guardCode === null
  };
}

function launchWps() {
  const query = execFile(
    "reg",
    ["query", "HKEY_CLASSES_ROOT\\KET.Sheet.12\\shell\\new\\command", "/ve"],
    { encoding: "utf8" },
    (error, stdout) => {
      if (error) {
        console.warn("[wps-dev] 无法从注册表解析 WPS 表格路径，请手动打开 WPS 表格。");
        return;
      }
      const match = /"([^"]+\.exe)"/i.exec(stdout);
      if (!match) {
        console.warn("[wps-dev] 注册表里没有可用的 WPS 表格可执行文件，请手动打开。");
        return;
      }
      const child = spawn(match[1], [], { detached: true, stdio: "ignore" });
      child.unref();
      console.log(`[wps-dev] 已拉起 ${match[1]}`);
    }
  );
  query.on("error", () => {
    console.warn("[wps-dev] 拉起 WPS 失败，请手动打开 WPS 表格。");
  });
}

async function run() {
  const options = parseArgs(process.argv.slice(2));
  if (options.help) {
    console.log(USAGE);
    return 0;
  }

  const config = readDevServerConfig(
    JSON.parse(fs.readFileSync(CAPABILITIES_PATH, "utf8"))
  );
  const identity = readAddonIdentity();

  const requested = resolveRequestedPort({
    requested: options.port ?? process.env.EB_WPS_DEV_PORT,
    candidates: config.wpsPortCandidates,
    source: options.port ? "--port" : "EB_WPS_DEV_PORT"
  });

  let selection;
  if (requested) {
    selection = { port: requested, host: "127.0.0.1", attempts: [] };
    console.log(`[wps-dev] 使用指定端口 ${requested}（已在候选列表内）。`);
  } else {
    try {
      selection = await selectDevPort({ candidates: config.wpsPortCandidates });
    } catch (error) {
      if (error instanceof DevPortUnavailableError) {
        console.error(
          describePortFailure({ attempts: error.attempts, candidates: config.wpsPortCandidates })
        );
        return 2;
      }
      throw error;
    }
    const skipped = selection.attempts
      .map(({ port, code }) => `${port} ${code === "EACCES" ? "被系统保留" : `被占用(${code})`}`)
      .join("、");
    console.log(
      `[wps-dev] 选定端口 ${selection.port}（候选 ${config.wpsPortCandidates.join("、")}` +
        `${skipped ? `；跳过 ${skipped}` : ""}）。`
    );
  }

  const port = selection.port;
  const origin = originForPort(port);

  // 登记之前先实测一次（`--port` 指定的端口也要过这一关）：端口不可用就尽快失败，
  // 不要把 publish.xml 指到一个根本起不来的地址上。真正的抢占竞态由 Vite 的
  // strictPort 兜住（见下面的 `waitForDevServer`）。
  const bind = await defaultTryBind(port);
  if (!bind.ok) {
    console.error(
      describePortFailure({
        attempts: [{ port, code: bind.code }],
        candidates: config.wpsPortCandidates
      })
    );
    if (requested) {
      console.error(`\n提示：${port} 是你指定的端口；去掉 --port / EB_WPS_DEV_PORT 让脚本按候选顺序挑选。`);
    }
    return 2;
  }

  // 1) 登记：只改本加载项那一条，其他插件原样保留；写回后读回校验。
  const xmlPath = publishXmlPath();
  const previous = fs.existsSync(xmlPath) ? fs.readFileSync(xmlPath, "utf8") : "";
  const entry = buildPluginEntry({ name: identity.name, type: identity.type, port });
  const registration = upsertPluginRegistration(previous, entry, { name: identity.name });
  if (registration.changed) {
    fs.mkdirSync(path.dirname(xmlPath), { recursive: true });
    fs.writeFileSync(xmlPath, registration.xml, "utf8");
  }
  const registeredUrl = readRegisteredUrl(fs.readFileSync(xmlPath, "utf8"), identity.name);
  if (registeredUrl !== registrationUrl(port)) {
    throw new Error(
      `登记读回失败：publish.xml 里 ${identity.name} 的地址是 ${registeredUrl}，预期 ${registrationUrl(port)}`
    );
  }
  console.log(
    `[wps-dev] 登记 ${identity.name}(${identity.type}) → ${registeredUrl}` +
      `${registration.changed ? "（已更新）" : "（无变化）"}。`
  );

  // 2) 起 Vite：端口只经环境变量传入，vite.config.ts 严格端口，不会换端口。
  //    直接 spawn vite 的可执行入口（不经过 `npm run` 的 cmd 外壳）：
  //    这样 Ctrl+C / kill 打到的就是 vite 本身，不会留下一个还在占端口的孤儿进程。
  const appRequire = createRequire(path.join(APP_DIR, "package.json"));
  const viteManifestPath = appRequire.resolve("vite/package.json");
  const viteManifest = JSON.parse(fs.readFileSync(viteManifestPath, "utf8"));
  const viteBin = path.resolve(
    path.dirname(viteManifestPath),
    typeof viteManifest.bin === "string" ? viteManifest.bin : viteManifest.bin.vite
  );
  const child = spawn(process.execPath, [viteBin], {
    cwd: APP_DIR,
    env: { ...process.env, EB_WPS_DEV_PORT: String(port) },
    stdio: "inherit"
  });

  const shutdown = (signal) => {
    if (child.exitCode === null) child.kill(signal);
  };
  process.on("SIGINT", () => shutdown("SIGINT"));
  process.on("SIGTERM", () => shutdown("SIGTERM"));

  try {
    await waitForDevServer(port, child);
  } catch (error) {
    console.error(`[wps-dev] 启动失败：${error.message}`);
    shutdown("SIGTERM");
    return 1;
  }

  // 3) 三处一致性 + 后端来源实测。
  const listening = `http://127.0.0.1:${port}/`;
  const registrationOk = registeredUrl === listening;
  let backend = null;
  if (options.backendCheck) {
    try {
      backend = await probeBackend(origin);
    } catch (error) {
      backend = { error: error.message };
    }
  }

  const line = (ok) => (ok ? "✅" : "❌");
  console.log("");
  console.log("================ WPS 开发链路自检 ================");
  console.log(`端口（配置候选）：${config.wpsPortCandidates.join("、")} → 选中 ${port}`);
  console.log(`登记地址（publish.xml）：${registeredUrl}  ${line(registrationOk)}`);
  console.log(`Vite 监听地址：${listening}  ✅（GET /index.html 200）`);
  console.log(`两者一致：${line(registrationOk)}${registrationOk ? "" : "（端口漂移，请重跑并检查 publish.xml）"}`);
  if (!options.backendCheck) {
    console.log("后端允许来源：未检查（--no-backend-check）");
  } else if (backend?.error) {
    console.log(`后端允许来源：无法连接 ${BACKEND_URL}（${backend.error}）`);
    console.log("  → 先启动后端：npm run dev:server；本机服务起来后重跑本命令只需重开任务窗格。");
  } else {
    console.log(
      `后端允许来源：${line(backend.corsOk && backend.guardOk)} ` +
        `（/health ACAO=${backend.allowOrigin ?? "缺失"}；来源守卫探测 HTTP ${backend.guardStatus}）`
    );
    if (!backend.corsOk || !backend.guardOk) {
      console.log("  → 后端不认识该来源：它还在用改动前的白名单。请**手动**重启本地服务");
      console.log("    （npm run dev:server），再重开任务窗格。本脚本不会停掉你正在跑的后端。");
    }
  }
  console.log("=================================================");
  console.log("");
  if (registration.changed) {
    console.log("登记已变更：请保存打开的文档并**完全退出** WPS 表格后重开（功能区/登记只在重启后生效）。");
  } else {
    console.log("登记未变更：只需刷新/重开任务窗格即可（功能区改动才需要重启 WPS）。");
  }
  console.log(`操作路径：WPS 表格 →「Excel Bro」选项卡 →「打开 Excel Bro」；任务窗格地址应为 ${listening}`);

  if (options.launch) {
    launchWps();
  }

  const code = await new Promise((resolve) => {
    child.on("exit", (exitCode, signal) => resolve(signal ? 0 : exitCode ?? 0));
  });
  return code;
}

run()
  .then((code) => {
    process.exitCode = code;
  })
  .catch((error) => {
    if (error instanceof DevConfigError) {
      console.error(`[wps-dev] ${error.message}`);
      process.exitCode = 2;
      return;
    }
    console.error(`[wps-dev] 启动失败：${error?.stack ?? error}`);
    process.exitCode = 1;
  });
