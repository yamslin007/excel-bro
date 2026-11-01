/**
 * WPS 开发端口/登记的**纯逻辑**（可被启动器和守卫脚本共同引用）。
 *
 * 为什么需要这一层：`wpsjs debug -p <端口>` 并不保证用你给的端口。
 * 它先用 `portfinder` 从 basePort 往上找"能 bind"的端口（EACCES/EADDRINUSE 都算
 * 不可用，于是 Windows 保留端口范围会被**静默跳过**），再把这个结果写进 WPS 的
 * publish.xml，然后把 `package.json` 的 `scripts.dev` 改写成 `vite --port <结果>`
 * 并 spawn 它。三处端口虽然此刻一致，但端口是"扫"出来的：换一台机器、重启后保留
 * 范围变化，跳到的端口就不在后端来源白名单里，任务窗格只剩 `Failed to fetch`。
 *
 * 这里改成"声明式"：端口候选只有 `config/capabilities.json` 一处
 * （`devServers.wpsPortCandidates`），后端白名单、Vite 监听、publish.xml 登记全部由
 * 它派生；本模块只负责**按声明顺序**挑第一个真正能绑定的端口，且绝不访问列表外的
 * 端口——挑不到就报错，不漂移。
 */
import net from "node:net";

export const LOOPBACK_HOST = "127.0.0.1";
export const MIN_PORT = 1024;
export const MAX_PORT = 65535;

export class DevConfigError extends Error {
  constructor(message) {
    super(message);
    this.name = "DevConfigError";
  }
}

export class DevPortUnavailableError extends Error {
  constructor(message, attempts) {
    super(message);
    this.name = "DevPortUnavailableError";
    /** 每个候选端口的实测结果：`[{port, code}]`，code 为 EACCES/EADDRINUSE/... */
    this.attempts = attempts;
  }
}

function assertPort(value, label) {
  if (typeof value !== "number" || !Number.isInteger(value)) {
    throw new DevConfigError(`${label} 必须是整数端口`);
  }
  if (value < MIN_PORT || value > MAX_PORT) {
    throw new DevConfigError(`${label} 必须落在 ${MIN_PORT}-${MAX_PORT} 之间`);
  }
  return value;
}

/**
 * 从 capabilities 解析开发端口声明，并拒绝一切"会打架"的配置。
 *
 * 只校验结构，不猜测机器状态：哪个候选真正能绑，由 `selectDevPort` 实测。
 */
export function readDevServerConfig(capabilities) {
  const section = capabilities?.devServers;
  if (!section || typeof section !== "object") {
    throw new DevConfigError("config/capabilities.json 缺少 devServers 配置段");
  }
  const excelPort = assertPort(section.excelPort, "devServers.excelPort");
  const candidates = section.wpsPortCandidates;
  if (!Array.isArray(candidates) || candidates.length === 0) {
    throw new DevConfigError("devServers.wpsPortCandidates 必须是非空端口数组");
  }
  const ports = candidates.map((port, index) =>
    assertPort(port, `devServers.wpsPortCandidates[${index}]`)
  );
  if (new Set(ports).size !== ports.length) {
    throw new DevConfigError("devServers.wpsPortCandidates 不允许出现重复端口");
  }
  if (ports.includes(excelPort)) {
    // 两端共用一个端口 = 后启动的一方必然失败或被迫漂移，等于回归。
    throw new DevConfigError(
      "devServers.wpsPortCandidates 不能包含 devServers.excelPort（会破坏 Excel 开发页）"
    );
  }
  return { excelPort, wpsPortCandidates: ports };
}

/**
 * 把"外部指定端口"（`--port` / `EB_WPS_DEV_PORT`）收窄到候选集合内。
 *
 * 后端白名单只认候选集合，所以列表外的端口一定会被 403 打回；这里提前失败，
 * 而不是让用户对着 `Failed to fetch` 猜。
 */
export function resolveRequestedPort({ requested, candidates, source = "EB_WPS_DEV_PORT" }) {
  if (requested === undefined || requested === null || requested === "") return null;
  const port = typeof requested === "number" ? requested : Number(String(requested).trim());
  if (!Number.isInteger(port)) {
    throw new DevConfigError(`${source} 必须是整数端口，收到 "${requested}"`);
  }
  if (!candidates.includes(port)) {
    throw new DevConfigError(
      `${source}=${port} 不在 devServers.wpsPortCandidates（${candidates.join("、")}）里。\n` +
        "后端只允许该列表内的来源，用列表外端口必然被拒。请先把该端口加入\n" +
        "config/capabilities.json → devServers.wpsPortCandidates，重启本地服务后重试。"
    );
  }
  return port;
}

export function originForPort(port) {
  return `http://${LOOPBACK_HOST}:${port}`;
}

/**
 * 实测"能不能在这个回环地址上 bind"。
 *
 * 只有 EADDRINUSE / EACCES 是**预期内**的"不可用"；其他错误（权限、协议栈异常）
 * 一律向上抛：把未知错误当成"端口不可用"会让我们悄悄换端口，正是要消灭的行为。
 */
export function defaultTryBind(port, host = LOOPBACK_HOST) {
  return new Promise((resolve, reject) => {
    const server = net.createServer();
    server.once("error", (error) => {
      if (error.code === "EADDRINUSE" || error.code === "EACCES") {
        resolve({ ok: false, code: error.code });
        return;
      }
      reject(error);
    });
    server.listen({ port, host, exclusive: true }, () => {
      server.close(() => resolve({ ok: true, code: null }));
    });
  });
}

/**
 * 按声明顺序挑第一个能绑定的端口。
 *
 * "不漂移"在这里是硬约束：只遍历传入的候选，返回的端口必然属于候选集合；
 * 一个都不可用时抛 `DevPortUnavailableError`（附每次尝试的原因），不自行扩范围。
 */
export async function selectDevPort({
  candidates,
  host = LOOPBACK_HOST,
  tryBind = defaultTryBind
}) {
  const attempts = [];
  for (const port of candidates) {
    const result = await tryBind(port, host);
    if (result.ok) {
      return { port, host, attempts };
    }
    attempts.push({ port, code: result.code });
  }
  throw new DevPortUnavailableError(
    `候选端口全部无法在 ${host} 上绑定：${attempts
      .map(({ port, code }) => `${port}(${code})`)
      .join("、")}`,
    attempts
  );
}

const CODE_HINTS = {
  EACCES:
    "EACCES = 该端口落在 Windows 的 TCP 排除（保留）范围内，本机任何进程都不能监听；" +
    "它常由 Hyper-V / WinNAT / Docker 动态保留，不是项目缺陷。",
  EADDRINUSE: "EADDRINUSE = 该端口已被别的进程监听（可能是上一次没退干净的开发服务器）。"
};

export function describePortFailure({ attempts, host = LOOPBACK_HOST, candidates }) {
  const lines = [];
  lines.push(`WPS 开发端口不可用：${host} 上的候选端口全部无法监听。`);
  for (const { port, code } of attempts) {
    lines.push(`  - ${port}：${code}${CODE_HINTS[code] ? `（${CODE_HINTS[code]}）` : ""}`);
  }
  lines.push("");
  lines.push("排查（复制即可）：");
  lines.push("  netsh int ipv4 show excludedportrange protocol=tcp");
  lines.push("  netstat -ano | findstr :<端口>");
  lines.push("");
  lines.push("处理方式：");
  lines.push("  1) 端口被别的进程占用：结束那个进程（多半是残留的 vite/node）后重试；");
  lines.push("  2) 端口在保留范围内：在 config/capabilities.json 的 devServers.wpsPortCandidates");
  lines.push("     末尾加一个不在保留范围内的回环端口（例如 3989），重启本地服务（白名单读同一份");
  lines.push("     配置）后重新执行 npm run dev:wps；");
  lines.push(`  3) 当前候选：${(candidates || attempts.map((a) => a.port)).join("、")}。`);
  lines.push("");
  lines.push("不会做的事：不改 node_modules、不删 Windows 保留端口、不关系统服务、");
  lines.push("不自动换到列表外的端口（那会让任务窗格被后端来源校验拒绝）。");
  return lines.join("\n");
}

function escapeAttribute(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/**
 * WPS 开发加载项的登记元素（与 `wpsjs` 写 publish.xml 的形状一致）。
 *
 * 字段含义来自 WPS 的 publish 调试机制：`enable_dev` 走开发端口、`url` 指向本地
 * 服务根路径。这里只生成这一条，其余插件条目原样保留。
 */
export function buildPluginEntry({ name, type, port }) {
  const url = `${originForPort(port)}/`;
  return (
    `<jspluginonline name="${escapeAttribute(name)}" type="${escapeAttribute(type)}" ` +
    `url="${escapeAttribute(url)}" debug="" enable="enable_dev" install="null"/>`
  );
}

export function registrationUrl(port) {
  return `${originForPort(port)}/`;
}

function pluginEntryPattern(name) {
  return new RegExp(
    `<jspluginonline\\b[^>]*\\bname="${escapeAttribute(name).replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"[^>]*?/>`
  );
}

export function readRegisteredUrl(xml, name) {
  const match = pluginEntryPattern(name).exec(xml || "");
  if (!match) return null;
  const url = /\burl="([^"]*)"/.exec(match[0]);
  return url ? url[1] : null;
}

/**
 * 只更新"本加载项"这一条登记，其他插件条目逐字保留。
 *
 * 定位失败（同一 name 的元素不是自闭合、或文件里没有 `</jsplugins>`）时**抛错**，
 * 不做"覆盖整份文件"的兜底：publish.xml 里还有别的加载项，猜错就毁掉别人的登记。
 */
export function upsertPluginRegistration(xml, entry, { name } = {}) {
  const pluginName = name ?? /name="([^"]*)"/.exec(entry)?.[1];
  if (!pluginName) {
    throw new DevConfigError("无法确定要登记的加载项名称");
  }
  const original = (xml || "").trim();
  const eol = original.includes("\r\n") ? "\r\n" : "\n";
  const base =
    original ||
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${eol}<jsplugins>${eol}</jsplugins>${eol}`;

  const pattern = pluginEntryPattern(pluginName);
  const match = pattern.exec(base);
  let next;
  if (match) {
    next = base.slice(0, match.index) + entry + base.slice(match.index + match[0].length);
  } else {
    if (new RegExp(`<jspluginonline\\b[^>]*\\bname="${pluginName}"`).test(base)) {
      throw new DevConfigError(
        `publish.xml 里 ${pluginName} 的登记不是自闭合元素，无法安全改写；请手动修正该条目后重试`
      );
    }
    const closing = base.lastIndexOf("</jsplugins>");
    if (closing < 0) {
      throw new DevConfigError("publish.xml 缺少 </jsplugins>，无法安全插入登记条目");
    }
    const before = base.slice(0, closing).replace(/\s*$/, "");
    const after = base.slice(closing);
    next = `${before}${eol}  ${entry}${eol}${after}`;
  }

  // 读回校验：写盘前先证明"这份 XML 被解析出来时，本加载项的 url 就是刚生成的那个"，
  // 否则后面"登记地址 == 监听地址"的结论就没有依据。
  const expectedUrl = /\burl="([^"]*)"/.exec(entry)?.[1];
  if (!expectedUrl || readRegisteredUrl(next, pluginName) !== expectedUrl) {
    throw new DevConfigError("改写后的登记地址与预期不一致，已中止（未写盘）");
  }
  return { xml: next, changed: next !== base, url: expectedUrl };
}
