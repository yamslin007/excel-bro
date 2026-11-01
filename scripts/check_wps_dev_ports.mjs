#!/usr/bin/env node
/**
 * WPS 开发端口的回归守卫（`npm run test:wps` 会先跑它，也可单独执行）。
 *
 * 钉住四条曾真实踩过的事实：
 *   1. 端口被 **Windows 保留**（EACCES）或被别的进程占用（EADDRINUSE）时，
 *      只在**声明过的候选**里选下一个，选不到就报错，绝不漂到列表外的端口；
 *   2. 登记地址（publish.xml）、监听地址、后端来源三者同源同值；
 *   3. 登记改写只动本加载项那一条，其他插件条目逐字保留；定位不确定时 fail-closed；
 *   4. 端口不再写进 `package.json`（wpsjs 曾把 `scripts.dev` 改成 `vite --port <扫出来的端口>`），
 *      后端也不再自带第二份端口清单。
 */
import assert from "node:assert/strict";
import fs from "node:fs";
import net from "node:net";
import path from "node:path";
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
const CAPABILITIES = JSON.parse(
  fs.readFileSync(path.join(REPO_ROOT, "config", "capabilities.json"), "utf8")
);
const WPS_PACKAGE = JSON.parse(
  fs.readFileSync(path.join(REPO_ROOT, "apps", "wps-addin", "package.json"), "utf8")
);

let passed = 0;
function check(name, fn) {
  fn();
  passed += 1;
  console.log(`  ok  ${name}`);
}

async function checkAsync(name, fn) {
  await fn();
  passed += 1;
  console.log(`  ok  ${name}`);
}

function listenOnEphemeral() {
  return new Promise((resolve, reject) => {
    const server = net.createServer();
    server.once("error", reject);
    server.listen({ port: 0, host: "127.0.0.1" }, () => resolve(server));
  });
}

/** 拿一个"刚刚释放、当前应可用"的回环端口。 */
async function takeFreePort() {
  const server = await listenOnEphemeral();
  const { port } = server.address();
  await new Promise((resolve) => server.close(resolve));
  return port;
}

const config = readDevServerConfig(CAPABILITIES);
const candidates = config.wpsPortCandidates;

console.log("WPS 开发端口守卫");

check("端口候选来自 config/capabilities.json（唯一来源），且不含 Excel 端口", () => {
  assert.ok(candidates.length > 0);
  assert.equal(new Set(candidates).size, candidates.length);
  assert.ok(!candidates.includes(config.excelPort));
  assert.equal(config.excelPort, 3000);
});

check("配置层面就拒绝 WPS/Excel 共用端口与重复端口", () => {
  assert.throws(
    () => readDevServerConfig({ devServers: { excelPort: 3000, wpsPortCandidates: [3000] } }),
    DevConfigError
  );
  assert.throws(
    () => readDevServerConfig({ devServers: { excelPort: 3000, wpsPortCandidates: [3987, 3987] } }),
    DevConfigError
  );
  assert.throws(() => readDevServerConfig({ devServers: { excelPort: 3000 } }), DevConfigError);
});

await checkAsync("系统保留端口（EACCES）被跳过：只在候选内按声明顺序取下一个", async () => {
  const [first, second] = [candidates[0], candidates[1] ?? candidates[0] + 1];
  const calls = [];
  const result = await selectDevPort({
    candidates: [first, second],
    tryBind: async (port) => {
      calls.push(port);
      return port === first ? { ok: false, code: "EACCES" } : { ok: true, code: null };
    }
  });
  assert.equal(result.port, second);
  assert.deepEqual(calls, [first, second]);
  assert.deepEqual(result.attempts, [{ port: first, code: "EACCES" }]);
});

await checkAsync("候选全被保留：报错而不是换到列表外的端口", async () => {
  const calls = [];
  await assert.rejects(
    () =>
      selectDevPort({
        candidates: [candidates[0]],
        tryBind: async (port) => {
          calls.push(port);
          return { ok: false, code: "EACCES" };
        }
      }),
    (error) => {
      assert.ok(error instanceof DevPortUnavailableError);
      assert.deepEqual(
        error.attempts.map((attempt) => attempt.port),
        [candidates[0]]
      );
      const message = describePortFailure({
        attempts: error.attempts,
        candidates: [candidates[0]]
      });
      assert.match(message, /netsh int ipv4 show excludedportrange protocol=tcp/);
      assert.match(message, /EACCES/);
      assert.match(message, /不自动换到列表外的端口/);
      assert.match(message, /config\/capabilities\.json/);
      // 报错文案里出现的端口只能是"候选本身"或文案里的示例端口，不许夹带一个替代端口
      const allowedMentions = new Set([candidates[0], 3989]);
      const mentioned = [...message.matchAll(/\b(\d{4,5})\b/g)].map((match) => Number(match[1]));
      for (const port of mentioned) {
        assert.ok(allowedMentions.has(port), `报错里出现了未声明的端口 ${port}`);
      }
      return true;
    }
  );
  // 只试了声明的那一个：没有"再往上扫一扫"
  assert.deepEqual(calls, [candidates[0]]);
});

await checkAsync("真实被占用（EADDRINUSE）时选中下一个可用候选，且端口不漂移", async () => {
  const occupied = await listenOnEphemeral();
  const { port: occupiedPort } = occupied.address();
  const freePort = await takeFreePort();
  try {
    const result = await selectDevPort({
      candidates: [occupiedPort, freePort],
      tryBind: defaultTryBind
    });
    assert.equal(result.port, freePort);
    assert.deepEqual(result.attempts, [{ port: occupiedPort, code: "EADDRINUSE" }]);
    assert.ok(result.port !== occupiedPort);
  } finally {
    await new Promise((resolve) => occupied.close(resolve));
  }
});

await checkAsync("真实空闲端口可直接选中", async () => {
  const freePort = await takeFreePort();
  const result = await selectDevPort({ candidates: [freePort], tryBind: defaultTryBind });
  assert.equal(result.port, freePort);
  assert.deepEqual(result.attempts, []);
});

check("--port / EB_WPS_DEV_PORT 只能是候选内端口（列表外直接拒绝）", () => {
  assert.equal(resolveRequestedPort({ requested: candidates[0], candidates }), candidates[0]);
  assert.equal(resolveRequestedPort({ requested: undefined, candidates }), null);
  assert.throws(
    () => resolveRequestedPort({ requested: "65533", candidates }),
    (error) => error instanceof DevConfigError && /不在 devServers\.wpsPortCandidates/.test(error.message)
  );
});

check("登记地址 == 监听地址（同一端口派生出同一个来源）", () => {
  const port = candidates[candidates.length - 1];
  const entry = buildPluginEntry({ name: "excel-bro-wps", type: "et", port });
  const xml = upsertPluginRegistration("", entry, { name: "excel-bro-wps" }).xml;
  const registered = readRegisteredUrl(xml, "excel-bro-wps");
  assert.equal(registered, registrationUrl(port));
  assert.equal(registered, `http://127.0.0.1:${port}/`);
  assert.equal(originForPort(port), `http://127.0.0.1:${port}`);
  assert.equal(registrationUrl(port).replace(/\/$/, ""), originForPort(port));
});

check("登记改写只动本加载项，其他插件逐字保留；重复执行不再改动", () => {
  const fixture = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    "<jsplugins>",
    '  <jspluginonline name="other-addon" type="wps" url="http://127.0.0.1:4123/" debug="" enable="enable_dev" install="null"/>',
    `  <jspluginonline name="excel-bro-wps" type="et" url="${registrationUrl(candidates[0])}" debug="" enable="enable_dev" install="null"/>`,
    "</jsplugins>",
    ""
  ].join("\n");
  const port = candidates[candidates.length - 1];
  const entry = buildPluginEntry({ name: "excel-bro-wps", type: "et", port });
  const first = upsertPluginRegistration(fixture, entry, { name: "excel-bro-wps" });
  assert.equal(first.changed, true);
  assert.ok(
    first.xml.includes(
      '  <jspluginonline name="other-addon" type="wps" url="http://127.0.0.1:4123/" debug="" enable="enable_dev" install="null"/>'
    ),
    "其他插件的登记条目被改动了"
  );
  assert.equal(readRegisteredUrl(first.xml, "excel-bro-wps"), registrationUrl(port));
  assert.equal(readRegisteredUrl(first.xml, "other-addon"), "http://127.0.0.1:4123/");

  const second = upsertPluginRegistration(first.xml, entry, { name: "excel-bro-wps" });
  assert.equal(second.changed, false);
  assert.equal(second.xml, first.xml);
});

check("publish.xml 结构不可安全改写时 fail-closed（不覆盖整份文件）", () => {
  const notSelfClosing = [
    "<jsplugins>",
    '<jspluginonline name="excel-bro-wps" type="et" url="http://127.0.0.1:3889/">',
    "</jspluginonline>",
    "</jsplugins>"
  ].join("\n");
  assert.throws(
    () =>
      upsertPluginRegistration(
        notSelfClosing,
        buildPluginEntry({ name: "excel-bro-wps", type: "et", port: candidates[0] }),
        { name: "excel-bro-wps" }
      ),
    DevConfigError
  );
  assert.throws(
    () =>
      upsertPluginRegistration(
        "<jsplugins>",
        buildPluginEntry({ name: "excel-bro-wps", type: "et", port: candidates[0] }),
        { name: "excel-bro-wps" }
      ),
    DevConfigError
  );
});

check("启动脚本不再把端口写进 package.json", () => {
  assert.equal(WPS_PACKAGE.scripts.dev, "vite");
  assert.ok(!/--port/.test(WPS_PACKAGE.scripts.dev));
  assert.equal(WPS_PACKAGE.scripts.debug, "node ../../scripts/start_wps_dev.mjs");
});

check("vite.config.ts 只从 EB_WPS_DEV_PORT 取端口，且保持 strictPort", () => {
  const source = fs.readFileSync(path.join(REPO_ROOT, "apps", "wps-addin", "vite.config.ts"), "utf8");
  assert.match(source, /EB_WPS_DEV_PORT/);
  assert.match(source, /devServers|wpsPortCandidates/);
  assert.match(source, /strictPort:\s*true/);
  assert.ok(!/port:\s*\d{3,5}\s*,/.test(source), "vite.config.ts 里仍有硬编码端口");
});

check("后端不再自带第二份端口清单（与启动器读同一份配置）", () => {
  const source = fs.readFileSync(path.join(REPO_ROOT, "server", "app", "main.py"), "utf8");
  assert.match(source, /capability_int_list\("devServers",\s*"wpsPortCandidates"\)/);
  for (const port of candidates) {
    assert.ok(
      !source.includes(`"http://127.0.0.1:${port}"`),
      `server/app/main.py 里仍硬编码了 http://127.0.0.1:${port}`
    );
  }
});

// 本机事实（不参与断言）：Windows 保留范围会随 Hyper-V/WinNAT 变化，故只打印。
const probe = await defaultTryBind(candidates[0]);
console.log(
  `  info 候选 ${candidates[0]} 本机实测：${probe.ok ? "可绑定" : probe.code}` +
    (probe.code === "EACCES" ? "（落在 Windows TCP 排除范围内，启动器会跳到下一个候选）" : "")
);

console.log(`WPS 开发端口守卫通过：${passed} 项`);
