// 反向变异脚本：WPS 任务窗格顶部「开发诊断导航条」已整条移除（2026-10-16）
//
// 背景：任务窗格顶部原来有一条仅 DEV 渲染的导航条（「主面板」/「宽幅审阅 · 能力探针」/
// 「WPS 格式能力 · 受控探测」三颗按钮）。它在真实宿主里独占一整行，只是开发诊断入口，
// 已按产品要求整条移除；诊断页保留（DEV 下改 hash 仍可进入）。
//
// 用法：node scripts/mutate_wps_dev_nav_removed.mjs [M1|M2] [restore]
// 不带参数：只打印目标文件的基线 sha256。
// M1 = 把那条导航条重新挂回主面板（三颗按钮原样回来）
// M2 = 诊断路由不再按 DEV 门控（非 DEV 也能靠改 hash 打开诊断页）
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";
const ROUTES = resolve(REPO, "apps", "wps-addin", "src", "AppRoutes.tsx");

const TARGETS = {
  M1: {
    file: ROUTES,
    seed: [
      "  // 主面板（含未知 hash / 非 DEV 的诊断地址）都兜底到主面板。",
      "  return <WpsProbePanel />;"
    ].join(LF),
    mutant: [
      "  // 主面板（含未知 hash / 非 DEV 的诊断地址）都兜底到主面板。",
      "  return (",
      "    <>",
      '      <nav className="dev-nav" aria-label="开发诊断入口">',
      '        <button type="button">主面板</button>',
      '        <button type="button">宽幅审阅 · 能力探针</button>',
      '        <button type="button">WPS 格式能力 · 受控探测</button>',
      "      </nav>",
      "      <WpsProbePanel />",
      "    </>",
      "  ); // [变异M1] 把开发诊断导航条挂回来"
    ].join(LF)
  },
  M2: {
    file: ROUTES,
    seed: [
      '  if (isDev() && hash === "/wide-view-poc") return <WideViewPocPanel />;',
      '  if (isDev() && hash === "/dev-format-probe") return <WpsFormatProbePanel />;'
    ].join(LF),
    mutant: [
      "  // [变异M2] 诊断路由不再按 DEV 门控",
      '  if (hash === "/wide-view-poc") return <WideViewPocPanel />;',
      '  if (hash === "/dev-format-probe") return <WpsFormatProbePanel />;'
    ].join(LF)
  }
};

/** 目标文件自带换行（CRLF/LF）都要先归一化，否则 includes/split 会误报 ALREADY-CLEAN。 */
function normalize(text) {
  return text.split(/\r\n/).join(LF);
}
/** 写回必须还原**原文件自己的**换行，否则"逐字节一致"这条判据会自己失败。 */
function denormalize(text, original) {
  const native = original.includes("\r\n") ? "\r\n" : LF;
  return text.split(LF).join(native);
}

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

const which = process.argv[2];
if (!which) {
  for (const [id, target] of Object.entries(TARGETS)) {
    console.log(`BASELINE ${id} ${target.file} sha256=${sha256(target.file)}`);
  }
  process.exit(0);
}
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_wps_dev_nav_removed.mjs [M1|M2] [restore]");
  process.exit(2);
}

const target = TARGETS[which];
const original = readFileSync(target.file, "utf8");

if (process.argv[3] === "restore") {
  const hits = normalize(original).split(target.mutant).length - 1;
  if (hits === 0) {
    console.log("ALREADY-CLEAN");
    process.exit(0);
  }
  writeFileSync(
    target.file,
    denormalize(normalize(original).split(target.mutant).join(target.seed), original),
    "utf8"
  );
  console.log(`restore: OK ${target.file}`);
  process.exit(0);
}

const occurrences = normalize(original).split(target.seed).length - 1;
if (occurrences !== 1) {
  console.log(`SKIP: SEED 出现 ${occurrences} 次（期望 1）于 ${target.file}`);
  process.exit(3);
}

writeFileSync(
  target.file,
  denormalize(normalize(original).replace(target.seed, target.mutant), original),
  "utf8"
);
console.log(`APPLIED ${which} → ${target.file}`);
