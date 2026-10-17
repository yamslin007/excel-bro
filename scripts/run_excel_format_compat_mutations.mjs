/**
 * 反向变异驱动：证明 Excel 2024（批量许可永久版，ExcelApi ≤ 1.17）格式兼容路径的**新增守卫**，
 * 以及第 19 轮「撤销按动作实际影响拆分恢复」的**范围/属性双层收窄**真有判别力。
 * 变异清单与判别点见 `scripts/mutate_excel_format_compat.mjs` 头部。
 *
 * 四条判据缺一不算过：
 * 1. 未变异时基线 GREEN（退出码 0 + 有效测试摘要 + 用例数达标）；
 * 2. 每个变异 apply 后其目标用例必须变红（退出码非 0 且**有测试摘要**）；
 * 3. restore 后**被改的那个文件**逐字节一致（sha256 比对）；
 * 4. 全部还原后复跑全绿，且五个文件最终都与基线一致。
 *
 * 起手另有前置状态闸（`check`）：文件停在上一次被中断的变异态时当场退出。
 *
 * 用法：node scripts/run_excel_format_compat_mutations.mjs
 */
import { spawnSync } from "node:child_process";
import { openSync, closeSync, readFileSync, rmSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");
const addinDir = resolve(repoRoot, "apps", "excel-addin");
const scratch = resolve(repoRoot, "outputs", "excel-format-compat-mutation.log");
const vitest = resolve(repoRoot, "node_modules", "vitest", "vitest.mjs");

function vitestSpec(name, file, minTests) {
  return {
    name,
    minTests,
    command: process.execPath,
    args: [vitest, "run", "--config", "vite.config.ts", file],
    cwd: addinDir
  };
}

const COMPAT_SPEC = vitestSpec("vitest:兼容路径闭环", "src/excelFormatCompatPath.test.ts", 13);
const MIXED_SPEC = vitestSpec("vitest:混合格式撤销", "src/mixedFormatUndo.test.ts", 21);
const CAPABILITY_SPEC = vitestSpec("vitest:格式能力判定", "src/excelFormatCapability.test.ts", 20);
const COPY_SPEC = vitestSpec("vitest:格式副本审阅", "src/excelFormatCopyReview.test.ts", 5);
const MANUAL_SPEC = vitestSpec("vitest:兼容副本投影", "src/manualReviewFile.test.ts", 10);
const PANEL_SPEC = vitestSpec(
  "vitest:预览阶段能力提示",
  "src/components/ExcelNativeReviewActions.test.tsx",
  19
);
const UNDO_SCOPE_SPEC = vitestSpec(
  "vitest:撤销按范围拆分恢复",
  "src/excelUndoScope.test.ts",
  13
);

const MUTATIONS = [
  { id: "C1", spec: COMPAT_SPEC, note: "需要 clear 时不再清空区域直接格式" },
  { id: "C2", spec: COMPAT_SPEC, note: "条件格式不再拒绝" },
  { id: "C3", spec: COMPAT_SPEC, note: "不再记录继承态属性" },
  { id: "C4", spec: COMPAT_SPEC, note: "预检不再拦缺 1.9 的宿主" },
  { id: "C5", spec: UNDO_SCOPE_SPEC, note: "采集门控失效：行高动作也逐格采格式" },
  { id: "C6", spec: UNDO_SCOPE_SPEC, note: "恢复门控失效：行高动作也走格式恢复" },
  { id: "C7", spec: UNDO_SCOPE_SPEC, note: "clear 后不并回非本动作属性的当前值" },
  { id: "C8", spec: CAPABILITY_SPEC, note: "1.17 判成 unavailable" },
  { id: "C9", spec: CAPABILITY_SPEC, note: "未直接设置被当成匹配" },
  { id: "C10", spec: COPY_SPEC, note: "副本格式被改过不再被发现" },
  { id: "C11", spec: MANUAL_SPEC, note: "格式改动不再标记" },
  { id: "C12", spec: PANEL_SPEC, note: "预览阶段不再提示能力缺失" },
  { id: "C13", spec: UNDO_SCOPE_SPEC, note: "恢复不按属性子集收窄（用全集）" },
  { id: "C14", spec: UNDO_SCOPE_SPEC, note: "冲突文案不指名具体属性" },
  { id: "C15", spec: UNDO_SCOPE_SPEC, note: "采集清单漏掉字体删除线" },
  { id: "C16", spec: UNDO_SCOPE_SPEC, note: "命名样式不再写回（被打回 Normal）" },
  { id: "C17", spec: UNDO_SCOPE_SPEC, note: "补回的当前值用一次性快照（同区域多动作撤销残留）" },
  { id: "C18", spec: COMPAT_SPEC, note: "区域含数据验证不再拒绝（clear(Formats) 会清掉验证）" }
];

const WATCHED = [
  resolve(addinDir, "src", "excel.ts"),
  resolve(addinDir, "src", "excelFormatCapability.ts"),
  resolve(addinDir, "src", "excelNativeComparison.ts"),
  resolve(addinDir, "src", "manualReviewFile.ts"),
  resolve(addinDir, "src", "components", "ExcelNativeReviewActions.tsx")
];

/**
 * 输出走**文件描述符**而不是管道：管道会被拦，而判据又必须读到真实摘要行。
 * vitest 的摘要有 ANSI 色码，解析前先剥掉。
 */
function runSpec(spec) {
  const fd = openSync(scratch, "w");
  const result = spawnSync(spec.command, spec.args, {
    cwd: spec.cwd,
    stdio: ["ignore", fd, fd]
  });
  closeSync(fd);
  const raw = readFileSync(scratch, "utf8").replace(/\x1b\[[0-9;]*m/g, "");
  const summaryLine =
    raw.split(/\r?\n/).find((line) => /Tests\s+\d+ (passed|failed)/.test(line)) ??
    raw.split(/\r?\n/).find((line) => /^\d+ (passed|failed)/.test(line)) ??
    raw
      .split(/\r?\n/)
      .reverse()
      .find((line) => /\b\d+ passed\b|\b\d+ failed\b/.test(line)) ??
    "";
  const passed = Number(/(\d+) passed/.exec(summaryLine)?.[1] ?? 0);
  return { status: result.status, summary: summaryLine.trim(), passed };
}

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

function invokeMutation(id, restore) {
  const args = [
    resolve(here, "mutate_excel_format_compat.mjs"),
    id,
    ...(restore ? ["restore"] : [])
  ];
  const fd = openSync(scratch, "w");
  const result = spawnSync(process.execPath, args, {
    cwd: repoRoot,
    stdio: ["ignore", fd, fd]
  });
  closeSync(fd);
  return { status: result.status, out: readFileSync(scratch, "utf8").trim() };
}

const baselineShas = new Map(WATCHED.map((path) => [path, sha256(path)]));
for (const [path, sha] of baselineShas) {
  console.log(`[mutate] 基线 ${path} sha256=${sha}`);
}

// 前置状态闸：停在上一次被中断的变异态时当场退出，不白跑二十几轮测试。
const preflight = invokeMutation("check");
if (preflight.status !== 0) {
  console.error(`[mutate] 前置状态闸未通过：\n${preflight.out}`);
  process.exit(1);
}
console.log(
  `[mutate] 前置状态闸通过：${MUTATIONS.length} 个 SEED 全部在位`
);
rmSync(scratch, { force: true });

const SPECS = [
  COMPAT_SPEC,
  MIXED_SPEC,
  CAPABILITY_SPEC,
  COPY_SPEC,
  MANUAL_SPEC,
  PANEL_SPEC,
  UNDO_SCOPE_SPEC
];
console.log("[mutate] 基线（未变异）应为 GREEN…");
for (const spec of SPECS) {
  const result = runSpec(spec);
  const ok = result.status === 0 && result.summary && result.passed >= spec.minTests;
  console.log(
    `[mutate]   基线 ${spec.name}: status=${result.status} 摘要="${result.summary}" ${
      ok ? "OK" : "不合格"
    }`
  );
  if (!ok) {
    console.error(`[mutate] 基线就不绿或没跑到预期用例数（要求 >= ${spec.minTests}），中止。`);
    process.exit(1);
  }
}

const failures = [];
for (const { id, spec, note } of MUTATIONS) {
  const applied = invokeMutation(id, false);
  if (applied.status !== 0) {
    failures.push(`${id} apply 失败：${applied.out}`);
    console.error(`[mutate] ${id} apply 失败：${applied.out}`);
    continue;
  }
  const red = runSpec(spec);
  // 判据 2 要求"变红"：退出码非 0 **且**有有效测试摘要 —— 编译/收集期直接崩掉
  // 不算判别力，那只能证明我们把代码写坏了。
  const wentRed = red.status !== 0 && Boolean(red.summary);
  console.log(
    `[mutate] ${id}（${note}）→ ${spec.name}: status=${red.status} 摘要="${red.summary}" ${
      wentRed ? "RED(合格)" : "未变红(不合格)"
    }`
  );
  if (!wentRed) {
    failures.push(
      `${id}（${note}）未让 ${spec.name} 变红（status=${red.status}, 摘要=${red.summary || "无"}）`
    );
  }
  const restored = invokeMutation(id, true);
  if (restored.status !== 0) {
    failures.push(`${id} restore 失败：${restored.out}`);
    console.error(`[mutate] ${id} restore 失败：${restored.out}`);
  }
}

// 判据 3：每个被改过的文件都必须逐字节回到基线。
for (const [path, sha] of baselineShas) {
  const now = sha256(path);
  if (now !== sha) {
    failures.push(`${path} 还原后与基线不一致：${now} != ${sha}`);
    console.error(`[mutate] 还原不一致 ${path}`);
  }
}
if (!failures.length) console.log("[mutate] 五个文件全部逐字节还原");

console.log("[mutate] 还原后复跑全绿…");
for (const spec of SPECS) {
  const result = runSpec(spec);
  const ok = result.status === 0 && result.summary && result.passed >= spec.minTests;
  console.log(
    `[mutate]   复跑 ${spec.name}: status=${result.status} 摘要="${result.summary}" ${
      ok ? "OK" : "不合格"
    }`
  );
  if (!ok) failures.push(`还原后复跑 ${spec.name} 不为绿（status=${result.status}）`);
}

rmSync(scratch, { force: true });

if (failures.length) {
  console.error("\n[mutate] 未通过：");
  for (const line of failures) console.error(`  - ${line}`);
  process.exit(1);
}
console.log(
  `\n[mutate] 全部通过：${MUTATIONS.length} 个变异各自让目标用例变红，五个文件逐字节还原，复跑全绿。`
);
