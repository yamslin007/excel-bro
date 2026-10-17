// 反向变异脚本：审阅布局「按整个显示器可用工作区 35/35/30」与「只在用户主动重排时才重排/恢复」。
//
// 用法：
//   node scripts/mutate_review_layout_ratio.mjs plan               # 打印每个 SEED 的 cwd/TEST/EXPECT（制表符分隔）
//   node scripts/mutate_review_layout_ratio.mjs apply   <ID>       # 只施加一个 SEED
//   node scripts/mutate_review_layout_ratio.mjs restore <ID>       # 只还原一个 SEED（逐字节 sha256 校验）
//
// 铁律（与仓库其它 mutate_*.mjs 一致）：
//   - 一次只 apply/restore 一个 SEED；先归一化 LF 匹配，写回按**原文件**的 EOL。
//   - 起手前置状态闸：文件已含该 SEED 的变异标记就拒绝 apply。
//   - 装 SIGINT/SIGTERM/SIGHUP + uncaughtException/unhandledRejection 还原钩子，异常退出也要还原。
//   - 还原按 apply 时保存的**原始字节**回写，并核对 sha256 一致（不通过即非零退出）。
//   - 报错必须印文件名。
//   - **测试由 bash 直调**，本脚本不 spawn：本沙箱里 node 起子进程会 EBUSY。
//     每个 SEED 必须复跑 `plan` 给出的 TEST，确认 EXPECT 用例**实跑且变红**（判据见仓库铁律四判据）。

import { createHash } from "node:crypto";
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import path from "node:path";

const REPO = path.resolve(import.meta.dirname, "..");
const STATE = path.join(REPO, ".tmp-review-layout-mutation");

const SEEDS = {
  // S1：把 AI 比例的参照换回**原表窗口内容区宽度**（这正是历史真实缺陷）。
  S1: {
    file: "apps/excel-addin/src/hooks/useReviewPaneRatio.ts",
    seed: "        const { workAreaWidthCssPx } = await measureLocalExcelWindow(identity!, sheet!, nonce!);",
    mutant: "        const { clientWidthCssPx: workAreaWidthCssPx } = await measureLocalExcelWindow(identity!, sheet!, nonce!); // [变异S1]",
    cwd: "apps/excel-addin",
    test: "src/hooks/useReviewPaneRatio.test.ts",
    expect: "derives the ratio from the whole monitor work area, never the original window or the iframe",
  },
  // S2：把 AI 窗格算成**右侧原表窗口**的 30%（用户明确禁止的算法）。
  S2: {
    file: "packages/core/src/reviewPaneSizing.ts",
    seed: [
      "  const wanted = Math.round(workAreaWidth * REVIEW_PANE_LIMITS.reviewAiWidthFraction);",
      "  // AI 窗格：不小于最小 AI 宽，也不吃掉原表窗口的最小表格宽，且不超过原表窗口的一半。",
    ].join("\n"),
    mutant: [
      "  const wanted = Math.round(originalWidth * REVIEW_PANE_LIMITS.reviewAiWidthFraction); // [变异S2]",
      "  // AI 窗格：不小于最小 AI 宽，也不吃掉原表窗口的最小表格宽，且不超过原表窗口的一半。",
    ].join("\n"),
    cwd: "packages/core",
    test: "src/reviewPaneSizing.test.ts",
    expect: "computes the 35/35/30 split from the whole usable work area",
  },
  // S3：measure 回执把工作区宽度写成内容区宽度（两端口径再次分叉）。
  S3: {
    file: "server/app/excel_window_layout.py",
    seed: "            return ExcelReviewPaneMeasurement(clientWidthCssPx=width, workAreaWidthCssPx=work_area)",
    mutant: "            return ExcelReviewPaneMeasurement(clientWidthCssPx=width, workAreaWidthCssPx=width)  # [变异S3]",
    cwd: ".",
    test: "server/tests/test_excel_window_layout.py",
    expect: "test_measure_only_verified_live_pair_no_mutations_or_tickets",
  },
  // S4：外侧比例回到旧的 60/40（AI 按内容区算时会把原表挤到约 24%）。
  S4: {
    file: "server/app/excel_window_layout.py",
    seed: '            fraction = 1.0 - capability_float("reviewPaneSizing", "reviewCopyWidthFraction")',
    mutant: "            fraction = 0.6  # [变异S4]",
    cwd: ".",
    test: "server/tests/test_excel_window_layout.py",
    expect: "test_native_positioning_exact_pair_geometry_and_rollback",
  },
  // S5：退出审阅时无视「用户手动拖过」，强行恢复成我们的临时布局。
  S5: {
    file: "apps/wps-addin/src/useWpsReviewPaneRatio.ts",
    seed: "          if (!manual && getWpsApplication().Hwnd === app.Hwnd) adapter.restore();",
    mutant: "          if (getWpsApplication().Hwnd === app.Hwnd) adapter.restore(); // [变异S5]",
    cwd: "apps/wps-addin",
    test: "src/wpsReviewPaneSizing.test.tsx",
    expect: "never force-restores over a width the user dragged during the review",
  },
  // S6：rearm 不再是依赖 ⇒ 用户主动「重新排列」不再重套默认比例。
  S6: {
    file: "apps/wps-addin/src/useWpsReviewPaneRatio.ts",
    seed: "  }, [review, rearm]);",
    mutant: "  }, [review]); // [变异S6]",
    cwd: "apps/wps-addin",
    test: "src/wpsReviewPaneSizing.test.tsx",
    expect: "re-arms the default pane ratio only on an explicit rearrange, never on its own",
  },
  // S7：没有活跃审阅也把「重新排列」报成功（假成功 / 顺手移动别人窗口）。
  S7: {
    file: "apps/excel-addin/src/hooks/useExcelNativeComparison.ts",
    seed: '      if (!current || current.messageId !== messageId || current.phase !== "review") throw new Error("这张卡片没有活跃审阅副本。");',
    mutant: '      if (!current || current.messageId !== messageId || current.phase !== "review") { ok = true; return; } // [变异S7]',
    cwd: "apps/excel-addin",
    test: "src/hooks/useExcelNativeComparison.rearrange.test.ts",
    expect: "refuses to arrange when the card has no live review, without contacting the bridge",
  },
  // S8：去掉原生调用**之后**的身份复检 ⇒ 中途换簿被报成已排列。
  S8: {
    file: "apps/wps-addin/src/wpsNativeComparison.ts",
    seed: [
      '    if (application.Windows.CompareSideBySideWith(caption) !== true) return "unavailable";',
      '    if (!identityHolds()) return "unavailable";',
    ].join("\n"),
    mutant: '    if (application.Windows.CompareSideBySideWith(caption) !== true) return "unavailable"; // [变异S8] 去掉调用后的身份复检',
    cwd: "apps/wps-addin",
    test: "src/wpsNativeComparison.test.ts",
    expect: "checks identity after the native call too, so a mid-call swap is not reported as arranged",
  },
};

const sha256 = (buffer) => createHash("sha256").update(buffer).digest("hex");
const eolOf = (text) => (text.includes("\r\n") ? "\r\n" : "\n");
const toLf = (text) => text.split("\r\n").join("\n");
const fromLf = (text, eol) => (eol === "\r\n" ? text.split("\n").join("\r\n") : text);
const statePath = (id, suffix) => path.join(STATE, `${id}.${suffix}`);
const markerOf = (id) => `[变异${id}]`;
const applied = new Set();

function requireSeed(id) {
  const target = id && SEEDS[id];
  if (!target) {
    console.error(`未知 SEED：${id}（可选：${Object.keys(SEEDS).join(", ")}）`);
    process.exit(2);
  }
  return target;
}

function apply(id) {
  const target = requireSeed(id);
  const absolute = path.join(REPO, target.file);
  mkdirSync(STATE, { recursive: true });
  const raw = readFileSync(absolute);
  const text = raw.toString("utf8");
  const eol = eolOf(text);
  const lf = toLf(text);
  if (lf.includes(markerOf(id))) {
    console.error(`前置状态闸：${target.file} 已含 ${markerOf(id)}，请先 restore ${id}。`);
    process.exit(3);
  }
  const hits = lf.split(target.seed).length - 1;
  if (hits !== 1) {
    console.error(`SEED ${id} 在 ${target.file} 命中 ${hits} 次（期望 1）；源码重写会打散旧 SEED。`);
    process.exit(3);
  }
  writeFileSync(statePath(id, "orig"), raw);
  writeFileSync(statePath(id, "sha256"), sha256(raw) + "\n", "utf8");
  writeFileSync(absolute, Buffer.from(fromLf(lf.replace(target.seed, target.mutant), eol), "utf8"));
  applied.add(id);
  console.log(`APPLIED ${id} → ${target.file}`);
}

function restore(id, { quiet = false } = {}) {
  const target = requireSeed(id);
  const absolute = path.join(REPO, target.file);
  const backup = statePath(id, "orig");
  if (!existsSync(backup)) {
    if (!quiet) console.log(`ALREADY-CLEAN：${target.file}（未施加过 ${id}）`);
    return;
  }
  const original = readFileSync(backup);
  const expected = readFileSync(statePath(id, "sha256"), "utf8").trim();
  writeFileSync(absolute, original);
  const actual = sha256(readFileSync(absolute));
  if (actual !== expected) {
    console.error(`还原后 sha256 不一致：${target.file}（期望 ${expected}，实际 ${actual}）`);
    process.exit(4);
  }
  rmSync(backup);
  rmSync(statePath(id, "sha256"));
  applied.delete(id);
  if (!quiet) console.log(`RESTORED ${id} → ${target.file}（sha256 ${actual.slice(0, 12)} 一致）`);
}

function restoreAll() {
  for (const id of [...applied]) {
    try { restore(id, { quiet: true }); console.log(`RESTORED ${id}（异常钩子）`); }
    catch (error) { console.error(`异常钩子还原 ${id} 失败：${error.message} —— 文件：${SEEDS[id].file}`); }
  }
}

for (const signal of ["SIGINT", "SIGTERM", "SIGHUP"]) {
  process.on(signal, () => { restoreAll(); process.exit(130); });
}
process.on("uncaughtException", (error) => {
  console.error(`未捕获异常：${error?.stack ?? error}`);
  restoreAll();
  process.exit(5);
});
process.on("unhandledRejection", (reason) => {
  console.error(`未处理的 rejection：${reason?.stack ?? reason}`);
  restoreAll();
  process.exit(5);
});

const [mode, id] = process.argv.slice(2);
if (mode === "apply") apply(id);
else if (mode === "restore") restore(id);
else if (mode === "plan") {
  for (const [seedId, target] of Object.entries(SEEDS)) {
    console.log([seedId, target.cwd, target.test, target.expect].join("\t"));
  }
} else {
  console.error("用法: node scripts/mutate_review_layout_ratio.mjs plan | apply <ID> | restore <ID>");
  process.exit(2);
}
