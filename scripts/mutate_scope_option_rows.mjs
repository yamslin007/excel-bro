// 反向变异脚本：数据范围面板「选项行」的四条口径（2026-10-16）
//
// 背景（都是无头浏览器/截图实测出来的）：
//   ① 列表行内不该再出现授权状态文案（「可分析 · 修改仍需预览确认」）；
//   ② 行内勾选框被 `.scope-mode-list i` 这条**后代**选择器吃掉，撑成 25px 灰方块
//      （实测 rgb(241,240,238)）—— 它只该匹配模式卡的直接子代图标；
//   ③ 工作表行与模式卡（跟随当前 / 选择多个 / 工作目录）必须同高（59px），
//      原先各行自定 38px / 57px；
//   ④ Excel 面板里块与块的间距被外层 fieldset 吃掉（实测 0），要跟 WPS 一样是 10px。
//
// 用法：node scripts/mutate_scope_option_rows.mjs [M1..M7] [restore]
// 不带参数：只打印四个目标文件的基线 sha256。
//
// 每个变异钉一处**互不相同**的属性，且都挂在该机制**真正管辖**的代码上：
//   M1（WPS 行内文案）  M2（Excel 行内文案）  M3（Excel 后代选择器泄漏）
//   M4（WPS 侧另写一套勾选框）  M5（WPS 行高）  M6（Excel 行高）  M7（块间距 token）
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";
const WPS_PANEL = resolve(REPO, "apps", "wps-addin", "src", "WpsProbePanel.tsx");
const WPS_STYLES = resolve(REPO, "apps", "wps-addin", "src", "styles.css");
const SHARED_STYLES = resolve(REPO, "apps", "excel-addin", "src", "styles.css");
const APP = resolve(REPO, "apps", "excel-addin", "src", "App.tsx");

/** 变异里重新插回去的那句授权文案（= SCOPE_ADMISSION_MESSAGES.none）。 */
const BADGE_TEXT = "可分析 · 修改仍需预览确认";
const BADGE_MARKUP = `<small className="scope-admission-badge">${BADGE_TEXT}</small>`;

// ── 种子 ──────────────────────────────────────────────────────────
const WPS_ROW_NAME = '                      <span>{sheet.name}</span>';
const EXCEL_ROW_CAPTION = [
  "                                <small>",
  "                                  {sheet.name === workbook.activeWorksheet",
  '                                    ? "当前工作表 · "',
  '                                    : ""}',
  '                                  {workbook.catalogOnly ? "字段按需读取" : `${sheet.rowCount} 行 · ${sheet.columnCount} 列`}',
  "                                </small>"
].join(LF);
const EXCEL_MODE_ICON = [
  ".scope-mode-list > button > i {",
  "  display: grid;"
].join(LF);
const WPS_ROW_SPAN_RULE = [
  ".wps-sheet-picker-options > button span {",
  "  overflow: hidden;",
  "  text-overflow: ellipsis;",
  "  white-space: nowrap;",
  "}"
].join(LF);
// 【定点锚】不能只写 `  min-height: var(--scope-option-height);` ——
// `.connection-shell > button` 也有 `min-height: 38px;`，那样"还原"时
// `split(mutant).join(seed)` 会把无关规则一起改掉（真的发生过：连接页按钮
// 被撑成 59px）。种子里必须带上一行**只属于这条规则**的声明。
const WPS_ROW_HEIGHT = [
  "  grid-template-columns: 20px minmax(0, 1fr) auto;",
  "  min-height: var(--scope-option-height);"
].join(LF);
const EXCEL_ROW_HEIGHT = [
  ".workbook-sheet-list > button {",
  "  display: flex;",
  "  width: 100%;",
  "  min-height: var(--scope-option-height);"
].join(LF);
const BLOCK_GAP = [
  ".scope-selection-fields {",
  "  display: grid;",
  "  gap: var(--scope-popover-block-gap);"
].join(LF);

const TARGETS = {
  M1: {
    file: WPS_PANEL,
    app: "wps",
    seed: WPS_ROW_NAME,
    // [变异 M1] 行内又把授权文案插回来 —— 用户明确说过不要它。
    mutant: `${WPS_ROW_NAME}${LF}                      {draftIndexes.includes(sheet.index) && (${LF}                        ${BADGE_MARKUP}${LF}                      )}`
  },
  M2: {
    file: APP,
    app: "excel",
    seed: EXCEL_ROW_CAPTION,
    // [变异 M2] 工作簿列表行内又把授权文案插回来。
    mutant: `${EXCEL_ROW_CAPTION}${LF}                                ${BADGE_MARKUP}`
  },
  M3: {
    file: SHARED_STYLES,
    app: "excel",
    seed: EXCEL_MODE_ICON,
    // [变异 M3] 模式图标退回**后代**选择器：行内勾选框又被吃成 25px 灰方块。
    mutant: [".scope-mode-list i {", "  display: grid;"].join(LF)
  },
  M4: {
    file: WPS_STYLES,
    app: "wps",
    seed: WPS_ROW_SPAN_RULE,
    // [变异 M4] WPS 侧又给行内勾选框另写一套（特异度盖掉共享的 `.sheet-checkbox`）。
    mutant: `${WPS_ROW_SPAN_RULE}${LF}${LF}.wps-sheet-picker-options > button i {${LF}  display: grid;${LF}  width: 20px;${LF}  height: 20px;${LF}  place-items: center;${LF}  border: 1px solid var(--color-border-strong);${LF}}`
  },
  M5: {
    file: WPS_STYLES,
    app: "wps",
    seed: WPS_ROW_HEIGHT,
    // [变异 M5] 工作表行退回自定的 38px（比模式卡矮一截）。
    mutant: WPS_ROW_HEIGHT.replace("  min-height: var(--scope-option-height);", "  min-height: 38px;")
  },
  M6: {
    file: SHARED_STYLES,
    app: "excel",
    seed: EXCEL_ROW_HEIGHT,
    // [变异 M6] Excel 工作表行退回内容高度（57px），不再跟模式卡对齐。
    mutant: EXCEL_ROW_HEIGHT.replace("  min-height: var(--scope-option-height);", "  min-height: 57px;")
  },
  M7: {
    file: SHARED_STYLES,
    app: "excel",
    seed: BLOCK_GAP,
    // [变异 M7] 内层 fieldset 不再读间距 token：Excel 的块间距又塌回 0。
    mutant: [".scope-selection-fields {", "  display: grid;", "  gap: 0;"].join(LF)
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
  for (const file of [WPS_PANEL, WPS_STYLES, SHARED_STYLES, APP]) {
    console.log(`BASELINE ${file} sha256=${sha256(file)}`);
  }
  process.exit(0);
}

/**
 * 前置状态闸：每个种子/变异串都必须**定点**。
 *  - 种子出现次数必须恰好 1（否则不知该改哪一处）；
 *  - 变异串在未变异时出现次数必须恰好 0（否则"还原"会把本来就存在的、
 *    与本次变异无关的那一处也一起改掉 —— `.connection-shell > button`
 *    被 `min-height: 38px;` 串误伤就是这条没查）。
 * 源码重写（换行、重排声明）会打散旧串，所以改完源码必跑 `check`。
 */
if (which === "check") {
  let bad = 0;
  for (const [id, target] of Object.entries(TARGETS)) {
    const current = normalize(readFileSync(target.file, "utf8"));
    // 变异未施加时种子在位；已施加时种子为 0、变异串为 1，两种都算健康。
    const seedHits = current.split(target.seed).length - 1;
    const mutantHits = current.split(target.mutant).length - 1;
    const ok = (seedHits === 1 && mutantHits === 0) || (seedHits === 0 && mutantHits === 1);
    if (!ok) {
      bad += 1;
      console.log(
        `MISS ${id} (${target.file}): SEED=${seedHits} MUTANT=${mutantHits}（期望 SEED=1,MUTANT=0）`
      );
    } else {
      console.log(`OK   ${id} SEED=${seedHits} MUTANT=${mutantHits}`);
    }
  }
  console.log(bad === 0 ? "check: OK" : `check: ${bad} 个变异串已失配`);
  process.exit(bad === 0 ? 0 : 3);
}

if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_scope_option_rows.mjs [M1..M7] [restore] | check");
  process.exit(2);
}

const target = TARGETS[which];
const original = readFileSync(target.file, "utf8");

if (process.argv[3] === "restore") {
  const text = normalize(original);
  const hits = text.split(target.mutant).length - 1;
  if (hits === 0) {
    console.log("ALREADY-CLEAN");
    process.exit(0);
  }
  if (hits !== 1) {
    // 不许盲替：命中多于一处说明串不唯一，"还原"会改到无关代码。
    console.error(`REFUSE: MUTANT 命中 ${hits} 处（期望 1）于 ${target.file}，拒绝还原`);
    process.exit(4);
  }
  writeFileSync(target.file, denormalize(text.replace(target.mutant, target.seed), original), "utf8");
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
