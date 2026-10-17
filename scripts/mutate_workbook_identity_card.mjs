// 反向变异脚本：数据范围弹层「当前文件卡 / 触发器的日期显示」（2026-10-16 夜）
//
// 口径（走过两步，别当成同一件事）：
//   ① 「文件名中未识别到报表日期，这个文字不应该出现呀」—— 只删否定式说明，
//      文件名里有日期时**照旧显示**「报表日期 2026-03-15」；
//   ② 「不需要 识别日期」—— `extractWorkbookDataPeriod` 与 `workbookDataPeriod` **整条删除**，
//      两个调用点（当前文件卡 / `.scope-trigger`）都只剩"哪个文件、哪张表"。
//
// 所以变异必须**两头夹**：
//   往下（H2/H3）证明"不许再出现否定式说明 / 不许整行删掉"；
//   往上（H1/H4）证明"不许把识别请回来" —— 只看 H2/H3 的话，一个"有则显示"的版本能全身而过。
//
// 用法：node scripts/mutate_workbook_identity_card.mjs [H1 H2 H3 H4] [restore] | check
// 不带参数：只打印目标文件当前的基线 sha256。
//
// 四个变异各钉一处**互不相同**的属性（H1/H4 还各钉一个**不同的调用点**）：
//   H1（把日期识别请回**当前文件卡**）—— 就地内联，不依赖已删掉的导出，
//        避免"变异靠一个不存在的标识符非 0 退出"这种**编译错误冒充判红**；
//   H2（退回「文件名中未识别到报表日期」否定式）—— 用户最早报告的那句话本身；
//   H3（整行 `<em>` 删掉）—— 证明负向断言不是靠"把整行删掉"就能满足；
//   H4（把日期加回**触发器**）—— 只钉住卡片的那条判据抓不到它。
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";
const APP = resolve(REPO, "apps", "excel-addin", "src", "App.tsx");

// ── 种子（改完 App.tsx 必跑 `check`，否则这里会静静失配）──────────────
/** 当前文件卡的第二行：整行单行、无三元 —— 这就是"识别整条已下线"的形状。 */
const CARD_LINE = "                  <em>当前工作表 {workbook.activeWorksheet}</em>";
/** 触发器：`auto` 那一支的文案（另一支是「已固定 N 个工作表」）。 */
const TRIGGER_AUTO_LINE = "                    ? `当前表 ${workbook.activeWorksheet}`";

/** H1 的目标：把识别就地写回卡片（不引用已删除的 `workbookDataPeriod`）。 */
const CARD_WITH_RECOGNITION = [
  "                  <em>",
  "                    {(() => {",
  "                      const match = workbook.name.match(/(\\d{4})-(\\d{2})-(\\d{2})/);",
  "                      return match ? `报表日期 ${match[0]} · ` : \"\";",
  "                    })()}",
  "                    当前工作表 {workbook.activeWorksheet}",
  "                  </em>"
].join(LF);
/** H2 的目标：用户最早报告的那句否定式说明。 */
const CARD_WITH_NEGATION =
  "                  <em>文件名中未识别到报表日期 · 当前工作表 {workbook.activeWorksheet}</em>";
/** H3 的目标：整行换成一句注释（`<em>` 不再存在）。 */
const CARD_REPLACED_BY_COMMENT = "                  {/* 第二行被整行删掉 */}";
/** H4 的目标：把日期前缀加回触发器（同样就地内联）。 */
const TRIGGER_AUTO_WITH_DATE =
  "                    ? `${(() => { const m = workbook.name.match(/\\d{4}-\\d{2}-\\d{2}/); return m ? m[0] + \" · \" : \"\"; })()}当前表 ${workbook.activeWorksheet}`";

const TARGETS = {
  H1: {
    note: "把日期识别请回当前文件卡（就地内联）",
    edits: [{ from: CARD_LINE, to: CARD_WITH_RECOGNITION }]
  },
  H2: {
    note: "退回「文件名中未识别到报表日期」否定式（用户最早报告的那句话）",
    edits: [{ from: CARD_LINE, to: CARD_WITH_NEGATION }]
  },
  H3: {
    note: "整行 `<em>` 被删掉（「不许出现」被误读成「这一行不要」）",
    edits: [{ from: CARD_LINE, to: CARD_REPLACED_BY_COMMENT }]
  },
  H4: {
    note: "把日期加回数据范围触发器（卡片判据抓不到的那半边）",
    edits: [{ from: TRIGGER_AUTO_LINE, to: TRIGGER_AUTO_WITH_DATE }]
  }
};

/** 目标文件自带 CRLF，先归一化，否则 includes 会误报 ALREADY-CLEAN。 */
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

function hits(text, needle) {
  return text.split(needle).length - 1;
}

/** `muted` 优先：插入/扩写型变异的 `from` 串常是 `to` 串的片段，先判 `unmuted` 会读反。 */
function state(current, target) {
  const muted = target.edits.every((edit) => hits(current, edit.to) === 1);
  if (muted) return { muted: true, unmuted: false, detail: "MUTATED" };
  const unmuted = target.edits.every((edit) => hits(current, edit.from) === 1);
  return { muted: false, unmuted, detail: unmuted ? "SEED" : "BROKEN" };
}

const which = process.argv[2];
if (!which) {
  console.log(`BASELINE ${APP} sha256=${sha256(APP)}`);
  process.exit(0);
}

/** 前置状态闸：每个变异要么完整在位、要么完整缺失；源码一改就必须先跑它。 */
if (which === "check") {
  let bad = 0;
  for (const [id, target] of Object.entries(TARGETS)) {
    const current = normalize(readFileSync(APP, "utf8"));
    const verdict = state(current, target);
    if (verdict.unmuted || verdict.muted) {
      console.log(`OK   ${id} ${verdict.detail}`);
    } else {
      bad += 1;
      const counts = target.edits
        .map((edit) => `from=${hits(current, edit.from)} to=${hits(current, edit.to)}`)
        .join("  ");
      console.log(`MISS ${id} (${APP}): ${counts}（期望每步 from=1,to=0 或 to=1）`);
    }
  }
  console.log(bad === 0 ? "check: OK" : `check: ${bad} 个变异串已失配`);
  process.exit(bad === 0 ? 0 : 3);
}

if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_workbook_identity_card.mjs [H1 H2 H3 H4] [restore] | check");
  process.exit(2);
}

const target = TARGETS[which];
const original = readFileSync(APP, "utf8");
let text = normalize(original);

if (process.argv[3] === "restore") {
  const verdict = state(text, target);
  if (verdict.unmuted) {
    console.log("ALREADY-CLEAN");
    process.exit(0);
  }
  if (!verdict.muted) {
    console.error(`REFUSE: ${which} 处于半截子状态（既非 SEED 也非 MUTANT），拒绝还原 ${APP}`);
    process.exit(4);
  }
  for (const edit of [...target.edits].reverse()) {
    const count = hits(text, edit.to);
    if (count !== 1) {
      console.error(`REFUSE: ${which} 还原步骤命中 ${count} 处（期望 1）于 ${APP}`);
      process.exit(4);
    }
    text = text.replace(edit.to, edit.from);
  }
  writeFileSync(APP, denormalize(text, original), "utf8");
  console.log(`restore: OK ${APP}`);
  process.exit(0);
}

const before = state(text, target);
if (before.muted) {
  console.log("ALREADY-MUTATED");
  process.exit(0);
}
if (!before.unmuted) {
  console.error(`REFUSE: ${which} 的种子已失配（SEED 命中数不为 1），先跑 check`);
  process.exit(3);
}
for (const edit of target.edits) {
  const count = hits(text, edit.from);
  if (count !== 1) {
    console.error(`REFUSE: ${which} 施加步骤命中 ${count} 处（期望 1）于 ${APP}`);
    process.exit(4);
  }
  text = text.replace(edit.from, edit.to);
}
writeFileSync(APP, denormalize(text, original), "utf8");
console.log(`APPLIED ${which} → ${APP}（${target.note}）`);
