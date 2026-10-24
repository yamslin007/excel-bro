// 反向变异脚本：进度浮标「透明底 + 主题色前景图形（三点 / 锚）」的
// **对比度、令牌与"不画底"守卫**
//
// 用法：node scripts/mutate_progress_dot_contrast.mjs M1|M2|M3|M4|M5|M6|M7|M8|check [restore]
//
// **2026-10-24 重新定点**：用户裁决「把这个锚的图标改一下，背景改为透明的」，
// 浮标那层品牌浅色渐变被删掉，所以旧 M2（渐变退回深色药丸）与旧 M3（把
// `brandHover` 改深、验证"最坏端点按所有背景令牌里最暗的算"）**都失去了对应的
// 被拆对象** —— 渐变没了，就没有"最暗端点"可算。旧 M2 顺延为 M4 家族（边框）
// 之外的新判别点，旧 M3 直接换成"悬停不许把底补回来"，另新增 M8（投影）。
//
//   M1（useTheme.ts）把 vivid-green 的前景色换回**纯品牌色** `#2d9b63`（不压深）。
//       打「前景色是否真的被压深到能过 3:1」—— 纯色下它是 2.93:1，掉到门槛以下。
//
//   M2（styles.css）把常态底色从 `transparent` 改成**不透明品牌浅色**
//       `background: var(--color-brand-bg);`。
//       打的是「值必须**真的是** `transparent`」这条 —— 而不是"有没有 `background`
//       声明"。刻意用**令牌**、不画渐变：这样"不许出现渐变/图片底"那条**不会**红，
//       红下来的只能是"底色必须是透明"这一条，判别点是唯一的。
//
//   M3（styles.css）给 `:hover` **补一层品牌浅色底**（`background: var(--color-brand-bg);`）。
//       打「悬停/按下不许重新着色」那条循环：旧写法是"渐变整体深一档"，底取消之后
//       再补底等于把用户刚要去掉的东西从 hover 里放回来。它只落在 `buttonRules[1]`，
//       常态那三条断言不受影响 —— 判别点唯一。
//
//   M4（styles.css）把用户裁决掉的**那圈描边加回来** —— 把"显式关掉边框"那条
//       声明换成"画一条 1px 实线"。打「外圈不许画边」这条负向断言 —— 就是
//       2026-10-12 用户说「去掉这个圆形边缘线条」「去掉边框颜色，改为透明即可」
//       两次要求去掉的东西。这里刻意用**令牌**而不是字面色值：那样"只准走令牌"
//       那条不会红，红下来的只能是"不许画边"这一条，判别点是唯一的。
//
//   M5（styles.css）把圆点退回**"降级行 + `color-mix()`"** 的旧写法。
//       打上次事故的本体：同一个属性上的 color-mix 一旦算不出值，`background`
//       会落到 `transparent`、圆点**整个消失**（降级行赢不了层叠）。守卫里那条
//       `not.toContain("color-mix(")` 必须把它拦下来。
//
//   M6（styles.css）把**锚图标**的填充换成另一个主题令牌（`--color-brand-text`）。
//       打「锚与圆点必须取**同一个**令牌」：同控件同底，换字形不该换颜色；
//       `brand-text` 正是守卫证明过过不了 3:1 的那个值。若守卫只查"是不是令牌"、
//       不查"是不是同一个令牌"，此变异不会红。
//
//   M7（styles.css）把锚图标的填充写成**字面量** `#265639`（与令牌同值）。
//       打「取值必须是 `var(--color-*)` 引用而不是字面色值」这条负向锚点。
//       注意它与 M6 的判别点不同：这里**值完全相同**，红的只能是"必须走令牌"；
//       同值冗余不算空变异 —— 它证明的恰恰是"同值也不行，必须引用令牌"。
//
//   M8（styles.css）把**投影加回来** `box-shadow: var(--shadow-md);`。
//       打「透明底上不能留投影」这条 —— `--shadow-*` 是按"有一个圆形底"算的，
//       底变透明之后它只剩一团没有主体的灰影，读成脏点。刻意用**令牌**，
//       所以"不许出现字面色值"那条不会红，判别点唯一。
//
// 八个变异落在**两个文件**上：M2/M3/M4/M5/M6/M7/M8 在 styles.css（全 LF），
// M1 在 useTheme.ts（全 CRLF）。所以 seed/mutant 的换行归一化与基线 sha256
// 都要**分文件**做，不能共用一份。
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const STYLES = resolve(REPO, "apps", "excel-addin", "src", "styles.css");
const THEME = resolve(REPO, "apps", "excel-addin", "src", "hooks", "useTheme.ts");

const TARGETS = {
  M1: {
    file: THEME,
    // `progressGlyph` 四个预设各一行，`#247c4f` 是 vivid-green 独有的。
    seed: '    progressGlyph: "#247c4f",',
    mutant: '    progressGlyph: "#2d9b63",'
  },
  M2: {
    file: STYLES,
    // 五条声明绑在一起才唯一：`background: transparent;` 全表 40 次、
    // `box-shadow: none;` 19 次，各自单独取都会 SKIP。
    // 这五行是浮标常态规则的**收尾段**（外面那圈 `border: none` 到 `cursor`），
    // 换掉哪一条都不会与别的规则撞车。
    seed: [
      "  border: none;",
      "  border-radius: var(--radius-full);",
      "  background: transparent;",
      "  box-shadow: none;",
      "  cursor: pointer;"
    ].join(LF),
    mutant: [
      "  border: none;",
      "  border-radius: var(--radius-full);",
      "  background: var(--color-brand-bg);",
      "  box-shadow: none;",
      "  cursor: pointer;"
    ].join(LF)
  },
  M3: {
    file: STYLES,
    // `:hover` 规则体里插一条底：只影响 `buttonRules[1]`（悬停/按下那条循环），
    // 常态的三条断言完全不动。旧 M3（brandHover 改深）随渐变删除而失去对象。
    seed: ".new-progress-button:hover {",
    mutant: [".new-progress-button:hover {", "  background: var(--color-brand-bg);"].join(LF)
  },
  M4: {
    file: STYLES,
    // 把用户裁决掉的那圈描边**加回来**（把"关掉边框"的声明换成"画一条实线"，
    // 用令牌，所以"只准走令牌"那条不会红 —— 红的必须是"外圈不许画边"这条负向
    // 断言本身，判别点是唯一的）。
    // seed 与 M2/M8 同一段五行（`border: none;` 单独一行在全表出现 5 次）。
    seed: [
      "  border: none;",
      "  border-radius: var(--radius-full);",
      "  background: transparent;",
      "  box-shadow: none;",
      "  cursor: pointer;"
    ].join(LF),
    mutant: [
      "  border: 1px solid var(--color-border-strong);",
      "  border-radius: var(--radius-full);",
      "  background: transparent;",
      "  box-shadow: none;",
      "  cursor: pointer;"
    ].join(LF)
  },
  M5: {
    file: STYLES,
    // 圆点规则里那唯一一条取色声明 —— 唯一。
    seed: "  background: var(--color-progress-glyph);",
    mutant: [
      "  background: var(--color-brand-text);",
      "  background: color-mix(in srgb, var(--color-brand-text) 80%, #000000);"
    ].join(LF)
  },
  M6: {
    file: STYLES,
    // 锚图标规则里那唯一一条 `fill: var(--color-*)` —— 全表只有这一处
    // （圆点走的是 `background`，其余 svg 的 `fill` 都是 `currentColor`/`none`）。
    seed: "  fill: var(--color-progress-glyph);",
    mutant: "  fill: var(--color-brand-text);"
  },
  M7: {
    file: STYLES,
    // 同一条声明，改成**与令牌同值**的字面量：判别点是"必须引用令牌"，
    // 不是"颜色必须不同"。
    seed: "  fill: var(--color-progress-glyph);",
    mutant: "  fill: #265639;"
  },
  M8: {
    file: STYLES,
    // 与 M2 同一段五行，只换 `box-shadow` 那一行：令牌而非字面量，
    // 所以"不许出现字面色值"不会红，红的只能是"透明底上不能留投影"。
    seed: [
      "  border: none;",
      "  border-radius: var(--radius-full);",
      "  background: transparent;",
      "  box-shadow: none;",
      "  cursor: pointer;"
    ].join(LF),
    mutant: [
      "  border: none;",
      "  border-radius: var(--radius-full);",
      "  background: transparent;",
      "  box-shadow: var(--shadow-md);",
      "  cursor: pointer;"
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

const which = process.argv[2];

/*
 * `check` 是**前置状态闸**：确认每个目标文件都处于 SEED 态（seed 恰好 1 次、
 * 变异态 0 次）。上一轮被 Ctrl-C 打断、或另有一个实例并发在改同一个文件时，
 * 文件会停在变异态，下一次运行的基线就会记到变异态上，最后误报成
 * "还原后与基线不一致" —— 指认方向完全是错的。起手就报出来，别白跑几轮 vitest。
 */
if (which === "check") {
  let bad = 0;
  for (const [name, target] of Object.entries(TARGETS)) {
    const text = normalize(readFileSync(target.file, "utf8"));
    const seedCount = text.split(target.seed).length - 1;
    const mutantCount = text.split(target.mutant).length - 1;
    if (seedCount !== 1 || mutantCount !== 0) {
      console.error(
        `CHECK FAIL ${name} ${target.file} SEED=${seedCount}（期望 1）变异态=${mutantCount}（期望 0）`
      );
      bad += 1;
    }
  }
  if (bad > 0) process.exit(4);
  console.log("CHECK OK");
  process.exit(0);
}

if (!which) {
  for (const [name, target] of Object.entries(TARGETS)) {
    const sha = createHash("sha256").update(readFileSync(target.file)).digest("hex");
    console.log(`BASELINE ${name} ${target.file} sha256=${sha}`);
  }
  process.exit(0);
}
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_progress_dot_contrast.mjs M1|M2|M3|M4|M5|M6|M7|M8|check [restore]");
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
  console.log(`S1 restore: OK ${target.file}`);
  process.exit(0);
}

const occurrences = normalize(original).split(target.seed).length - 1;
if (occurrences !== 1) {
  console.log(`SKIP: SEED 出现 ${occurrences} 次（期望 1）。`);
  process.exit(3);
}

writeFileSync(
  target.file,
  denormalize(normalize(original).replace(target.seed, target.mutant), original),
  "utf8"
);
console.log(`APPLIED ${which} → ${target.file}`);
