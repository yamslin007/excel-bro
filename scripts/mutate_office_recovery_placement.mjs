// 反向变异：**唯一恢复入口**这一组断言真的能红（2026-10-29）
//
// 为什么必须做：本轮把恢复卡从"产生它的那条消息下面"搬到"输入区上方"，判据是
// `excelWriteEndToEnd.test.tsx` 的三条结构断言 + 真实布局脚本。结构断言这种东西
// 极易假绿（元素还在、文案还在，只是位置变回去了）。所以每个变异把机制**拆掉**，
// 看对应用例是不是真的红、失败原文里有没有它自己那处证据。
//
// P1（两步结构搬迁）：把输入区上方的恢复卡整块搬回那条消息里 —— 全页仍然只有一张
//    可点入口，所以"按钮数量 = 1"那条仍是绿的，只有"入口不许住在滚动容器里"必须红。
//    这正好证明两条断言测的是**两个**属性，不是同一句。JSX 里挪一块结构必然是两步
//    （先删、再插），所以 `edits[]` 有序多步，每步都要求命中恰好 1 处。
// P2：往只读记录里补一个可点按钮 —— "历史消息里没有按钮""全页只有一个续跑入口"必须红。
// P3：共享选择器从"取最后一条"改成"取第一条" —— core 的"新的执行覆盖旧的"必须红。
//    App 侧不变红是**对的**：那里只有一条卡，选谁都一样；所以判别点挂在 core。
//
// 用法：
//   node scripts/mutate_office_recovery_placement.mjs              # 打印三个基线 sha256
//   node scripts/mutate_office_recovery_placement.mjs check P1     # 是否已干净地变异（0=是）
//   node scripts/mutate_office_recovery_placement.mjs P1           # 应用
//   node scripts/mutate_office_recovery_placement.mjs P1 restore   # 还原（逐字节校验）
import { createHash } from "node:crypto";
import { readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const APP = resolve(REPO, "apps", "excel-addin", "src", "App.tsx");
const PANEL_COMPONENT = resolve(
  REPO, "apps", "excel-addin", "src", "components", "OfficeRecoveryPanel.tsx"
);
const CORE_VIEW = resolve(REPO, "packages", "core", "src", "executionRecoveryView.ts");

/** 输入区上方那一块（原样，8 空格缩进）。P1 把它整块删掉。 */
const PANEL_BLOCK = [
  "        {officeRecoveryEntry && (",
  "          <OfficeRecoveryPanel",
  "            card={officeRecoveryEntry.card}",
  "            busy={busy}",
  "            onAction={(action) => void runOfficeRecoveryAction(officeRecoveryEntry.row, action)}",
  "            onReverify={() => void reverifyOfficeRecoveryMessage(officeRecoveryEntry.row)}",
  "          />",
  "        )}"
].join(LF);

/**
 * 旧形态：把同一张卡画在**产生它的那条消息**里（14 空格缩进，消息流内部）。
 * 这是本轮回退要复现的形状 —— 不是"再插一张"，否则"唯一入口"那条也会红，
 * 判别点就不唯一了。
 */
const LEGACY_BLOCK = [
  "              {message.officeResume && (",
  "                <OfficeRecoveryPanel",
  "                  card={message.officeResume}",
  "                  busy={busy}",
  "                  onAction={(action) => void runOfficeRecoveryAction(message, action)}",
  "                  onReverify={() => void reverifyOfficeRecoveryMessage(message)}",
  "                />",
  "              )}"
].join(LF);

/** 消息里那段只读记录的起点（插在它前面 = 卡片排在记录上面，与旧形态一致）。 */
const RECORD_ANCHOR = [
  "              {message.officeResume && (",
  "                /*"
].join(LF);

/**
 * P1 第 1 步留一个**占位注释**而不是删空。
 *
 * 删空之后 `restore` 只能把原文补回文件末尾 —— 那会把 App.tsx 改坏，也过不了
 * "还原后逐字节 sha256 一致"这条判据。占位注释让还原点的位置永远唯一。
 */
const PANEL_PLACEHOLDER = "        {/* MUTANT P1：恢复入口搬回消息里 */}";

const TARGETS = {
  P1: [
    { file: APP, from: PANEL_BLOCK, to: PANEL_PLACEHOLDER },
    { file: APP, from: RECORD_ANCHOR, to: LEGACY_BLOCK + LF + RECORD_ANCHOR }
  ],
  P2: [
    {
      file: PANEL_COMPONENT,
      from: "      <strong>执行记录</strong>",
      to: [
        "      <strong>执行记录</strong>",
        '      <button type="button">继续剩余动作</button>'
      ].join(LF)
    }
  ],
  P3: [
    {
      file: CORE_VIEW,
      from: "  for (let index = rows.length - 1; index >= 0; index -= 1) {",
      to: "  for (let index = 0; index < rows.length; index += 1) {"
    }
  ]
};

function normalize(text) {
  return text.split(/\r\n/).join(LF);
}
function denormalize(text, original) {
  const native = original.includes("\r\n") ? "\r\n" : LF;
  return text.split(LF).join(native);
}
function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

/** 写回必须读回校验并重试：变异体留在工作区比变异失败更糟。 */
function writeVerified(path, text) {
  const expected = normalize(text);
  let last = "";
  for (let attempt = 1; attempt <= 5; attempt += 1) {
    writeFileSync(path, text, "utf8");
    last = normalize(readFileSync(path, "utf8"));
    if (last === expected) return true;
  }
  console.error(`写回校验失败（5 次）：${path}`);
  console.error(`期望片段：${expected.slice(0, 80)}`);
  console.error(`实际片段：${last.slice(0, 80)}`);
  return false;
}

/**
 * 这一步现在是否处于"已变异"。
 *
 * **`muted` 优先**：插入型变异的 `to` 天然包含 `from`（把 B 插到 A 前面，A 还在里面），
 * 先判 "from 还在" 会把已变异误报成未变异。删除型（`to === ""`）反过来：
 * 只有 `from` 彻底不在了才算已变异。
 */
function isMuted(edit) {
  return normalize(readFileSync(edit.file, "utf8")).includes(edit.to);
}

// 支持两种写法：`P1` / `P1 restore` / `check P1`
const checkMode = process.argv[2] === "check";
const which = checkMode ? process.argv[3] : process.argv[2];
if (!which) {
  for (const [id, edits] of Object.entries(TARGETS)) {
    for (const edit of edits) {
      console.log(`BASELINE ${id} ${edit.file} sha256=${sha256(edit.file)}`);
    }
  }
  process.exit(0);
}
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_office_recovery_placement.mjs [P1|P2|P3] [restore|check]");
  process.exit(2);
}

const edits = TARGETS[which];
const mode = checkMode ? "check" : process.argv[3];

if (mode === "check") {
  const states = edits.map((edit) => ({
    file: edit.file,
    muted: isMuted(edit),
    occurrences: normalize(readFileSync(edit.file, "utf8")).split(edit.from).length - 1
  }));
  for (const state of states) {
    console.log(
      `CHECK ${which} ${state.file} muted=${state.muted} seedOccurrences=${state.occurrences}`
    );
  }
  process.exit(states.every((state) => state.muted) ? 0 : 1);
}

if (mode === "restore") {
  let restored = false;
  for (const edit of [...edits].reverse()) {
    if (!isMuted(edit)) continue;
    const original = readFileSync(edit.file, "utf8");
    const next = normalize(original).split(edit.to).join(edit.from);
    if (!writeVerified(edit.file, denormalize(next, original))) process.exit(4);
    restored = true;
  }
  if (!restored) console.log("ALREADY-CLEAN");
  for (const edit of edits) {
    console.log(`restore: ${edit.file} sha256=${sha256(edit.file)}`);
  }
  process.exit(0);
}

// 应用：先整轮校验 SEED 唯一，再动手（半途停下会留下只改了一半的结构）。
for (const [index, edit] of edits.entries()) {
  const occurrences = normalize(readFileSync(edit.file, "utf8")).split(edit.from).length - 1;
  if (occurrences !== 1) {
    console.log(
      `SKIP: 第 ${index + 1} 步 SEED 出现 ${occurrences} 次（期望 1）于 ${edit.file}`
    );
    process.exit(3);
  }
}
for (const edit of edits) {
  const original = readFileSync(edit.file, "utf8");
  const mutated = denormalize(normalize(original).replace(edit.from, edit.to), original);
  if (!writeVerified(edit.file, mutated)) process.exit(4);
}
for (const edit of edits) {
  console.log(`APPLIED ${which} → ${edit.file} sha256=${sha256(edit.file)}`);
}
