// 反向变异脚本：数据范围弹层「动作条的层 + 按钮集合」（2026-10-16）
//
// 背景（用户实测提出）：Excel 的「完成」按钮和 WPS 的「确认」按钮**不在同一层**。
// 根因不是 padding 差几个像素，而是**结构**：
//   · WPS  `div.wps-scope-actions` 是 `.scope-popover-scroll` 的**兄弟** ⇒ 固定页脚；
//   · Excel `div.scope-action-buttons` 原来在滚动容器**里面**（scroll → fieldset
//     → scope-controls → 动作条），所以只是"内容末尾的一行 flex"，会跟着内容滚走。
// 修法：把动作条移到滚动容器之外（同为兄弟），两端读同一份页脚声明，
// 按钮集合同步成 刷新 / 确认 / 取消。
//
// 用法：node scripts/mutate_scope_action_bar_layer.mjs [E1 E2 B1 B2 B3 C1 C2 F1 L1] [restore] | check
// 不带参数：只打印四个目标文件的基线 sha256。
//
// 九个变异各钉一处**互不相同**的属性：
//   E1（Excel 动作条搬回滚动容器里）  E2（WPS 动作条搬回滚动容器里）
//   B1（工作簿动作条的「取消」出现条件写反）  B2（确认按钮文案退回「完成」）
//   B3（工作目录页脚的「取消」出现条件写反 —— 与 B1 是两个调用点）
//   C1（WPS 侧又自己写一遍页脚）      C2（共享声明被拆成只管 Excel 一条）
//   F1（错误条被塞进动作条里）  L1（宽屏两栏覆盖插回共享样式表）
//
// 多步变异的表达方式：`edits` 是**有序**的若干次定点替换，每步都要求
// `from` 命中**恰好 1 处**，否则拒绝执行 —— 这正是「不许盲替」那条纪律。
// 一次结构搬迁（把 `</div>` 从 A 处挪到 B 处）必然是两步：先删、后插。
// 两步的 `from`/`to` 各自独立定点，所以不存在"删对了插错了"还能静默通过的情况。
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";
const APP = resolve(REPO, "apps", "excel-addin", "src", "App.tsx");
const WPS_PANEL = resolve(REPO, "apps", "wps-addin", "src", "WpsProbePanel.tsx");
const SHARED_STYLES = resolve(REPO, "apps", "excel-addin", "src", "styles.css");
const WPS_STYLES = resolve(REPO, "apps", "wps-addin", "src", "styles.css");

// ── 种子：每个串都必须在本文件里**只出现一次** ────────────────────────
/** 滚动容器在 Excel 侧的收口：外层 fieldset 关闭 → 滚动容器关闭 → 紧跟一段 JSX 注释。 */
const EXCEL_SCROLL_CLOSE = [
  "            </fieldset>",
  "            </div>",
  "",
  "            {/*"
].join(LF);
/** 同上去掉滚动容器那一行 —— 即 E1 的第一步结果。 */
const EXCEL_SCROLL_CLOSE_DROPPED = ["            </fieldset>", "", "            {/*"].join(LF);
/** 工作簿错误条的开头（E1 的第二步要在这里**之前**补回滚动容器的关闭）。 */
const EXCEL_SCOPE_ERROR_HEAD = [
  "            )}",
  '            {sourceMode === "workbook" && scopeConfirmationError &&'
].join(LF);
const EXCEL_SCOPE_ERROR_HEAD_WITH_CLOSE = [
  "            )}",
  "            </div>",
  '            {sourceMode === "workbook" && scopeConfirmationError &&'
].join(LF);
/** 动作条 + 错误块的收尾五连（F1 用：把错误条挪进动作条里面）。 */
const EXCEL_BAR_TAIL = [
  "              </div>",
  "            )}",
  '            {sourceMode === "workbook" && scopeConfirmationError &&',
  '              <p role="alert" className="authorized-target-error">{targetGuard.error || scopeConfirmationError}</p>}'
].join(LF);
const EXCEL_BAR_TAIL_ERROR_INSIDE = [
  '              {sourceMode === "workbook" && scopeConfirmationError &&',
  '                <p role="alert" className="authorized-target-error">{targetGuard.error || scopeConfirmationError}</p>}',
  "              </div>",
  "            )}"
].join(LF);
/** 「取消」的出现条件（B1 把它写反）。 */
const EXCEL_DISCARD_GATE = [
  "                {scopeDraftDirty && (",
  '                  <button type="button" className="scope-discard" onClick={discardWorkbookScopeEdit}>'
].join(LF);
/**
 * 工作目录页脚的「取消」出现条件（B3 把它写反）。
 *
 * 与 B1 钉的**不是同一个调用点**：B1 在工作簿动作条（判据是名字草稿 `scopeDraftDirty`），
 * B3 在工作目录固定页脚（判据是键草稿 `folderDraftDirty`）。两者各有各的已提交副本 ——
 * 一边绿一边红不能互相代替。
 */
const EXCEL_FOLDER_DISCARD_GATE = [
  "                {folderDraftDirty && (",
  '                  <button type="button" className="scope-discard" onClick={cancelFolderScopeEdit}>'
].join(LF);
/** 工作簿动作条里确认按钮的文案（B2 退回旧文案）。 */
const EXCEL_CONFIRM_LABEL =
  '{targetGuard.confirming || scopeCheckingCommit ? "核对中..." : status === "scanning" ? "读取中..." : "确认"}';

/**
 * WPS 侧的滚动容器收口：工作目录勾选区的收口 → 滚动容器关闭 → 紧跟工作目录页脚那段注释。
 *
 * 重定点记录（两处结构重写各打散过一次旧锚，**每次都只改锚、不改判别点**）：
 * 2026-10-16 夜第一次：本轮在滚动容器收口与动作条之间插进 `.folder-selection-footer`
 * 与工作目录侧的 `ScopeAdmissionSummary`（两者都在滚动容器**外**，与 Excel 同构），
 * 旧锚（`checking={admission.checking} />` + 两个 `</div>`）失配；
 * 2026-10-16 夜第二次：工作目录模式按 Excel 的格式重排（页脚 = 计数 + 确认同排、
 * 承认汇总在页脚下方、不再有 `.wps-scope-actions`），页脚前多了一段说明注释。
 * 新锚仍取**同一个**判别点 —— 那一行 `</div>` 就是 `.scope-popover-scroll` 的收口。
 */
const WPS_SCROLL_CLOSE = [
  "          )}",
  "          </div>",
  "          {/*",
  "            工作目录侧的固定页脚"
].join(LF);
const WPS_SCROLL_CLOSE_DROPPED = [
  "          )}",
  "          {/*",
  "            工作目录侧的固定页脚"
].join(LF);
/** 补回滚动容器的收口：落在动作条**之后**，于是动作条被留在滚动容器里面。 */
const WPS_SCOPE_ERROR = [
  '          {folderError && <p role="alert" className="authorized-target-error">{folderError}</p>}',
  '          {scopeError && <p role="alert" className="authorized-target-error">{scopeError}</p>}'
].join(LF);
const WPS_SCOPE_ERROR_WITH_CLOSE = [
  '          {folderError && <p role="alert" className="authorized-target-error">{folderError}</p>}',
  "          </div>",
  '          {scopeError && <p role="alert" className="authorized-target-error">{scopeError}</p>}'
].join(LF);

/** 共享页脚声明（C2 把两个类拆开，WPS 那一端就再也没有页脚样式）。 */
const SHARED_FOOTER_CO_LIST = [".scope-action-buttons,", ".wps-scope-actions {"].join(LF);
const SHARED_FOOTER_EXCEL_ONLY = [".scope-action-buttons {"].join(LF);
/**
 * 宽屏两栏覆盖（L1 把它插回来）。
 *
 * 2026-10-16 按用户口径把 `@media (min-width: 520px)` 那段整段删了（见 styles.css 里的说明注释）：
 * 宽面板不该是"左边一直挂着文件列表、右边勾子表"，而是和窄屏一样的「点文件 → 进入 → 勾子表」。
 * L1 把这段原样插回 `.confirm-sheets` 之前 —— 判据（scopeActionBarLayer 的 CSS 契约用例）必须变红。
 * 插入型变异：`to` 里含 `from` 串，`state()` 先判 muted 才不会误报。
 */
const EXCEL_FOLDER_TWO_COLUMN = [
  "@media (min-width: 520px) {",
  "  .folder-master-detail {",
  "    grid-template-columns: minmax(160px, 0.8fr) minmax(230px, 1.2fr);",
  "    gap: 9px;",
  "  }",
  "",
  "  .folder-master-detail .folder-file-list,",
  "  .folder-master-detail.folder-picker-view-sheets .folder-file-list,",
  "  .folder-master-detail .folder-sheet-detail {",
  "    display: grid;",
  "  }",
  "",
  "  .folder-back-button {",
  "    display: none;",
  "  }",
  "}",
  "",
  ".confirm-sheets {"
].join(LF);
const EXCEL_FOLDER_SINGLE_ANCHOR = [".confirm-sheets {"].join(LF);
/** WPS 样式表里紧随我那段说明注释的第一条规则（C1 在它前面插回一条同名规则）。 */
const WPS_BACKUP_RULE_HEAD = [".wps-backup-content {", "  min-height: 0;"].join(LF);
const WPS_DUPLICATE_FOOTER = [
  ".wps-scope-actions {",
  "  display: flex;",
  "  flex: 0 0 auto;",
  "  justify-content: flex-end;",
  "  gap: 7px;",
  "  border-top: 1px solid var(--color-border);",
  "  padding: 10px var(--spacing-md);",
  "}",
  "",
  WPS_BACKUP_RULE_HEAD
].join(LF);

const TARGETS = {
  E1: {
    file: APP,
    app: "excel",
    note: "Excel 动作条搬回滚动容器里（层的原始缺陷）",
    edits: [
      { from: EXCEL_SCROLL_CLOSE, to: EXCEL_SCROLL_CLOSE_DROPPED },
      { from: EXCEL_SCOPE_ERROR_HEAD, to: EXCEL_SCOPE_ERROR_HEAD_WITH_CLOSE }
    ]
  },
  E2: {
    file: WPS_PANEL,
    app: "wps",
    note: "WPS 动作条搬回滚动容器里（另一端同一判据）",
    edits: [
      { from: WPS_SCROLL_CLOSE, to: WPS_SCROLL_CLOSE_DROPPED },
      { from: WPS_SCOPE_ERROR, to: WPS_SCOPE_ERROR_WITH_CLOSE }
    ]
  },
  B1: {
    file: APP,
    app: "excel",
    note: "「取消」出现条件写反：有草稿差异时反而没有出口",
    edits: [{ from: EXCEL_DISCARD_GATE, to: EXCEL_DISCARD_GATE.replace("scopeDraftDirty", "!scopeDraftDirty") }]
  },
  B2: {
    file: APP,
    app: "excel",
    note: "确认按钮文案退回「完成」，两端叫法又分叉",
    edits: [{ from: EXCEL_CONFIRM_LABEL, to: EXCEL_CONFIRM_LABEL.replace('"确认"', '"完成"') }]
  },
  B3: {
    file: APP,
    app: "excel",
    note: "工作目录页脚的「取消」出现条件写反：草稿与已提交一致时才有出口、真有差异时反而没有",
    edits: [
      {
        from: EXCEL_FOLDER_DISCARD_GATE,
        to: EXCEL_FOLDER_DISCARD_GATE.replace("folderDraftDirty", "!folderDraftDirty")
      }
    ]
  },
  C1: {
    file: WPS_STYLES,
    app: "excel",
    note: "WPS 侧又自己写一遍页脚（跨端 CSS 契约在 Excel 侧判红）",
    edits: [{ from: WPS_BACKUP_RULE_HEAD, to: WPS_DUPLICATE_FOOTER }]
  },
  C2: {
    file: SHARED_STYLES,
    app: "excel",
    note: "共享页脚声明被拆成只管 Excel 一条",
    edits: [{ from: SHARED_FOOTER_CO_LIST, to: SHARED_FOOTER_EXCEL_ONLY }]
  },
  F1: {
    file: APP,
    app: "excel",
    note: "错误条被塞进动作条里面（位置对了、层错了）",
    edits: [{ from: EXCEL_BAR_TAIL, to: EXCEL_BAR_TAIL_ERROR_INSIDE }]
  },
  L1: {
    file: SHARED_STYLES,
    app: "excel",
    note: "宽屏两栏覆盖插回共享样式表（进入式单栏被翻回「左文件右子表」，返回按钮重新被藏）",
    edits: [{ from: EXCEL_FOLDER_SINGLE_ANCHOR, to: EXCEL_FOLDER_TWO_COLUMN }]
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

function hits(text, needle) {
  return text.split(needle).length - 1;
}

/**
 * 判定一个变异此刻处在哪个状态。
 *
 * `muted` 优先：插入型变异的 `from` 串天然是 `to` 串的子串（把 B 插到 A 前面，
 * 原来的 A 还在里面），若先判 `unmuted` 会把"已变异"误报成"未变异"。
 * 真正防误改的是 apply/restore 每一步的 `命中恰好 1 处` 拒绝，不是这里。
 */
function state(current, target) {
  const muted = target.edits.every((edit) => hits(current, edit.to) === 1);
  if (muted) return { muted: true, unmuted: false, detail: "MUTATED" };
  const unmuted = target.edits.every((edit) => hits(current, edit.from) === 1);
  return { muted: false, unmuted, detail: unmuted ? "SEED" : "BROKEN" };
}

const which = process.argv[2];
if (!which) {
  for (const file of [APP, WPS_PANEL, SHARED_STYLES, WPS_STYLES]) {
    console.log(`BASELINE ${file} sha256=${sha256(file)}`);
  }
  process.exit(0);
}

/**
 * 前置状态闸：每个变异要么**完整在位**、要么**完整缺失**，不允许半截子状态。
 * 源码重写（改注释、重排声明、换行）会打散旧串，所以改完源码必跑 `check`。
 */
if (which === "check") {
  let bad = 0;
  for (const [id, target] of Object.entries(TARGETS)) {
    const current = normalize(readFileSync(target.file, "utf8"));
    const verdict = state(current, target);
    const ok = verdict.unmuted || verdict.muted;
    if (!ok) {
      bad += 1;
      const counts = target.edits
        .map((edit) => `from=${hits(current, edit.from)} to=${hits(current, edit.to)}`)
        .join("  ");
      console.log(`MISS ${id} (${target.file}): ${counts}（期望每步 from=1,to=0 或 to=1）`);
    } else {
      console.log(`OK   ${id} ${verdict.detail}`);
    }
  }
  console.log(bad === 0 ? "check: OK" : `check: ${bad} 个变异串已失配`);
  process.exit(bad === 0 ? 0 : 3);
}

if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_scope_action_bar_layer.mjs [E1 E2 B1 B2 B3 C1 C2 F1 L1] [restore] | check");
  process.exit(2);
}

const target = TARGETS[which];
const original = readFileSync(target.file, "utf8");
let text = normalize(original);

if (process.argv[3] === "restore") {
  const verdict = state(text, target);
  if (verdict.unmuted) {
    console.log("ALREADY-CLEAN");
    process.exit(0);
  }
  if (!verdict.muted) {
    console.error(`REFUSE: ${which} 处于半截子状态（既非 SEED 也非 MUTANT）于 ${target.file}，拒绝还原`);
    process.exit(4);
  }
  // 逆序撤销，每一步都要求命中恰好 1 处。
  for (const edit of [...target.edits].reverse()) {
    const count = hits(text, edit.to);
    if (count !== 1) {
      console.error(`REFUSE: ${which} 还原步骤命中 ${count} 处（期望 1）于 ${target.file}`);
      process.exit(4);
    }
    text = text.replace(edit.to, edit.from);
  }
  writeFileSync(target.file, denormalize(text, original), "utf8");
  console.log(`restore: OK ${target.file}`);
  process.exit(0);
}

const before = state(text, target);
if (before.muted) {
  console.log("ALREADY-MUTATED");
  process.exit(0);
}
if (!before.unmuted) {
  console.error(`REFUSE: ${which} 的种子已失配（SEED 命中数不为 1）于 ${target.file}，先跑 check`);
  process.exit(3);
}
for (const edit of target.edits) {
  const count = hits(text, edit.from);
  if (count !== 1) {
    console.error(`REFUSE: ${which} 施加步骤命中 ${count} 处（期望 1）于 ${target.file}`);
    process.exit(4);
  }
  text = text.replace(edit.from, edit.to);
}
writeFileSync(target.file, denormalize(text, original), "utf8");
console.log(`APPLIED ${which} → ${target.file}（${target.note}）`);
