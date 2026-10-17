// 反向变异脚本：WPS 数据范围面板的**第三种来源 —— 工作目录**（2026-10-16 晚）
//
// 本轮把 WPS 上那颗 `disabled` 的「工作目录」卡按 Excel 的同一套口径打开：
// 选目录 → 文件/表两级勾选 → 承认汇总 → 确认提交授权 → 草稿与取消 →
// 刷新增表不作废、换目录/目标消失即作废；发送侧在工作目录范围下**明确拒绝**，
// 绝不拿当前打开的 WPS 工作簿顶上。
//
// 每个变异**只钉一处属性**，而且要挂在那个机制**真正管辖**的属性上：
//   W1 把旧的能力说明（"WPS 当前仅支持打开的工作簿"）请回卡片副标题；
//   W2 把「勾选」直接当授权提交（草稿/已提交分离被拆掉）；
//   W3 刷新改当同意（目标消失也不再作废授权）；
//   W4 换目录不作废旧授权（旧根上的勾选被带到新根）；
//   W5 EB 系统遗留表也算"可勾选"；
//   W6 发送闸读不到来源（等于永远放行 → 偷用当前工作簿）；
//   W7 触发器在工作目录模式下顺带显示当前工作簿名；
//   C1 存在性名单也过滤 EB 表（把"两个判据"合成一个）。
//
// 用法：node scripts/mutate_wps_work_directory.mjs [W1..W7 C1] [restore] | check
// 不带参数：只打印各目标文件当前的基线 sha256。
//
// 改完源码必跑 `check`（源码重写会打散这里的 SEED）。
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const PANEL = resolve(REPO, "apps", "wps-addin", "src", "WpsProbePanel.tsx");
const AGENT = resolve(REPO, "apps", "wps-addin", "src", "WpsAgentPanel.tsx");
const CORE = resolve(REPO, "packages", "core", "src", "folderScope.ts");

// ── 种子 ──────────────────────────────────────────────────────────────
const FOLDER_CAPTION_LINE =
  "                <span><strong>工作目录</strong><small>批量处理多个本地工作簿</small></span>";

const TOGGLE_FOLDER_SHEET = [
  "  const toggleFolderSheet = (fileId: string, sheetName: string) => {",
  "    const key = folderSheetKey(fileId, sheetName);",
  "    setDraftFolderKeys((current) =>",
  "      current.includes(key) ? current.filter((item) => item !== key) : [...current, key]",
  "    );",
  "  };"
].join(LF);

const PRESERVE_LINE = "      if (!preserveAuthorization) setFolderScope(null);";

const NEW_ROOT_BLOCK = [
  "      if (!sameRoot) {",
  "        setDraftFolderKeys([]);",
  "        setFolderScope(null);",
  "      }"
].join(LF);

const SEND_GUARD_SOURCE_LINE = [
  "    const folderRefusal = workDirectorySendRefusal({",
  "      sourceMode: scopeSource,"
].join(LF);

const TRIGGER_SMALL_GUARD = '            {scopeSource === "workbook" && (';

const EB_SHEET_SET_LINE =
  'const EB_SYSTEM_SHEETS = new Set(["#EB_RULES", "#EB_RULES_BACKUP"]);';

const EXISTENCE_KEYS_FN = [
  "export function folderSheetKeysOf(catalog: FolderScopeCatalog): string[] {",
  "  return catalog.files.flatMap((file) =>",
  "    file.worksheets.map((sheet) => folderSheetKey(file.id, sheet.name))",
  "  );",
  "}"
].join(LF);

/**
 * W8：滚动容器的收口（`.scope-popover-scroll` 那一行 `</div>`）与工作目录页脚之间。
 * W8 把**这一行 `</div>` 往后挪到工作簿动作条之前**，于是工作目录页脚与承认汇总被卷进滚动容器，
 * 而工作簿动作条仍在容器外。
 *
 * 为什么要这条：`wpsScopeDraft.test.tsx` 里那条层次判据原先只写了**更严的那半句**
 * （`actions.previousElementSibling === scroll`），它在工作簿来源下才成立；
 * 工作目录来源下的固定页脚（`.folder-selection-footer`）需要单独一条判据。
 * E2（`mutate_scope_action_bar_layer.mjs`）会把动作条一起卷进去，所以**总是**先触发
 * 工作簿那半句，工作目录那半句永远轮不到 —— 那条判据没有被单独证明过。
 *
 * 2026-10-16 晚重定点：工作目录模式按 Excel 的格式重排（固定页脚 = 计数 + 确认同排，
 * 承认汇总在页脚下方，替换掉了原来的 `.wps-scope-actions`），旧锚（
 * `)}` + `<div className="wps-scope-actions">`）因此失配。新锚仍取**同一个**判别点。
 */
const SCROLL_CLOSE_BEFORE_FOLDER_BLOCK = [
  "          )}",
  "          </div>",
  "          {/*",
  "            工作目录侧的固定页脚"
].join(LF);
const SCROLL_CLOSE_BEFORE_FOLDER_BLOCK_DROPPED = [
  "          )}",
  "          {/*",
  "            工作目录侧的固定页脚"
].join(LF);
const ADMISSION_BEFORE_WORKBOOK_ACTIONS = [
  "              count={draftFolderKeys.length} checking={admission.checking} />",
  "          )}",
  "          {/* 工作簿模式的动作条"
].join(LF);
const ADMISSION_BEFORE_WORKBOOK_ACTIONS_WITH_CLOSE = [
  "              count={draftFolderKeys.length} checking={admission.checking} />",
  "          )}",
  "          </div>",
  "          {/* 工作簿模式的动作条"
].join(LF);

// ── 变异目标 ──────────────────────────────────────────────────────────
const TARGETS = {
  W1: {
    note: "把「WPS 当前仅支持打开的工作簿」请回卡片副标题（能力说明复活）",
    edits: [
      {
        file: PANEL,
        from: FOLDER_CAPTION_LINE,
        to: "                <span><strong>工作目录</strong><small>WPS 当前仅支持打开的工作簿</small></span>"
      }
    ]
  },
  W2: {
    note: "勾选即提交授权（草稿 / 已提交的分离被拆掉）",
    edits: [
      {
        file: PANEL,
        from: TOGGLE_FOLDER_SHEET,
        to: [
          "  const toggleFolderSheet = (fileId: string, sheetName: string) => {",
          "    const key = folderSheetKey(fileId, sheetName);",
          "    setDraftFolderKeys((current) => {",
          "      const next = current.includes(key)",
          "        ? current.filter((item) => item !== key)",
          "        : [...current, key];",
          "      const root = folderCatalog ? folderRootIdentity(folderCatalog.folderPath) : null;",
          "      if (root) setFolderScope({ root, keys: [...next] });",
          "      return next;",
          "    });",
          "  };"
        ].join(LF)
      }
    ]
  },
  W3: {
    note: "刷新被当成同意（目标消失也不再作废授权）",
    edits: [{ file: PANEL, from: PRESERVE_LINE, to: "      // MUTANT: 刷新即同意，不再作废授权。" }]
  },
  W4: {
    note: "换目录不作废旧授权",
    edits: [
      {
        file: PANEL,
        from: NEW_ROOT_BLOCK,
        to: ["      if (!sameRoot) {", "        setDraftFolderKeys([]);", "      }"].join(LF)
      }
    ]
  },
  W5: {
    note: "EB 系统遗留表也算可勾选",
    edits: [{ file: CORE, from: EB_SHEET_SET_LINE, to: "const EB_SYSTEM_SHEETS = new Set<string>([]);" }]
  },
  W6: {
    note: "发送闸读不到来源（等于放行 → 偷用当前工作簿）",
    edits: [
      {
        file: AGENT,
        from: SEND_GUARD_SOURCE_LINE,
        to: ["    const folderRefusal = workDirectorySendRefusal({", '      sourceMode: "workbook",'].join(LF)
      }
    ]
  },
  W7: {
    note: "触发器在工作目录模式下顺带显示当前工作簿名",
    edits: [{ file: PANEL, from: TRIGGER_SMALL_GUARD, to: "            {true && (" }]
  },
  C1: {
    note: "存在性名单也过滤 EB 表（两个判据合成一个）",
    edits: [
      {
        file: CORE,
        from: EXISTENCE_KEYS_FN,
        to: [
          "export function folderSheetKeysOf(catalog: FolderScopeCatalog): string[] {",
          "  return catalog.files.flatMap((file) =>",
          "    file.worksheets",
          "      .filter((sheet) => !isEBSystemSheet(sheet.name))",
          "      .map((sheet) => folderSheetKey(file.id, sheet.name))",
          "  );",
          "}"
        ].join(LF)
      }
    ]
  },
  W8: {
    note: "工作目录页脚与承认汇总被卷进滚动容器里（工作簿动作条仍在容器外）",
    edits: [
      {
        file: PANEL,
        from: SCROLL_CLOSE_BEFORE_FOLDER_BLOCK,
        to: SCROLL_CLOSE_BEFORE_FOLDER_BLOCK_DROPPED
      },
      {
        file: PANEL,
        from: ADMISSION_BEFORE_WORKBOOK_ACTIONS,
        to: ADMISSION_BEFORE_WORKBOOK_ACTIONS_WITH_CLOSE
      }
    ]
  }
};

const ALL_FILES = [...new Set(Object.values(TARGETS).flatMap((t) => t.edits.map((e) => e.file)))];

/** 目标文件里有的带 CRLF，先归一化，否则 includes 会误报 ALREADY-CLEAN。 */
function normalize(text) {
  return text.split(/\r\n/).join(LF);
}
/** 写回必须还原**那个文件自己的**换行，否则"逐字节一致"这条判据会自己失败。 */
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

/** 一个变异要读写的全部文件（去重）。 */
function filesOf(target) {
  return [...new Set(target.edits.map((edit) => edit.file))];
}

function readAll(files) {
  const map = new Map();
  for (const file of files) {
    const original = readFileSync(file, "utf8");
    map.set(file, { original, text: normalize(original) });
  }
  return map;
}

/** `muted` 优先：插入/扩写型变异的 `from` 串常是 `to` 串的片段，先判 `unmuted` 会读反。 */
function state(target, map) {
  const muted = target.edits.every((edit) => hits(map.get(edit.file).text, edit.to) === 1);
  if (muted) return { muted: true, unmuted: false, detail: "MUTATED" };
  const unmuted = target.edits.every((edit) => hits(map.get(edit.file).text, edit.from) === 1);
  return { muted: false, unmuted, detail: unmuted ? "SEED" : "BROKEN" };
}

const which = process.argv[2];
if (!which) {
  for (const file of ALL_FILES) console.log(`BASELINE ${file} sha256=${sha256(file)}`);
  process.exit(0);
}

/** 前置状态闸：每个变异要么完整在位、要么完整缺失；源码一改就必须先跑它。 */
if (which === "check") {
  const map = readAll(ALL_FILES);
  let bad = 0;
  for (const [id, target] of Object.entries(TARGETS)) {
    const verdict = state(target, map);
    if (verdict.detail !== "BROKEN") {
      console.log(`OK   ${id} ${verdict.detail}`);
    } else {
      bad += 1;
      const counts = target.edits
        .map(
          (edit) =>
            `${edit.file.split(/[\\/]/).pop()} from=${hits(map.get(edit.file).text, edit.from)}` +
            ` to=${hits(map.get(edit.file).text, edit.to)}`
        )
        .join("  ");
      console.log(`MISS ${id}: ${counts}（期望每步 from=1,to=0 或 to=1）`);
    }
  }
  console.log(bad === 0 ? "check: OK" : `check: ${bad} 个变异串已失配`);
  process.exit(bad === 0 ? 0 : 3);
}

if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_wps_work_directory.mjs [W1..W7 C1] [restore] | check");
  process.exit(2);
}

const target = TARGETS[which];
const files = filesOf(target);
const map = readAll(files);

/** 每个文件都必须自己处于"整体 SEED"或"整体 MUTATED"；半截子状态一律拒绝。 */
function perFileState(edit) {
  const text = map.get(edit.file).text;
  const muted = hits(text, edit.to) === 1;
  const unmuted = hits(text, edit.from) === 1;
  return { muted, unmuted };
}

if (process.argv[3] === "restore") {
  const verdict = state(target, map);
  if (verdict.unmuted) {
    console.log("ALREADY-CLEAN");
    process.exit(0);
  }
  if (!verdict.muted) {
    console.error(`REFUSE: ${which} 处于半截子状态（既非 SEED 也非 MUTANT），拒绝还原`);
    process.exit(4);
  }
  for (const edit of target.edits) {
    const { muted } = perFileState(edit);
    if (!muted) {
      console.error(`REFUSE: ${which} 还原步骤在 ${edit.file} 命中 0 处（期望 1）`);
      process.exit(4);
    }
    const entry = map.get(edit.file);
    entry.text = entry.text.replace(edit.to, edit.from);
  }
  for (const [file, entry] of map) writeFileSync(file, denormalize(entry.text, entry.original), "utf8");
  console.log(`restore: OK ${files.map((file) => file.split(/[\\/]/).pop()).join(", ")}`);
  process.exit(0);
}

const before = state(target, map);
if (before.muted) {
  console.log("ALREADY-MUTATED");
  process.exit(0);
}
if (!before.unmuted) {
  console.error(`REFUSE: ${which} 的种子已失配（SEED 命中数不为 1），先跑 check`);
  process.exit(3);
}
for (const edit of target.edits) {
  const { unmuted } = perFileState(edit);
  if (!unmuted) {
    console.error(`REFUSE: ${which} 施加步骤在 ${edit.file} 命中 0 处（期望 1）`);
    process.exit(4);
  }
  const entry = map.get(edit.file);
  entry.text = entry.text.replace(edit.from, edit.to);
}
for (const [file, entry] of map) writeFileSync(file, denormalize(entry.text, entry.original), "utf8");
console.log(
  `APPLIED ${which} → ${files.map((file) => file.split(/[\\/]/).pop()).join(", ")}（${target.note}）`
);
