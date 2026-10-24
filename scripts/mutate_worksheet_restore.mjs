// 反向变异脚本：删表「在原地撤销」能力的判据判别力（2026-10-20 第四轮）。
//
// 用法：
//   node scripts/mutate_worksheet_restore.mjs <SEED>            # 应用变异
//   node scripts/mutate_worksheet_restore.mjs <SEED> restore    # 还原
//   node scripts/mutate_worksheet_restore.mjs restore-all       # 全部还原
//
// SEED 一览（括号里是"必须变红"的用例）：
//   R1  位置一律追加到末尾        （server: test_restore_puts_sheets_back_in_place_with_objects_and_visibility）
//   R2  不再按归档原文修 #REF!    （server: test_restore_repairs_ref_formulas_and_workbook_names）
//   R3  不再修工作簿级名称定义    （server: test_restore_repairs_ref_formulas_and_workbook_names）
//   R4  同名表不再算冲突          （server: test_restore_refuses_when_same_name_sheet_exists）
//   R5  撤销范围核对放过"多出来的表"（core:   refuses sheets that this restore cannot explain）
//   R6  目录扩张一律认下来        （excel: worksheet restore target rebinding）
//   R7  撤销不再防重复点击        （excel: a double click cannot dispatch a second restore）
//   R8  撤销回执不再要求覆盖全部删除表（excel: worksheet restore receipt）
//
// 铁律：脚本只做 apply / restore，测试由 bash 直调；SEED 一律按 LF 写，写回时按文件原换行符；
//       还原后必须逐字节与备份一致（见 run_worksheet_restore_mutations.mjs）。
import { readFileSync, writeFileSync } from "node:fs";

const REPO = "E:\\codex\\excel-bro";

const TARGETS = {
  // R1：位置一律追加到末尾 —— "放回原位置"退化成"塞到最后"
  R1: {
    file: `${REPO}\\server\\app\\worksheet_restore.py`,
    seed: [
      "    index = list(archive_order).index(name)",
      "    for previous in reversed(archive_order[:index]):",
      "        if previous in current:",
      '            return ("after", previous)',
      "    for following in archive_order[index + 1:]:",
      "        if following in current:",
      '            return ("before", following)',
      '    return ("end", None)',
    ].join("\n"),
    mutant: [
      "    # [变异R1] 位置一律追加到末尾",
      "    _ = (archive_order, name, current)",
      '    return ("end", None)',
    ].join("\n"),
  },
  // R2：不再按归档原文修被删除打断的跨表公式
  R2: {
    file: `${REPO}\\server\\app\\worksheet_restore.py`,
    seed: [
      "                if _repair_formula_cell(live_sheet, top + row_index, left + column_index, formula):",
      "                    repaired += 1",
    ].join("\n"),
    mutant: [
      "                # [变异R2] 不再按归档原文修 #REF!",
      "                if False:",
      "                    repaired += 1",
    ].join("\n"),
  },
  // R3：不再修工作簿级名称定义（`=SUM(Sheet7!A1)` → `=#REF!` 保持原样）
  R3: {
    file: `${REPO}\\server\\app\\worksheet_restore.py`,
    seed: ["        if \"#REF!\" in current and current != expected:", "            try:",
      "                live.RefersTo = expected", "                repaired += 1"].join("\n"),
    mutant: ["        # [变异R3] 不再修工作簿级名称定义",
      "        if False and \"#REF!\" in current and current != expected:", "            try:",
      "                live.RefersTo = expected", "                repaired += 1"].join("\n"),
  },
  // R4：同名表不再算冲突 —— 会覆盖用户后来新建的同名表
  R4: {
    file: `${REPO}\\server\\app\\worksheet_restore.py`,
    seed: ["    conflicts = [name for name in pending if name in live]", "    if conflicts:",
      "        raise RestoreConflict("].join("\n"),
    mutant: ["    conflicts = [name for name in pending if name in live]", "    # [变异R4] 同名不再当冲突",
      "    if False and conflicts:", "        raise RestoreConflict("].join("\n"),
  },
  // R5：共享判据放过"这次撤销解释不了的新增表"
  R5: {
    file: `${REPO}\\packages\\core\\src\\worksheetRestore.ts`,
    seed: [
      "  if (actual.some((name) => !authorizedSet.has(name) && !restoredSet.has(name))) {",
      '    return { ok: false, reason: "unexpected_sheets" };',
      "  }",
    ].join("\n"),
    mutant: "  // [变异R5] 目录里多出来的表一律放过",
  },
  // R6：守卫不再检查"目录里多出来的表是不是这次撤销放回的"
  R6: {
    file: `${REPO}\\apps\\excel-addin\\src\\hooks\\useAuthorizedTarget.ts`,
    seed:
      "    if (target.sheets.some((sheet) => !knownBefore.has(sheet.name) && !restored.has(sheet.name))) return false;",
    mutant:
      "    // [变异R6] 不再检查多出来的表是否来自本次撤销",
  },
  // R7：撤销不再防重复点击（在途期间能再发一次，可能重复造表）
  R7: {
    file: `${REPO}\\apps\\excel-addin\\src\\hooks\\useWorksheetCheckpoint.ts`,
    seed: [
      "  async function restore(operationId: string): Promise<WorksheetRestoreOutcome> {",
      '    if (inFlight.current) return { ok: false, error: "上一项检查点操作尚未结束。" };',
    ].join("\n"),
    mutant: [
      "  async function restore(operationId: string): Promise<WorksheetRestoreOutcome> {",
      "    // [变异R7] 不再防重复点击",
    ].join("\n"),
  },
  // R8：撤销回执不再要求"覆盖本次删除的全部表"
  R8: {
    file: `${REPO}\\apps\\excel-addin\\src\\contracts.ts`,
    seed:
      "  const covered = deletedSheets.every(name => item.restoredSheets!.includes(name));",
    mutant:
      "  const covered = deletedSheets.every(name => item.restoredSheets!.includes(name)) || true; // [变异R8] 覆盖关系不生效",
  },
};

function fileNewline(text) {
  return text.includes("\r\n") ? "\r\n" : "\n";
}
function toFileNewline(text, newline) {
  const lf = text.replaceAll("\r\n", "\n");
  return newline === "\n" ? lf : lf.replaceAll("\n", newline);
}

const which = process.argv[2];

if (which === "restore-all") {
  let restored = 0;
  for (const key of Object.keys(TARGETS)) {
    const entry = TARGETS[key];
    const text = readFileSync(entry.file, "utf8");
    const normalized = text.replaceAll("\r\n", "\n");
    const mutant = entry.mutant.replaceAll("\r\n", "\n");
    if (!normalized.includes(mutant)) { console.log(`clean: ${key}`); continue; }
    writeFileSync(entry.file,
      toFileNewline(normalized.split(mutant).join(entry.seed.replaceAll("\r\n", "\n")), fileNewline(text)),
      "utf8");
    restored += 1;
    console.log(`restored: ${key}`);
  }
  process.exit(restored === 0 ? 0 : 0);
}

if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_worksheet_restore.mjs R1..R8|restore-all [restore]");
  process.exit(2);
}

const target = TARGETS[which];
const original = readFileSync(target.file, "utf8");
const newline = fileNewline(original);
const normalized = original.replaceAll("\r\n", "\n");
const seed = target.seed.replaceAll("\r\n", "\n");
const mutant = target.mutant.replaceAll("\r\n", "\n");

if (process.argv[3] === "restore") {
  if (!normalized.includes(mutant)) {
    console.log("ALREADY-CLEAN：" + target.file);
    process.exit(0);
  }
  writeFileSync(target.file, toFileNewline(normalized.split(mutant).join(seed), newline), "utf8");
  console.log("RESTORED：" + target.file);
  process.exit(0);
}

const occurrences = normalized.split(seed).length - 1;
if (occurrences !== 1) {
  console.log(`SKIP: SEED 出现 ${occurrences} 次（期望 1）→ ${which}`);
  process.exit(3);
}
const applied = toFileNewline(normalized.replace(seed, mutant), newline);
if (applied === original) {
  console.log("SKIP: 替换未生效（换行符或占位符不匹配）。");
  process.exit(3);
}
writeFileSync(target.file, applied, "utf8");
console.log(`APPLIED ${which} → ${target.file}`);
