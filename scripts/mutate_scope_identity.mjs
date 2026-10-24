// 反向变异脚本：授权集合身份「顺序无关」判据的判别力（2026-10-19 六续）。
//
// 用法：
//   node scripts/mutate_scope_identity.mjs <SEED>            # 应用变异
//   node scripts/mutate_scope_identity.mjs <SEED> restore    # 还原
//   node scripts/mutate_scope_identity.mjs restore-all       # 全部还原
//
// SEED 一览（括号里是"必须变红"的用例；每个 SEED 只落在**一处**判别点上）：
//   I1  Excel 证据键按点选顺序拼      （excel e2e: A.按非目录顺序确认三张表后…）
//   I2  WPS 绑定键按点选顺序拼       （wps: 同一组工作表换个点选顺序…）
//   I3  共享层规范形不排序           （core: 同一组 sheet id 换个顺序仍是同一范围）
//   I4  规范形顺手去重（Set）        （core: 顺序无关 ≠ 合并成员）
//
// 铁律：脚本只做 apply / restore，测试由 runner 直调；SEED 一律按 LF 写，
//       写回时按文件原换行符；还原后必须逐字节与备份一致（见 runner）。
import { readFileSync, writeFileSync } from "node:fs";

const REPO = "E:\\codex\\excel-bro";

const TARGETS = {
  // I1：证据键回到"直接 stringify selected" —— 同一组表只因点选顺序不同就产出
  //     两个键，checkForAnalysis 判成"证据冲突"，请求卡在「暂时无法核实当前目标」。
  I1: {
    file: `${REPO}\\apps\\excel-addin\\src\\officeAuthorizedTarget.ts`,
    seed: '  return { key: authorizedTargetEvidenceKey(identity ?? "document-reference", target.selected), conclusive: true };',
    mutant: [
      "  // [变异I1] 键受点选顺序影响",
      '  return { key: JSON.stringify([identity ?? "document-reference", target.selected]), conclusive: true };',
    ].join("\n"),
  },
  // I2：WPS 侧同样的写法。两端各拼一份"看起来一样"的键，一端忘了规范化就对不上。
  I2: {
    file: `${REPO}\\apps\\wps-addin\\src\\wpsClarificationTargetConfirmation.ts`,
    seed: [
      "  return authorizedTargetEvidenceKey(",
      "    JSON.stringify([target.workbookFullName, target.workbookPath, target.workbookName]),",
      "    target.sheets.map(({ index, name, id }) => JSON.stringify([index, name, id ?? null]))",
      "  );",
    ].join("\n"),
    mutant: [
      "  return JSON.stringify([target.workbookFullName, target.workbookPath, target.workbookName,",
      "    target.sheets.map(({ index, name, id }) => [index, name, id ?? null])]); // [变异I2] 键受点选顺序影响",
    ].join("\n"),
  },
  // I3：规范形退化成"原样返回" —— 顺序重新成为身份的一部分。
  I3: {
    file: `${REPO}\\packages\\core\\src\\authorizedTarget.ts`,
    seed: [
      "export function canonicalTargetSelection(selected: readonly string[]): string[] {",
      "  return [...selected].sort();",
      "}",
    ].join("\n"),
    mutant: [
      "export function canonicalTargetSelection(selected: readonly string[]): string[] {",
      "  return [...selected]; // [变异I3] 不排序：顺序重新成为身份的一部分",
      "}",
    ].join("\n"),
  },
  // I4：规范化顺手去重 —— 重复 ID 被"排序无关"一起吞掉，缺失的授权也查不出来。
  I4: {
    file: `${REPO}\\packages\\core\\src\\authorizedTarget.ts`,
    seed: [
      "export function canonicalTargetSelection(selected: readonly string[]): string[] {",
      "  return [...selected].sort();",
      "}",
    ].join("\n"),
    mutant: [
      "export function canonicalTargetSelection(selected: readonly string[]): string[] {",
      "  return [...new Set(selected)].sort(); // [变异I4] 去重：重复 ID 不再是差异",
      "}",
    ].join("\n"),
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
  for (const key of Object.keys(TARGETS)) {
    const entry = TARGETS[key];
    const text = readFileSync(entry.file, "utf8");
    const normalized = text.replaceAll("\r\n", "\n");
    const mutant = entry.mutant.replaceAll("\r\n", "\n");
    if (!normalized.includes(mutant)) { console.log(`clean: ${key}`); continue; }
    writeFileSync(entry.file,
      toFileNewline(normalized.split(mutant).join(entry.seed.replaceAll("\r\n", "\n")), fileNewline(text)),
      "utf8");
    console.log(`restored: ${key}`);
  }
  process.exit(0);
}

if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_scope_identity.mjs I1..I4|restore-all [restore]");
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
