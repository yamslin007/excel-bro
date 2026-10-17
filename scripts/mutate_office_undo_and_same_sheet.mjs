// 反向变异脚本：Excel 写侧撤销闸与「同表不重叠」的恢复支持判据（2026-10-10 D 要求4）
// 用法：node scripts/mutate_office_undo_and_same_sheet.mjs M1..M6b|restore-all [restore]
// M1  = 撤销前不再核对执行后状态（直接覆盖用户的后续编辑）
// M2  = 撤销前不再要求 postStateCoverage === "complete"（界面失去"撤销需核对"判据）
// M3  = officeRecoverySupport 不再看来源依赖（同表也放行断点恢复 = 放松恢复闸）
// M4  = WPS prepareWpsPlan 不再按快照值核对来源矩阵
// M5  = WPS 执行写入前不再逐区域核对来源
// M6a/M6b = 一次拆全 WPS 结果预览的源变化**冗余层**（两道闸必须同时拆）
//
// 铁律：脚本只做 apply/restore；测试由 bash 直调。
import { readFileSync, writeFileSync } from "node:fs";

const REPO = "E:\\codex\\excel-bro";
const FILE = `${REPO}\\apps\\excel-addin\\src\\excel.ts`;
// SEED 一律按 LF 书写；文件若是 CRLF，apply/restore 时再换回（见 apply/restore）。
const NL = "\n";

const TARGETS = {
  // M1：拆掉"逐区域核对执行后指纹"这一道闸（用户编辑会被静默覆盖）
  M1: {
    seed: [
      '    if (loadedRangePostFingerprint(range) !== saved.postFingerprint) {',
      '      throw new Error(',
      '        `「${saved.sheet}」${saved.range} 在执行后又发生了变化，已停止撤销以免覆盖后续编辑。`',
      '      );',
      '    }',
    ].join(NL),
    mutant: '    // [变异M1] 不再核对执行后状态',
  },
  // M2：拆掉覆盖度前置闸（"部分覆盖"的撤销数据也会被拿去直接覆盖）
  M2: {
    seed: [
      '  if (snapshot.postStateCoverage !== "complete") {',
      '    throw new Error(',
      '      "这份撤销数据缺少完整的执行后状态校验，已停止自动覆盖；请先人工核对当前工作簿。"',
      '    );',
      '  }',
    ].join(NL),
    mutant: '  // [变异M2] 不再要求执行后状态完整核对',
  },
  // M3：officeRecoverySupport 忽略来源依赖 → "同表不重叠"被当成"支持断点恢复"
  M3: {
    seed: [
      '  const declaredDependency = plan.sourceDependencySheets?.length',
      '    ? plan.sourceDependencySheets',
      '    : null;',
      '  const sourceSheets = new Set(',
      '    (declaredDependency ?? plan.sourceFingerprintSheets ?? []).map((name) =>',
      '      name.toLocaleLowerCase()',
      '    )',
      '  );',
    ].join(NL),
    mutant: [
      '  const declaredDependency = plan.sourceDependencySheets?.length',
      '    ? plan.sourceDependencySheets',
      '    : null;',
      '  // [变异M3] 不再按来源依赖判重叠（同表也放行断点恢复）',
      '  const sourceSheets = new Set<string>();',
    ].join(NL),
  },
  // M4：prepareWpsPlan 不再核对源区域是否在计算后变化（同表/异表都不再看来源）
  M4: {
    file: `${REPO}\\apps\\wps-addin\\src\\wpsPlanExecutor.ts`,
    nl: "\n",
    /*
     * SEED 里要出现源码的模板字面量 `${...}`，但直接写在 JS 字符串里会被本脚本
     * 自己插值。开括号用占位串 `<<>>` 顶替，apply/restore 时再换回。
     */
    seed: [
      "    const current = readMatrices(rangeFor(sheet, snapshot.range), rowCount, columnCount);",
      "    const emptyFormulas = blankCellMatrix(rowCount, columnCount);",
      "    if (matrixFingerprintV2(current.values, emptyFormulas) !== matrixFingerprintV2(snapshot.values, emptyFormulas)) {",
      "      throw new Error(`源区域「$<<>>snapshot.sheetName}!$<<>>snapshot.range}」在计算后发生变化，请重新生成计划。`);",
      "    }",
    ].join("\n"),
    mutant: "    // [变异M4] 不再核对来源区域是否在计算后变化\n    const current = readMatrices(rangeFor(sheet, snapshot.range), rowCount, columnCount);",
  },
  // M5：WPS 执行写入前不再逐区域核对来源（"同表不重叠"被当成"来源不用看"）
  M5: {
    file: `${REPO}\\apps\\wps-addin\\src\\wpsPlanExecutor.ts`,
    nl: "\n",
    seed: [
      "    if (current.fingerprint !== source.previewFingerprint) {",
      "      throw new Error(`源区域「$<<>>source.sheetName}!$<<>>source.range}」在预览后发生变化，请重新预览并确认。`);",
      "    }",
    ].join("\n"),
    mutant: "    // [变异M5] 写入前不再逐区域核对来源",
  },
  /*
   * M6：**一次拆全** WPS 结果预览的源变化**冗余层**。
   *
   * "来源变了 → 预览失效"由两道闸共保：① `prepareWpsPlan` 按快照值比对来源矩阵
   * （= M4）；② `prepareWpsPythonResult.assertCurrent` 比对两次 `prepareWpsPlan`
   * 的 `sources[].previewFingerprint`。单拆任一道另一道仍然抛错 ⇒ GREEN(NG)，
   * 那是**判别点选错**而非实现缺陷。所以 M4 必须与 M6b **同时**拆掉才算一次完整变异。
   */
  M6b: {
    file: `${REPO}\\apps\\wps-addin\\src\\wpsPythonResult.ts`,
    nl: "\n",
    seed: [
      "    if (current.actions.some((action,index) => action.previewFingerprint !== preview.prepared.actions[index].previewFingerprint) ||",
      "        current.sources.some((source,index) => source.previewFingerprint !== preview.prepared.sources[index].previewFingerprint)) throw new PythonResultUnavailable();",
    ].join("\n"),
    mutant: "    // [变异M6b] 不再比对重算后的来源/动作指纹",
  }
};

/** 占位串 `<<>>` 换回 `{`（见 M4 注释）。 */
function unplace(text) {
  return text.replace(/<<>>/g, "{");
}

const which = process.argv[2];

/**
 * 文件实际用的换行符以**文件本身**为准（仓库里同一目录下 CRLF/LF 混用）。
 * SEED 一律按 LF 书写；写回时再按文件原本的换行符还原 —— 用错换行符
 * `replace` 会静默命中 0 次却打印 APPLIED，把变异留在源码里。
 */
function fileNewline(text) {
  return text.includes("\r\n") ? "\r\n" : "\n";
}

function seedText(entry) {
  return unplace(entry.seed.replaceAll("\r\n", "\n"));
}

/**
 * 把文本按目标换行符写回。
 *
 * **先把已有 CRLF 归一成 LF 再统一转换** —— 否则 CRLF 文件里再做一次
 * `replaceAll("\n","\r\n")` 会产出 `\r\r\n`，把文件改坏（且全文件 diff 爆炸）。
 */
function toFileNewline(text, newline) {
  const lf = text.replaceAll("\r\n", "\n");
  return newline === "\n" ? lf : lf.replaceAll("\n", newline);
}

if (which === "restore-all") {
  for (const key of Object.keys(TARGETS)) {
    const entry = TARGETS[key];
    const f = entry.file ?? FILE;
    const text = readFileSync(f, "utf8");
    const newline = fileNewline(text);
    // 比较/替换一律在 LF 归一化文本上做：文件是 CRLF 时，拿 LF 拼的 MUTANT
    // 直接 `includes`/`split` 会永远不命中，变异就永远还原不掉。
    const normalizedText = text.replaceAll("\r\n", "\n");
    const mutant = entry.mutant.replaceAll("\r\n", "\n");
    if (!normalizedText.includes(mutant)) {
      console.log(`clean: ${key}`);
      continue;
    }
    writeFileSync(
      f,
      toFileNewline(
        normalizedText.split(mutant).join(seedText(entry)),
        newline
      ),
      "utf8"
    );
    console.log(`restored: ${key}`);
  }
  process.exit(0);
}
if (!TARGETS[which]) {
  console.error(
    "用法: node scripts/mutate_office_undo_and_same_sheet.mjs M1|M2|M3|M4|M5|M6b|restore-all [restore]"
  );
  process.exit(2);
}
const target = TARGETS[which];
const file = target.file ?? FILE;
const original = readFileSync(file, "utf8");
const newline = fileNewline(original);
const normalizedOriginal = original.replaceAll("\r\n", "\n");
const mutantText = target.mutant.replaceAll("\r\n", "\n");

if (process.argv[3] === "restore") {
  if (!normalizedOriginal.includes(mutantText)) {
    console.log("ALREADY-CLEAN：" + file);
    process.exit(0);
  }
  writeFileSync(
    file,
    toFileNewline(
      normalizedOriginal.split(mutantText).join(seedText(target)),
      newline
    ),
    "utf8"
  );
  console.log("RESTORED：" + file);
  process.exit(0);
}

const seed = seedText(target);
const occurrences = normalizedOriginal.split(seed).length - 1;
if (occurrences !== 1) {
  console.log(`SKIP: SEED 出现 ${occurrences} 次（期望 1）。`);
  process.exit(3);
}
const applied = toFileNewline(
  normalizedOriginal.replace(seed, mutantText),
  newline
);
if (applied === original) {
  console.log("SKIP: 替换未生效（换行符或占位符不匹配）。");
  process.exit(3);
}
writeFileSync(file, applied, "utf8");
console.log(`APPLIED ${which} → ${file}`);

