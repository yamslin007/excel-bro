// 反向变异脚本：通用合成案例（要求5）的关键守卫。
//
// 用法：node scripts/mutate_synthetic_case_guards.mjs S1..S4|restore-all [restore]
//
// S1 = `compilePythonValuesPlan` 不再要求"目标表必须在声明来源里"
//      （结果矩阵就能往**未声明**的表上写 = 白名单编译形同虚设）
// S2 = `stagePythonResultPreview` 不再核对"来源快照与 scope.sources 一致"
//      （来源被换掉也能生成预览 = 结果据以算出的输入不再可信）
// S3 = Excel `prepareOfficePythonResult.assertCurrent` 不再在写前重读来源
//      （预览到确认之间来源被改，照样放行）
// S4 = `recordsFromSheet` 缺字段时不再"整张跳过 + 警告"，而是继续返回记录
//      （"少读一张表"变成静默算出的错答案，且 complete 不再为 false）
//
// 铁律：脚本只做 apply/restore；测试由 bash 直调；一次只 apply 一个 SEED。
import { readFileSync, writeFileSync } from "node:fs";

const REPO = "E:\\codex\\excel-bro";
const CORE_RESULT = `${REPO}\\packages\\core\\src\\pythonResult.ts`;
const CORE_DATA = `${REPO}\\packages\\core\\src\\pythonData.ts`;
const OFFICE_RESULT = `${REPO}\\apps\\excel-addin\\src\\officePythonResult.ts`;
const DATA_TOOLS = `${REPO}\\apps\\excel-addin\\src\\dataTools.ts`;

/** 目标文件的换行符（CRLF/LF）——apply/restore 都要先归一化再换回。 */
function fileNewline(text) {
  return text.includes("\r\n") ? "\r\n" : "\n";
}
/** 把 SEED 文本（可能含 `<<>>` 占位）从占位形式换回 `{`。 */
function unplace(text) {
  return text.replaceAll("<<>>", "{");
}
/** SEED 文本按 LF 书写；写入前按目标文件换行符转换。 */
function toFileNewline(text, newline) {
  const lf = text.replace(/\r\n/g, "\n");
  return newline === "\r\n" ? lf.replace(/\n/g, "\r\n") : lf;
}

const TARGETS = {
  // S1：白名单编译不再要求目标表在声明来源里
  S1: {
    file: CORE_RESULT,
    seed: [
      "  if (!scope.sources.some((source) => source.sheetId === target.sheetId && source.sheetName === target.sheetName) ||",
      "      !rect || rect.rows !== receipt.rows || rect.columns !== receipt.columns ||",
      "      target.range !== target.range.toUpperCase() || !/^[A-Z]{1,3}[1-9][0-9]{0,6}(?::[A-Z]{1,3}[1-9][0-9]{0,6})?$/.test(target.range)) throw new PythonResultUnavailable();",
    ].join("\n"),
    mutant: [
      "  // [变异S1] 不再要求目标表在声明来源里",
      "  if (!rect || rect.rows !== receipt.rows || rect.columns !== receipt.columns ||",
      "      target.range !== target.range.toUpperCase() || !/^[A-Z]{1,3}[1-9][0-9]{0,6}(?::[A-Z]{1,3}[1-9][0-9]{0,6})?$/.test(target.range)) throw new PythonResultUnavailable();",
    ].join("\n")
  },

  // S2：预览不再核对来源快照与 scope.sources 一致
  S2: {
    file: CORE_RESULT,
    seed: [
      "  if (JSON.stringify(sources.map((snapshot) => snapshot.source)) !== JSON.stringify(scope.sources)) throw new PythonResultUnavailable();",
    ].join("\n"),
    mutant: "  // [变异S2] 不再核对来源快照与 scope 一致"
  },

  // S3：Excel 写前复核不再重读来源
  //     —— 注意：`captureSources()` 自己也会因"来源值变化"抛错，所以这一道是
  //     **冗余层**。单独拆没判别力（GREEN(NG)），必须与 S3b 同时拆才算一次完整变异。
  S3: {
    file: OFFICE_RESULT,
    seed: [
      "    const current = await session.captureSources();",
      "    const expected = preview.prepared.sources;",
      "    if (current.length !== expected.length ||",
      "        current.some((snapshot, index) =>",
      "          snapshot.source.sheetId !== expected[index].source.sheetId ||",
      "          snapshot.source.sheetName !== expected[index].source.sheetName ||",
      "          snapshot.source.range !== expected[index].source.range ||",
      "          JSON.stringify(snapshot.values) !== JSON.stringify(expected[index].values)",
      "        )) throw new PythonResultUnavailable();",
    ].join("\n"),
    mutant: "    // [变异S3] 不再在写前重读来源"
  },

  // S3b：拆掉 `captureSources()` 内部"来源值变化即抛错"这一层（与 S3 合拆）
  S3b: {
    file: CORE_DATA,
    seed: [
      "        for (let index = 0; index < sources.length; index++) {",
      "          const serialized = snapshot(await host.read(sources[index]), sources[index]);",
      "          if (await hash(serialized) !== localFingerprints[index]) throw new PythonDataUnavailable();",
      "          await check();",
      "          snapshots.push({source:{...sources[index]},values:JSON.parse(serialized) as CellValue[][]});",
      "        }",
    ].join("\n"),
    mutant: [
      "        for (let index = 0; index < sources.length; index++) {",
      "          const serialized = snapshot(await host.read(sources[index]), sources[index]);",
      "          // [变异S3b] 不再核对来源值是否变化",
      "          await check();",
      "          snapshots.push({source:{...sources[index]},values:JSON.parse(serialized) as CellValue[][]});",
      "        }",
    ].join("\n")
  },

  // S4：缺字段的表不再整张跳过
  S4: {
    file: DATA_TOOLS,
    seed: [
      "  if (missing.length > 0) {",
      "    warnings.push(",
      "      `「${sheet.name}」未找到字段：${[...new Set(missing)].join(\"、\")}`",
      "    );",
      "    return;",
      "  }",
    ].join("\n"),
    mutant: [
      "  if (missing.length > 0) {",
      "    // [变异S4] 缺字段也继续返回记录（不整张跳过）",
      "    warnings.push(",
      "      `「${sheet.name}」未找到字段：${[...new Set(missing)].join(\"、\")}`",
      "    );",
      "  }",
    ].join("\n")
  }
};

const which = process.argv[2];
const restoreMode = process.argv[3] === "restore";

if (which === "restore-all") {
  for (const id of Object.keys(TARGETS)) {
    applyOrRestore(id, true);
  }
  console.log("restore-all 完成");
  process.exit(0);
}

function applyOrRestore(id, restore) {
  const target = TARGETS[id];
  if (!target) throw new Error(`未知变异：${id}`);
  const raw = readFileSync(target.file, "utf8");
  const newline = fileNewline(raw);
  const seed = toFileNewline(unplace(target.seed), newline);
  const mutant = toFileNewline(unplace(target.mutant), newline);
  if (restore) {
    const normalized = raw.replace(/\r\n/g, "\n");
    const mutantLf = mutant.replace(/\r\n/g, "\n");
    if (!normalized.includes(mutantLf)) {
      console.log(`${id} restore: ALREADY-CLEAN`);
      return;
    }
    const restored = normalized.replace(mutantLf, seed.replace(/\r\n/g, "\n"));
    writeFileSync(target.file, toFileNewline(restored, newline));
    console.log(`${id} restore: OK`);
    return;
  }
  const normalized = raw.replace(/\r\n/g, "\n");
  const seedLf = seed.replace(/\r\n/g, "\n");
  if (!normalized.includes(seedLf)) {
    throw new Error(`${id} SEED 未命中（先 od -c 核对换行符与缩进）`);
  }
  const mutated = normalized.replace(seedLf, mutant.replace(/\r\n/g, "\n"));
  writeFileSync(target.file, toFileNewline(mutated, newline));
  console.log(`${id} apply: OK`);
}

applyOrRestore(which, restoreMode);
