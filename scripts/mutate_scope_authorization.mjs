// 反向变异脚本：数据范围「授权同步」判据的判别力（2026-10-19 第五轮）。
//
// 用法：
//   node scripts/mutate_scope_authorization.mjs <SEED>            # 应用变异
//   node scripts/mutate_scope_authorization.mjs <SEED> restore    # 还原
//   node scripts/mutate_scope_authorization.mjs restore-all       # 全部还原
//
// SEED 一览（括号里是"必须变红"的用例；每个 SEED 只落在**一处**判别点上）：
//   A1  守卫按**草稿口径**读当前范围   （excel: ⑨…后台不误报）
//   A2  「取消」不回退草稿口径         （excel: ④…已提交授权与守卫基线不变）
//   A3  重读宿主即改写已提交授权       （excel hook: 重读宿主只刷新草稿）
//   A4  旧读数照旧用来比新基线         （excel hook: 旧读数不得判成范围变化）
//   A5  采用当前工作簿一律按 manual    （excel: ⑧…跟随当前口径）
//   A6  口径差异不算草稿差异           （excel: ⑨…发送也不静默改读）
//
// 铁律：脚本只做 apply / restore，测试由 runner 直调；SEED 一律按 LF 写，
//       写回时按文件原换行符；还原后必须逐字节与备份一致（见 runner）。
import { readFileSync, writeFileSync } from "node:fs";

const REPO = "E:\\codex\\excel-bro";

const TARGETS = {
  // A1：守卫读的是弹层里的草稿口径 —— 用户只是点了「跟随当前」看一眼，
  //     后台检查就按活动表去比已提交的多表范围，弹「工作表范围发生变化」。
  A1: {
    file: `${REPO}\\apps\\excel-addin\\src\\App.tsx`,
    seed: [
      "    inspect: () => inspectOfficeAuthorizedTarget(",
      '      committedScopeMode === "auto" ? null : committedSheetNames',
      "    ),",
    ].join("\n"),
    mutant: [
      "    inspect: () => inspectOfficeAuthorizedTarget(",
      '      workbookScopeMode === "auto" ? null : committedSheetNames // [变异A1] 按草稿口径读',
      "    ),",
    ].join("\n"),
  },
  // A2：「取消」只回退名字、不回退口径 —— 草稿口径留在原地，触发器与守卫
  //     都改按"跟随当前"说话，已提交的多表范围被当成"变了"。
  A2: {
    file: `${REPO}\\apps\\excel-addin\\src\\hooks\\useScopeSelection.ts`,
    seed: "    setWorkbookScopeMode(committedScopeMode);",
    mutant: "    // [变异A2] 取消不回退口径",
  },
  // A3：重读宿主（scan / 刷新）直接把"目录里还剩什么"抄成已提交授权 ——
  //     后台读一次就换掉用户的授权，紧接着的请求被判成"范围变化"。
  A3: {
    file: `${REPO}\\apps\\excel-addin\\src\\hooks\\useScopeSelection.ts`,
    seed: [
      "      if (!committedSeededRef.current && committedNamesRef.current.length === 0) {",
      "        committedSeededRef.current = true;",
      "        setCommittedSheetNames(fallback);",
      "        setCommittedScopeMode(workbookScopeMode);",
      "      }",
    ].join("\n"),
    mutant: [
      "      // [变异A3] 每次重读都改写已提交授权",
      "      committedSeededRef.current = true;",
      "      setCommittedSheetNames(fallback);",
      "      setCommittedScopeMode(workbookScopeMode);",
    ].join("\n"),
  },
  // A4：去掉读数与判断之间的代数隔离 —— 读数在飞期间基线被推进（删表收尾 /
  //     原地撤销的直接 bind），旧读数回来后被拿去和新基线比。
  A4: {
    file: `${REPO}\\apps\\excel-addin\\src\\hooks\\useAuthorizedTarget.ts`,
    seed: "        if (bindGeneration.current !== generation) continue;",
    mutant: "        // [变异A4] 旧读数照旧用来比新基线",
  },
  // A5：采用当前工作簿时把口径写死成 manual —— 「跟随当前」的确认变成
  //     "固定这几张表"，之后切活动表要么误报、要么静默改读。
  A5: {
    file: `${REPO}\\apps\\excel-addin\\src\\hooks\\useScopeSelection.ts`,
    seed: [
      "      setWorkbookScopeMode(mode);",
      "      setSelectedSheetNames(sheetNames);",
      "      // 草稿与已提交**同批**落地：守卫在这一帧之后立刻就会读到它们。",
      "      commitAuthorizedScope(sheetNames, mode);",
    ].join("\n"),
    mutant: [
      '      setWorkbookScopeMode("manual"); // [变异A5] 口径写死成 manual',
      "      setSelectedSheetNames(sheetNames);",
      '      commitAuthorizedScope(sheetNames, "manual");',
    ].join("\n"),
  },
  // A6：草稿与已提交的差异只比名字、不比口径 —— 切到"跟随当前"而不确认，
  //     发送时按草稿口径去读活动表，读到已授权范围之外的表也没人拦。
  A6: {
    file: `${REPO}\\apps\\excel-addin\\src\\hooks\\useScopeSelection.ts`,
    seed: '    (sourceMode === "workbook" && workbookScopeMode !== committedScopeMode);',
    mutant: "    false; // [变异A6] 口径差异不算草稿差异",
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
  console.error("用法: node scripts/mutate_scope_authorization.mjs A1..A6|restore-all [restore]");
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
