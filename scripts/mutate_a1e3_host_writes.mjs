// 反向变异脚本：A1-E3 第三批三个问题的**关键判据**判别力（2026-10-25 第三批）。
//
// 用法：
//   node scripts/mutate_a1e3_host_writes.mjs <SEED>            # 应用变异
//   node scripts/mutate_a1e3_host_writes.mjs <SEED> restore    # 还原
//   node scripts/mutate_a1e3_host_writes.mjs restore-all       # 全部还原
//
// SEED 一览（括号里是"必须变红"的用例；每个 SEED 只落在**一处**判别点上）：
//   H1  空上报当成 zero            （server: test_no_report_is_not_zero_writes）
//   H2  忽略边界外的上报           （server: test_report_outside_the_ledger_boundary_is_flagged_not_counted_as_zero）
//   H4  展示层照字面读 zero        （core:   矛盾的线上数据不许被升级）
//   H5  WPS 计划写不计数           （wps:    一个动作只写一块时…）
//   H6  WPS 辅助改格式算计划写     （wps:    文本保真用的 NumberFormat…）
//   H7  没身份就按名字认簿         （server: test_dev_fault_refuses_to_arm_without_a_workbook_identity）
//   H8  同名多身份取最近那个       （server: test_dev_fault_refuses_when_the_name_carries_two_identities）
//   H9  任务侧只报名字也照样绑     （server: test_dev_fault_does_not_bind_a_task_that_only_carries_a_name）
//
// 铁律：脚本只做 apply / restore，测试由 runner 直调；SEED 一律按 LF 写，
//       写回时按文件原换行符；还原后必须逐字节与备份一致（见 runner）。
import { readFileSync, writeFileSync } from "node:fs";

const REPO = "E:\\codex\\excel-bro";

const TARGETS = {
  // H1：没有上报时直接宣布"全覆盖且零调用" —— 把"没收到"读成"没发生"。
  H1: {
    file: `${REPO}\\server\\app\\main.py`,
    seed: [
      "    if not entries:",
      '        observation = "not_collected"',
      '        call_evidence = "none"',
    ].join("\n"),
    mutant: [
      "    if not entries:",
      '        observation = "complete"  # [变异H1] 空数组当成"全覆盖且零调用"',
      '        call_evidence = "zero"',
    ].join("\n"),
  },
  // H2：账本没认领过的动作却收到它的观测 —— 不再当成矛盾证据，反而落进 zero。
  H2: {
    file: `${REPO}\\server\\app\\main.py`,
    seed: "        elif missing or unexpected or undetermined:",
    mutant: '        elif missing or undetermined:  # [变异H2] 忽略落在执行边界之外的上报',
  },
  // H4：展示层把 `zero` 当结论，不再核对覆盖面是否自洽。
  H4: {
    file: `${REPO}\\packages\\core\\src\\hostWriteAudit.ts`,
    seed: [
      "    const selfConsistent =",
      '      observation.observation === "complete" &&',
      "      observation.missingActions === 0 &&",
      "      observation.unexpectedReports === 0;",
    ].join("\n"),
    mutant: "    const selfConsistent = true; // [变异H4] 照字面读 zero，不核对覆盖是否自洽",
  },
  // H5：真正赋值给宿主的那一次不记计划写 ⇒ 覆盖在、计数恒为 0。
  H5: {
    file: `${REPO}\\apps\\wps-addin\\src\\wpsWriteMatrix.ts`,
    seed: [
      "    countHostWrite(true);",
      "    range[property] = single ? values[0][0] : values;",
    ].join("\n"),
    mutant: [
      "    // [变异H5] 计划写不计数",
      "    range[property] = single ? values[0][0] : values;",
    ].join("\n"),
  },
  // H6：为保文本原样改的 NumberFormat 也算成计划写 ⇒ 辅助调用混进重放判据。
  H6: {
    file: `${REPO}\\apps\\wps-addin\\src\\wpsWriteMatrix.ts`,
    seed: [
      "      countHostWrite(false);",
      '      item.cell.NumberFormat = "@";',
    ].join("\n"),
    mutant: [
      "      countHostWrite(true);  // [变异H6] 保文本用的辅助改格式也算计划写",
      '      item.cell.NumberFormat = "@";',
    ].join("\n"),
  },
  // H7：服务端一个身份都没见过时按名字补一个 —— 先到先得回来了。
  H7: {
    file: `${REPO}\\server\\app\\dev_fault_target.py`,
    seed: [
      "    seen = _identities_for(host, name)",
      "    if not seen:",
    ].join("\n"),
    mutant: [
      '    seen = _identities_for(host, name) or [f"name:{name}"]  # [变异H7] 没身份就按名字认簿',
      "    if not seen:",
    ].join("\n"),
  },
  // H8：同一个名字下见过多个身份时不再拒绝，退回"最近见过的那一个"。
  H8: {
    file: `${REPO}\\server\\app\\dev_fault_target.py`,
    seed: "    if len(seen) > 1:",
    mutant: "    if len(seen) > 1 and False:  # [变异H8] 同名多身份照样启用（取最近那个）",
  },
  // H9：任务侧的 `name:` 回退值不再被拒 ⇒ 按名字张冠李戴。
  H9: {
    file: `${REPO}\\server\\app\\dev_source_changed_fault.py`,
    seed: [
      "        if not dev_fault_target.is_trusted_identity(target):",
      "            return False",
      "        if target != armed.workbook_identity:",
      "            return False",
    ].join("\n"),
    mutant: [
      "        # [变异H9] 任务侧只报名字（name: 回退值）也照样绑",
      "        if target and dev_fault_target.is_trusted_identity(target) and target != armed.workbook_identity:",
      "            return False",
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

/** 写回必须**读回校验**：Windows 上偶发 Errno 22 会把变异留在工作区。 */
function writeVerified(file, text) {
  let lastError;
  for (let attempt = 0; attempt < 3; attempt += 1) {
    try {
      writeFileSync(file, text, "utf8");
      if (readFileSync(file, "utf8") === text) return;
      lastError = new Error("写回内容与预期不一致");
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError ?? new Error(`写回失败：${file}`);
}

const which = process.argv[2];

if (which === "restore-all") {
  for (const key of Object.keys(TARGETS)) {
    const entry = TARGETS[key];
    const text = readFileSync(entry.file, "utf8");
    const normalized = text.replaceAll("\r\n", "\n");
    const mutant = entry.mutant.replaceAll("\r\n", "\n");
    if (!normalized.includes(mutant)) { console.log(`clean: ${key}`); continue; }
    writeVerified(entry.file,
      toFileNewline(normalized.split(mutant).join(entry.seed.replaceAll("\r\n", "\n")), fileNewline(text)));
    console.log(`restored: ${key}`);
  }
  process.exit(0);
}

if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_a1e3_host_writes.mjs H1|H2|H4..H9|restore-all [restore]");
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
  writeVerified(target.file, toFileNewline(normalized.split(mutant).join(seed), newline));
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
writeVerified(target.file, applied);
console.log(`APPLIED ${which} → ${target.file}`);
