// 反向变异脚本：Excel 2024（批量许可永久版，ExcelApi ≤ 1.17）格式兼容路径的**新增守卫**，
// 以及第 19 轮「撤销按动作实际影响拆分恢复」的**范围/属性双层收窄**。
//
// 用法：node scripts/mutate_excel_format_compat.mjs C1|…|C18|check [restore]
//
// 被测机制分六处，各自打不同的判别点：
//
//   ── 兼容捕获/恢复 apps/excel-addin/src/excel.ts ──
//   C1 `if (needsClear) range.clear(Excel.ClearApplyTo.formats)` 变成空操作
//      → 打「需要 clear 时必须真的清空区域直接格式，否则"未直接设置"的属性留在被改过的新值上」。
//   C2 `if (conditionalFormats && conditionalFormats.value > 0)` 恒假
//      → 打「目标区域含条件格式时必须明确拒绝」。
//   C3 inherited 记录改成空数组
//      → 打「按实际修改属性记录哪些是继承态（这是"未采集 ≠ 默认值"的来源记录）」。
//   C4 预检的 `if (formatCapability.path === "unavailable")` 恒假
//      → 打「能力检测前移：连 1.9 都没有的宿主在执行前检查就被点名拒绝」。
//   C5 `const needCellFormat = scope.cellFormat` 改成恒真
//      → 打「采集按动作实际影响门控：只改行高/只改值的动作不该逐格采格式
//        （行高动作被这么一改，会因区域含条件格式而误拒）」。
//   C6 `if (scope.cellFormat) restoreCellFormats(...)` 改成恒真
//      → 打「恢复按动作实际影响门控：只改行高时**不得**调用任何单元格格式恢复」。
//   C7 `const current = observedDirectCells?.[…] ?? {}` 改成恒 `{}`
//      → 打「clear 之前必须把"非本动作属性"的**当前值**并回去，否则原有删除线/图案色/
//        斜线边框会被 clear 一并清掉」。
//
//   ── 第 19 轮：范围 / 属性双层收窄 ──
//   C13 `const directKeys = directKeysOfFields(fields)` 改成按**全集**取
//      → 打「恢复按属性子集收窄：用全集会把用户在本动作之后对别的属性的修改抹掉」。
//   C14 冲突文案里 `(field) => (savedFormatFields?.[field] ?? null) !== current[field]` 恒 false
//      → 打「冲突必须指名**具体属性**（"单元格格式（加粗）"），不能只说"单元格格式"」。
//   C15 采集清单里 `strikethrough: true` 改成 false
//      → 打「凡是 `clear(Formats)` 会清掉、我们又要恢复的属性，都必须进采集清单」。
//   C16 `if (direct.style != null) properties.style = direct.style` 变成空操作
//      → 打「命名样式（非 Normal）必须写回，否则 clear(Formats) 会把它打回 Normal」。
//   C17 `readCurrentDirectCells` 改成按 `RequestContext` 缓存首次读取
//      → 打「`clear(Formats)` 后补回的"非本动作属性当前值"必须读**恢复那一刻**的现场。
//        缓存成一次性快照后，反序撤销里较早的动作会把较晚动作刚撤销掉的属性又写回来
//        （同区域"先加粗、后居中"撤销后居中残留）」。
//   C18 数据验证 fail-closed 判据恒假
//      → 打「兼容路径下目标区域含数据验证（`type !== "None"`）时必须在写入前拒绝。
//        `clear(Formats)` 会清除数据验证，而数据验证不在格式动作管辖范围、也读不回
//        逐格验证来源 ⇒ 不拒绝就是静默丢验证」。
//
//   ── 能力判定 apps/excel-addin/src/excelFormatCapability.ts ──
//   C8 `if (supportsDirect)` 恒假
//      → 打「1.9–1.17 走兼容路径，而不是一刀切判成 unavailable」。这是整个特性的支点。
//   C9 `formatDirectValueMatches` 把「未直接设置（null）」当成匹配
//      → 打「未直接设置 ≠ 已设置成默认值」。
//
//   ── 副本审阅 apps/excel-addin/src/excelNativeComparison.ts ──
//   C10 assertExcelNativeComparisonUnchanged 里删掉 copyFormatEvidence 的比对
//      → 打「副本里的格式被改过必须在写回前被发现」。
//
//   ── 兼容副本投影 apps/excel-addin/src/manualReviewFile.ts ──
//   C11 删掉 `marked.add(address)`
//      → 打「格式动作真正改变了哪些格，要如实标出来」。
//
//   ── 预览阶段提示 apps/excel-addin/src/components/ExcelNativeReviewActions.tsx ──
//   C12 不再渲染能力提示
//      → 打「能力检测前移到预览阶段，不能等用户确认后才报笼统错误」。
//
// 每次只 apply/restore 一个 SEED；脚本只做 apply / restore 两件事，判据由
// run_excel_format_compat_mutations.mjs 负责。
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const EXCEL = resolve(REPO, "apps", "excel-addin", "src", "excel.ts");
const CAPABILITY = resolve(REPO, "apps", "excel-addin", "src", "excelFormatCapability.ts");
const COMPARISON = resolve(REPO, "apps", "excel-addin", "src", "excelNativeComparison.ts");
const MANUAL = resolve(REPO, "apps", "excel-addin", "src", "manualReviewFile.ts");
const COMPONENT = resolve(
  REPO,
  "apps",
  "excel-addin",
  "src",
  "components",
  "ExcelNativeReviewActions.tsx"
);

// 模板字面量里的 `<<>>` 是 `${` 的占位 —— 直接写 `${` 会被本脚本自己插值掉。
const DOLLAR = "<<>>";

const MUTATIONS = {
  C1: {
    file: EXCEL,
    seed: `    if (needsClear) range.clear(Excel.ClearApplyTo.formats);`,
    mutant: `    if (needsClear) void Excel.ClearApplyTo.formats;`
  },
  C2: {
    file: EXCEL,
    seed: `  if (conditionalFormats && conditionalFormats.value > 0) {`,
    mutant: `  if (false) {`
  },
  C3: {
    file: EXCEL,
    seed: `        const inherited = modifiedFields.filter(
          (field) => !directFormatHasField(direct, field)
        );`,
    mutant: `        const inherited = [] as ExcelFormatField[];`
  },
  C4: {
    file: EXCEL,
    seed: `  if (formatCapability.path === "unavailable") {`,
    mutant: `  if (false) {`
  },
  C5: {
    file: EXCEL,
    seed: `  const needCellFormat = scope.cellFormat;`,
    mutant: `  const needCellFormat = true;`
  },
  C6: {
    file: EXCEL,
    seed: `  if (scope.cellFormat) restoreCellFormats(range, saved, formatFields, observedDirectCells);`,
    mutant: `  if (true) restoreCellFormats(range, saved, formatFields, observedDirectCells);`
  },
  C7: {
    file: EXCEL,
    seed: `        const current = observedDirectCells?.[rowIndex]?.[columnIndex] ?? {};`,
    mutant: `        const current: Record<string, unknown> = observedDirectCells?.[rowIndex]?.[columnIndex] ? {} : {};`
  },
  C8: {
    file: CAPABILITY,
    seed: `  if (supportsDirect) {`,
    mutant: `  if (false) {`
  },
  C9: {
    file: CAPABILITY,
    seed: `    case "fontBold":
    case "fontItalic":
    case "wrapText":
      return actual === expected;`,
    mutant: `    case "fontBold":
    case "fontItalic":
    case "wrapText":
      return actual === expected || actual == null;`
  },
  C10: {
    file: COMPARISON,
    seed: `    if (current !== session.copyFormatEvidence) throw error("副本中的格式已变化，请重新生成审阅副本。");`,
    mutant: `    void current;`
  },
  C11: {
    file: MANUAL,
    seed: `        if (digest(before) !== digest(after)) marked.add(address);`,
    mutant: `        void digest(before);
        void digest(after);`
  },
  C12: {
    file: COMPONENT,
    // 重定点记录（2026-10-16）：组件里的能力提示被压成单行（`details` 的
    // `excel-format-capability-detail` 类名与 tone 类一并去掉），旧的多行种子失配。
    // 只改锚、不改判别点 —— 判别点仍是「这段能力提示一个字都不渲染」（mutant 还是 `{null}`）。
    seed: `    {formatNotice && <div className="excel-format-capability-notice" role="status">
      <p className="excel-format-capability-summary">{formatNotice.summary}</p>
      <details><summary>技术原因</summary><p>{formatNotice.detail}</p></details>
    </div>}`,
    mutant: `    {null}`
  },
  C13: {
    file: EXCEL,
    seed: `    const directKeys = directKeysOfFields(fields);`,
    mutant: `    const directKeys = directKeysOfFields(ALL_DIRECT_FORMAT_FIELDS);`
  },
  C14: {
    file: EXCEL,
    seed: `        (field) => (savedFormatFields?.[field] ?? null) !== current[field]`,
    mutant: `        () => false`
  },
  C15: {
    file: EXCEL,
    seed: `    underline: true,
    strikethrough: true,`,
    mutant: `    underline: true,
    strikethrough: false,`
  },
  C16: {
    file: EXCEL,
    seed: `  if (direct.style != null) properties.style = direct.style;`,
    mutant: `  void direct.style;`
  },
  C17: {
    file: EXCEL,
    seed: `async function readCurrentDirectCells(
  context: Excel.RequestContext,
  range: Excel.Range
): Promise<DirectCellFormatValues[][] | null> {
  // 只有兼容（直接格式）路径需要它；显示值路径与不可用宿主都返回 null。
  if (currentExcelFormatCapability().path !== "direct") return null;
  const properties = range.getCellProperties(DIRECT_CELL_PROPERTIES_LOAD);
  await context.sync();
  return properties.value.map((row) =>
    row.map((cell) => {
      const typed = cell as Excel.CellProperties;
      return readDirectCellFormat(typed.format, typed.style);
    })
  );
}`,
    mutant: `const __staleDirectByContext = new WeakMap<
  Excel.RequestContext,
  DirectCellFormatValues[][]
>();
async function readCurrentDirectCells(
  context: Excel.RequestContext,
  range: Excel.Range
): Promise<DirectCellFormatValues[][] | null> {
  const cached = __staleDirectByContext.get(context);
  if (cached) return cached;
  if (currentExcelFormatCapability().path !== "direct") return null;
  const properties = range.getCellProperties(DIRECT_CELL_PROPERTIES_LOAD);
  await context.sync();
  const rows = properties.value.map((row) =>
    row.map((cell) => {
      const typed = cell as Excel.CellProperties;
      return readDirectCellFormat(typed.format, typed.style);
    })
  );
  __staleDirectByContext.set(context, rows);
  return rows;
}`
  },
  C18: {
    file: EXCEL,
    seed: `  if (dataValidation && dataValidation.type !== "None") {
    throw new FormatRangeUnrecoverableError(
      "目标区域含数据验证；恢复用的 clear(Formats) 会清除数据验证且无法逐格还原，本机缺少可验证的保留路径"
    );
  }`,
    mutant: `  if (false) {
    throw new FormatRangeUnrecoverableError(
      "目标区域含数据验证；恢复用的 clear(Formats) 会清除数据验证且无法逐格还原，本机缺少可验证的保留路径"
    );
  }`
  }
};

function newlineOf(text) {
  return text.includes("\r\n") ? "\r\n" : LF;
}

function normalize(text) {
  return text.replace(/\r\n/g, LF);
}

function expand(value) {
  return String(value).replaceAll(DOLLAR, "${");
}

function readTarget(mutation) {
  return readFileSync(mutation.file, "utf8");
}

function apply(name) {
  const mutation = MUTATIONS[name];
  if (!mutation) throw new Error(`未知变异：${name}`);
  const raw = readTarget(mutation);
  const eol = newlineOf(raw);
  const source = normalize(raw);
  const seed = normalize(expand(mutation.seed));
  const mutant = normalize(expand(mutation.mutant));
  const occurrences = source.split(seed).length - 1;
  if (occurrences !== 1) {
    throw new Error(
      `${mutation.file}：SEED 命中 ${occurrences} 次（要求恰好 1 次），变异 ${name} 无法 apply`
    );
  }
  const next = source.replace(seed, mutant);
  writeFileSync(mutation.file, eol === LF ? next : next.replace(/\n/g, eol), "utf8");
  return { file: mutation.file };
}

function restore(name) {
  const mutation = MUTATIONS[name];
  if (!mutation) throw new Error(`未知变异：${name}`);
  const raw = readTarget(mutation);
  const eol = newlineOf(raw);
  const source = normalize(raw);
  const seed = normalize(expand(mutation.seed));
  const mutant = normalize(expand(mutation.mutant));
  const occurrences = source.split(mutant).length - 1;
  if (occurrences !== 1) {
    throw new Error(
      `${mutation.file}：MUTANT 命中 ${occurrences} 次（要求恰好 1 次），变异 ${name} 无法 restore`
    );
  }
  const next = source.replace(mutant, seed);
  writeFileSync(mutation.file, eol === LF ? next : next.replace(/\n/g, eol), "utf8");
  return { file: mutation.file };
}

/** 前置状态闸：每个文件停在 SEED 态（未变异）才允许继续。 */
function check() {
  for (const [name, mutation] of Object.entries(MUTATIONS)) {
    const raw = readTarget(mutation);
    const source = normalize(raw);
    const seed = normalize(expand(mutation.seed));
    const mutant = normalize(expand(mutation.mutant));
    if (!source.includes(seed)) {
      return { ok: false, name, file: mutation.file, reason: "找不到 SEED（可能停在上一次被中断的变异态）" };
    }
    if (source.includes(mutant)) {
      return { ok: false, name, file: mutation.file, reason: "已经处在 MUTANT 态（上一次 restore 没跑完）" };
    }
  }
  return { ok: true };
}

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

const mode = process.argv[2];
const argument = process.argv[3];

// 中断也要还原：把已 apply 的变异原路退回。
let applied = null;
const guard = (signal) => {
  if (applied) {
    try {
      restore(applied);
      process.stderr.write(`[restore:${signal}] ${MUTATIONS[applied].file} 已还原\n`);
    } catch (error) {
      process.stderr.write(`[restore:${signal}] 还原失败：${error.message}\n`);
    }
  }
  process.exit(1);
};
for (const signal of ["SIGINT", "SIGTERM", "SIGHUP"]) {
  process.on(signal, () => guard(signal));
}
process.on("uncaughtException", (error) => {
  guard(`uncaught:${error.message}`);
});

try {
  if (mode === "check") {
    const result = check();
    if (!result.ok) {
      process.stdout.write(`CHECK FAILED ${result.file}: ${result.reason}（变异 ${result.name}）\n`);
      process.exit(1);
    }
    process.stdout.write(`CHECK OK 全部 ${Object.keys(MUTATIONS).length} 个 SEED 都在位\n`);
    process.exit(0);
  }
  if (!MUTATIONS[mode]) throw new Error(`未知用法：${mode}`);
  if (argument === "restore") {
    const result = restore(mode);
    process.stdout.write(`RESTORED ${mode} → ${result.file} sha256=${sha256(result.file)}\n`);
    process.exit(0);
  }
  const result = apply(mode);
  applied = mode;
  process.stdout.write(`APPLIED ${mode} → ${result.file} sha256=${sha256(result.file)}\n`);
  process.exit(0);
} catch (error) {
  process.stderr.write(`${error.message}\n`);
  process.exit(1);
}
