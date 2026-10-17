// 反向变异脚本：第 3 阶段「常用格式操作」两端闭环的**新增守卫**。
//
// 用法：node scripts/mutate_format_actions.mjs V1|V2|…|V8|P1|…|P8|check [restore]
//
// 被测机制分四处，各自打不同的判别点：
//
//   ── 共享层 packages/core/src/formatActions.ts（vitest: formatActions.test.ts）──
//   V1 行高改动退回 { kind: "range" }        → 打「行高/列宽必须声明整行整列」。
//      这是本轮最容易写错的一条：只改一行代码，预览就从"影响整行"退回"只影响这块"。
//   V5 颜色比较改成大小写敏感                → 打「宿主回读改大小写不算改动」。
//   V6 `buildFormatChangeReview` 把"缺键"当已采集 → 打「未采集 ≠ 默认值」。
//   V7 `formatActionEdits` 丢掉 italic 一项   → 打「五属性一项不漏」。
//   V8 预览的 impactNotice 恒为 null          → 打「行高列宽的提示必须真的渲染出来」。
//
//   ── Excel 端字段校验 apps/excel-addin/src/contracts.ts ──
//   V2 删掉 setFont 的"空属性"拒绝           → 打「空格式动作必须在动单元格之前被拒」。
//      由 formatActionFields.test.ts 接住；其中一条走 assertEditablePlan，
//      所以调用链被拆掉同样会红。
//
//   ── WPS 端 fail-closed 与配置漂移 apps/wps-addin/src/wpsPlanExecutor.ts / config ──
//   V3 往 WPS_SUPPORTED_ACTION_KEYS 里加 "setFont"（配置未同步）
//   V4 往 capabilities.wpsWrite.allowedActions 里加 "setFont"（执行器未实现）
//      这两条互为反向：只查一半的守卫会漏掉另一半，所以必须各证一次。
//
//   ── Python 协议 server/app/models.py（pytest: test_protocol_consistency.py）──
//   P1/P2/P3 拆掉 setFont / setAlignment / resizeRange 的空属性拒绝
//   P4 拆掉 setBorders 的重复边拒绝
//      P1–P4 一律把守卫的 `if` 条件改写成 `if False:` —— 保留原语句结构，
//      只拆掉"该拒的没拒"这一个判别点；不用 `void (...)` 之类会抛 NameError
//      的替身，否则测试变红的原因就变成"抛了别的异常"，而不是"守卫没了"。
//
//   ── 跨语言漂移检测（同一批 pytest，读的是 TS 源）──
//   P5 TS 的 BORDER_WEIGHTS 改一个取值     → 打「闭集两端一致」
//   P6 TS 的行高上限 409→408                → 打「值域两端一致」
//   P7 TS 的字体名上限 255→254              → 打「长度上限两端一致」
//   P8 TS 把 italic 从可选改成必填          → 打「可选性两端一致」
//      P5–P8 的判别点在**另一端**：改的是 TS，红的是 Python 用例。
//      若检测器退化成"只读自己这边的常量"，这四条全都不会红。
//
// 十二个变异落在五个文件上，换行符不统一（models.py / contracts.ts /
// wpsPlanExecutor.ts / capabilities.json 是 CRLF，formatActions.ts 是 LF），
// 所以 seed/mutant 的归一化与写回都要**按文件**取目标换行。
//
// 脚本只做 apply / restore 两件事，判据由 run_format_action_mutations.mjs 负责。
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const CORE = resolve(REPO, "packages", "core", "src", "formatActions.ts");
const CONTRACTS = resolve(REPO, "apps", "excel-addin", "src", "contracts.ts");
const WPS_EXECUTOR = resolve(REPO, "apps", "wps-addin", "src", "wpsPlanExecutor.ts");
const CAPABILITIES = resolve(REPO, "config", "capabilities.json");
const MODELS = resolve(REPO, "server", "app", "models.py");

// 模板字面量里的 `<<>>` 是 `${` 的占位 —— 直接写 `${` 会被本脚本自己插值掉。
const DOLLAR = "<<>>";

const MUTATIONS = {
  V1: {
    file: CORE,
    seed: `    if (changesRows) return { kind: "rows", ...base, rows };`,
    mutant: `    if (changesRows) return { kind: "range", ...base };`
  },
  V2: {
    file: CONTRACTS,
    seed: `        throw new Error(\`${DOLLAR}where}的 setFont 没有指定任何要修改的字体属性\`);`,
    mutant: `        void where;`
  },
  V3: {
    file: WPS_EXECUTOR,
    seed: `  "clearRange:contents"`,
    mutant: `  "clearRange:contents",\n  "setFont"`
  },
  V4: {
    file: CAPABILITIES,
    seed: `"allowedActions": ["writeValues", "writeFormulas", "clearRange:contents"]`,
    mutant: `"allowedActions": ["writeValues", "writeFormulas", "clearRange:contents", "setFont"]`
  },
  V5: {
    file: CORE,
    seed: `    if (property === "fontColor" || property === "fillColor") {`,
    mutant: `    if (false) {`
  },
  V6: {
    file: CORE,
    seed: `    const beforeCaptured = input.before.captured && property in input.before.properties;`,
    mutant: `    const beforeCaptured = input.before.captured;`
  },
  V7: {
    file: CORE,
    seed: `      if (action.italic != null) edits.push({ property: "fontItalic", value: action.italic });`,
    mutant: `      void action;`
  },
  V8: {
    file: CORE,
    seed: `      impactNotice = formatImpactNotice(group.impact);`,
    mutant: `      impactNotice = null;`
  },
  P1: {
    file: MODELS,
    seed: `        if all(
            value is None
            for value in (self.bold, self.italic, self.name, self.size, self.color)
        ):`,
    mutant: `        if False:`
  },
  P2: {
    file: MODELS,
    seed: `        if self.horizontal is None and self.vertical is None and self.wrapText is None:`,
    mutant: `        if False:`
  },
  P3: {
    file: MODELS,
    seed: `        if self.rowHeight is None and self.columnWidth is None:`,
    mutant: `        if False:`
  },
  P4: {
    file: MODELS,
    seed: `        if len(set(self.sides)) != len(self.sides):`,
    mutant: `        if False:`
  },
  P5: {
    file: CONTRACTS,
    seed: `  "hairline",`,
    mutant: `  "hairlineX",`
  },
  P6: {
    file: CONTRACTS,
    seed: `export const ROW_HEIGHT_RANGE = { min: 0, max: 409 } as const;`,
    mutant: `export const ROW_HEIGHT_RANGE = { min: 0, max: 408 } as const;`
  },
  P7: {
    file: CONTRACTS,
    seed: `export const FONT_NAME_MAX_LENGTH = 255;`,
    mutant: `export const FONT_NAME_MAX_LENGTH = 254;`
  },
  P8: {
    file: CONTRACTS,
    seed: `      italic?: boolean | null;`,
    mutant: `      italic: boolean | null;`
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
    throw new Error(`${name} 的 SEED 在 ${mutation.file} 里命中 ${occurrences} 次（要求恰好 1 次）`);
  }
  const next = source.replace(seed, mutant);
  writeFileSync(mutation.file, eol === LF ? next : next.replace(/\n/g, eol), "utf8");
  return { file: mutation.file, eol };
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
    throw new Error(`${name} 的 MUTANT 在 ${mutation.file} 里命中 ${occurrences} 次（要求恰好 1 次）`);
  }
  const next = source.replace(mutant, seed);
  writeFileSync(mutation.file, eol === LF ? next : next.replace(/\n/g, eol), "utf8");
  return { file: mutation.file, eol };
}

/** 前置状态闸：文件停在 SEED 态（未变异）才允许继续。 */
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
      process.stderr.write(`[restore:${signal}] ${applied} 已还原\n`);
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
      process.stdout.write(`CHECK FAILED ${result.name} ${result.file}: ${result.reason}\n`);
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
