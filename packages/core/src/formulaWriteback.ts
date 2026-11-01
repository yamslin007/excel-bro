/**
 * 公式写入的**共享判据**（宿主无关）。
 *
 * 存在理由（2026-10-30）：精确续写以前只认 `writeValues`。要放开 `writeFormulas`，必须先回答
 * 两个问题，而这两个问题的答案**两端必须一致**，否则同一份物料在 Excel 说"已写入"、
 * 在 WPS 说"写了一半"：
 *
 * 1. **一格算不算"已经写好了"？** —— 只看公式文本，不看计算值。
 *    两个不同的公式可以算出同一个数（`=1+1` 与 `=2` 都是 2），显示值相同**不能**证明
 *    写进去的是我们要的那条公式；反过来，公式文本一致而计算值不同（跨表数据变了、
 *    还没重算）仍然是"公式已写入"。
 *    这条纪律靠**签名**保证：本模块所有判定函数的入参里**没有计算值矩阵**，
 *    想拿计算值兜底就得先改类型 —— 不是靠注释劝。
 *
 * 2. **这块区域现在到底是写前、写后，还是"在途不明"？** —— 三种现场必须分得开。
 *    续跑只允许在"写前"上动手、"写后"跳过；"在途不明"一律停下来核对，
 *    绝不猜着重写一遍（那是重复写入），也绝不假装已完成。
 *
 * 刻意**不**做公式归一化：`=A1*2` 与 `=$A$1*2` 是两个不同的写入目标，`=SUM(A1:A3)` 与
 * 平移后的 `=SUM(B1:B3)` 更是。归一化会让"写错了列"看起来"一模一样"。
 *
 * 宿主差异（怎么读 `Formula`/`Value2`、怎么写回）留在各自适配层：
 * WPS 走 `wpsPlanExecutor` 的矩阵读取，Excel 走 `excel.ts` 的
 * `formulasEqual` / `formulasR1C1Equal` 两条 criterion。
 */

/** 公式矩阵：每格是公式文本（`=` 开头）或 `null`（该格不是公式）。 */
export type FormulaMatrix = readonly (readonly (string | null)[])[];

/** 只有 `=` 开头才算公式 —— 与 WPS `range.Formula` 对常量返回常量本身的口径一致。 */
export function isFormulaText(cell: unknown): cell is string {
  return typeof cell === "string" && cell.startsWith("=");
}

/**
 * 文本身份：**逐字符比较**，不 trim、不改大小写、不补 `$`、不换算 R1C1。
 *
 * 为什么连大小写都不改：Excel 函数名不区分大小写，但**公式文本**区分（我们写什么就该读到什么）。
 * 这里放宽会让"写进去的其实是另一条公式"过关。
 */
export function formulaTextEqual(left: unknown, right: unknown): boolean {
  return isFormulaText(left) && isFormulaText(right) && left === right;
}

export type FormulaWriteMatrixIssue = {
  /** 0 基行列，便于直接定位。 */
  row: number;
  column: number;
  /** `not_formula` = 不是 `=` 开头的文本；`not_text` = 根本不是字符串。 */
  kind: "not_text" | "not_formula";
};

/**
 * 写后目标矩阵的形状/类型诊断（返回 `null` = 合格）。
 *
 * 用途：把"记录里要写的是公式"这件事钉死。一份 `writeFormulas` 的物料如果退化成常量
 * （比如只存了计算值、或者被投影成值矩阵），**哪怕算出来的数一模一样**，也不能拿去写 ——
 * 那是把公式换成常量，与用户确认过的计划不是一回事。
 */
export function formulaWriteMatrixIssue(matrix: readonly (readonly unknown[])[]): FormulaWriteMatrixIssue | null {
  for (let row = 0; row < matrix.length; row += 1) {
    const line = matrix[row];
    if (!Array.isArray(line)) return { row, column: 0, kind: "not_text" };
    for (let column = 0; column < line.length; column += 1) {
      const cell = line[column];
      if (typeof cell !== "string") return { row, column, kind: "not_text" };
      if (!isFormulaText(cell)) return { row, column, kind: "not_formula" };
    }
  }
  return null;
}

/**
 * 公式文本里引用的**工作表名**（去重、保序、去引号）。
 *
 * 为什么共享：公式的依赖关系藏在文本里，服务端登记的 `sourceDependencySheets` 是另一份
 * 声明。续跑时必须能回答"这条公式引用的表，是否仍在这次授权与来源复核范围内" ——
 * 否则一条 `=Other!A1*2` 的公式可以绕过"来源变了就不许续跑"。
 *
 * 处理规则（刻意保守）：
 * - `Sheet1!A1` → `Sheet1`；`'My Sheet'!A1` → `My Sheet`（引号里的 `''` 还原成一个 `'`）；
 * - `[Book1.xlsx]Sheet1!A1` → `[Book1.xlsx]Sheet1`（**外部引用原样保留**，交给调用方拒绝）；
 * - 双引号字符串字面量里出现的 `!` 不算引用（`="Sheet1!A1"` 是文本，不是引用）；
 * - 不解析 `INDIRECT` / `OFFSET` 之类的动态引用 —— 解析不出来就不假装知道，
 *   调用方看到 `dynamic` 标记应当 fail-closed。
 */
export function formulaSheetReferences(matrix: FormulaMatrix): { sheets: string[]; dynamic: boolean } {
  const seen = new Set<string>();
  let dynamic = false;
  for (const line of matrix) {
    for (const cell of line) {
      if (!isFormulaText(cell)) continue;
      const found = referencesInFormula(cell);
      if (found.dynamic) dynamic = true;
      for (const name of found.sheets) seen.add(name);
    }
  }
  return { sheets: [...seen], dynamic };
}

function referencesInFormula(formula: string): { sheets: string[]; dynamic: boolean } {
  const sheets: string[] = [];
  let dynamic = false;
  let index = 0;
  while (index < formula.length) {
    const char = formula[index];
    // 双引号字符串字面量：跳过（Excel 里 "" 表示一个引号）。
    if (char === '"') {
      index += 1;
      while (index < formula.length) {
        if (formula[index] === '"') {
          if (formula[index + 1] === '"') {
            index += 2;
            continue;
          }
          index += 1;
          break;
        }
        index += 1;
      }
      continue;
    }
    if (char === "'") {
      // 带引号的工作表名：成对的 ' 之间，'' 表示一个 '。
      let name = "";
      let cursor = index + 1;
      let closed = false;
      while (cursor < formula.length) {
        if (formula[cursor] === "'") {
          if (formula[cursor + 1] === "'") {
            name += "'";
            cursor += 2;
            continue;
          }
          closed = true;
          cursor += 1;
          break;
        }
        name += formula[cursor];
        cursor += 1;
      }
      if (closed && formula[cursor] === "!") {
        sheets.push(name);
        index = cursor + 1;
        continue;
      }
      index = cursor;
      continue;
    }
    const upper = formula.toUpperCase();
    if (upper.startsWith("INDIRECT(", index) || upper.startsWith("OFFSET(", index)) dynamic = true;
    if (char === "!") {
      // 向左回读：`[book]sheet!` 或 `sheet!`（不含引号、不含运算符与括号）。
      let start = index;
      while (start > 0 && /[^\s()+\-*/,&^=<>{};]/.test(formula[start - 1])) start -= 1;
      const token = formula.slice(start, index);
      if (token) sheets.push(token);
    }
    index += 1;
  }
  return { sheets, dynamic };
}

/**
 * 公式的**依赖面** —— 恢复时"来源证据必须覆盖哪些表"的唯一判据输入。
 *
 * 为什么不直接用 `formulaSheetReferences`：它只回答"文本里点名了哪些表"，看不见
 * **未限定表名**的引用（`=A1*2` 里的 `A1` 就在写入目标表身上）。恢复时那是最要命的
 * 一种 —— 证据按表登记，而写入目标表的值每写一格就变，再也分不清"我们自己写的"
 * 和"别人改的"。所以把依赖拆成四档，让调用方按"能不能登记证据"决定放不放行续跑：
 * 既不把所有公式一律拒绝，也不把看不见的那部分当成"没有依赖"。
 */
export type FormulaDependencies = {
  /** 显式跨表引用的表名（去重保序）；外部工作簿引用不在此列。 */
  sheets: string[];
  /** 引用了**本表**（未限定表名的单元格/整列/整行）—— 与写入目标同表。 */
  local: boolean;
  /** 解析不出依赖（`INDIRECT` / `OFFSET`）。 */
  dynamic: boolean;
  /** 外部工作簿引用（`[Book.xlsx]Sheet1!A1`）。 */
  external: boolean;
  /** 认不出的裸标识符（很可能是**定义名称**）：可以指向任意表 ⇒ 依赖不确定。 */
  unresolved: boolean;
};

const CELL_REFERENCE = /^\$?[A-Za-z]{1,3}\$?\d+$/;
const COLUMN_REFERENCE = /^\$?[A-Za-z]{1,3}$/;
const ROW_REFERENCE = /^\$?\d+$/;
const BOOLEAN_LITERALS = new Set(["TRUE", "FALSE"]);

/** 词字符：字母（含中文）、数字、`$`、`_`、`.`。 */
function isWordChar(char: string | undefined): boolean {
  return char !== undefined && /[0-9A-Za-z_$.\u4e00-\u9fff]/.test(char);
}

/** 词首：字母（含中文）、`$`、`_`。数字开头的是字面量，不进词。 */
function isWordStart(char: string | undefined): boolean {
  return char !== undefined && /[A-Za-z_$\u4e00-\u9fff]/.test(char);
}

/** `Sheet!` 之后那一截单元格/区域引用（`A1`、`$A$1`、`A1:B2`、`#REF!`）。 */
function skipReferenceAfterBang(formula: string, from: number): number {
  let index = from;
  while (index < formula.length && /[0-9A-Za-z$:.[\]]/.test(formula[index])) index += 1;
  return index;
}

/**
 * 扫“有没有读**本表**”和“有没有认不出的名字”。
 *
 * 只做两件事，且都往**保守**方向错：
 * - 未限定表名的单元格/整列/整行引用 ⇒ `local = true`（哪怕它其实写在别的表的公式里，
 *   我们只按文本判，判不准就拒绝恢复）；
 * - 认不出的裸标识符 ⇒ `unresolved = true`（定义名称可以指向任意表）。
 * 双引号字面量、带引号/方括号的表名限定引用都跳过 —— 那些不算"读本表"。
 */
function scanLocalReferences(matrix: FormulaMatrix): {
  local: boolean;
  unresolved: boolean;
} {
  let local = false;
  let unresolved = false;
  const visit = (formula: string) => {
    let index = 0;
    while (index < formula.length) {
      const char = formula[index];
      if (char === '"') {
        index += 1;
        while (index < formula.length) {
          if (formula[index] === '"') {
            if (formula[index + 1] === '"') {
              index += 2;
              continue;
            }
            index += 1;
            break;
          }
          index += 1;
        }
        continue;
      }
      if (char === "'") {
        let cursor = index + 1;
        let closed = false;
        while (cursor < formula.length) {
          if (formula[cursor] === "'") {
            if (formula[cursor + 1] === "'") {
              cursor += 2;
              continue;
            }
            closed = true;
            cursor += 1;
            break;
          }
          cursor += 1;
        }
        if (closed && formula[cursor] === "!") {
          index = skipReferenceAfterBang(formula, cursor + 1);
          continue;
        }
        index = cursor;
        continue;
      }
      if (char === "[") {
        // 外部工作簿引用：[Book.xlsx]Sheet1!A1 —— 归 external，不算本表。
        const close = formula.indexOf("]", index);
        if (close < 0) {
          index += 1;
          continue;
        }
        let cursor = close + 1;
        while (isWordChar(formula[cursor])) cursor += 1;
        if (formula[cursor] === "!") {
          index = skipReferenceAfterBang(formula, cursor + 1);
          continue;
        }
        index = cursor;
        continue;
      }
      // 整行区间 `3:5`：两端都以数字开头，进不了下面的"词"分支，单独认。
      // `1E5` 这类科学计数在这里只吃到 `1`，后面 `E5` 会被下面的词首判断跳过。
      if (/[0-9]/.test(char) && !isWordChar(formula[index - 1])) {
        let cursor = index;
        while (/[0-9]/.test(formula[cursor])) cursor += 1;
        if (formula[cursor] === ":" && /[0-9$]/.test(formula[cursor + 1] ?? "")) {
          local = true;
        }
        index = cursor;
        continue;
      }
      // `1E5` 这类：数字本身不是词，后面的 `E5` 不能当成单元格引用。
      if (isWordStart(char) && !isWordChar(formula[index - 1])) {
        const start = index;
        let cursor = index;
        while (isWordChar(formula[cursor])) cursor += 1;
        const token = formula.slice(start, cursor);
        const next = formula[cursor];
        if (next === "!") {
          index = skipReferenceAfterBang(formula, cursor + 1);
          continue;
        }
        if (next === "(") {
          index = cursor + 1;
          continue;
        }
        if (next === ":") {
          // 区域引用：`A1:B2`、`A:A`、`$A$1:$B$2`。左半是引用还不够，
          // 右半也要一起吃掉 —— 否则 `A:A` 的后一个 `A` 会被当成"认不出的名字"。
          if (CELL_REFERENCE.test(token) || COLUMN_REFERENCE.test(token) || ROW_REFERENCE.test(token)) {
            local = true;
            let end = cursor + 1;
            while (isWordChar(formula[end])) end += 1;
            const right = formula.slice(cursor + 1, end);
            if (CELL_REFERENCE.test(right) || COLUMN_REFERENCE.test(right) || ROW_REFERENCE.test(right)) {
              index = end;
              continue;
            }
            index = cursor + 1;
            continue;
          }
          index = cursor + 1;
          continue;
        }
        if (CELL_REFERENCE.test(token)) {
          local = true;
          index = cursor;
          continue;
        }
        if (BOOLEAN_LITERALS.has(token.toUpperCase())) {
          index = cursor;
          continue;
        }
        unresolved = true;
        index = cursor;
        continue;
      }
      index += 1;
    }
  };
  for (const line of matrix) {
    for (const cell of line) {
      if (isFormulaText(cell)) visit(cell);
    }
  }
  return { local, unresolved };
}

/** 公式矩阵的完整依赖面（见 `FormulaDependencies`）。 */
export function formulaDependencies(matrix: FormulaMatrix): FormulaDependencies {
  const base = formulaSheetReferences(matrix);
  const sheets: string[] = [];
  let external = false;
  for (const name of base.sheets) {
    // `formulaSheetReferences` 把外部引用**原样**返回（`[Book1.xlsx]Sheet1`），交给调用方拒。
    if (name.includes("[") || name.includes("]")) {
      external = true;
      continue;
    }
    if (!sheets.some((item) => item.toLocaleLowerCase() === name.toLocaleLowerCase())) {
      sheets.push(name);
    }
  }
  const scan = scanLocalReferences(matrix);
  return {
    sheets,
    local: scan.local,
    dynamic: base.dynamic,
    external,
    unresolved: scan.unresolved
  };
}

/** 「公式计划的来源证据」要求：ok 时给出**必须被证据覆盖**的表集合。 */
export type FormulaSourceEvidenceRequirement =
  | { ok: true; sheets: string[] }
  | { ok: false; reason: string };

/**
 * 公式计划续跑要求登记哪些来源证据 —— **两端唯一的判据**。
 *
 * 产品规则（2026-10-30 拍板）：公式精确恢复要求"登记的来源证据仍然一致"；来源值变了
 * 就停止续跑并明确提示，不自动重新规划，也不放宽任何一端。于是能续跑的前提变成
 * "依赖的每一张表都能被登记的证据覆盖"：
 *
 * - 动态引用 / 外部工作簿 / 未限定表名的本表引用 / 认不出的名称 ⇒ **不可续跑**
 *   （前两者解析不出依赖，后两者与写入目标同表或落在别处，都没法把"我们自己写的"
 *   与"别人改的"分开）；
 * - 引用的表若也是本次写入目标 ⇒ 不可续跑（同样的分离问题）；
 * - 其余情况返回需要登记的表集合；空集合（如 `=1+1`）表示不读任何来源，不需要证据。
 *
 * 它只回答"能不能登记证据"，不回答"授权有没有覆盖" —— 授权是另一道闸，不许互相冒充。
 */
export function formulaSourceEvidenceRequirement(plan: {
  actions?: readonly {
    type?: string | null;
    sheet?: string | null;
    formulas?: FormulaMatrix | null;
  }[] | null;
} | null | undefined): FormulaSourceEvidenceRequirement {
  const actions = plan?.actions ?? [];
  const targets = new Set(
    actions
      .map((action) => String(action?.sheet ?? "").trim().toLocaleLowerCase())
      .filter(Boolean)
  );
  const sheets: string[] = [];
  for (const [index, action] of actions.entries()) {
    if (action?.type !== "writeFormulas") continue;
    const matrix = Array.isArray(action.formulas) ? action.formulas : [];
    const deps = formulaDependencies(matrix);
    const label = `第 ${index + 1} 步公式`;
    if (deps.dynamic) {
      return {
        ok: false,
        reason: `${label}含动态引用（INDIRECT / OFFSET），解析不出它读哪张表`
      };
    }
    if (deps.external) {
      return { ok: false, reason: `${label}引用了外部工作簿，跨重启恢复不核对它` };
    }
    if (deps.local) {
      return {
        ok: false,
        reason: `${label}引用了写入目标表自身的单元格，恢复时无法把本次写入与外部修改分开核对`
      };
    }
    if (deps.unresolved) {
      return {
        ok: false,
        reason: `${label}含无法判定的名称引用（可能是定义名称），依赖不确定`
      };
    }
    for (const name of deps.sheets) {
      if (targets.has(name.toLocaleLowerCase())) {
        return {
          ok: false,
          reason: `${label}依赖的来源表「${name}」也是本次写入目标`
        };
      }
      if (!sheets.some((item) => item.toLocaleLowerCase() === name.toLocaleLowerCase())) {
        sheets.push(name);
      }
    }
  }
  return { ok: true, sheets };
}

/*
 * 续跑时的三句话，**两端共用**（Excel 与 WPS 不许各写一份文案）：
 * 三种成因不同，用户要做的事也不同 —— 值真变了要重新读数据；没证据要去核对；
 * 读不出来先恢复读取能力。合并成一句"不能续跑"等于把处置方法也藏起来。
 */
export const FORMULA_SOURCE_CHANGED_MESSAGE = "来源已变化，不能沿用旧结果续跑。";
export const FORMULA_SOURCE_EVIDENCE_MISSING_MESSAGE =
  "这次执行没有登记公式来源的证据，无法确认来源未变，不能续跑。";
export const FORMULA_SOURCE_UNVERIFIABLE_MESSAGE =
  "读不到公式依赖的来源值，无法确认来源未变，不能续跑。";

/** 续跑前对来源证据的结论。`not-applicable` = 这份计划没有需要证据的公式。 */
export type FormulaResumeSourceVerdict =
  | { verdict: "not-applicable" }
  | { verdict: "ok"; sheets: string[] }
  | { verdict: "refused"; reason: string };

/**
 * 续跑前比对**登记的**来源值与现场 —— 公式精确恢复的唯一判据（两端共用）。
 *
 * 比对对象是**公式真正依赖的那几张表**（由公式文本算出，见
 * `formulaSourceEvidenceRequirement`），不是"本次授权范围的整张快照"：
 * 拿一个与公式依赖无关的快照去比，既会在无关表被改时误拒，也证明不了公式的输入没变。
 *
 * 指纹口径由计划自身决定（= 依赖表集合），因此不额外登记一份表清单：
 * 登记时与续跑时算出的是同一个集合。谁把口径写歪了（比如改成授权范围），
 * 两端口径就会分叉 —— 这正是"两端读同一份判据"要防的事。
 *
 * 读取回调由宿主提供（Excel `captureWorkbookSourceFingerprint` / WPS `readMatrices`）；
 * 它**抛错或返回空**一律按"读不到"拒绝 —— 读不到不等于没变。
 * `sourceDependencySheets`（本机实际读过哪些表）是**登记闸**的事，不在这里重复判：
 * 三条判据分开说，别互相冒充。
 */
export async function verifyFormulaResumeSources(
  plan: Parameters<typeof formulaSourceEvidenceRequirement>[0] & {
    sourceDependencyFingerprint?: string | null;
  },
  readFingerprint: (sheets: readonly string[]) => Promise<string | null>
): Promise<FormulaResumeSourceVerdict> {
  const requirement = formulaSourceEvidenceRequirement(plan);
  if (!requirement.ok) return { verdict: "refused", reason: requirement.reason };
  if (requirement.sheets.length === 0) return { verdict: "not-applicable" };
  const recorded =
    typeof plan?.sourceDependencyFingerprint === "string"
      ? plan.sourceDependencyFingerprint.trim()
      : "";
  if (!recorded) {
    return { verdict: "refused", reason: FORMULA_SOURCE_EVIDENCE_MISSING_MESSAGE };
  }
  let live: string | null = null;
  try {
    live = await readFingerprint(requirement.sheets);
  } catch {
    live = null;
  }
  const current = typeof live === "string" ? live.trim() : "";
  if (!current) {
    return { verdict: "refused", reason: FORMULA_SOURCE_UNVERIFIABLE_MESSAGE };
  }
  // 比的是**来源值**，不是公式文本：公式文本一个字符都没变时，来源值照样可能变
  // （`=数据!A1*2` 写下去之后 数据!A1 从 2 改成 99）。所以这里没有任何"公式文本相同"
  // 的旁路 —— 那会被读成"结果一致"，而公式结果恰恰取决于源值。
  return current === recorded
    ? { verdict: "ok", sheets: requirement.sheets }
    : { verdict: "refused", reason: FORMULA_SOURCE_CHANGED_MESSAGE };
}

/** 单格的现场归类。 */
export type FormulaCellSite = "before" | "written" | "indeterminate" | "unknown";

/**
 * 一格现在是"写前 / 写后 / 判不出来 / 都不是"。
 *
 * - `written`：与写后文本**逐字符相同**，且写前写后本来就不同；
 * - `before`：与写前文本相同，且写前写后本来就不同；
 * - `indeterminate`：写前与写后的公式文本本来就一样（例如把 `=A1` 再写成 `=A1`）——
 *   光看文本**无法**判断是否写过。**不许**报成 `written`；
 * - `unknown`：既不是写前也不是写后（被人改过、或写了别的东西）。
 *
 * 入参里刻意没有"计算值"：显示值相同不构成任何一类证据。
 */
export function formulaCellSite(before: unknown, written: unknown, observed: unknown): FormulaCellSite {
  const beforeText = isFormulaText(before) ? before : null;
  const writtenText = isFormulaText(written) ? written : null;
  const observedText = isFormulaText(observed) ? observed : null;
  if (beforeText === writtenText) {
    return observedText === beforeText ? "indeterminate" : "unknown";
  }
  if (observedText === writtenText) return "written";
  if (observedText === beforeText) return "before";
  return "unknown";
}

export type FormulaMatrixSiteKind = "before" | "written" | "partial" | "indeterminate" | "unknown";

export type FormulaMatrixSite = {
  kind: FormulaMatrixSiteKind;
  /** 逐格归类（与入参同形，便于界面/日志点名到格）。 */
  cells: FormulaCellSite[][];
  /** 第一个不是 `written` 也不在预期里的格（`kind` 为 `partial`/`unknown` 时给人看）。 */
  detail: { row: number; column: number; site: FormulaCellSite } | null;
};

/**
 * 整块区域的现场归类。
 *
 * 汇总口径（`partial` 就是"在途"最典型的样子）：
 * - 全部 `written` ⇒ `written`；
 * - 全部 `before` ⇒ `before`；
 * - 有 `unknown` ⇒ `unknown`（先怀疑现场，别猜）；
 * - 有 `indeterminate` 且没有 `unknown` ⇒ `indeterminate`（判不出来，也要停）；
 * - 其余（既有 written 又有 before）⇒ `partial`。
 *
 * 形状不一致一律 `unknown`：读回来都不是同一块区域，谈不上"写没写"。
 */
export function formulaMatrixSite(before: FormulaMatrix, written: FormulaMatrix, observed: FormulaMatrix): FormulaMatrixSite {
  const rows = written.length;
  const columns = written[0]?.length ?? 0;
  const sameShape = (matrix: FormulaMatrix) =>
    matrix.length === rows && matrix.every((line) => line.length === columns);
  if (!sameShape(before) || !sameShape(observed) || rows === 0 || columns === 0) {
    return { kind: "unknown", cells: [], detail: null };
  }
  const cells: FormulaCellSite[][] = [];
  let writtenCount = 0;
  let beforeCount = 0;
  let indeterminate = false;
  let unknown = false;
  let detail: FormulaMatrixSite["detail"] = null;
  for (let row = 0; row < rows; row += 1) {
    const line: FormulaCellSite[] = [];
    for (let column = 0; column < columns; column += 1) {
      const site = formulaCellSite(before[row][column], written[row][column], observed[row][column]);
      line.push(site);
      if (site === "written") writtenCount += 1;
      else if (site === "before") beforeCount += 1;
      else if (site === "indeterminate") indeterminate = true;
      else {
        unknown = true;
        detail ??= { row, column, site };
      }
    }
    cells.push(line);
  }
  const kind: FormulaMatrixSiteKind = unknown
    ? "unknown"
    : indeterminate
      ? "indeterminate"
      : writtenCount === rows * columns
        ? "written"
        : beforeCount === rows * columns
          ? "before"
          : "partial";
  if (!detail) {
    for (let row = 0; row < rows && !detail; row += 1) {
      const column = cells[row].findIndex((site) => site !== "written");
      if (column >= 0) detail = { row, column, site: cells[row][column] };
    }
  }
  return { kind, cells, detail };
}

/**
 * 续跑闸：只放行两种现场，其余一律停。
 *
 * - `before` ⇒ 这个动作还没写，可以按计划写；
 * - `written` ⇒ 已经写好，**跳过**，不重写；
 * - 其余（`partial` / `unknown` / `indeterminate`）⇒ 停下来核对，不给"继续"按钮。
 */
export function formulaResumeGate(site: FormulaMatrixSite): "write" | "skip" | "reconcile" {
  if (site.kind === "before") return "write";
  if (site.kind === "written") return "skip";
  return "reconcile";
}
