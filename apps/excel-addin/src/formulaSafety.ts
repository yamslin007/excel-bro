/**
 * 公式安全护栏（前端镜像 server/app/safety.py）
 * 写入工作簿前拦截危险公式函数，防止 LLM 生成/提示注入的公式外传数据或钓鱼。
 */

import capabilities from "../../../config/capabilities.json";

const DANGEROUS_FORMULA_FUNCTIONS: Array<{
  name: string;
  pattern: RegExp;
}> = [
  { name: "WEBSERVICE", pattern: /\bWEBSERVICE\s*\(/i },
  { name: "HYPERLINK", pattern: /\bHYPERLINK\s*\(/i },
  { name: "IMPORTXML", pattern: /\bIMPORTXML\s*\(/i },
  { name: "IMPORTDATA", pattern: /\bIMPORTDATA\s*\(/i },
  { name: "IMPORTHTML", pattern: /\bIMPORTHTML\s*\(/i },
  { name: "FILTERXML", pattern: /\bFILTERXML\s*\(/i },
  { name: "EXEC", pattern: /\bEXEC\s*\(/i },
  { name: "CALL", pattern: /\bCALL\s*\(/i },
  { name: "REGISTER", pattern: /\bREGISTER\s*\(/i }
];

const DDE_PATTERN = /^\s*=\s*cmd\s*[|'`]/i;

// 外部工作簿 / UNC 引用：函数名黑名单查不到，但打开工作簿或点击超链接时会让
// Excel/Windows 对 \\host\share 发起 SMB 连接，泄露 NTLM 哈希。
const UNC_PREFIX_PATTERN = /\\\\/;
const EXTERNAL_WORKBOOK_REF_PATTERN = /\[[^\]\s]+\.(?:xlsx|xlsm|xlsb|xls|csv|xlw)\]/i;
const HYPERLINK_NETWORK_PREFIX_PATTERN = /^(?:\\\\|\/\/)/i;
const FILE_SCHEME_PATTERN = /^file:/i;

/** 是否放行 HYPERLINK 公式函数（config safety.allowHyperlink）。默认拦截。 */
function hyperlinkEnabled(): boolean {
  const safety = (capabilities as { safety?: { allowHyperlink?: boolean } })
    .safety;
  return safety?.allowHyperlink === true;
}

/** 提取公式中所有外部工作簿引用（方括号内的文件名）。 */
function externalWorkbookFileNames(formula: string): string[] {
  const pattern = new RegExp(EXTERNAL_WORKBOOK_REF_PATTERN.source, "gi");
  const matches = formula.match(pattern);
  if (!matches) return [];
  return matches.map((ref) => ref.slice(1, -1));
}

/** 公式引用的外部工作簿文件名列表（供验算拦截判断是否外部引用导致报错）。 */
export function formulaExternalFileNames(formula: string): string[] {
  return externalWorkbookFileNames(formula);
}

/** 提取公式中 name( ... ) 的顶层参数（支持嵌套括号与字符串引号）。 */
function extractFunctionArgs(formula: string, name: string): string[][] {
  const regex = new RegExp(`\\b${name}\\s*\\(`, "gi");
  const results: string[][] = [];
  let match: RegExpExecArray | null;
  while ((match = regex.exec(formula)) !== null) {
    const args: string[] = [];
    let index = match.index + match[0].length;
    let depth = 1;
    let inString = false;
    let current = "";
    while (index < formula.length && depth > 0) {
      const ch = formula[index];
      if (inString) {
        if (ch === '"' && formula[index - 1] !== "\\") inString = false;
        current += ch;
      } else if (ch === '"') {
        inString = true;
        current += ch;
      } else if (ch === "(") {
        depth += 1;
        current += ch;
      } else if (ch === ")") {
        depth -= 1;
        if (depth === 0) {
          args.push(current.trim());
          break;
        }
        current += ch;
      } else if (ch === "," && depth === 1) {
        args.push(current.trim());
        current = "";
      } else {
        current += ch;
      }
      index += 1;
    }
    results.push(args);
  }
  return results;
}

function isExactMatchFlag(value: string | undefined): boolean {
  if (value === undefined) return false;
  const trimmed = value.trim();
  return trimmed === "FALSE" || trimmed === "0" || trimmed === "false";
}

/**
 * 检测公式中的近似/模糊匹配用法，返回命中的函数名；全部精确匹配则返回 null。
 * VLOOKUP/HLOOKUP 省略第四参数默认近似（TRUE）；MATCH 省略第三参数默认 1；
 * LOOKUP 本身只有近似语义；XLOOKUP 的 match_mode 为 1/-1/2 都是非精确。
 */
export function fuzzyLookupMatch(formula: string): string | null {
  if (!formula) return null;
  for (const name of ["VLOOKUP", "HLOOKUP", "LOOKUP", "MATCH", "XLOOKUP"]) {
    const calls = extractFunctionArgs(formula, name);
    for (const args of calls) {
      if (args.length < 2) continue;
      if (name === "VLOOKUP" || name === "HLOOKUP") {
        if (!isExactMatchFlag(args[3])) return name;
      } else if (name === "MATCH") {
        if (!isExactMatchFlag(args[2])) return name;
      } else if (name === "LOOKUP") {
        return name;
      } else if (name === "XLOOKUP") {
        if (args[4] !== undefined && !isExactMatchFlag(args[4])) return name;
      }
    }
  }
  return null;
}

/** 校验公式必须精确匹配，命中模糊匹配则抛错。label 用于错误提示定位。 */
export function assertExactLookup(formula: string, label: string): void {
  const matched = fuzzyLookupMatch(formula);
  if (matched !== null) {
    throw new Error(
      `${label}包含模糊匹配（${matched}）：近似匹配会静默返回错误数据，已拒绝；请改用精确匹配。`
    );
  }
}

/**
 * 公式含危险函数/注入载体时返回命中名称，否则返回 null。
 * allowHyperlink 不传时按配置决定；显式传值可覆盖（供测试）。
 * allowedExternal 为本次会话已勾选的外部工作簿文件名集合：方括号内文件名
 * 全部 ∈ allowedExternal 才放行外部引用，否则仍返回 "EXTERNAL_REF"。
 */
export function dangerousFormula(
  formula: string,
  allowHyperlink = hyperlinkEnabled(),
  allowedExternal?: ReadonlySet<string>
): string | null {
  if (!formula) return null;
  for (const { name, pattern } of DANGEROUS_FORMULA_FUNCTIONS) {
    if (name === "HYPERLINK" && allowHyperlink) continue;
    if (pattern.test(formula)) return name;
  }
  if (DDE_PATTERN.test(formula)) return "DDE";
  if (UNC_PREFIX_PATTERN.test(formula)) return "UNC";
  if (EXTERNAL_WORKBOOK_REF_PATTERN.test(formula)) {
    const refs = externalWorkbookFileNames(formula);
    const allAllowed =
      refs.length > 0 &&
      refs.every((fileName) => allowedExternal?.has(fileName) ?? false);
    if (!allAllowed) return "EXTERNAL_REF";
  }
  return null;
}

/** 超链接地址含命令注入载体/共享路径时返回命中名称，否则返回 null。 */
export function dangerousHyperlinkAddress(address: string): string | null {
  if (!address) return null;
  if (address.trimStart().startsWith("=")) return "DDE";
  if (DDE_PATTERN.test(address)) return "DDE";
  const stripped = address.trimStart();
  if (HYPERLINK_NETWORK_PREFIX_PATTERN.test(stripped)) return "UNC";
  if (FILE_SCHEME_PATTERN.test(stripped)) return "FILE";
  return null;
}

/** 校验单个公式，危险则抛错。label 用于错误提示定位。 */
export function assertSafeFormula(
  formula: string,
  label: string,
  allowedExternal?: ReadonlySet<string>
): void {
  const matched = dangerousFormula(formula, hyperlinkEnabled(), allowedExternal);
  if (matched !== null) {
    throw new Error(`${label}包含被禁用的函数：${matched}，已拒绝写入`);
  }
}

/** 校验超链接地址，危险则抛错。 */
export function assertSafeHyperlink(address: string, label: string): void {
  const matched = dangerousHyperlinkAddress(address);
  if (matched !== null) {
    throw new Error(`${label}包含被禁用的注入载体：${matched}`);
  }
}
