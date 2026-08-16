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

/**
 * 公式含危险函数/注入载体时返回命中名称，否则返回 null。
 * allowHyperlink 不传时按配置决定；显式传值可覆盖（供测试）。
 */
export function dangerousFormula(
  formula: string,
  allowHyperlink = hyperlinkEnabled()
): string | null {
  if (!formula) return null;
  for (const { name, pattern } of DANGEROUS_FORMULA_FUNCTIONS) {
    if (name === "HYPERLINK" && allowHyperlink) continue;
    if (pattern.test(formula)) return name;
  }
  if (DDE_PATTERN.test(formula)) return "DDE";
  if (UNC_PREFIX_PATTERN.test(formula)) return "UNC";
  if (EXTERNAL_WORKBOOK_REF_PATTERN.test(formula)) return "EXTERNAL_REF";
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
export function assertSafeFormula(formula: string, label: string): void {
  const matched = dangerousFormula(formula);
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
