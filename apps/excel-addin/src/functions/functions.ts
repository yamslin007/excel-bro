/**
 * Excel Bro Custom Functions
 * EB() 函数的 Custom Functions runtime
 */

/// <reference path="./custom-functions.d.ts" />

import { loadRulesFromSheet } from '../ebStorage';
import { getBuiltinFunction, EBRule, BUILTIN_RULES } from '../ebRules';

// 规则缓存
let rulesCache: Map<string, EBRule> | null = null;
let lastLoadTime = 0;
const CACHE_TTL = 5000; // 5 秒缓存

/**
 * 加载规则（带缓存）
 */
async function loadRules(): Promise<Map<string, EBRule>> {
  const now = Date.now();

  if (rulesCache && (now - lastLoadTime) < CACHE_TTL) {
    return rulesCache;
  }

  rulesCache = await loadRulesFromSheet();
  lastLoadTime = now;
  return rulesCache;
}

/**
 * 执行规则
 */
function executeRuleLogic(rule: EBRule, args: any[]): any {
  // 仅执行内置预制规则。曾经存在"用户自定义规则"分支（用 new Function 求值
  // rule.compiled），已成为不可达死代码并构成任意代码执行隐患，已移除。
  const builtinFn = getBuiltinFunction(rule.name);
  if (builtinFn) {
    return builtinFn(...args);
  }
  throw new Error(`规则 "${rule.name}" 不是内置预制规则，无法执行`);
}

/**
 * EB 主函数
 * @customfunction EB
 * @param {string} ruleName 规则名称
 * @param {any[][]} args 规则参数
 * @returns {any} 计算结果
 */
async function EB(ruleName: string, ...args: any[]): Promise<any> {
  try {
    // 加载规则
    const rules = await loadRules();

    // 查找规则
    const rule = rules.get(ruleName);
    if (!rule) {
      return `#ERROR: 规则 "${ruleName}" 不存在`;
    }

    // 验证参数数量
    const requiredParams = rule.params.filter(p => p.required);
    if (args.length < requiredParams.length) {
      return `#ERROR: 规则 "${ruleName}" 需要至少 ${requiredParams.length} 个参数`;
    }

    // 执行规则
    return executeRuleLogic(rule, args);
  } catch (error) {
    return `#ERROR: ${error}`;
  }
}

/**
 * 注册别名函数
 * 注意：associate 的 key 必须与 functions.json 里的 id 一致（如 EBCLEAN），不是显示名 EB.CLEAN
 */
function registerAliasFunction(rule: EBRule, fn: (...args: any[]) => any) {
  if (!rule.alias) return;
  CustomFunctions.associate(`EB${rule.alias.toUpperCase()}`, fn);
}

/**
 * 注册自定义函数
 */
if (typeof CustomFunctions !== 'undefined') {
  CustomFunctions.associate('EB', EB);

  // 预制规则注册为独立函数 EB.<alias>，支持 =EB. 下拉补全
  BUILTIN_RULES.forEach(rule => {
    const fn = rule.alias ? getBuiltinFunction(rule.name) : null;
    if (fn) {
      registerAliasFunction(rule, (...args: any[]) => fn(...args));
    }
  });
}
