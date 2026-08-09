/**
 * EB 函数加载
 * 用户不再新增自定义规则；只提供内置预制规则。
 */

import { EBRule, BUILTIN_RULES } from './ebRules';

/**
 * 加载所有可用规则（仅内置预制规则）
 */
export async function loadRulesFromSheet(): Promise<Map<string, EBRule>> {
  const rules = new Map<string, EBRule>();
  for (const rule of BUILTIN_RULES) {
    rules.set(rule.name, rule);
  }
  return rules;
}
