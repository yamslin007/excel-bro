/**
 * EB 函数规则系统 - 类型定义和预制规则
 */

export type EBParamType = 'text' | 'number' | 'any';
export type EBCategory = 'builtin';

export interface EBParam {
  name: string;
  type: EBParamType;
  required: boolean;
  description?: string;
}

export interface EBRule {
  // 基础信息
  id: string;
  name: string;
  alias?: string;          // 英文别名：预制规则会注册为独立函数 EB.<alias>
  description: string;
  category: EBCategory;

  // 参数定义
  params: EBParam[];

  // 执行逻辑
  logic: string;           // 原始表达式（用于显示）
  compiled: string;        // 编译后的函数体

  // 依赖关系
  dependencies: string[];  // 依赖的其他规则名称

  // 元数据
  createdAt: string;       // ISO 8601
  updatedAt: string;
  version: number;
  readonly: boolean;       // 预制规则不可修改
}

/**
 * 隐形字符清理
 * 清除 CHAR(160)、零宽空格、控制字符、方向标记
 */
function cleanInvisibleChars(text: any): string {
  if (text === null || text === undefined) return '';

  return String(text)
    // 控制字符 (ASCII 0-31, 127, 128-159)
    .replace(/[\u0000-\u001F\u007F-\u009F]/g, '')
    // 零宽字符（直接删除，不替换为空格）
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    // 方向标记
    .replace(/[\u200E\u200F\u202A-\u202E]/g, '')
    // 各种空格字符 → 普通空格
    .replace(/[\u00A0\u1680\u2000-\u2009\u202F\u205F\u3000]/g, ' ')
    // 去除前后空格
    .trim()
    // 多个连续空格 → 单个空格
    .replace(/\s+/g, ' ');
}

/**
 * 空白检测
 * 检测伪空白（看着空但实际不空的单元格）
 */
function isBlankCell(value: any): boolean {
  if (value === null || value === undefined) return true;
  if (value === '') return true;

  // 清理后检查
  const cleaned = cleanInvisibleChars(value);
  return cleaned === '';
}

/**
 * 预制规则定义
 */
export const BUILTIN_RULES: EBRule[] = [
  {
    id: 'builtin-clean-invisible',
    name: '隐形字符清理',
    alias: 'CLEAN',
    description: '清除 CHAR(160)、零宽空格、控制字符、方向标记等隐形字符',
    category: 'builtin',
    params: [
      {
        name: 'text',
        type: 'text',
        required: true,
        description: '要清理的文本'
      }
    ],
    logic: 'cleanInvisibleChars(text)',
    compiled: cleanInvisibleChars.toString(),
    dependencies: [],
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    version: 1,
    readonly: true
  },
  {
    id: 'builtin-is-blank',
    name: '空白检测',
    alias: 'EMPTY',
    description: '检测伪空白单元格（空格、公式空字符串、残留标记）',
    category: 'builtin',
    params: [
      {
        name: 'value',
        type: 'any',
        required: true,
        description: '要检测的单元格值'
      }
    ],
    logic: 'isBlankCell(value)',
    compiled: isBlankCell.toString(),
    dependencies: ['隐形字符清理'], // 依赖清理函数
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    version: 1,
    readonly: true
  }
];

/**
 * 预制函数映射表
 */
export const BUILTIN_FUNCTIONS: Record<string, Function> = {
  '隐形字符清理': cleanInvisibleChars,
  '空白检测': isBlankCell
};

/**
 * 从规则名称获取预制函数
 */
export function getBuiltinFunction(ruleName: string): Function | null {
  return BUILTIN_FUNCTIONS[ruleName] || null;
}

/**
 * 检查是否为预制规则
 */
export function isBuiltinRule(ruleName: string): boolean {
  return ruleName in BUILTIN_FUNCTIONS;
}

/**
 * 英文别名校验
 * 别名会注册为独立函数 EB.<alias>，需满足 Excel 函数名规范
 */
export const ALIAS_PATTERN = /^[A-Za-z][A-Za-z0-9_]{0,29}$/;

/**
 * 校验别名是否可用，返回错误信息；合法返回 null
 */
export function validateAlias(alias: string, existingRules: EBRule[]): string | null {
  if (!ALIAS_PATTERN.test(alias)) {
    return '别名需以字母开头，只能包含字母、数字、下划线（最长 30 字符）';
  }

  const upper = alias.toUpperCase();
  const conflict = [...BUILTIN_RULES, ...existingRules].some(
    r => r.alias && r.alias.toUpperCase() === upper
  );
  if (conflict) {
    return `别名 "${alias}" 已被占用`;
  }

  return null;
}

/**
 * 执行规则
 */
export function executeRule(ruleName: string, args: any[]): any {
  const fn = getBuiltinFunction(ruleName);
  if (!fn) {
    throw new Error(`规则 "${ruleName}" 不存在`);
  }

  try {
    return fn(...args);
  } catch (error) {
    throw new Error(`执行规则 "${ruleName}" 失败: ${error}`);
  }
}

/**
 * 获取规则示例公式
 * 有别名的规则优先显示独立函数形式 =EB.ALIAS(...)，否则用分发器形式
 */
export function getRuleExample(rule: EBRule): string {
  const paramPlaceholders = rule.params.map((p, i) =>
    String.fromCharCode(65 + i) + '1'  // A1, B1, C1...
  ).join(', ');

  if (rule.alias) {
    return `=EB.${rule.alias}(${paramPlaceholders})`;
  }

  return `=EB("${rule.name}"${paramPlaceholders ? ', ' + paramPlaceholders : ''})`;
}
