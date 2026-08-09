import { describe, it, expect } from 'vitest';
import {
  BUILTIN_RULES,
  getBuiltinFunction,
  isBuiltinRule,
  executeRule,
  getRuleExample,
  validateAlias
} from './ebRules';

describe('预制 EB 规则', () => {
  describe('隐形字符清理', () => {
    const fn = getBuiltinFunction('隐形字符清理')!;

    it('应该清除 CHAR(160)', () => {
      const input = 'hello\u00A0world';
      expect(fn(input)).toBe('hello world');
    });

    it('应该清除零宽空格', () => {
      const input = 'hello\u200Bworld';
      expect(fn(input)).toBe('helloworld');
    });

    it('应该清除控制字符', () => {
      const input = 'hello\u0001\u001Fworld';
      expect(fn(input)).toBe('helloworld');
    });

    it('应该清除方向标记', () => {
      const input = 'hello\u200Eworld\u200F';
      expect(fn(input)).toBe('helloworld');
    });

    it('应该合并多个连续空格', () => {
      const input = 'hello    world';
      expect(fn(input)).toBe('hello world');
    });

    it('应该去除前后空格', () => {
      const input = '  hello world  ';
      expect(fn(input)).toBe('hello world');
    });

    it('应该处理 null 和 undefined', () => {
      expect(fn(null)).toBe('');
      expect(fn(undefined)).toBe('');
    });

    it('应该处理混合隐形字符', () => {
      const input = '  hello\u00A0\u200B\u0001world  ';
      expect(fn(input)).toBe('hello world');
    });
  });

  describe('空白检测', () => {
    const fn = getBuiltinFunction('空白检测')!;

    it('应该识别 null', () => {
      expect(fn(null)).toBe(true);
    });

    it('应该识别 undefined', () => {
      expect(fn(undefined)).toBe(true);
    });

    it('应该识别空字符串', () => {
      expect(fn('')).toBe(true);
    });

    it('应该识别只有空格的字符串', () => {
      expect(fn('   ')).toBe(true);
    });

    it('应该识别只有隐形字符的字符串', () => {
      expect(fn('\u00A0\u200B')).toBe(true);
    });

    it('应该识别混合空格和隐形字符', () => {
      expect(fn('  \u00A0\u200B  ')).toBe(true);
    });

    it('应该识别非空白文本', () => {
      expect(fn('hello')).toBe(false);
    });

    it('应该识别包含文本的字符串', () => {
      expect(fn('  hello  ')).toBe(false);
    });

    it('应该识别数字 0', () => {
      expect(fn(0)).toBe(false);
    });
  });

  describe('预制规则配置', () => {
    it('应该有 2 个预制规则', () => {
      expect(BUILTIN_RULES).toHaveLength(2);
    });

    it('所有预制规则应该标记为只读', () => {
      BUILTIN_RULES.forEach(rule => {
        expect(rule.readonly).toBe(true);
        expect(rule.category).toBe('builtin');
      });
    });

    it('应该能通过名称识别预制规则', () => {
      expect(isBuiltinRule('隐形字符清理')).toBe(true);
      expect(isBuiltinRule('空白检测')).toBe(true);
      expect(isBuiltinRule('不存在的规则')).toBe(false);
    });

    it('应该能获取预制函数', () => {
      expect(getBuiltinFunction('隐形字符清理')).toBeTruthy();
      expect(getBuiltinFunction('空白检测')).toBeTruthy();
      expect(getBuiltinFunction('不存在的规则')).toBeNull();
    });
  });

  describe('executeRule', () => {
    it('应该执行隐形字符清理', () => {
      expect(executeRule('隐形字符清理', ['hello\u00A0world'])).toBe('hello world');
    });

    it('应该执行空白检测', () => {
      expect(executeRule('空白检测', ['   '])).toBe(true);
    });

    it('应该抛出不存在规则的错误', () => {
      expect(() => executeRule('不存在', [])).toThrow('规则 "不存在" 不存在');
    });
  });

  describe('getRuleExample', () => {
    it('有别名的规则应该生成独立函数示例', () => {
      const rule = BUILTIN_RULES.find(r => r.name === '隐形字符清理')!;
      expect(getRuleExample(rule)).toBe('=EB.CLEAN(A1)');
    });

    it('无别名的规则应该生成分发器示例', () => {
      const mockRule = {
        ...BUILTIN_RULES[0],
        name: '测试规则',
        alias: undefined,
        params: [
          { name: 'a', type: 'number' as const, required: true },
          { name: 'b', type: 'number' as const, required: true },
          { name: 'c', type: 'number' as const, required: true }
        ]
      };
      expect(getRuleExample(mockRule)).toBe('=EB("测试规则", A1, B1, C1)');
    });
  });

  describe('validateAlias', () => {
    it('应该接受合法别名', () => {
      expect(validateAlias('SCORE', [])).toBeNull();
      expect(validateAlias('my_func1', [])).toBeNull();
    });

    it('应该拒绝非法格式', () => {
      expect(validateAlias('1ABC', [])).toBeTruthy();
      expect(validateAlias('含中文', [])).toBeTruthy();
      expect(validateAlias('HAS-DASH', [])).toBeTruthy();
      expect(validateAlias('A'.repeat(31), [])).toBeTruthy();
    });

    it('应该拒绝与预制规则冲突的别名（忽略大小写）', () => {
      expect(validateAlias('clean', [])).toBeTruthy();
      expect(validateAlias('EMPTY', [])).toBeTruthy();
    });

    it('应该拒绝与现有规则冲突的别名', () => {
      const existing = [{ ...BUILTIN_RULES[0], alias: 'SCORE' }];
      expect(validateAlias('score', existing)).toBeTruthy();
    });
  });

  describe('依赖关系', () => {
    it('空白检测应该依赖隐形字符清理', () => {
      const rule = BUILTIN_RULES.find(r => r.name === '空白检测')!;
      expect(rule.dependencies).toContain('隐形字符清理');
    });

    it('其他规则不应该有依赖', () => {
      const cleanRule = BUILTIN_RULES.find(r => r.name === '隐形字符清理')!;
      expect(cleanRule.dependencies).toHaveLength(0);
    });
  });
});
