/**
 * EB 函数说明（只读）
 * 展示内置的预制函数及其用法，用户在单元格用 =EB(...) 调用。
 */

import { useState, useEffect } from 'react';
import { loadRulesFromSheet } from './ebStorage';
import { EBRule, getRuleExample } from './ebRules';

interface RuleManagerProps {
  isOpen: boolean;
  onClose: () => void;
}

export function RuleManager({ isOpen, onClose }: RuleManagerProps) {
  const [rules, setRules] = useState<Map<string, EBRule>>(new Map());
  const [searchQuery, setSearchQuery] = useState('');

  useEffect(() => {
    if (isOpen) {
      loadRules();
    }
  }, [isOpen]);

  async function loadRules() {
    try {
      const loadedRules = await loadRulesFromSheet();
      setRules(loadedRules);
    } catch (error) {
      console.error('加载函数失败:', error);
    }
  }

  const filteredRules = Array.from(rules.values()).filter(rule =>
    rule.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
    rule.description.toLowerCase().includes(searchQuery.toLowerCase())
  );

  if (!isOpen) return null;

  return (
    <aside className="tool-drawer" aria-label="EB 函数说明">
      <div className="tool-drawer-header">
        <div>
          <strong>函数说明</strong>
          <span>共 {rules.size} 个内置函数，在单元格用 =EB(...) 调用</span>
        </div>
        <button onClick={onClose} aria-label="关闭">
          ×
        </button>
      </div>

      <div className="tool-library rule-manager-body">
        <div className="rule-toolbar">
          <input
            type="text"
            className="rule-search"
            placeholder="搜索函数名称或描述"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
          />
        </div>

        {filteredRules.length > 0 ? (
          <section className="rule-section">
            <h3 className="rule-section-title">内置函数（{filteredRules.length}）</h3>
            {filteredRules.map(rule => (
              <RuleCard key={rule.id} rule={rule} />
            ))}
          </section>
        ) : (
          <div className="tool-empty">
            <i>▦</i>
            <strong>{searchQuery ? '未找到匹配的函数' : '暂无函数'}</strong>
            <span>这些是内置函数，可在单元格用 =EB(...) 调用。</span>
          </div>
        )}

        <p className="rule-tip">提示：复杂文本处理用 =EB(...) 内置函数；一般计算请用 /function 生成原生公式。</p>
      </div>
    </aside>
  );
}

// 函数卡片（只读）
function RuleCard({ rule }: { rule: EBRule }) {
  const example = getRuleExample(rule);

  return (
    <article className="rule-card">
      <div className="rule-card-head">
        <div className="rule-card-title">
          <strong>{rule.name}</strong>
          {rule.readonly && <span className="rule-tag">内置</span>}
        </div>
      </div>

      <p className="rule-card-desc">{rule.description}</p>

      <code className="rule-example">{example}</code>

      {rule.dependencies.length > 0 && (
        <div className="rule-deps">
          <span>依赖：</span>
          {rule.dependencies.map(dep => (
            <span key={dep} className="rule-tag rule-tag-dep">
              {dep}
            </span>
          ))}
        </div>
      )}
    </article>
  );
}
