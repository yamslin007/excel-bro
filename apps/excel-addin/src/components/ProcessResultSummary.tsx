import type { ProcessOutcomeInput } from "@excel-bro/core/processOutcome";
import { processOutcomeSummary } from "@excel-bro/core/processOutcome";
import "./processResultSummary.css";

/**
 * 结果摘要 —— **过程框之外**的那三行（Excel / WPS 共用）。
 *
 * 为什么必须在过程框外面：过程框回答"刚才在做什么"，一旦收起就只剩一行；
 * 用户真正要读的是"改了哪里 / 验证结果 / 下一步点哪"。把这三句留在会折叠的
 * 区域里，等于让用户为了看结论去展开一堆技术步骤。
 *
 * 三段内容全部由共享层从结构化事实派生（`processOutcome`）：
 *  - 改了哪里 ← 真实读回的前后差异；
 *  - 验证结果 ← 独立验收回执（没有就是没有，不写成"已验证"）；
 *  - 下一步 ← 当前阶段对应的那个真实入口。
 * 这里**不做任何文案判断**，只把不会为 null 的那几行画出来。
 */
export function ProcessResultSummary({
  input,
  className = ""
}: {
  input: ProcessOutcomeInput;
  className?: string;
}) {
  const summary = processOutcomeSummary(input);
  if (!summary) return null;
  const { changedWhere, verification, nextStep } = summary;
  return (
    <section
      className={`process-result-summary ${className}`.trim()}
      aria-label="本次结果摘要"
      data-outcome-phase={input.phase}
    >
      {changedWhere && (
        <p data-result-line="changed">{changedWhere}</p>
      )}
      {verification && (
        <p data-result-line="verification">{verification}</p>
      )}
      <p data-result-line="next">{nextStep}</p>
    </section>
  );
}

export default ProcessResultSummary;
