import { useMemo, useState } from "react";
import type { BlankPolicy, CellValue } from "@excel-bro/core";
import type { ColumnMappingRule } from "../contracts";
import {
  addExactMapping,
  addSplitMapping,
  applyBlankPolicy,
  localMappingSourceKey,
  mappingTargets,
  suggestMappingDelimiter,
  uncoveredSplitItems
} from "../localMappingResolution";

/**
 * 界面上的四种选择 → 结构化空白策略。
 *
 * 「写入空字符串」也是 `kind:"value"`，只是 `value` 是 `""` —— 与 `clear`
 * 不是一回事：前者留下常量空字符串（`ISBLANK()` 为假），后者是真空单元格。
 * 用两个不同的选项表达，正是为了让这个差别可见。
 */
function blankPolicyOf(
  choice: "keep" | "clear" | "value" | "empty",
  value: string
): BlankPolicy {
  switch (choice) {
    case "keep":
      return { kind: "keep" };
    case "clear":
      return { kind: "clear" };
    case "empty":
      return { kind: "value", value: "" };
    case "value":
      return { kind: "value", value };
  }
}

/** 待答来源值与补充规则只在当前面板内存中；这里不把它们挂到澄清消息上。 */
export function LocalMappingResolution(props: {
  rule: ColumnMappingRule;
  values: readonly CellValue[];
  allowExactMapping: boolean;
  busy: boolean;
  /**
   * 停在哪种缺口上。`"blank"` 是**读到真实空白之后**才出现的，
   * 与类别/分隔符那两类**不共用一组选项**。
   */
  reason?: "unmapped" | "separator_or_mapping" | "blank";
  /** `reason === "blank"` 时真实读到的空白格数量。 */
  blankCount?: number;
  onResolve: (rule: ColumnMappingRule) => Promise<void>;
  onCancel: () => void;
}) {
  const reason = props.reason ?? "unmapped";
  const [mode, setMode] = useState<"exact" | "split">(
    props.allowExactMapping ? "exact" : "split"
  );
  const [blankChoice, setBlankChoice] = useState<
    "keep" | "clear" | "value" | "empty"
  >("keep");
  const [blankValue, setBlankValue] = useState("");
  const [exactTargets, setExactTargets] = useState<Record<number, string>>({});
  const [delimiter, setDelimiter] = useState(() =>
    suggestMappingDelimiter(props.rule, props.values)
  );
  const [splitTargets, setSplitTargets] = useState<Record<string, string>>({});
  const [priorityOrder, setPriorityOrder] = useState<string[]>(() => mappingTargets(props.rule));
  const [priorityConfirmed, setPriorityConfirmed] = useState(false);
  const [error, setError] = useState("");
  const missingItems = useMemo(
    () => uncoveredSplitItems(props.rule, props.values, delimiter),
    [props.rule, props.values, delimiter]
  );
  const targets = useMemo(() => {
    const all = [
      ...mappingTargets(props.rule),
      ...missingItems.map((item) => splitTargets[item]?.trim() ?? "").filter(Boolean)
    ];
    const unique = [...new Set(all)];
    return [
      ...priorityOrder.filter((target) => unique.includes(target)),
      ...unique.filter((target) => !priorityOrder.includes(target))
    ];
  }, [props.rule, missingItems, splitTargets, priorityOrder]);

  const submit = async () => {
    let rule: ColumnMappingRule;
    try {
      rule = reason === "blank"
        ? applyBlankPolicy(props.rule, blankPolicyOf(blankChoice, blankValue))
        : mode === "exact"
        ? addExactMapping(
            props.rule,
            props.values,
            props.values.map((_, index) => exactTargets[index] ?? "")
          )
        : addSplitMapping(props.rule, {
            delimiter,
            missingItems,
            targets: missingItems.map((item) => splitTargets[item] ?? ""),
            priorityTargets: priorityConfirmed ? targets : []
          });
    } catch (failure) {
      // 这里只显示本地输入校验错误，不能把宿主异常原文混进持久化界面。
      setError(failure instanceof Error ? failure.message : "请检查补充的规则。");
      return;
    }
    setError("");
    try {
      await props.onResolve(rule);
    } catch {
      // 宿主错误可能含源单元格原文，不在表单里显示 Error.message。
      setError("重新读取或生成预览失败，原表未写入。请检查授权范围后重试。");
    }
  };

  return (
    <div className="wps-local-mapping-values local-mapping-resolution">
      {reason === "blank" ? (
        <>
          <p>
            读源区域时发现 <strong>{props.blankCount ?? 0}</strong> 个空白格。
            这些位置的目标格应当怎么处理？四种做法互不相同，选定后重新生成预览。
          </p>
          <label className="editable">
            <input
              type="radio"
              name="blank-policy"
              checked={blankChoice === "keep"}
              disabled={props.busy}
              onChange={() => setBlankChoice("keep")}
            />
            <span>保留目标原值 —— 这些位置不写入，公式、值与格式都不动</span>
          </label>
          <label className="editable">
            <input
              type="radio"
              name="blank-policy"
              checked={blankChoice === "clear"}
              disabled={props.busy}
              onChange={() => setBlankChoice("clear")}
            />
            <span>清空目标内容 —— 清成真空单元格，不是写一个空字符串</span>
          </label>
          <label className="editable">
            <input
              type="radio"
              name="blank-policy"
              checked={blankChoice === "value"}
              disabled={props.busy}
              onChange={() => setBlankChoice("value")}
            />
            <span>写入下面填的指定值</span>
          </label>
          {blankChoice === "value" && (
            <label className="editable">
              <span>空白格写入的值</span>
              <input
                aria-label="空白格写入的值"
                type="text"
                value={blankValue}
                disabled={props.busy}
                onChange={(event) => setBlankValue(event.target.value)}
                placeholder="例如：未分类"
              />
            </label>
          )}
          <label className="editable">
            <input
              type="radio"
              name="blank-policy"
              checked={blankChoice === "empty"}
              disabled={props.busy}
              onChange={() => setBlankChoice("empty")}
            />
            <span>写入空字符串 —— 明确写一个空字符串常量，与「清空」不同</span>
          </label>
          {error && <p role="alert" className="local-mapping-error">{error}</p>}
          <div className="clarification-actions">
            <button
              type="button"
              disabled={props.busy || (blankChoice === "value" && blankValue.trim() === "")}
              onClick={() => void submit()}
            >
              按选定的空白处理重新生成预览
            </button>
            <button type="button" disabled={props.busy} onClick={props.onCancel}>暂不修改</button>
          </div>
        </>
      ) : mode === "exact" ? (
        <>
          <p>逐项精确映射：每种当前来源值分别指定目标；后续出现新组合仍会再次询问。</p>
          {props.values.map((value, index) => (
            <label key={index} className="editable">
              <span>来源值：{localMappingSourceKey(value)}</span>
              <input
                aria-label={`来源值 ${localMappingSourceKey(value)} 的目标值`}
                type="text"
                value={exactTargets[index] ?? ""}
                placeholder="填写该整格对应的目标值"
                disabled={props.busy}
                onChange={(event) => setExactTargets((current) => ({ ...current, [index]: event.target.value }))}
              />
            </label>
          ))}
        </>
      ) : (
        <>
          <p>拆分映射：新增你确认的分隔符，并保留规则已有的分隔符；多个类别按你确认的优先顺序计算，不会自动猜测。</p>
          <label className="editable">
            <span>分隔符（候选需你确认）</span>
            <input
              aria-label="拆分符"
              type="text"
              value={delimiter}
              disabled={props.busy}
              onChange={(event) => { setDelimiter(event.target.value); setPriorityConfirmed(false); }}
              placeholder="输入实际使用的分隔符"
            />
          </label>
          {missingItems.map((item) => (
            <label key={item} className="editable">
              <span>未覆盖的拆分项：{item}</span>
              <input
                aria-label={`拆分项 ${item} 的目标值`}
                type="text"
                value={splitTargets[item] ?? ""}
                disabled={props.busy}
                onChange={(event) => {
                  setSplitTargets((current) => ({ ...current, [item]: event.target.value }));
                  setPriorityConfirmed(false);
                }}
                placeholder="填写此项的目标值"
              />
            </label>
          ))}
          <div className="local-mapping-priority" aria-label="目标优先顺序">
            <strong>多个类别同时出现时，从高到低取目标：</strong>
            <span>初始顺序只是草稿，不代表业务上的高低；请按实际规则调整后再确认。</span>
            {targets.map((target, index) => (
              <div key={target} className="local-mapping-priority-item">
                <span>{index + 1}. {target}</span>
                <button type="button" disabled={props.busy || index === 0} onClick={() => {
                  const next = [...targets];
                  [next[index - 1], next[index]] = [next[index], next[index - 1]];
                  setPriorityOrder(next);
                  setPriorityConfirmed(false);
                }}>上移</button>
                <button type="button" disabled={props.busy || index === targets.length - 1} onClick={() => {
                  const next = [...targets];
                  [next[index], next[index + 1]] = [next[index + 1], next[index]];
                  setPriorityOrder(next);
                  setPriorityConfirmed(false);
                }}>下移</button>
              </div>
            ))}
            <label>
              <input type="checkbox" checked={priorityConfirmed} disabled={props.busy || targets.length === 0} onChange={(event) => setPriorityConfirmed(event.target.checked)} />
              我已核对上述优先顺序（最上方优先）
            </label>
          </div>
        </>
      )}
      {/*
       * 空白那一支自带错误提示与操作区。**不能**再叠一套"按补充映射重新生成
       * 预览"：那张卡上没有任何映射表缺口，两个提交按钮并排会让用户不知道
       * 自己按的是哪一种补救，也容易误以为"补映射"能顺带处理空白。
       */}
      {reason !== "blank" && (
        <>
          {error && <p role="alert" className="local-mapping-error">{error}</p>}
          <div className="clarification-actions">
            <button type="button" disabled={props.busy || (mode === "split" && (!delimiter || !priorityConfirmed))} onClick={() => void submit()}>
              按补充映射重新生成预览
            </button>
            {mode === "exact" ? (
              <button type="button" disabled={props.busy} onClick={() => { setMode("split"); setError(""); }}>改为说明拆分规则</button>
            ) : props.allowExactMapping ? (
              <button type="button" disabled={props.busy} onClick={() => { setMode("exact"); setError(""); }}>将整格分别映射</button>
            ) : null}
            <button type="button" disabled={props.busy} onClick={props.onCancel}>暂不修改</button>
          </div>
        </>
      )}
    </div>
  );
}
