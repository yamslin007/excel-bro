import type { ChatMessage } from "../types/chat";

type ScopeSelectionGuidanceData = NonNullable<
  ChatMessage["scopeSelectionGuidance"]
>;

export function ScopeSelectionGuidance({
  keeperName,
  missingSheetNames
}: ScopeSelectionGuidanceData) {
  const missingPreview = missingSheetNames.slice(0, 6);
  const hasMore = missingSheetNames.length > missingPreview.length;

  return (
    <section className="scope-selection-guidance" aria-label="工作表范围操作引导">
      <p>当前授权范围不包含全部工作表，因此不能删除尚未授权的工作表。</p>
      <p>
        请点击输入框上方的
        <span className="scope-guidance-target">
          <i aria-hidden="true">▦</i>
          <strong>工作表范围</strong>
        </span>
        ，勾选当前工作簿的全部工作表，然后再次发送这条指令。请同时勾选需要保留的
        「{keeperName}」。
      </p>
      <p className="scope-guidance-missing">
        <strong>尚未授权：</strong>
        {missingPreview.map((name) => `「${name}」`).join("、")}
        {hasMore ? `等 ${missingSheetNames.length} 张工作表` : ""}。
      </p>
      <p>本次未生成执行计划，也未删除任何工作表。</p>
    </section>
  );
}
