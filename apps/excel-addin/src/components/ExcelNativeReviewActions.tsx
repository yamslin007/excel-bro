import type { AnalysisPlan } from "../contracts";
import type { ExcelNativeComparison } from "../excelNativeComparison";
import type { useExcelNativeComparison } from "../hooks/useExcelNativeComparison";
import { useState } from "react";
import { excelNativeReviewMode } from "../excelNativeComparison";
import capabilities from "../../../../config/capabilities.json";

export function ExcelNativeReviewActions({ messageId, plan, authorization, controller, disabled, commit, paneNeedsResize = false, beforeOperation }: {
  messageId: string; plan: AnalysisPlan; authorization: readonly string[];
  controller: ReturnType<typeof useExcelNativeComparison>; disabled: boolean;
  commit: (session: ExcelNativeComparison) => Promise<boolean>;
  paneNeedsResize?: boolean;
  beforeOperation?: () => Promise<boolean>;
}) {
  const [acknowledgedCopy, setAcknowledgedCopy] = useState<string | null>(null);
  const review = controller.review?.messageId === messageId ? controller.review : null;
  const feedback = controller.feedback?.messageId === messageId ? controller.feedback.text : null;
  const manual = review?.session.mode === "manual";
  const manualReviewed = Boolean(review && acknowledgedCopy === review.session.copySheetName);
  const mode = excelNativeReviewMode();
  return <section className="excel-native-review" aria-label="Excel 原生并排审阅">
    {review ? <>
      {paneNeedsResize && review.phase === "review" && <p role="status">AI 窗格自动比例尚未确认。请拖动「Excel Bro」左侧分隔线，将 AI 调到右侧 Excel 窗口约 {Math.round(capabilities.reviewPaneSizing.reviewTaskPaneWidthFraction * 100)}%，为原表留出空间。若刚更新代码，请先重启 Python 服务并刷新加载项。</p>}
      <p>{review.phase === "review" ? manual
        ? `原表未修改 · 已请求打开兼容副本，工作表「${review.session.copySheetName}」，标记 ${review.session.changedCells} 处计划变化。${review.session.windowLayout === 'arranged' ? '已自动排列：左侧审阅副本，右侧原表与 AI 对话；请核对副本。' : '自动排列未完成。请切到副本，在 Excel「视图 → 并排查看」中选择原工作簿；若窗口分属不同实例，请手动将副本放左侧、原表与 AI 放右侧。'}`
        : `原表未修改 · 独立副本已标记 ${review.session.changedCells} 处变化。请在 Excel 原生并排窗口中核对。` : review.phase === "committed" ? "本次计划已写回原表，不可重复提交。" : "审阅已停止。请核对执行结果，再关闭副本并重新生成预览。"}</p>
      {manual && <p>公式由 Excel 打开副本后重算；加载项不能自动回读或关闭兼容副本。副本内的人工修改不会写回，保存只执行本卡已审阅的计划，随后验证原表实际结果。</p>}
      {review.phase === "review" && review.session.windowLayout === 'arranged' && <p>{review.session.synchronousScrolling === 'enabled'
        ? '已启用 Excel 原生并排比较与同步滚动，拖动任一表的滚动条可联动查看。'
        : review.session.synchronousScrolling === 'different_instance'
          ? '两份工作簿属于不同 Excel 实例：目前仅贴齐排列，未启用同步滚动。请将审阅副本保存到本机后，在原表所在 Excel 实例中打开，再选择「视图 → 并排查看 → 同步滚动」。不要将副本另存为原文件。'
          : '窗口已贴齐排列，但同步滚动未能确认。请在 Excel「视图 → 并排查看」中选择本次副本，并启用「同步滚动」。若还有其他工作簿打开，请先保存并关闭无关工作簿再生成副本。'}</p>}
      {manual && review.session.windowLayout === 'service_outdated' && <p role="status">本地窗口排列服务版本过旧。请停止并重启 Python 服务、刷新加载项，关闭旧副本且不保存，再重新生成副本；新版排列为左侧副本、右侧原表与 AI。</p>}
      {manual && review.phase === "review" && <label className="excel-manual-review-confirm">
        <input type="checkbox" checked={manualReviewed} disabled={disabled}
          onChange={event => setAcknowledgedCopy(event.target.checked ? review.session.copySheetName : null)} />
        我已在 Excel 中核对副本，确认按此计划写回原表
      </label>}
      <div className="excel-native-review-buttons">
        <button className="secondary-button" disabled={disabled} onClick={() => void controller.discard(messageId)}>{review.phase === "review" ? "放弃变更" : "关闭副本"}</button>
        {review.phase === "review" && <button className="primary-button" disabled={disabled || (manual && !manualReviewed)} onClick={() => void (async () => {
          if (beforeOperation && !(await beforeOperation())) return;
          await controller.save(messageId, plan, authorization, commit, manualReviewed);
        })()}>保存变更到原表</button>}
      </div>
    </> : <>
      <button className="primary-button" disabled={disabled || Boolean(controller.review) || mode === "unavailable"} onClick={() => void (async () => {
        if (beforeOperation && !(await beforeOperation())) return;
        await controller.stage(messageId, plan, authorization);
      })()}>生成副本并排审阅</button>
      <p>原表保持不变，副本试算并标记变化；确认后仅写回已审阅的计划，不导入副本中的人工修改。</p>
      {mode === "manual" && <p>本地服务将尝试把副本放左侧、原表与 AI 放右侧。窗口无法确认时会提示手动并排；审阅后勾选确认，再写回原表。</p>}
      {mode === "unavailable" && <p>当前 Excel 缺少创建兼容副本所需的基础 API，未生成副本、未写入原表。</p>}
    </>}
    {feedback && <p role="status">{feedback}</p>}
  </section>;
}
