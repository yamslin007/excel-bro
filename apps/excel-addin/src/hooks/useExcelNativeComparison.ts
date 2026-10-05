import { useRef, useState } from "react";
import type { AnalysisPlan } from "../contracts";
import {
  assertExcelNativeComparisonUnchanged, closeExcelNativeComparison,
  stageExcelNativeComparison, type ExcelNativeComparison
} from "../excelNativeComparison";

type Review = { messageId: string; session: ExcelNativeComparison; phase: "review" | "committed" | "blocked" };
type Feedback = { messageId: string; text: string } | null;
export function useExcelNativeComparison() {
  // Never attach snapshots to ChatMessage/history/storage. One live copy per panel.
  const live = useRef<Review | null>(null);
  const running = useRef(false);
  const [review, setReview] = useState<Review | null>(null);
  const [pending, setPending] = useState(false);
  const [feedback, setFeedback] = useState<Feedback>(null);
  const publish = (value: Review | null) => { live.current = value; setReview(value); };
  async function operation(messageId: string, job: () => Promise<void>) {
    if (running.current) return;
    running.current = true; setPending(true); setFeedback(null);
    try { await job(); } catch (reason) {
      setFeedback({ messageId, text: reason instanceof Error ? reason.message : "原生并排审阅失败。" });
    } finally { running.current = false; setPending(false); }
  }
  async function stage(messageId: string, plan: AnalysisPlan, authorization: readonly string[]) {
    return operation(messageId, async () => {
      if (live.current) throw new Error("请先保存或放弃上一份审阅副本，再生成新的副本。");
      let session: ExcelNativeComparison;
      try { session = await stageExcelNativeComparison(plan, authorization); }
      catch (reason) { throw new Error(`${reason instanceof Error ? reason.message : "副本打开失败。"} 原表未写入；若临时副本已打开，请关闭且不保存后重试。`); }
      publish({ messageId, session, phase: "review" });
    });
  }
  async function discard(messageId: string) {
    return operation(messageId, async () => {
      const current = live.current;
      if (!current || current.messageId !== messageId) throw new Error("这张卡片没有活跃审阅副本。");
      if (current.session.mode === "manual") {
        publish(null);
        setFeedback({ messageId, text: current.phase === "blocked"
          ? "已结束本次兼容审阅。关闭副本不撤销既有写入，请按执行结果与恢复点核对原表；请手动关闭临时副本且不保存。"
          : "已放弃本次兼容审阅，未写回原表。请手动关闭临时副本且不保存；加载项不能自动关闭此窗口。" });
        return;
      }
      if (!await closeExcelNativeComparison(current.session)) throw new Error("副本尚未关闭；如 Excel 询问是否保存临时副本，请选择不保存。原表未因此改变。");
      publish(null);
      setFeedback({ messageId, text: current.phase === "committed" ? "原表已写回，临时副本已关闭。" : current.phase === "blocked" ? "临时副本已关闭。关闭副本不撤销既有写入，请按执行结果与恢复点核对原表。" : "已放弃审阅副本，未写回原表。" });
    });
  }
  async function save(messageId: string, plan: AnalysisPlan, authorization: readonly string[],
    commit: (session: ExcelNativeComparison) => Promise<boolean>, manualReviewed = false) {
    return operation(messageId, async () => {
      const current = live.current;
      if (!current || current.messageId !== messageId || current.phase !== "review") throw new Error("没有可保存的已审阅副本。");
      if (current.session.mode === "manual" && !manualReviewed) throw new Error("请先在 Excel 中审阅兼容副本，并勾选确认；原表未写入。");
      try {
        await Excel.run(context => assertExcelNativeComparisonUnchanged(context, current.session, authorization, plan));
        const completed = await commit(current.session);
        if (!completed) { publish({ ...current, phase: "blocked" }); throw new Error("写回未完成，不能重复提交。请查看执行结果与恢复点，核对原表后再生成新预览。"); }
      } catch (reason) { publish({ ...current, phase: "blocked" }); throw reason; }
      publish({ ...current, phase: "committed" });
      if (current.session.mode === "manual") {
        publish(null);
        setFeedback({ messageId, text: "变更已按审阅计划写回原表。兼容副本未自动回读或关闭，请手动关闭且不保存；原文件是否保存请按 Excel 的正常保存流程操作。" });
        return;
      }
      // A close failure must not make an already committed write look failed or allow replay.
      try {
        if (await closeExcelNativeComparison(current.session)) {
          publish(null); setFeedback({ messageId, text: "变更已写回原表，审阅副本已关闭。是否保存文件请按 Excel 的正常保存流程操作。" });
        } else setFeedback({ messageId, text: "变更已写回原表；副本尚未关闭，请关闭且不保存临时副本。" });
      } catch { setFeedback({ messageId, text: "变更已写回原表；请手动关闭临时副本。" }); }
    });
  }
  return { review, pending, feedback, stage, save, discard };
}
