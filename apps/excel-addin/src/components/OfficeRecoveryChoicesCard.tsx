import {
  PENDING_DISCOVERY_CHOOSE_DETAIL,
  PENDING_DISCOVERY_TRUNCATED_DETAIL,
  pendingExecutionCandidateLabel,
  type PendingExecutionCandidate
} from "@excel-bro/core/pendingExecutionDiscovery";

/**
 * 「发现这个工作簿里未完成的写入」——多候选时的**选择入口**。
 *
 * 三条纪律都落在这张卡上：
 *
 * 1. **不替用户选**：每条候选各一个按钮，先点哪条就是哪条；没有"默认选中"、
 *    没有"最新的一条"高亮，也不按更新时间自动排第一（排序在共享层，最危险的在最前）。
 * 2. **不写入**：这里只有读取与判定，按钮点击后走的是"GET 状态 + 本机判据"那条路；
 *    真正写入要等用户再看一次卡片并按续跑。
 * 3. **不折叠失败**：服务不可达时显示"暂时读不到"，绝不显示成"没有未完成任务"。
 *
 * 展示只用到共享层给的一句话标签（进度 + 状态），不碰单元格内容。
 */
export default function OfficeRecoveryChoicesCard({
  choices,
  failure,
  busy,
  onSelect
}: {
  choices: { candidates: PendingExecutionCandidate[]; truncated: boolean } | null;
  failure: string | null;
  busy: boolean;
  onSelect: (candidate: PendingExecutionCandidate) => void;
}) {
  if (!choices && !failure) return null;
  return (
    <section className="office-recovery-choices" aria-label="发现未完成的 Office 写入">
      {choices && (
        <>
          <strong>发现 {choices.candidates.length} 条未完成的写入</strong>
          <p>{PENDING_DISCOVERY_CHOOSE_DETAIL}</p>
          <ul>
            {choices.candidates.map((candidate) => (
              <li key={candidate.operationId}>
                <span>{pendingExecutionCandidateLabel(candidate)}</span>
                <button
                  type="button"
                  className="secondary-button"
                  disabled={busy}
                  onClick={() => onSelect(candidate)}
                >
                  核对这一次
                </button>
              </li>
            ))}
          </ul>
          {choices.truncated && (
            <p role="status">{PENDING_DISCOVERY_TRUNCATED_DETAIL}</p>
          )}
        </>
      )}
      {failure && <p role="status">{failure}</p>}
    </section>
  );
}
