export function PrePlanningPauseCard({ busy, error, onContinue }: {
  busy: boolean; error?: string; onContinue: () => void;
}) {
  return <div className="analysis-pause-inline" aria-label="继续未完成的需求分析">
      <button type="button" disabled={busy} onClick={onContinue}>继续分析</button>
      {error && <span role="alert">{error}</span>}
  </div>;
}
