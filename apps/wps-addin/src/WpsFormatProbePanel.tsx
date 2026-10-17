/**
 * WPS 格式能力探测 · 开发诊断面板（任务窗格内 #/dev-format-probe）。
 *
 * 这是 Task #47 的**可操作入口**：不用手写源码函数，打开面板 → 确认目标 → 点按钮 → 复制结果。
 *
 * 面板职责：
 *   1. 展示当前工作簿 / 工作表 / 探测目标（单格地址），并**明示整行/整列连带影响**。
 *   2. 运行前让用户**明确确认**"这是可丢弃的独立测试工作簿"；确认**绑定到目标身份**
 *      （工作簿路径 + 工作表名 + 目标地址），切换工作簿/工作表后必须重新确认。
 *   3. 运行前校验目标资格（.xlsx + 单个空白单元格 + 无值无公式）。
 *   4. 逐项展示每个格式成员的五态结论 + 整轮基线恢复状态 + 是否提前停止。
 *   5. 提供"一键复制"逐项报告（JSON），并同步经 `__wps_probe` 打到 dev server console。
 *
 * **不读业务数据、不碰任何业务区域**：探测只作用于用户指定的单个空白单元格。
 */
import { useEffect, useMemo, useState } from "react";
import { brandLogoUrl } from "../../excel-addin/src/brandLogo";
import { getWpsApplication } from "./wpsHost";
import {
  runWpsFormatProbe,
  assertWpsFormatProbeTarget,
  readWpsFormatProbeTargetIdentity,
  sameProbeTarget,
  WPS_FORMAT_PROBE_GROUPS,
  WPS_FORMAT_PROBE_GROUP_LABELS,
  type WpsFormatProbeGroup,
  type WpsFormatProbeReport,
  type WpsFormatMemberProbe,
  type WpsFormatProbeTarget
} from "./wpsFormatProbe";
import { reportDevProbeDetail } from "./devProbe";

type State =
  | { phase: "idle"; targetOk: boolean; reason: string }
  | { phase: "confirm"; target: WpsFormatProbeTarget }
  | { phase: "running"; target: WpsFormatProbeTarget }
  | { phase: "done"; report: WpsFormatProbeReport; target: WpsFormatProbeTarget }
  | { phase: "error"; message: string };

const STATUS_TEXT: Record<WpsFormatMemberProbe["status"], string> = {
  member_exists: "成员存在",
  write_verified: "写入验证通过",
  restore_verified: "恢复验证通过",
  skipped: "未执行",
  failed: "失败"
};

function summarize(report: WpsFormatProbeReport): string {
  const s = report.summary;
  return [
    `共 ${s.total} 项`,
    `恢复验证通过 ${s.restoreVerified}`,
    `写入验证通过 ${s.writeVerified}`,
    `成员存在 ${s.memberExists}`,
    `未执行 ${s.skipped}`,
    `失败 ${s.failed}`,
    report.baselineRestored ? "整轮基线恢复通过" : "整轮基线未通过"
  ].join(" · ");
}

export default function WpsFormatProbePanel() {
  const [probeCell, setProbeCell] = useState("A1");
  const [selectedGroups, setSelectedGroups] = useState<readonly WpsFormatProbeGroup[]>([...WPS_FORMAT_PROBE_GROUPS]);
  const [confirmedTarget, setConfirmedTarget] = useState<WpsFormatProbeTarget | null>(null);
  const [state, setState] = useState<State>(() => {
    const check = assertWpsFormatProbeTarget(getWpsApplication(), "A1");
    return check.ok
      ? { phase: "idle", targetOk: true, reason: "" }
      : { phase: "idle", targetOk: false, reason: check.reason };
  });
  const [copied, setCopied] = useState(false);

  // 每次探测目标变化时重新校验资格，并清空已确认（换地址/换表后旧确认失效）。
  const refreshTarget = (cell: string) => {
    setProbeCell(cell);
    setConfirmedTarget(null);
    const check = assertWpsFormatProbeTarget(getWpsApplication(), cell);
    setState(check.ok
      ? { phase: "idle", targetOk: true, reason: "" }
      : { phase: "idle", targetOk: false, reason: check.reason });
  };

  const toggleGroup = (group: WpsFormatProbeGroup) => {
    setSelectedGroups((current) => {
      const has = current.includes(group);
      // 保持 WPS_FORMAT_PROBE_GROUPS 的顺序，只增删目标组。
      return has
        ? current.filter((g) => g !== group)
        : WPS_FORMAT_PROBE_GROUPS.filter((g) => current.includes(g) || g === group);
    });
  };

  // 进入确认态：重新读取当前真实目标身份，展示给用户确认。
  const beginConfirm = () => {
    const app = getWpsApplication();
    const check = assertWpsFormatProbeTarget(app, probeCell);
    if (!check.ok) {
      setState({ phase: "idle", targetOk: false, reason: check.reason });
      return;
    }
    setConfirmedTarget(check.target);
    setState({ phase: "confirm", target: check.target });
  };

  // 用户点"确认这是可丢弃的独立测试工作簿"后，绑定到目标身份并运行。
  const run = () => {
    if (state.phase !== "confirm") return;
    // 运行前重读目标身份，与确认时绑定的一致才执行（换簿/换表后失效）。
    const app = getWpsApplication();
    const currentIdentity = readWpsFormatProbeTargetIdentity(app, probeCell);
    if (!currentIdentity || !sameProbeTarget(currentIdentity, state.target)) {
      setConfirmedTarget(null);
      setState({ phase: "idle", targetOk: false, reason: "工作簿或工作表已切换，请重新确认这是可丢弃的独立测试工作簿。" });
      return;
    }
    const boundTarget = currentIdentity;
    setState({ phase: "running", target: boundTarget });
    try {
      const report = runWpsFormatProbe(app, { probeCell, groups: selectedGroups });
      setState({ phase: "done", report, target: boundTarget });
      try {
        reportDevProbeDetail("taskpane", { formatProbeJson: JSON.stringify(report) });
      } catch {
        // 上报失败不影响面板展示。
      }
    } catch (error) {
      setState({
        phase: "error",
        message: error instanceof Error ? error.message : "无法读取 WPS 宿主。"
      });
    }
  };

  const copyReport = async () => {
    if (state.phase !== "done") return;
    const text = JSON.stringify(state.report, null, 2);
    try {
      await navigator.clipboard.writeText(text);
      setCopied(true);
      window.setTimeout(() => setCopied(false), 1500);
    } catch {
      // 剪贴板不可用时，退化为选中文本提示。
    }
  };

  // 换簿/换表监听：确认后若目标身份变化，立即失效（回到 idle）。
  useEffect(() => {
    if (state.phase !== "confirm" && state.phase !== "running" && state.phase !== "done") return;
    const app = getWpsApplication();
    const current = readWpsFormatProbeTargetIdentity(app, probeCell);
    if (!current || !sameProbeTarget(current, state.target)) {
      setConfirmedTarget(null);
      setState({ phase: "idle", targetOk: false, reason: "工作簿或工作表已切换，请重新确认这是可丢弃的独立测试工作簿。" });
    }
  }, [state, probeCell]);

  const reportText = useMemo(() => {
    if (state.phase !== "done") return "";
    return JSON.stringify(state.report, null, 2);
  }, [state]);

  return (
    <main className="poc-panel-shell">
      <header>
        <img className="probe-logo" src={brandLogoUrl} alt="" />
        <div>
          <strong>WPS 格式能力 · 受控探测</strong>
          <p>只探测单个空白单元格，逐项验证写入与恢复，不碰业务数据</p>
        </div>
      </header>

      <section className="poc-dialog-facts">
        <div>
          <dt>工作簿</dt>
          <dd>{state.phase === "confirm" || state.phase === "running" || state.phase === "done" ? state.target.workbookName : (readWpsFormatProbeTargetIdentity(getWpsApplication(), probeCell)?.workbookName ?? "未识别")}</dd>
        </div>
        <div>
          <dt>工作表</dt>
          <dd>{state.phase === "confirm" || state.phase === "running" || state.phase === "done" ? state.target.sheetName : (readWpsFormatProbeTargetIdentity(getWpsApplication(), probeCell)?.sheetName ?? "未识别")}</dd>
        </div>
        <div>
          <dt>探测目标</dt>
          <dd>
            <input
              aria-label="探测单元格"
              value={probeCell}
              onChange={(event) => refreshTarget(event.target.value)}
              disabled={state.phase === "running" || state.phase === "confirm"}
            />
            <small>单个空白单元格地址（如 A1、$B$3）。</small>
          </dd>
        </div>
      </section>

      <section className="poc-panel-note" role="status">
        ⚠️ 行高（RowHeight）会改变**整行**、列宽（ColumnWidth）会改变**整列**；
        填充（Interior）与边框（Borders）是关联属性，改其一可能连带改动同组其它字段。
        请确认这是**可丢弃的独立测试工作簿**，探测会在写入后尽力恢复原值。
      </section>

      {state.phase === "idle" && state.targetOk === false && (
        <section className="poc-panel-error" role="alert">
          <strong>目标不可用</strong>
          <p>{state.reason}</p>
        </section>
      )}

      {state.phase === "error" && (
        <section className="poc-panel-error" role="alert">
          <strong>宿主不可用</strong>
          <p>{state.message}</p>
        </section>
      )}

      {state.phase === "confirm" && (
        <section className="poc-panel-confirm" role="dialog">
          <strong>确认这是可丢弃的独立测试工作簿</strong>
          <p>
            即将对「{state.target.workbookName}」的「{state.target.sheetName}」表
            「{state.target.probeCell}」单元格执行格式能力探测。探测会临时写入哨兵值并
            尝试恢复；行高/列宽会改变整行/整列。请确认该工作簿可丢弃。
          </p>
          <div className="poc-panel-result-actions">
            <button type="button" className="primary" onClick={run}>
              确认并运行探测
            </button>
            <button type="button" onClick={() => { setConfirmedTarget(null); setState({ phase: "idle", targetOk: true, reason: "" }); }}>
              取消
            </button>
          </div>
        </section>
      )}

      {(state.phase === "idle" || state.phase === "confirm") && (
        <>
          <section className="poc-panel-note">
            <strong>选择要探测的能力组</strong>
            <p>
              可在新的独立测试基线上按组单独运行（例如填充/边框此前因前置失败未测到，
              可单独勾选重跑）；同一轮内任一成员恢复失败仍会立即停止后续写入，不允许在
              已受污染的状态上继续。
            </p>
            <div className="probe-group-select" role="group" aria-label="能力组选择">
              {WPS_FORMAT_PROBE_GROUPS.map((group) => {
                const checked = selectedGroups.includes(group);
                return (
                  <label key={group} className={checked ? "probe-group-checked" : ""}>
                    <input
                      type="checkbox"
                      checked={checked}
                      disabled={state.phase !== "idle"}
                      onChange={() => toggleGroup(group)}
                    />
                    <span>{WPS_FORMAT_PROBE_GROUP_LABELS[group]}</span>
                  </label>
                );
              })}
            </div>
          </section>

          <button
            type="button"
            className="primary"
            disabled={
              state.phase !== "idle" || !state.targetOk || selectedGroups.length === 0
            }
            onClick={beginConfirm}
          >
            {state.phase === "confirm" ? "请在上方确认" : "选择目标并确认"}
          </button>
        </>
      )}

      {state.phase === "running" && (
        <p className="poc-panel-note" role="status">正在探测…</p>
      )}

      {state.phase === "done" && (
        <section className="poc-panel-result" role="status">
          <div className="poc-panel-result-actions">
            <strong>{summarize(state.report)}</strong>
            <button type="button" onClick={copyReport}>
              {copied ? "已复制" : "复制逐项报告"}
            </button>
          </div>

          {state.report.aborted && (
            <p className="poc-panel-error" role="alert">
              探测因恢复失败/不确定而提前停止，后续成员已标「未执行」，零写入。
            </p>
          )}

          <table className="format-probe-table">
            <thead>
              <tr>
                <th>成员</th>
                <th>状态</th>
                <th>哨兵</th>
                <th>原值</th>
                <th>说明</th>
              </tr>
            </thead>
            <tbody>
              {state.report.members.map((member) => (
                <tr key={member.member} className={`probe-${member.status}`}>
                  <td>{member.member}</td>
                  <td>{STATUS_TEXT[member.status]}</td>
                  <td>{String(member.sentinel)}</td>
                  <td>{String(member.original)}</td>
                  <td>{member.note || "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>

          <details>
            <summary>逐项 JSON（可复制）</summary>
            <pre className="format-probe-json">{reportText}</pre>
          </details>

          <p className="poc-panel-note">
            这只证明"该成员在**本机这个 WPS 版本**上可读、可写、回读一致、且能恢复"。
            未执行（AutoFit、或前置失败停止的成员）不计入通过；失败项请记录现象。
            不能据此推断其它版本，也不等于"WPS 格式动作已接通"——那仍需要执行器实现与端到端验收。
          </p>
        </section>
      )}
    </main>
  );
}
