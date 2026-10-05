/**
 * 宽视图 PoC 面板（任务窗格内 #/wide-view-poc）。
 *
 * 只做一件事：在当前 WPS 实机上探测"独立宽窗口"是否可用，并把结果如实回报。
 * **不读工作簿、不带单元格值**，所以可以在任何工作簿里安全运行。
 *
 * 探针失败时这里必须显式报错，**不允许**偷偷回退到 localStorage 快照或
 * 窗格内浮层 —— 那会把"没有这个能力"粉饰成"能打开"。
 */

import { useState } from "react";
import { getWpsApplication } from "./wpsHost";
import {
  probeWideViewDialog,
  type WideViewDialogProbe
} from "./wpsWideViewPoc";
import { reportDevProbe } from "./devProbe";

type State =
  | { phase: "idle" }
  | { phase: "running" }
  | { phase: "done"; probe: WideViewDialogProbe }
  | { phase: "host-error"; message: string };

const STATUS_TEXT: Record<WideViewDialogProbe["status"], string> = {
  requested: "已请求打开宽窗口，请核对实际效果",
  missing: "❌ 宿主未暴露该 API",
  error: "❌ 调用失败"
};

export default function WideViewPocPanel() {
  const [state, setState] = useState<State>({ phase: "idle" });

  const run = () => {
    setState({ phase: "running" });
    try {
      const probe = probeWideViewDialog(getWpsApplication());
      setState({ phase: "done", probe });
      reportDevProbe("taskpane", {
        wideViewPocRequested: probe.status === "requested",
        wideViewPocMissing: probe.status === "missing",
        wideViewPocError: probe.status === "error"
      });
    } catch (error) {
      setState({
        phase: "host-error",
        message: error instanceof Error ? error.message : "无法读取 WPS 宿主。"
      });
    }
  };

  return (
    <main className="poc-panel-shell">
      <header>
        <span className="probe-logo">EB</span>
        <div>
          <strong>宽幅审阅 · 能力探针</strong>
          <p>只读探测，不读取任何工作表内容</p>
        </div>
      </header>

      <button type="button" className="primary" disabled={state.phase === "running"} onClick={run}>
        {state.phase === "running" ? "正在探测…" : "探测独立宽窗口"}
      </button>

      {state.phase === "host-error" && (
        <section className="poc-panel-error" role="alert">
          <strong>宿主不可用</strong>
          <p>{state.message}</p>
        </section>
      )}

      {state.phase === "done" && (
        <section className={`poc-panel-result ${state.probe.status}`} role="status">
          <strong>{STATUS_TEXT[state.probe.status]}</strong>
          <dl className="poc-dialog-facts">
            <div>
              <dt>宿主版本</dt>
              <dd>{state.probe.hostVersion ?? "未报告"}</dd>
            </div>
            <div>
              <dt>存在的 API</dt>
              <dd>{state.probe.presentApis.join("、") || "无"}</dd>
            </div>
            <div>
              <dt>实际打开</dt>
              <dd>{state.probe.openedWith ?? "未成功"}</dd>
            </div>
            <div>
              <dt>窗口句柄成员</dt>
              <dd>{state.probe.dialogMembers.join("、") || "无"}</dd>
            </div>
            <div>
              <dt>请求尺寸</dt>
              <dd>
                {state.probe.requestedSize.width}×{state.probe.requestedSize.height}（逻辑像素）
              </dd>
            </div>
            <div>
              <dt>关闭</dt>
              <dd>探针不自动关闭；请在宽窗口观察后手动关闭</dd>
            </div>
          </dl>
          {state.probe.error && (
            <p className="poc-panel-error-detail" role="alert">
              {state.probe.error}
            </p>
          )}
          {state.probe.status === "requested" && (
            <p className="poc-panel-note">
              这只证明宿主接受打开请求。请记录宽窗口中显示的 CSS 宽度、是否真正出现两栏可用空间，
              原任务窗格是否仍可同时查看和操作，以及手动关闭后能否返回。
              本探针路由暂时替换了聊天内容；与真实对话同屏须在正式接线后再验。
            </p>
          )}
          {state.probe.status !== "requested" && (
            <p className="poc-panel-note">
              当前调用未能打开窗口。请记录上面的 API、版本与现象；不能据此推断所有
              WPS 版本都不具备独立宽窗口能力。
            </p>
          )}
        </section>
      )}
    </main>
  );
}
