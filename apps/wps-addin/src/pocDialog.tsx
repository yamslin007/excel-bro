/**
 * 宽视图 PoC 用的**只读测试页**。刻意不加载任何工作簿数据、
 * 不访问 Application.ActiveWorkbook —— 它只用来回答两个问题：
 * 1. 对话框能不能开到 1280 宽并正常渲染；
 * 2. 同源 WebView 之间的无数据内存握手能否成功；
 * 3. 关窗后能否回到任务窗格。
 *
 * 页面自己会量一下 `window.innerWidth`（它已经是 CSS 逻辑像素），
 * 把"实际拿到了多宽"显示给用户；不能再除以 devicePixelRatio。
 */

import { useEffect, useState } from "react";
import { reportDevProbe } from "./devProbe";
import { startChildChannelProbe, type ChannelProbeStatus } from "./wpsWideViewChannelPoc";

type Probe = {
  innerWidth: number;
  innerHeight: number;
  devicePixelRatio: number;
  cssWidth: number;
};

function measure(): Probe {
  const ratio = Number(window.devicePixelRatio) || 1;
  return {
    innerWidth: window.innerWidth,
    innerHeight: window.innerHeight,
    devicePixelRatio: ratio,
    cssWidth: window.innerWidth
  };
}

export default function PocDialog() {
  const [probe, setProbe] = useState<Probe>(() => measure());
  const [channelStatus, setChannelStatus] = useState<ChannelProbeStatus | null>(null);

  useEffect(() => {
    const next = measure();
    setProbe(next);
    reportDevProbe("taskpane", {
      dialogOpened: true,
      dialogCssWidth: next.cssWidth,
      dialogInnerWidth: next.innerWidth
    });
  }, []);

  useEffect(() => {
    const nonce = new URLSearchParams(window.location.search).get("probeNonce");
    if (!nonce) return;
    return startChildChannelProbe(nonce, setChannelStatus);
  }, []);

  return (
    <main className="poc-dialog-shell">
      <header>
        <span className="probe-logo">EB</span>
        <strong>变更审阅宽视图 · 能力探针</strong>
      </header>
      <p>
        这是只读测试页，<b>没有读取任何工作表内容</b>。若你能看到它、
        请记录下面的实际宽度与跨窗口通信结果。目标是在这个宽窗口内放置
        「修改前表格｜修改后表格｜对话」三栏，而不是依赖原任务窗格露在窗口旁边；
        仅仅出现此页面还不能证明目标布局可用。
      </p>
      <dl className="poc-dialog-facts">
        <div>
          <dt>CSS 宽度</dt>
          <dd>{probe.cssWidth}px</dd>
        </div>
        <div>
          <dt>估计物理宽度</dt>
          <dd>{Math.round(probe.innerWidth * probe.devicePixelRatio)}px</dd>
        </div>
        <div>
          <dt>devicePixelRatio</dt>
          <dd>{probe.devicePixelRatio}</dd>
        </div>
        <div>
          <dt>高度</dt>
          <dd>{probe.innerHeight}px</dd>
        </div>
      </dl>
      <p className="poc-dialog-hint">
        跨窗口内存通信：{channelStatus
          ? { waiting: "正在握手", connected: "已双向连通", unsupported: "不可用", timeout: "超时" }[channelStatus]
          : "未发起"}。探针只交换随机校验码，不传工作簿数据。
      </p>
      <p className="poc-dialog-hint">
        请点击对话框的关闭按钮，确认原任务窗格仍可使用；若无法关闭，
        请记录具体表现，不要据此直接推断是哪一个宿主 API 缺失。
      </p>
    </main>
  );
}
