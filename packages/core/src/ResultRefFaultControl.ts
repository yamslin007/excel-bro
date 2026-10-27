import {createElement as h, useState} from "react";
import {
  canArmResultRefFault,
  canCancelResultRefFault,
  describeResultRefFault,
  type ResultRefFaultStatus,
  type ResultRefFaultTarget,
} from "./resultRefFault";

/**
 * 两端共用的「结果引用过期」验收入口（开发专用）。
 *
 * 只负责把用户的**显式启用**送到服务端，并把服务端状态如实说出来。它自己不做任何
 * 失效判断：失效发生在服务端解析引用的那一刻（结果正文只在服务端登记器里）。
 * 因此这个控件既不改变任务、也不改变写入流程，确认/放弃仍然只有一处入口。
 *
 * `load` 只在用户展开这块区域时调用一次：开发入口不该在每次挂载时打请求。
 */
export function ResultRefFaultControl({
  visible, disabled, target, load, arm, cancel,
}: {
  visible: boolean;
  disabled: boolean;
  target: ResultRefFaultTarget | null;
  load: () => Promise<ResultRefFaultStatus>;
  arm: (target: ResultRefFaultTarget) => Promise<ResultRefFaultStatus>;
  cancel: () => Promise<ResultRefFaultStatus>;
}) {
  const [status, setStatus] = useState<ResultRefFaultStatus | null>(null);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  /**
   * 是否已展开。
   *
   * 收起时**不渲染状态播报**：一个折叠着的开发入口没有"当前状态"要对用户说，
   * 而 `role="status"` 是活区 —— 面板里任何常驻的 status 都会让"这一轮还有没有
   * 进度提示"这类判据失去意义（既有的"新对话后不残留进度提示"用例就是这么被
   * 一条常驻 status 弄红的）。状态只在展开之后出现。
   */
  const [opened, setOpened] = useState(false);

  const run = async (fn: () => Promise<ResultRefFaultStatus>) => {
    setBusy(true);
    try {
      const next = await fn();
      setStatus(next);
      setError("");
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason));
    } finally {
      setBusy(false);
    }
  };

  if (!visible) return null;
  const blocked = disabled || busy || !target;
  return h("details", {
      className: "result-ref-fault-control",
      // 展开时才读一次状态：开发入口不在挂载时就打服务端。
      onToggle: (event: {currentTarget: HTMLDetailsElement}) => {
        const open = event.currentTarget.open;
        setOpened(open);
        if (open && !busy) void run(load);
      },
    },
    h("summary", null, "结果引用过期（开发专用）"),
    h("p", null,
      "仅在独立测试工作簿中使用。启用后，本工作簿的下一次规划任务的第一个 Python 结果引用会在“已登记、尚未首次写入”处失效：模型必须重新读算，写入才会发生。只影响这一个引用，不改结果保留期。"),
    h("button", {
      type: "button",
      disabled: blocked || !canArmResultRefFault(status, target, busy),
      onClick: () => { if (target) void run(() => arm(target)); },
    }, "启用下一次引用过期"),
    h("button", {
      type: "button",
      disabled: blocked || !canCancelResultRefFault(status, busy),
      onClick: () => void run(cancel),
    }, "取消验收"),
    opened
      ? h("p", {role: error ? "alert" : "status"},
          error || describeResultRefFault(status, target))
      : null
  );
}
