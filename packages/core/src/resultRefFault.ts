/**
 * 「结果引用过期」验收入口（A1-E1）的**共享判据**。
 *
 * 这一层只做确定性的判断与文案，不发请求、不碰宿主、不读工作簿：
 *
 * - 入口是否可见（仅开发模式 / 显式开启的验收构建）；
 * - 启用请求要带哪些字段（**只有**宿主与工作簿身份，绝不带任务或引用）；
 * - 当前状态该怎么说、按钮该不该可按。
 *
 * 真正的失效发生在服务端（结果正文只在本机易失登记器里，模型和界面都只有不透明
 * 句柄），两端面板只负责把用户的显式启用送过去。所以这里既不是"UI 开关即失效"，
 * 也不允许任何一端自己伪造一个"已过期"。
 */

/** 与服务端 `DevResultRefFaultStatus.phase` 同一组取值。 */
export type ResultRefFaultPhase =
  | "off"
  | "idle"
  | "armed"
  | "bound"
  | "consumed"
  | "expired";

export type ResultRefFaultStatus = {
  enabled: boolean;
  phase: ResultRefFaultPhase;
  expiresInSeconds: number;
  taskBound: boolean;
  refBound: boolean;
};

export type ResultRefFaultHostId = "excel-office" | "wps-et";

/** 启用请求体：刻意**不含** taskId / resultRef（服务端 extra=forbid 会直接拒）。 */
export type ResultRefFaultTarget = {
  hostId: ResultRefFaultHostId;
  workbookName: string;
  workbookIdentity: string;
};

/**
 * 入口可见性。`devMode` 由各端自己算（`import.meta.env.DEV` 或验收构建的显式开关），
 * 两端读同一个函数，避免一端在某类构建里"还能看见"而另一端看不见。
 */
export function resultRefFaultVisible(devMode: boolean): boolean {
  return devMode === true;
}

/**
 * 构造启用请求。工作簿身份缺失时返回 `null` —— 绑定的语义是"这本工作簿的下一次
 * 规划任务"，没有身份就无从绑定，宁可不可用也不做一个全局开关。
 */
export function resultRefFaultTarget(
  hostId: ResultRefFaultHostId,
  workbookName: string | null | undefined,
  workbookIdentity: string | null | undefined
): ResultRefFaultTarget | null {
  const name = (workbookName ?? "").trim();
  if (!name) return null;
  return { hostId, workbookName: name, workbookIdentity: (workbookIdentity ?? "").trim() };
}

/** 状态文案的唯一来源（两端面板不再各写一套）。 */
export function describeResultRefFault(
  status: ResultRefFaultStatus | null,
  target: ResultRefFaultTarget | null
): string {
  if (!target) return "当前工作簿尚未保存，无法绑定验收目标；未启用。";
  if (!status) return "尚未读取验收入口状态。";
  if (!status.enabled) return "服务端没有开启结果引用过期入口（需要开发模式）。";
  switch (status.phase) {
    case "off":
      return "服务端没有开启结果引用过期入口（需要开发模式）。";
    case "idle":
      return "未启用：启用后只影响本工作簿的下一次规划任务。";
    case "armed":
      return `已启用：本工作簿的下一次规划任务的第一个 Python 结果引用将失效（一次性，${status.expiresInSeconds} 秒内有效）。`;
    case "bound":
      return "已绑定本次任务，等待它的第一个 Python 结果登记。";
    case "consumed":
      return "已触发：旧引用被拒绝并已删除；后续任务不受影响。";
    case "expired":
      return "启用已超时，未触发；可重新启用。";
  }
}

/** 已绑定到某次任务后不能再改主意（否则"单次"就失去意义）。 */
export function canArmResultRefFault(
  status: ResultRefFaultStatus | null,
  target: ResultRefFaultTarget | null,
  busy: boolean
): boolean {
  if (!target || busy) return false;
  if (!status) return true;
  return status.enabled && (status.phase === "idle" || status.phase === "armed" || status.phase === "expired" || status.phase === "consumed");
}

export function canCancelResultRefFault(
  status: ResultRefFaultStatus | null,
  busy: boolean
): boolean {
  if (busy || !status) return false;
  return status.phase === "armed";
}
