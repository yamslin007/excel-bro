import { tailOf } from "@excel-bro/core/pauseControl";
import type { WpsReconcileVerdict } from "./wpsExecutionReconcile";
import { reportDevProbeRecovery } from "./devProbe";
import { isDev } from "./isDev";

/*
 * 恢复核对的最小诊断（**仅 DEV**）。
 *
 * 起因（2026-10-28 夜间跨进程轮）：完整退出 WPS 后重开，面板先显示「证据不足」，
 * 点一次「重新核对（只读）」就恢复。那段正文唯一对应
 * `wpsExecutionReconcile.ts` 的 `localInspection == null || "unknown"`，
 * 但 `unknown` 有十几条不同的产生路径 —— 缺了这一层，现场只剩"证据不足"四个字，
 * 无法判断是宿主没就绪、授权没恢复、描述不符，还是读值本身失败。
 *
 * 这一层的边界（刻意写死，改之前先读）：
 * - **只观察，不参与判定**。`wpsReconcileVerdict` 的输入输出不因它改变一个字节；
 *   没有重试、没有降级、没有"拿不到就再试一次"。
 * - **只在 DEV 记录**。非 DEV 直接 return，不留缓冲区、不打日志、不发请求 ——
 *   生产里这条路径与加它之前完全一样。
 * - **只收结构化标量**：枚举、布尔、计数、序号、时间。不收自由文本，因此
 *   单元格值、工作表名、工作簿路径、`ownerToken`、恢复包正文、模型密钥都进不来。
 *   异常也只留 `name`（正则白名单），**不留 message** —— WPS 的异常消息经常把
 *   文件路径整段带出来。
 * - 关联信息只留 ID 尾部若干位（沿用 `tailOf`），够把"首次发现"和"只读重试"
 *   两次记录串起来，又不把完整执行身份留在日志里。
 */

/** 哪一次核对触发的：首次发现 / 用户点选候选 / 只读重试 / 已选卡重算 / 点击动作前复核。 */
export type WpsRecoveryDiagnosticStage =
  | "discovery"
  | "select"
  | "recheck"
  | "reverify"
  | "pre-action";

/**
 * 现场核对的**阶段**。异常被吞掉时靠它说清"在哪一步断的"，
 * 而不是把所有失败都写成同一句 `unknown`。
 */
export type WpsInspectionPhase =
  | "authorize"
  | "identity"
  | "descriptor"
  | "sheets"
  | "actions"
  | "read"
  | "format"
  | "values";

/** `inspectWpsRecoveryState` 的结论。与既有字符串返回值一字不差。 */
export type WpsInspectionOutcome = "completed" | "not_started" | "conflict" | "unknown";

/**
 * 原因码。
 *
 * 命名规则：`<阶段>_<具体原因>`。每个码都对应 `inspectWpsRecoveryStateDetail`
 * 里**唯一一处**返回点 —— 加码时不要复用"看起来差不多"的码，否则现场又分不开。
 */
export type WpsRecoveryDiagnosticCode =
  /** 面板还没拿到服务端状态，谈不上核验现场。 */
  | "state_missing"
  /** 本次核验拿到的已选表是**空集**（授权范围还没恢复到可核验的状态）。 */
  | "authorize_no_selection"
  /** `authorizeWpsSheets` 抛错（宿主没就绪、序号失效、系统表等）。 */
  | "authorize_threw"
  /** 当前工作簿身份读不出来（未保存 / 宿主未就绪）。 */
  | "workbook_identity_unreadable"
  /** 当前工作簿身份与账本记录不同。 */
  | "workbook_identity_mismatch"
  /** 描述里没有任何动作。 */
  | "descriptor_empty"
  /** 描述的动作数与 `totalActions` 不符。 */
  | "descriptor_action_count"
  /** 账本登记过的授权表不在当前授权集合里。 */
  | "authorized_sheet_not_allowed"
  /** 动作的目标表不在当前授权集合里。 */
  | "action_sheet_not_allowed"
  /** 动作的目标表不在**登记**的授权表名里。 */
  | "action_sheet_not_authorized"
  /** 动作序号与数组下标不符（描述本身错位）。 */
  | "action_index_mismatch"
  /** 目标区域地址解析不出来。 */
  | "action_range_unparsable"
  /** 登记的行列数与区域地址不符。 */
  | "action_range_shape_mismatch"
  /** 登记格数与行列数不符。 */
  | "action_cell_count_mismatch"
  /** 登记缺 `expectedFingerprint`（旧描述）。 */
  | "action_fingerprint_missing"
  /** 目标总格数超出写入上限，拒绝核对。 */
  | "cell_budget_exceeded"
  /** 拿不到工作簿对象。 */
  | "workbook_unreadable"
  /** `sheet.Range(address)` 拿不到区域对象。 */
  | "host_range_unreadable"
  /** 宿主交出的区域行列数与登记不符。 */
  | "host_range_shape_mismatch"
  /** 格式动作没声明属性。 */
  | "format_fields_undeclared"
  /** 格式动作的属性名里有认不出的（不拿剩下的凑哈希）。 */
  | "format_fields_unrecognized"
  /** 格式快照不完整（有格读不到）。 */
  | "format_snapshot_incomplete"
  /** `Value2` / `Formula` 读回 `undefined`。 */
  | "value_read_undefined"
  /** 读回的形状不是登记的矩阵形状。 */
  | "value_shape_unexpected"
  /** 在途动作既不是写后也不是写前。 */
  | "in_flight_unexpected"
  /** 非在途动作与它该在的一侧不符（疑似人工修改）。 */
  | "completed_action_mismatch"
  /** 核验函数自身抛错（外层 `catch`）。 */
  | "inspection_threw"
  /** 面板侧调用外层抛错（`inspectWpsRecoveryStateDetail` 之外的 try）。 */
  | "panel_threw";

/**
 * 一条诊断记录。**全部是标量**，没有自由文本字段 —— 这是"不记录敏感内容"的
 * 结构性保证，不是靠调用方自觉。
 */
export type WpsRecoveryDiagnosticEvent = {
  /** 自增序号，用于"首次发现"与"只读重试"两条记录对照。 */
  seq: number;
  /** `Date.now()`。报告里按 UTC 打印，避免跨时区读错。 */
  at: number;
  stage: WpsRecoveryDiagnosticStage;
  /** 用户当时看到的结论种类（诊断只是旁注）。 */
  verdict: WpsReconcileVerdict["kind"];
  outcome?: WpsInspectionOutcome;
  code?: WpsRecoveryDiagnosticCode;
  phase?: WpsInspectionPhase;
  /** 执行身份尾部（6 位）。 */
  operationIdTail?: string;
  /** 任务身份尾部（6 位）。 */
  taskIdTail?: string;
  revision?: number;
  nextActionIndex?: number;
  totalActions?: number;
  inFlightActionIndex?: number | null;
  /** 本次核验看到的**已选表数量**（不记表名）。 */
  selectedSheetCount?: number;
  /** 本次核验看到的**授权表数量**（不记表名）。 */
  authorizedSheetCount?: number;
  /** 出问题的动作序号（有的话）。 */
  actionIndex?: number;
  /** 异常 `name`（白名单正则），不是 message。 */
  errorName?: string;
};

const DIAGNOSTIC_LIMIT = 40;
const events: WpsRecoveryDiagnosticEvent[] = [];
let sequence = 0;

/** 异常 `name` 的白名单：ASCII 标识符。带空格/路径/中文的一律丢弃。 */
export function safeDiagnosticErrorName(error: unknown): string | undefined {
  if (!error || typeof error !== "object") return undefined;
  const name = (error as { name?: unknown }).name;
  return typeof name === "string" && /^[A-Za-z_$][A-Za-z0-9_$]{0,40}$/.test(name) ? name : undefined;
}

/** 记录一条诊断。**非 DEV 直接丢弃**（不缓冲、不打日志、不发请求）。 */
export function recordWpsRecoveryDiagnostic(
  event: Omit<WpsRecoveryDiagnosticEvent, "seq" | "at"> & { at?: number }
): void {
  if (!isDev()) return;
  sequence += 1;
  const record: WpsRecoveryDiagnosticEvent = { ...event, seq: sequence, at: event.at ?? Date.now() };
  events.push(record);
  if (events.length > DIAGNOSTIC_LIMIT) events.splice(0, events.length - DIAGNOSTIC_LIMIT);
  /*
   * 控制台输出只给"真的在看的人"：vitest 下 MODE 是 `test`，否则每跑一遍面板用例
   * 就会往 stdout 刷几十行原因码，把真正的失败信息埋掉。缓冲区照记 —— 测试断言读的是
   * `wpsRecoveryDiagnosticEvents()`，不是 stdout。
   */
  if (import.meta.env.MODE !== "test") {
    try {
      console.debug(`[wps-recovery] ${formatWpsRecoveryDiagnosticEvent(record)}`);
    } catch {
      // 控制台不可用不影响判定（诊断永远不是判据的一部分）。
    }
  }
  /*
   * 真机上 WPS 窗格很难开 devtools，所以再留一条**显式开关**的出口：
   * `VITE_ENABLE_RECOVERY_DIAG=true` 启动 dev server 时，才把同一条记录
   * POST 到 `/__wps_probe`，由 dev server 控制台打印出来供复制。
   * 默认关着 —— 测试与日常开发不会因此往外发请求。
   */
  if (import.meta.env.VITE_ENABLE_RECOVERY_DIAG === "true") {
    try {
      reportDevProbeRecovery({
        seq: record.seq,
        stage: record.stage,
        verdict: record.verdict,
        ...(record.outcome ? { outcome: record.outcome } : {}),
        ...(record.code ? { code: record.code } : {}),
        ...(record.phase ? { phase: record.phase } : {}),
        ...(record.operationIdTail ? { operationIdTail: record.operationIdTail } : {}),
        ...(record.taskIdTail ? { taskIdTail: record.taskIdTail } : {}),
        ...(record.revision != null ? { revision: record.revision } : {}),
        ...(record.nextActionIndex != null ? { nextActionIndex: record.nextActionIndex } : {}),
        ...(record.totalActions != null ? { totalActions: record.totalActions } : {}),
        ...(record.inFlightActionIndex != null ? { inFlightActionIndex: record.inFlightActionIndex } : {}),
        ...(record.selectedSheetCount != null ? { selectedSheetCount: record.selectedSheetCount } : {}),
        ...(record.authorizedSheetCount != null ? { authorizedSheetCount: record.authorizedSheetCount } : {}),
        ...(record.actionIndex != null ? { actionIndex: record.actionIndex } : {}),
        ...(record.errorName ? { errorName: record.errorName } : {})
      });
    } catch {
      // 发不出去就只在本地留着，不因为遥测失败去打断恢复核对。
    }
  }
}

/** 只读快照（测试与"开发诊断"面板共用）。 */
export function wpsRecoveryDiagnosticEvents(): readonly WpsRecoveryDiagnosticEvent[] {
  return events.map((event) => ({ ...event }));
}

/** 测试专用：清空缓冲区与序号。 */
export function resetWpsRecoveryDiagnostics(): void {
  events.length = 0;
  sequence = 0;
}

function pad(value: number, width = 2): string {
  return String(value).padStart(width, "0");
}

/** 单条记录的紧凑单行文本。字段顺序固定 —— 两次记录能直接上下对照。 */
export function formatWpsRecoveryDiagnosticEvent(event: WpsRecoveryDiagnosticEvent): string {
  const at = new Date(event.at);
  const clock = `${pad(at.getUTCHours())}:${pad(at.getUTCMinutes())}:${pad(at.getUTCSeconds())}` +
    `.${pad(at.getUTCMilliseconds(), 3)}`;
  const parts = [
    `#${event.seq}`,
    clock,
    event.stage,
    `verdict=${event.verdict}`,
    ...(event.outcome ? [`outcome=${event.outcome}`] : []),
    ...(event.code ? [`code=${event.code}`] : []),
    ...(event.phase ? [`phase=${event.phase}`] : []),
    ...(event.operationIdTail ? [`op=…${event.operationIdTail}`] : []),
    ...(event.taskIdTail ? [`task=…${event.taskIdTail}`] : []),
    ...(event.revision != null ? [`rev=${event.revision}`] : []),
    ...(event.nextActionIndex != null || event.totalActions != null
      ? [`cursor=${event.nextActionIndex ?? "?"}/${event.totalActions ?? "?"}`]
      : []),
    ...(event.inFlightActionIndex != null ? [`inFlight=${event.inFlightActionIndex}`] : []),
    ...(event.actionIndex != null ? [`action=${event.actionIndex}`] : []),
    ...(event.selectedSheetCount != null ? [`selected=${event.selectedSheetCount}`] : []),
    ...(event.authorizedSheetCount != null ? [`authorized=${event.authorizedSheetCount}`] : []),
    ...(event.errorName ? [`error=${event.errorName}`] : [])
  ];
  return parts.join(" ");
}

/**
 * 可直接复制给人看的报告。
 *
 * 头一行是口径声明（UTC 时间、只含原因码），后面每行一条记录，最早的在最前 ——
 * 现场要看的就是"首次发现那一行"与"只读重试那一行"的 `code` 差在哪。
 */
export function wpsRecoveryDiagnosticReport(limit = 12): string {
  const snapshot = wpsRecoveryDiagnosticEvents();
  if (!snapshot.length) return "（DEV 恢复诊断：暂无记录）";
  const shown = snapshot.slice(-Math.max(1, limit));
  return [
    `DEV 恢复诊断（共 ${snapshot.length} 条，显示最近 ${shown.length} 条；时间为 UTC；只含原因码，不含表名/单元格值/路径）`,
    ...shown.map((event) => formatWpsRecoveryDiagnosticEvent(event))
  ].join("\n");
}

/* ------------------------------------------------------------------------- *
 * 现场核验结果的**详细**形态（只有原因码与计数，没有业务数据）
 * ------------------------------------------------------------------------- */

export type WpsInspectionResult = {
  outcome: WpsInspectionOutcome;
  /** `unknown` / `conflict` 的原因；`completed` / `not_started` 也可能带（无需码时为 undefined）。 */
  code?: WpsRecoveryDiagnosticCode;
  /** 断在哪一步。异常被吞掉时这是唯一能说明位置的信息。 */
  phase: WpsInspectionPhase;
  /** 本次授权解析出的表数量（不记表名）。 */
  selectedSheetCount: number;
  authorizedSheetCount: number;
  /** 出问题的动作序号（有的话）。 */
  actionIndex?: number;
  errorName?: string;
};
