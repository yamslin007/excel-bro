import { describe, expect, it } from "vitest";
import {
  TASK_RESUME_SCHEMA_VERSION,
  TASK_WRITE_STATE_LABEL,
  isTaskResumeRecord,
  newTaskId,
  sanitizeTaskResumeRecord,
  taskResumeDecision,
  withTaskResumeConfirmation,
  withTaskResumeState,
  type TaskResumeRecord
} from "./taskResume";

/** 一条合法记录的最小形态。各用例只改自己关心的那一位。 */
function baseRecord(overrides: Partial<TaskResumeRecord> = {}): TaskResumeRecord {
  const now = new Date().toISOString();
  return {
    schemaVersion: TASK_RESUME_SCHEMA_VERSION,
    taskId: "task-abc",
    phase: "awaiting_clarification",
    blocker: "user_input",
    resumability: "reauthorize",
    writeState: "not_written",
    originalPrompt: "把订单表的金额列按币种换算",
    confirmations: [],
    authorizationRevision: 2,
    createdAt: now,
    updatedAt: now,
    ...overrides
  };
}

describe("sanitizeTaskResumeRecord 只认显式白名单", () => {
  it("保留全部合法字段", () => {
    const record = sanitizeTaskResumeRecord(
      baseRecord({
        turnId: "turn-1",
        operationId: "op-1",
        requestId: "req-1",
        confirmations: [{ label: "币种列", value: "人民币" }]
      })
    );
    expect(record).not.toBeNull();
    expect(record?.taskId).toBe("task-abc");
    expect(record?.turnId).toBe("turn-1");
    expect(record?.operationId).toBe("op-1");
    expect(record?.requestId).toBe("req-1");
    expect(record?.authorizationRevision).toBe(2);
    expect(record?.confirmations).toHaveLength(1);
  });

  it("丢弃单元格内容、矩阵、宿主对象引用与完整路径（结构性隐私保证）", () => {
    // 负向断言必须配正向锚点：先证明"合法字段确实被保留"，
    // 再证明"这些越界字段一个都没活下来"，否则可能只是整体返回了 null。
    const dirty = {
      ...baseRecord({ confirmations: [{ label: "币种列", value: "人民币" }] }),
      // 以下都不允许存在于可持久化的恢复记录里
      values: [["机密来源值A", "机密来源值B"]],
      plan: { id: "plan-1", actions: [{ type: "writeValues", values: [["x"]] }] },
      targetMatrix: [["机密目标值"]],
      sourceSnapshots: [{ reference: "订单!A1:B2", values: [["机密来源值A"]] }],
      workbookPath: "C:\\Users\\somebody\\私密目录\\工作簿.xlsx",
      workbookRef: { __hostObject: true },
      sheetRefs: [{ ref: { __hostObject: true } }],
      __proto__: { polluted: true }
    };
    const record = sanitizeTaskResumeRecord(dirty);
    expect(record).not.toBeNull();
    // 正向锚点：合法内容确实被保留下来
    expect(record?.confirmations?.[0]?.value).toBe("人民币");
    // 负向：越界字段的键一个都不在，序列化后也不含任何单元格内容
    const serialized = JSON.stringify(record);
    for (const forbidden of [
      "values",
      "plan",
      "targetMatrix",
      "sourceSnapshots",
      "workbookPath",
      "workbookRef",
      "sheetRefs",
      "机密来源值A",
      "机密目标值",
      "C:\\\\Users\\\\somebody"
    ]) {
      expect(serialized).not.toContain(forbidden);
      expect(Object.keys(record ?? {})).not.toContain(forbidden);
    }
    expect(Object.keys(record ?? {}).sort()).toEqual(
      [
        "authorizationRevision",
        "blocker",
        "confirmations",
        "createdAt",
        "originalPrompt",
        "phase",
        "resumability",
        "schemaVersion",
        "taskId",
        "updatedAt",
        "writeState"
      ].sort()
    );
  });

  it("缺少原需求时返回 null，不用问题摘要冒充完整任务", () => {
    expect(
      sanitizeTaskResumeRecord(baseRecord({ originalPrompt: "" }))
    ).toBeNull();
    expect(
      sanitizeTaskResumeRecord({
        ...baseRecord(),
        originalPrompt: 123
      })
    ).toBeNull();
    expect(
      sanitizeTaskResumeRecord({ ...baseRecord(), originalPrompt: "   " })
    ).toBeNull();
  });

  it("未知 schemaVersion / 未知枚举一律拒绝，不猜", () => {
    expect(
      sanitizeTaskResumeRecord({ ...baseRecord(), schemaVersion: 2 })
    ).toBeNull();
    expect(
      sanitizeTaskResumeRecord({ ...baseRecord(), phase: "some_new_phase" })
    ).toBeNull();
    expect(
      sanitizeTaskResumeRecord({ ...baseRecord(), blocker: "some_new_blocker" })
    ).toBeNull();
    expect(
      sanitizeTaskResumeRecord({ ...baseRecord(), writeState: "maybe" })
    ).toBeNull();
    expect(
      sanitizeTaskResumeRecord({ ...baseRecord(), resumability: "hope" })
    ).toBeNull();
  });

  it("非对象输入返回 null", () => {
    for (const value of [null, undefined, 0, "", [], "task"]) {
      expect(sanitizeTaskResumeRecord(value)).toBeNull();
    }
  });

  it("确认项去重、限条数、丢弃空标签或空内容", () => {
    const record = sanitizeTaskResumeRecord(
      baseRecord({
        confirmations: [
          { label: "币种列", value: "人民币" },
          { label: "币种列", value: "人民币" },
          { label: "", value: "无标签" },
          { label: "空内容", value: "  " },
          { label: "分隔符", value: "、" },
          ...Array.from({ length: 20 }, (_unused, index) => ({
            label: `填充${index}`,
            value: `值${index}`
          }))
        ]
      })
    );
    expect(record?.confirmations).toHaveLength(8);
    expect(record?.confirmations.map((entry) => entry.label)).toEqual([
      "币种列",
      "分隔符",
      "填充0",
      "填充1",
      "填充2",
      "填充3",
      "填充4",
      "填充5"
    ]);
  });

  it("可选身份字段是空串时整个丢掉，不留下空键", () => {
    const record = sanitizeTaskResumeRecord(
      baseRecord({ turnId: "   ", operationId: "", requestId: undefined })
    );
    expect(record).not.toBeNull();
    expect("turnId" in (record ?? {})).toBe(false);
    expect("operationId" in (record ?? {})).toBe(false);
    expect("requestId" in (record ?? {})).toBe(false);
  });

  it("isTaskResumeRecord 与消毒结果一致", () => {
    expect(isTaskResumeRecord(baseRecord())).toBe(true);
    expect(isTaskResumeRecord({ ...baseRecord(), schemaVersion: 9 })).toBe(
      false
    );
  });
});

describe("taskResumeDecision 只让有检查点的那条说「继续」", () => {
  it("有安全检查点：说按当前授权重新读取并分析", () => {
    const decision = taskResumeDecision(
      baseRecord({
        phase: "planning",
        blocker: "network",
        resumability: "checkpoint"
      })
    );
    expect(decision.action).toBe("continue_checkpoint");
    expect(decision.canContinue).toBe(true);
    expect(decision.actionLabel).toBe("继续分析");
    // 正向锚点：只有这条文案允许出现"按当前授权重新读取并分析"
    expect(decision.detail).toContain("按当前授权重新读取并分析");
    // 负向：不得承诺"不重做已完成步骤"—— 引用可能已过期，取不到就要重读。
    expect(decision.detail).not.toContain("不会重做");
    expect(decision.detail).not.toContain("原样续跑");
  });

  it("无检查点：只能重新生成预览，且不得冒充精确续跑", () => {
    const decision = taskResumeDecision(
      baseRecord({
        phase: "planning",
        blocker: "network",
        resumability: "replan"
      })
    );
    expect(decision.action).toBe("regenerate_preview");
    // 正向锚点：文案必须点明"重新生成预览"
    expect(decision.detail).toContain("重新生成预览");
    // 负向：不得出现只有真检查点才能说的话
    expect(decision.detail).not.toContain("上次保存的进度");
    expect(decision.detail).not.toContain("按当前授权重新读取并分析");
    expect(decision.actionLabel).toBe("重新分析");
  });

  it("等待回答 + 需要重新授权：走 reauthorize，不是「已失效」", () => {
    const decision = taskResumeDecision(
      baseRecord({
        phase: "awaiting_clarification",
        blocker: "user_input",
        resumability: "reauthorize"
      })
    );
    expect(decision.action).toBe("reauthorize");
    expect(decision.canContinue).toBe(true);
    expect(decision.detail).toContain("重新确认数据范围");
  });

  it("执行结果未知：先核对，不盲目重跑", () => {
    const decision = taskResumeDecision(
      baseRecord({
        phase: "executing",
        blocker: "execution_unknown",
        resumability: "reconcile",
        writeState: "unknown"
      })
    );
    expect(decision.action).toBe("reconcile");
    expect(decision.detail).toContain("核对");
    expect(decision.detail).toContain("不会盲目重跑");
  });

  it("本机保存失败：明确不可继续，且不给主按钮", () => {
    const decision = taskResumeDecision(
      baseRecord({ blocker: "storage", resumability: "replan" })
    );
    expect(decision.action).toBe("unavailable");
    expect(decision.canContinue).toBe(false);
  });

  it("来源变化：仍可继续，但要重新读取", () => {
    const decision = taskResumeDecision(
      baseRecord({
        phase: "preview_ready",
        blocker: "source_changed",
        resumability: "replan"
      })
    );
    expect(decision.action).toBe("regenerate_preview");
    expect(decision.detail).toContain("重新读取");
  });
});

describe("写入状态文案不把「未知」说成「未修改」", () => {
  it("unknown 与 not_written 是两句不同的话", () => {
    expect(TASK_WRITE_STATE_LABEL.unknown).not.toBe(
      TASK_WRITE_STATE_LABEL.not_written
    );
    expect(TASK_WRITE_STATE_LABEL.unknown).toContain("未知");
    expect(TASK_WRITE_STATE_LABEL.not_written).toBe("原表未写入");
  });
});

describe("任务身份与状态推进", () => {
  it("newTaskId 稳定带前缀且互不相同", () => {
    const first = newTaskId();
    const second = newTaskId();
    expect(first.startsWith("task-")).toBe(true);
    expect(first).not.toBe(second);
  });

  it("withTaskResumeState 保持 taskId，只改状态位", () => {
    const record = baseRecord({ turnId: "turn-1" });
    const next = withTaskResumeState(record, {
      phase: "answer_submitted",
      blocker: "network",
      resumability: "replan"
    });
    expect(next.taskId).toBe(record.taskId);
    expect(next.turnId).toBe("turn-1");
    expect(next.phase).toBe("answer_submitted");
    expect(next.blocker).toBe("network");
  });

  it("withTaskResumeConfirmation 追加去重且仍满足消毒约束", () => {
    const record = baseRecord();
    const once = withTaskResumeConfirmation(record, {
      label: "分隔符",
      value: "、"
    });
    const twice = withTaskResumeConfirmation(once, {
      label: "分隔符",
      value: "、"
    });
    expect(twice.confirmations).toHaveLength(1);
    expect(sanitizeTaskResumeRecord(twice)).not.toBeNull();
  });
});
