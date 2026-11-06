import { describe, expect, it } from "vitest";
import capabilities from "../../../config/capabilities.json";
import {
  PENDING_DISCOVERY_LIMIT,
  PENDING_DISCOVERY_SCHEMA_VERSION,
  assertPendingExecutionDiscovery,
  orderPendingExecutionCandidates,
  pendingExecutionCandidateLabel,
  pendingExecutionSelection,
  type PendingExecutionCandidate,
  type PendingExecutionDiscoveryResponse
} from "./pendingExecutionDiscovery";

/**
 * 「重新发现未完成执行」的共享判据（2026-10-27 真实 WPS 现场）。
 *
 * 缺口：浏览器 `localStorage` 里的单条发现索引是**唯一**入口，完整退出宿主后它丢了，
 * 账本里的进度和原表现场都在，界面却再也找不到这次执行。修法是把索引降级为缓存，
 * 由后端按**当前工作簿稳定身份**精确查找。这组用例钉住的是发现结果的判定：
 * 谁算候选、能不能自动选、截断怎么说、响应里不许夹带什么。
 */

const WPS_IDENTITY = "wps-path:c:\\data\\恢复验收.xlsx";
const OTHER_IDENTITY = "wps-path:d:\\备份\\恢复验收.xlsx";

function candidate(overrides: Partial<PendingExecutionCandidate> = {}): PendingExecutionCandidate {
  return {
    operationId: "wps-execution-1fd938fd",
    taskId: "task-1",
    kind: "resume",
    status: "needs_reconciliation",
    reviewState: null,
    createdAt: 1_760_000_000_000,
    updatedAt: 1_760_000_100_000,
    nextActionIndex: 1,
    totalActions: 3,
    inFlightActionIndex: 1,
    ownerActive: false,
    requiresReconciliation: true,
    canResume: true,
    ...overrides
  };
}

function response(
  candidates: PendingExecutionCandidate[],
  overrides: Partial<PendingExecutionDiscoveryResponse> = {}
): PendingExecutionDiscoveryResponse {
  return {
    schemaVersion: PENDING_DISCOVERY_SCHEMA_VERSION,
    namespace: "wps",
    workbookIdentity: WPS_IDENTITY,
    limit: PENDING_DISCOVERY_LIMIT,
    truncated: false,
    candidates,
    ...overrides
  };
}

describe("pending execution discovery", () => {
  it("没有候选 ⇒ none，且不编造入口", () => {
    expect(pendingExecutionSelection(response([]), WPS_IDENTITY)).toEqual({ kind: "none" });
  });

  it("恰好一条候选 ⇒ single，原样带出这一条（不是「最近一条」而是唯一一条）", () => {
    const only = candidate();
    const selection = pendingExecutionSelection(response([only]), WPS_IDENTITY);
    expect(selection.kind).toBe("single");
    expect(selection.kind === "single" && selection.candidate.operationId).toBe(only.operationId);
  });

  it("多条候选 ⇒ 必须选择，绝不自动挑最新的一条", () => {
    const older = candidate({ operationId: "wps-execution-older", updatedAt: 1_760_000_050_000 });
    const newer = candidate({
      operationId: "wps-execution-newer",
      status: "failed",
      inFlightActionIndex: null,
      nextActionIndex: 0,
      updatedAt: 1_760_999_999_000
    });
    const selection = pendingExecutionSelection(response([older, newer]), WPS_IDENTITY);
    expect(selection.kind).toBe("choose");
    if (selection.kind !== "choose") throw new Error("必须要求用户选择");
    // 选择入口要能区分执行身份、时间与进度 —— 不是一条"最新任务"。
    expect(selection.candidates.map((item) => item.operationId).sort()).toEqual(
      [older.operationId, newer.operationId].sort()
    );
    expect(selection.truncated).toBe(false);
  });

  it("截断时即使只带回一条也必须让用户确认，不能冒充唯一", () => {
    const selection = pendingExecutionSelection(response([candidate()], { truncated: true }), WPS_IDENTITY);
    expect(selection.kind).toBe("choose");
  });

  it("未完成的排序：在途/待核对优先，其次按更新时间倒序，再按身份定序", () => {
    const writeBeforeFailure = candidate({
      operationId: "wps-execution-failed",
      status: "failed",
      inFlightActionIndex: null,
      nextActionIndex: 0,
      requiresReconciliation: false,
      updatedAt: 1_760_999_999_000
    });
    const inFlight = candidate({ operationId: "wps-execution-inflight", updatedAt: 1_760_000_000_000 });
    const other = candidate({
      operationId: "wps-execution-mid",
      status: "running",
      inFlightActionIndex: null,
      requiresReconciliation: false,
      updatedAt: 1_760_500_000_000
    });
    expect(orderPendingExecutionCandidates([writeBeforeFailure, other, inFlight]).map((item) => item.operationId))
      .toEqual(["wps-execution-inflight", "wps-execution-mid", "wps-execution-failed"]);
  });

  it("同一个身份出现两次只算一条（去重按 operationId）", () => {
    const duplicated = candidate();
    const selection = pendingExecutionSelection(response([duplicated, { ...duplicated }]), WPS_IDENTITY);
    expect(selection.kind).toBe("single");
  });

  it("已完成执行不许作为待续写任务出现", () => {
    /*
     * `completed` **不在** `PendingExecutionStatus` 里（准入闸会拒），这里刻意用一个
     * 类型之外的状态模拟"服务端旧记录/别家载荷"，验证选择器不会把它当成待续写任务。
     * 用 `as unknown as` 是表明"这本来就不是合法值"，不是把类型当摆设。
     */
    const completed = candidate({
      status: "completed" as unknown as PendingExecutionCandidate["status"],
      inFlightActionIndex: null,
      requiresReconciliation: false
    });
    expect(pendingExecutionSelection(response([completed]), WPS_IDENTITY)).toEqual({ kind: "none" });
  });

  it("同名不同路径：身份不是完整相等就一条都不认", () => {
    const selection = pendingExecutionSelection(
      response([candidate()], { workbookIdentity: OTHER_IDENTITY }),
      WPS_IDENTITY
    );
    expect(selection.kind).toBe("mismatch");
    expect(pendingExecutionSelection(response([candidate()]), OTHER_IDENTITY).kind).toBe("mismatch");
    // 只有文件名相同也仍然不认。
    expect(pendingExecutionSelection(response([candidate()]), "wps-path:c:\\other\\恢复验收.xlsx").kind)
      .toBe("mismatch");
  });

  it("响应里夹带 ownerToken / 恢复包正文 / 值矩阵 / 完整计划一律拒绝", () => {
    const withOwner = { ...response([candidate()]), ownerToken: "owner-token" };
    expect(() => assertPendingExecutionDiscovery(withOwner)).toThrow(/不允许|未知字段/);
    for (const leaked of ["descriptor", "plan", "packageJson", "values", "material", "beforeSnapshot"]) {
      const payload = { ...response([candidate()]), [leaked]: { anything: true } };
      expect(() => assertPendingExecutionDiscovery(payload)).toThrow();
    }
    const candidateLeak = response([{ ...candidate(), values: [[111]] } as unknown as PendingExecutionCandidate]);
    expect(() => assertPendingExecutionDiscovery(candidateLeak)).toThrow();
  });

  it("候选的形状与进度必须自洽（在途动作落在外，游标越界，都拒绝）", () => {
    const badCursor = response([candidate({ nextActionIndex: 4, totalActions: 3 })]);
    expect(() => assertPendingExecutionDiscovery(badCursor)).toThrow();
    const badInFlight = response([candidate({ inFlightActionIndex: 3 })]);
    expect(() => assertPendingExecutionDiscovery(badInFlight)).toThrow();
  });

  it("候选只给定位、时间与进度，文案里不出现任何单元格内容", () => {
    const label = pendingExecutionCandidateLabel(candidate());
    expect(label).toContain("第 2 步");
    expect(label).toContain("共 3 步");
    expect(label).toContain("待核对");
    // 111/222 是用户现场的真实值；候选文案里一个都不许出现。
    expect(label).not.toMatch(/111|222|333/);
  });

  it("限额来自 config/capabilities.json，不在代码里另写一个数", () => {
    expect(PENDING_DISCOVERY_LIMIT).toBe(capabilities.taskRecovery.pendingDiscoveryMaxCandidates);
    expect(PENDING_DISCOVERY_LIMIT).toBeGreaterThan(1);
  });

  it("候选超过限额时按限额收敛并把截断说清楚", () => {
    const many = Array.from({ length: PENDING_DISCOVERY_LIMIT + 3 }, (_, index) =>
      candidate({ operationId: `wps-execution-${String(index).padStart(4, "0")}`, updatedAt: 1_760_000_000_000 + index })
    );
    const selection = pendingExecutionSelection(response(many, { truncated: true }), WPS_IDENTITY);
    if (selection.kind !== "choose") throw new Error("必须要求用户选择");
    expect(selection.candidates).toHaveLength(PENDING_DISCOVERY_LIMIT);
    expect(selection.truncated).toBe(true);
  });
});

describe('待审阅也是"还没处理完"，但不是"待续写"（2026-11-03）', () => {
  /** 写完、回读通过、用户还没点确认的那一次执行。 */
  function reviewCandidate(overrides: Partial<PendingExecutionCandidate> = {}): PendingExecutionCandidate {
    return candidate({
      kind: 'review',
      status: 'completed',
      reviewState: 'pending',
      nextActionIndex: 3,
      totalActions: 3,
      inFlightActionIndex: null,
      requiresReconciliation: false,
      canResume: false,
      ...overrides
    });
  }

  function response(
    withCandidates: PendingExecutionCandidate[],
    truncated = false
  ): PendingExecutionDiscoveryResponse {
    return {
      schemaVersion: PENDING_DISCOVERY_SCHEMA_VERSION,
      namespace: 'wps' as const,
      workbookIdentity: WPS_IDENTITY,
      limit: 5,
      truncated,
      candidates: withCandidates
    };
  }

  it('已完成 + 显式待审阅 ⇒ 一条可用的审阅入口（不会被当成"没有未完成任务"）', () => {
    const selection = pendingExecutionSelection(response([reviewCandidate()]), WPS_IDENTITY);
    expect(selection.kind).toBe('single');
    if (selection.kind !== 'single') return;
    expect(selection.candidate.kind).toBe('review');
    expect(selection.candidate.reviewState).toBe('pending');
    // 待审阅**没有**续跑资格：一步都不能再写。
    expect(selection.candidate.canResume).toBe(false);
  });

  it('已完成但没有显式审阅状态 ⇒ 一律丢弃（老记录不许批量变成待审阅）', () => {
    // `kind: 'resume'` 却 `completed`：自相矛盾，两张皮都不认。
    const contradictory = reviewCandidate({ kind: 'resume', reviewState: null });
    expect(pendingExecutionSelection(response([contradictory]), WPS_IDENTITY)).toEqual({ kind: 'none' });
    // `kind: 'review'` 却没有审阅状态：不是"待审阅"，是"不知道"。
    const unknown = reviewCandidate({ reviewState: null });
    expect(pendingExecutionSelection(response([unknown]), WPS_IDENTITY)).toEqual({ kind: 'none' });
  });

  it('准入闸拒绝自相矛盾的候选，且不会把 completed 混进待续写', () => {
    expect(() => assertPendingExecutionDiscovery(response([reviewCandidate({ kind: 'resume' })])))
      .toThrow(/待续写/);
    expect(() => assertPendingExecutionDiscovery(response([reviewCandidate({ reviewState: null })])))
      .toThrow(/待审阅/);
    // 缺少性质标记：既不能当续写也不能当审阅。
    const missingKind = { ...reviewCandidate() } as Record<string, unknown>;
    delete missingKind.kind;
    expect(() => assertPendingExecutionDiscovery(response([missingKind as never])))
      .toThrow(/性质标记/);
  });

  it('待审阅与"有在途动作"同档最前：内容已经落表，必须先看见', () => {
    const failedBeforeWrite = candidate({ operationId: 'zzz', status: 'failed', nextActionIndex: 0, inFlightActionIndex: null, updatedAt: 1 });
    const ordered = orderPendingExecutionCandidates([failedBeforeWrite, reviewCandidate({ operationId: 'aaa', updatedAt: 1 })]);
    expect(ordered[0].kind).toBe('review');
  });

  it('候选行的说明不许用续跑口径（那会让人以为还能"继续执行"）', () => {
    const text = pendingExecutionCandidateLabel(reviewCandidate());
    expect(text).toContain('等待你确认或放弃');
    expect(text).not.toContain('已写 3 步');
    expect(pendingExecutionCandidateLabel(reviewCandidate({ reviewState: 'blocked' })))
      .toContain('未通过独立验证');
  });
});
