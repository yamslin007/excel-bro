/**
 * 自动审阅工作流：生成计划后自动生成副本展示结果（Excel / WPS 共用语义）。
 *
 * ## 核心变化
 *
 * **旧流程**：
 * 1. 用户选择"直接执行"或"先生成副本比对"
 * 2. 选择副本比对时，点击"生成副本并排审阅"按钮
 * 3. 副本生成后展示修改后的结果
 * 4. 用户核对后点击"确认写回原表"
 *
 * **新流程**：
 * 1. 合法计划通过授权校验后，**自动生成并执行到审阅副本**
 * 2. **默认展示修改后的待确认结果**，标记变化区域
 * 3. 可选"双视图对比"查看修改前后对比
 * 4. 确认后写回原表，或放弃变更
 * 5. 确认成功后显示"保存为工具"和"撤销本次变更"
 *
 * ## 自动生成的触发条件
 *
 * 满足以下**所有**条件时自动生成：
 * 1. 计划通过白名单校验
 * 2. 计划通过授权范围检查
 * 3. 宿主支持副本生成（原生或兼容模式可用）
 * 4. 本次计划尚未生成过副本（避免重复生成）
 *
 * ## 单次生成保证
 *
 * 只在**首次满足条件**时生成一次：
 * - 流式更新、组件重渲染不重复生成
 * - 展开明细、切换视图不重复生成
 * - 计划卡滚动出视口再回来不重复生成
 *
 * ## 副本失效条件
 *
 * 以下情况旧结果与旧确认资格失效，必须重新生成：
 * - 计划被编辑（planId 变化）
 * - 授权范围变化（授权的工作表集合变化）
 * - 切换工作簿（workbookIdentity 变化）
 * - 用户明确选择"放弃变更"
 *
 * ## 视图切换
 *
 * 默认单视图（展示修改后的结果），可选双视图对比：
 * - 单视图：只显示审阅副本（修改后）
 * - 双视图：原表与副本并排对比
 *
 * 切换视图**只改变展示方式**，不：
 * - 重新规划
 * - 重复执行
 * - 重新创建副本
 * - 改变用户授权
 * - 改变确认状态
 */

/**
 * 审阅副本的自动生成状态。
 *
 * 追踪"这份计划是否已经自动生成过副本"。
 */
export type AutoReviewState = {
  /** 已自动生成副本的 planId 集合（避免重复生成）。 */
  generated: Set<string>;

  /** 当前审阅会话的身份（用于失效检测）。 */
  current: AutoReviewIdentity | null;
};

/**
 * 审阅会话身份：用于检测副本是否失效。
 */
export type AutoReviewIdentity = {
  /** 计划唯一标识。 */
  planId: string;

  /** 工作簿身份。 */
  workbookIdentity: string;

  /** 授权的工作表集合（字符串化的 JSON）。 */
  authorizationKey: string;
};

/**
 * 视图模式。
 */
export type ReviewViewMode =
  | 'single'      // 单视图：只显示修改后的结果
  | 'dual';       // 双视图：原表与副本并排对比

/**
 * 自动审阅工作流的可用性。
 */
export type AutoReviewAvailability = {
  /** 是否可以生成审阅副本。 */
  available: boolean;

  /** 不可用时的原因。 */
  reason: string | null;
};

/**
 * 检查是否应该自动生成审阅副本。
 *
 * @param state 当前自动审阅状态
 * @param identity 本次审阅身份
 * @param availability 副本生成可用性
 * @returns 是否应该生成
 */
export function shouldAutoGenerateReview(
  state: AutoReviewState,
  identity: AutoReviewIdentity,
  availability: AutoReviewAvailability
): boolean {
  // 宿主不支持副本生成
  if (!availability.available) {
    return false;
  }

  // 已经生成过这个 planId
  if (state.generated.has(identity.planId)) {
    return false;
  }

  // 应该生成
  return true;
}

/**
 * 记录已自动生成的副本。
 *
 * @param state 当前状态
 * @param identity 审阅身份
 * @returns 新状态
 */
export function markReviewGenerated(
  state: AutoReviewState,
  identity: AutoReviewIdentity
): AutoReviewState {
  const generated = new Set(state.generated);
  generated.add(identity.planId);

  return {
    generated,
    current: identity
  };
}

/**
 * 检查审阅副本是否已失效。
 *
 * @param state 当前状态
 * @param identity 新的审阅身份
 * @returns 是否失效
 */
export function isReviewInvalidated(
  state: AutoReviewState,
  identity: AutoReviewIdentity
): boolean {
  if (!state.current) {
    return false; // 没有当前审阅，谈不上失效
  }

  // 任何身份字段变化都导致失效
  return state.current.planId !== identity.planId ||
    state.current.workbookIdentity !== identity.workbookIdentity ||
    state.current.authorizationKey !== identity.authorizationKey;
}

/**
 * 使当前审阅失效（用户放弃变更、或计划被编辑等场景）。
 *
 * @param state 当前状态
 * @returns 新状态（清除当前审阅）
 */
export function invalidateCurrentReview(
  state: AutoReviewState
): AutoReviewState {
  return {
    generated: state.generated, // 保留生成记录，避免重复生成同一 planId
    current: null
  };
}

/**
 * 重置自动审阅状态（用于工作簿切换等场景）。
 */
export function resetAutoReviewState(): AutoReviewState {
  return {
    generated: new Set(),
    current: null
  };
}

/**
 * 序列化授权集合为可比较的键。
 *
 * @param authorization 授权的工作表名称数组
 * @returns 序列化的键
 */
export function serializeAuthorizationKey(authorization: readonly string[]): string {
  return JSON.stringify([...authorization].sort());
}
