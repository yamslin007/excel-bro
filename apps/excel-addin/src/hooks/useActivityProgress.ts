import { useCallback, useEffect, useState } from "react";

export interface ActivityStep {
  label: string;
  elapsedMs: number;
  // note：模型对这一步的真实判断（第一层，默认可见），例如「它把需求理解成了什么」。
  note?: string;
  // detail：原始明细（第二层，展开「查看详情」才显示），例如锁定的执行指令、结构化返回。
  detail?: string;
}

export interface ActivityProgress {
  title: string;
  detail: string;
  completed: ActivityStep[];
  startedAt: number;
  lastStepAt: number;
}

export interface ActivityLog {
  steps: ActivityStep[];
  totalMs: number;
}

interface UseActivityProgressOptions {
  onPersistLog: (log: ActivityLog) => void;
}

/**
 * 活动进度管理 Hook
 *
 * 职责：
 * - 管理当前执行/规划进度的标题、明细和已完成步骤
 * - 管理进度计时器（按秒刷新）
 * - 将当前 activity 的步骤固化为日志并通知上层
 */
export function useActivityProgress({
  onPersistLog
}: UseActivityProgressOptions) {
  const [activity, setActivity] = useState<ActivityProgress | null>(null);
  const [activitySeconds, setActivitySeconds] = useState(0);

  const startActivity = useCallback((title: string, detail: string) => {
    const now = Date.now();
    setActivity({
      title,
      detail,
      completed: [],
      startedAt: now,
      lastStepAt: now
    });
    setActivitySeconds(0);
  }, []);

  const advanceActivity = useCallback(
    (
      title: string,
      detail: string,
      completedStep?: string,
      stepInsight?: { note?: string; detail?: string }
    ) => {
      setActivity((current) => {
        const now = Date.now();
        const startedAt = current?.startedAt ?? now;
        const lastStepAt = current?.lastStepAt ?? startedAt;
        const completed =
          completedStep && completedStep.trim()
            ? [
                ...(current?.completed ?? []),
                {
                  label: completedStep,
                  elapsedMs: now - lastStepAt,
                  note: stepInsight?.note,
                  detail: stepInsight?.detail
                }
              ]
            : current?.completed ?? [];
        return {
          title,
          detail,
          completed,
          startedAt,
          lastStepAt:
            completedStep && completedStep.trim() ? now : lastStepAt
        };
      });
    },
    []
  );

  const updateActivityDetail = useCallback((detail: string) => {
    setActivity((current) =>
      current
        ? {
            ...current,
            detail
          }
        : current
    );
  }, []);

  const completeActivity = useCallback(() => {
    if (activity && activity.completed.length > 0) {
      onPersistLog({
        steps: activity.completed,
        totalMs: Date.now() - activity.startedAt
      });
    }
    setActivity(null);
    setActivitySeconds(0);
  }, [activity, onPersistLog]);

  useEffect(() => {
    if (!activity) return;
    const updateElapsed = () =>
      setActivitySeconds(
        Math.max(0, Math.floor((Date.now() - activity.startedAt) / 1000))
      );
    updateElapsed();
    const timer = window.setInterval(updateElapsed, 1000);
    return () => window.clearInterval(timer);
  }, [activity?.startedAt]);

  return {
    activity,
    activitySeconds,
    startActivity,
    advanceActivity,
    updateActivityDetail,
    completeActivity
  };
}
