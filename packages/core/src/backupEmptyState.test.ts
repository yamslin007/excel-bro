import { describe, expect, it } from "vitest";
import { backupEmptyState } from "./backupEmptyState";

describe("备份抽屉空态：两种真实含义不许混成一句", () => {
  it("有可选中恢复点时不显示空态", () => {
    expect(backupEmptyState({ hasUndoRecord: true })).toBeNull();
    // 即便最近一次是"已撤销"，只要还留着更早的可撤销恢复点，也不算空态。
    expect(backupEmptyState({ hasUndoRecord: true, latestExecutionUndone: true })).toBeNull();
  });

  it("没有恢复点、最近一次也没撤销 ⇒「暂无恢复点」，且说明写在**写入前**", () => {
    const view = backupEmptyState({ hasUndoRecord: false });
    expect(view).not.toBeNull();
    expect(view!.kind).toBe("none");
    expect(view!.title).toBe("暂无恢复点");
    /*
     * 这条断言是这次修改的核心：文案必须说"写入**前**就保存快照"。
     * 旧文案是"只有确认并完成写入后才会创建"，把保护说成了确认之后的事 —— 说反了。
     */
    expect(view!.hint).toContain("写入前");
    expect(view!.hint).not.toContain("确认");
    expect(view!.hint).not.toContain("确认变更后");
  });

  it("最近一次执行已撤销 ⇒「本次执行已撤销」，并说清恢复点是被用掉而不是没有过", () => {
    const view = backupEmptyState({ hasUndoRecord: false, latestExecutionUndone: true });
    expect(view).not.toBeNull();
    expect(view!.kind).toBe("undone");
    expect(view!.title).toBe("本次执行已撤销");
    expect(view!.hint).toContain("写入之前");
    expect(view!.hint).toContain("回收");
    // 不许把它说成"没有保护过"。
    expect(view!.hint).not.toContain("还没有恢复点");
  });

  it("两种空态的标题不同（反向变异：合并成一句必须变红）", () => {
    const none = backupEmptyState({ hasUndoRecord: false })!;
    const undone = backupEmptyState({ hasUndoRecord: false, latestExecutionUndone: true })!;
    expect(none.title).not.toBe(undone.title);
    expect(none.kind).not.toBe(undone.kind);
  });
});
