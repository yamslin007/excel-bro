import { describe, expect, it } from "vitest";
import type { AnalysisPlan, ExcelAction } from "../../excel-addin/src/contracts";
import type { HostWriteAudit } from "@excel-bro/core/hostWriteAudit";
import { summarizeHostWrites } from "@excel-bro/core/hostWriteAudit";
import { capabilityHost } from "./wpsCapabilityTestHost";
import { executeWpsPlan, prepareWpsPlan, wpsUndoRecordFor } from "./wpsPlanExecutor";
import { beginHostWriteScope, countHostWrite, endHostWriteScope, hostWriteScopeActive } from "./wpsHostWriteCounter";
import { writeWpsMatrix } from "./wpsWriteMatrix";

/**
 * WPS 端「宿主观测」的判据（A1-E3 第三批）。
 *
 * 过去 WPS 只能给"请求次数 + 认领结果 + 回读回执"三档 —— `hostWriteCalls` 那一档
 * 是空的（❌ 未实现，不是宿主限制）。这里钉住补上之后能证明什么：
 *
 * 1. 计数发生在**真正赋值给宿主对象**的地方：一个动作分块写几次就是几次；
 * 2. `mutatingCalls` 只数承载**计划内容**的调用；同一动作里为保持文本原样而临时
 *    改 `NumberFormat` 的辅助调用只进 `hostWriteCalls`；
 * 3. 写前失败 → `failed_before_write` 且 `hostWriteCalls === 0`（不许并进"已写入"）；
 *    但"动作没被认领"（执行边界之外）连条目都不该有 —— 缺席 ≠ 记零；
 * 4. 宿主写 API 抛错 → `failed_after_write`（可能已部分生效），不是零；
 * 5. 撤销 / 审阅副本 / 恢复点这些**辅助写**一律不计（作用域之外）；
 * 6. 恢复只写剩余动作时，只有被真正执行的那几步有观测。
 */

const plan = (actions: ExcelAction[]): AnalysisPlan => ({
  id: "wps-host-writes",
  title: "宿主观测",
  summary: "宿主观测",
  actions,
  assumptions: [],
  warnings: [],
  acceptanceCriteria: []
});

const values = (range: string, matrix: (number | string)[][]): ExcelAction => ({
  type: "writeValues",
  sheet: "明细",
  range,
  values: matrix as ExcelAction extends { values: infer V } ? V : never
});

function collect() {
  const entries: HostWriteAudit[] = [];
  return { entries, onHostWrite: (entry: HostWriteAudit) => { entries.push(entry); } };
}

describe("WPS 宿主写调用观测", () => {
  it("一个动作只写一块时：主机调用 1 次，结果 written + 回读通过", async () => {
    const host = capabilityHost(["明细"], 4, 4);
    const prepared = prepareWpsPlan(host.application, [1], plan([values("A1:B1", [[10, 20]])]));
    const { entries, onHostWrite } = collect();

    await executeWpsPlan(host.application, prepared, undefined, undefined, undefined, "exec-hw-1", undefined, onHostWrite);

    expect(entries).toHaveLength(1);
    expect(entries[0]).toMatchObject({
      actionIndex: 0,
      actionType: "writeValues",
      hostWriteCalls: 1,
      mutatingCalls: 1,
      outcome: "written",
      readback: "verified"
    });
    // 与"单元格写入个数"不是一回事：2 个单元格只走了 1 次宿主写调用。
    expect(host.writes()).toBe(2);
  });

  it("分块写：块数就是宿主写调用次数（不是单元格数，也不是动作数）", async () => {
    const host = capabilityHost(["明细"], 120, 4);
    const rows = Array.from({ length: 120 }, (_, r) => [r + 1] as number[]);
    const prepared = prepareWpsPlan(host.application, [1], plan([values("A1:A120", rows)]));
    const { entries, onHostWrite } = collect();

    await executeWpsPlan(host.application, prepared, undefined, undefined, undefined, "exec-hw-2", undefined, onHostWrite);

    expect(entries).toHaveLength(1);
    // chunkRows = 100 ⇒ 120 行分两块各一次。
    expect(entries[0].hostWriteCalls).toBe(2);
    expect(entries[0].mutatingCalls).toBe(2);
    expect(host.writes()).toBe(120);
  });

  it("文本保真用的 NumberFormat 只进主机调用数，不进计划写调用数", async () => {
    const host = capabilityHost(["明细"], 4, 4);
    const prepared = prepareWpsPlan(host.application, [1], plan([values("A1:A2", [["001"], ["002"]])]));
    const { entries, onHostWrite } = collect();

    await executeWpsPlan(host.application, prepared, undefined, undefined, undefined, "exec-hw-3", undefined, onHostWrite);

    expect(entries).toHaveLength(1);
    // 1 次值写入 + 2 格 @ 格式 + 2 格恢复 = 5 次宿主写调用；计划写只有 1 次。
    expect(entries[0].hostWriteCalls).toBe(5);
    expect(entries[0].mutatingCalls).toBe(1);
  });

  it("预览后目标就变了：连动作都没认领，一条观测都不该有（不是 failed_before_write）", async () => {
    const host = capabilityHost(["明细"], 4, 4);
    const prepared = prepareWpsPlan(host.application, [1], plan([values("A1:A1", [[1]])]));
    const { entries, onHostWrite } = collect();
    // 预览之后有人改了目标区域：恢复快照构建阶段就拒，进不到动作循环。
    host.grids[0][0][0].value = "someone-else";

    await expect(
      executeWpsPlan(host.application, prepared, undefined, undefined, undefined, "exec-hw-4", undefined, onHostWrite)
    ).rejects.toThrow("在预览后发生变化");

    /*
     * 这里**不能**记一条 `failed_before_write`：那个状态的含义是"动作已认领、
     * 进入了执行边界、随后在写 API 之前失败"。本条走的是执行边界之外，
     * 它的事实是"这次执行没有任何动作被认领"，观测应当缺席而不是记零。
     */
    expect(entries).toHaveLength(0);
    expect(summarizeHostWrites(entries).hostWriteCalls).toBe(0);
    expect(host.writes()).toBe(0);
  });

  it("认领被拒：failed_before_write 且主机调用 0 次（不算已写入）", async () => {
    const host = capabilityHost(["明细"], 4, 4);
    const prepared = prepareWpsPlan(host.application, [1], plan([values("A1:A1", [[1]])]));
    const { entries, onHostWrite } = collect();

    await expect(
      executeWpsPlan(host.application, prepared, undefined, undefined, async (phase) => {
        // 写前检查点 = 向账本认领这个动作；服务端拒绝时一步都不许写。
        if (phase === "before") throw new Error("claim rejected by ledger");
      }, "exec-hw-4", { taskId: "task-hw" }, onHostWrite)
    ).rejects.toThrow("claim rejected by ledger");

    expect(entries).toHaveLength(1);
    expect(entries[0]).toMatchObject({
      actionIndex: 0,
      hostWriteCalls: 0,
      outcome: "failed_before_write",
      readback: "not_applicable"
    });
    expect(summarizeHostWrites(entries).hostWriteCalls).toBe(0);
    expect(host.writes()).toBe(0);
  });

  it("宿主写 API 抛错：按 failed_after_write 记（可能已部分生效），不是零", async () => {
    const host = capabilityHost(["明细"], 4, 4);
    const prepared = prepareWpsPlan(host.application, [1], plan([values("A1:A1", [[7]])]));
    const { entries, onHostWrite } = collect();
    const base = host.sheets[0].Range!;
    // 只让写抛错：读（指纹核对）照旧。
    host.sheets[0].Range = (address: string) =>
      new Proxy(base(address), {
        set(target, property, value) {
          if (property === "Value2") throw new Error("host write refused");
          return Reflect.set(target, property, value);
        }
      });

    await expect(
      executeWpsPlan(host.application, prepared, undefined, undefined, undefined, "exec-hw-5", undefined, onHostWrite)
    ).rejects.toThrow("host write refused");

    expect(entries).toHaveLength(1);
    expect(entries[0].outcome).toBe("failed_after_write");
    expect(entries[0].hostWriteCalls).toBe(1);
    expect(entries[0].errorCode).toBe("WPS_ACTION_FAILED");
  });

  it("撤销与辅助写不进观测（作用域之外一次都不计）", async () => {
    const host = capabilityHost(["明细"], 4, 4);
    const prepared = prepareWpsPlan(host.application, [1], plan([values("A1:A1", [[9]])]));
    const { entries, onHostWrite } = collect();

    const record = await executeWpsPlan(
      host.application, prepared, undefined, undefined, undefined, "exec-hw-6", undefined, onHostWrite
    );
    const afterExecute = entries.length;
    expect(afterExecute).toBe(1);

    // 辅助写直接调用（撤销/审阅副本走的就是这条），不在任何作用域里。
    expect(hostWriteScopeActive()).toBe(false);
    writeWpsMatrix(host.sheets[0].Range!("A1:A1"), [[99]]);
    expect(entries).toHaveLength(afterExecute);
    expect(wpsUndoRecordFor(host.application, record.executionId)?.state).toBe("verified");
  });

  it("观测回调抛错不影响执行：这是统计，不是判据", async () => {
    const host = capabilityHost(["明细"], 4, 4);
    const prepared = prepareWpsPlan(host.application, [1], plan([values("A1:A1", [[11]])]));

    const record = await executeWpsPlan(
      host.application, prepared, undefined, undefined, undefined, "exec-hw-7", undefined,
      () => { throw new Error("collector blew up"); }
    );

    expect(record.state).toBe("verified");
    expect(host.grids[0][0][0].value).toBe(11);
  });

  it("恢复只写剩余动作时，只有被执行的那几步有观测", async () => {
    const host = capabilityHost(["明细"], 4, 4);
    const prepared = prepareWpsPlan(host.application, [1], plan([
      values("A1:A1", [[1]]),
      values("A2:A2", [[2]]),
      values("A3:A3", [[3]])
    ]));
    const first = collect();
    // 第二步写完之后回执丢了：动作已写，账本还没收到写后检查点。
    await expect(
      executeWpsPlan(host.application, prepared, undefined, undefined, async (phase, checkpoint) => {
        if (phase === "after" && checkpoint.actionIndex === 1) throw new Error("lost receipt");
      }, "exec-hw-8", { taskId: "task-hw" }, first.onHostWrite)
    ).rejects.toThrow("lost receipt");
    expect(first.entries.map(entry => entry.actionIndex)).toEqual([0, 1]);
    expect(first.entries[1].outcome).toBe("failed_after_write");

    const record = wpsUndoRecordFor(host.application, "exec-hw-8");
    expect(record).not.toBeNull();
    const second = collect();
    await executeWpsPlan(host.application, prepared, undefined, undefined, undefined, "exec-hw-8",
      { taskId: "task-hw", resume: { record: record!, startActionIndex: 2, check: () => undefined } },
      second.onHostWrite);

    // 只续第三步：前两步**一次都没有**再被记成计划写调用。
    expect(second.entries.map(entry => entry.actionIndex)).toEqual([2]);
    expect(second.entries[0].hostWriteCalls).toBe(1);
  });

  it("作用域可嵌套且退出后恢复外层计数", () => {
    const outer = beginHostWriteScope();
    countHostWrite(true);
    const inner = beginHostWriteScope();
    countHostWrite(true);
    countHostWrite(false);
    expect(endHostWriteScope(inner)).toEqual({ calls: 2, mutating: 1 });
    countHostWrite(true);
    expect(endHostWriteScope(outer)).toEqual({ calls: 2, mutating: 2 });
    expect(hostWriteScopeActive()).toBe(false);
  });
});
