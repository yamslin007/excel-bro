import {reportWpsHostWrites} from "../../excel-addin/src/api";
import {armRecoveryProbe,bindRecoveryProbe,interruptRecoveryProbe} from "@excel-bro/core/recoveryProbe";
import {inspectWpsRecoveryState} from "./wpsRecoveryInspection";
import {fixture, operationId, recoveryChannel} from "./wpsResume.testHost";
import {backendRecovery, resetBackendRecovery} from "./wpsRecoveryStore.testHost";
import {buildWpsRecoveryPackage, ensureWpsRecoveryMaterial, retireWpsRecoveryPackage} from "./wpsRecoveryPackage";
import {beforeEach, describe, expect, it, vi} from "vitest";
import type {WpsExecutionState} from "../../excel-addin/src/contracts";
import {assertWpsRecoveryDescriptor, assertWpsExecutionResult} from "../../excel-addin/src/contracts";
import {adoptWpsUndoRecord, executeWpsPlan, inspectWpsUndoRecord, latestWpsUndo, matrixFingerprintV2, prepareWpsPlan, prepareWpsResume, undoWpsExecution, wpsResumeWindow, wpsUndoRecordFor, wpsResumeContext} from "./wpsPlanExecutor";
import {canResumeWpsExecution, resumeWpsExecution, wpsResumeUnavailableReason} from "./wpsExecutionResume";
import {wpsReconcileVerdict, wpsReconcileActions} from "./wpsExecutionReconcile";
import * as ledger from "./wpsExecutionLedger";
import capabilities from "../../../config/capabilities.json";
import {stampLiteralPlanOrigin} from "@excel-bro/core/literalPlanOrigin";

vi.mock("../../excel-addin/src/api", async original => {
  const actual = await original<typeof import("../../excel-addin/src/api")>();
  const store = await import("./wpsRecoveryStore.testHost");
  /*
   * 只把"本机后端那三个接口"换成本机替身，其余照旧。
   *
   * 必须**真的**走 `saveExecutionRecoveryPackage` / `loadExecutionRecoveryPackage` 的
   * 语义（正文文本进出、读回逐字节比对），否则"写前落盘失败必须零宿主写入"
   * "落盘后读回不一致就停"这些判据会被一个"永远成功"的假函数顺手抹掉。
   */
  return {...actual,
    reportWpsHostWrites: vi.fn(async () => ({})),
    saveExecutionRecoveryPackage: vi.fn(async (id: string, text: string) => store.backendSave(id, text)),
    loadExecutionRecoveryPackage: vi.fn(async (id: string) => store.backendLoad(id)),
    deleteExecutionRecoveryPackage: vi.fn(async (id: string) => store.backendRemove(id))};
});
vi.mock("./wpsExecutionLedger", async original => ({...await original<typeof import("./wpsExecutionLedger")>(),
  claimWpsExecutionResume: vi.fn(), startWpsAction: vi.fn(), completeWpsAction: vi.fn(),
  finishWpsExecution: vi.fn(), suspendWpsExecution: vi.fn()}));
async function interrupted(next=2) {
  const f=fixture();
  await expect(executeWpsPlan(f.application,f.prepared,undefined,undefined,async (phase,c) => {
    if (phase === "after" && c.actionIndex === next-1) throw new Error("lost receipt");
  },operationId,{taskId:"task-1",...recoveryChannel})).rejects.toThrow("lost receipt");
  f.state.nextActionIndex=next; // Production reconciliation has proved the last applied action.
  f.writes.length=0;
  return f;
}
function wire(state: WpsExecutionState, onClaim?:()=>void, onStart?:()=>void) {
  let owned={...state,ownerActive:true,ownerToken:"owner-new",revision:state.revision+1};
  vi.mocked(ledger.claimWpsExecutionResume).mockImplementation(async()=>{onClaim?.();return owned;});
  vi.mocked(ledger.startWpsAction).mockImplementation(async(_id,req)=>{onStart?.();
    expect(req.ownerToken).toBe("owner-new"); expect(req.actionIndex).toBe(owned.nextActionIndex);
    owned={...owned,inFlightActionIndex:req.actionIndex,revision:owned.revision+1};return owned;});
  vi.mocked(ledger.completeWpsAction).mockImplementation(async(_id,req)=>{
    expect(req.actionIndex).toBe(owned.inFlightActionIndex);
    owned={...owned,inFlightActionIndex:null,nextActionIndex:req.actionIndex+1,revision:owned.revision+1};return owned;});
  vi.mocked(ledger.finishWpsExecution).mockImplementation(async(_id, request)=>{assertWpsExecutionResult(request.result);return {...owned,status:"completed"};});
  vi.mocked(ledger.suspendWpsExecution).mockImplementation(async()=>({...owned,ownerActive:false}));
}
beforeEach(()=>{vi.clearAllMocks();resetBackendRecovery();});
describe("WPS exact remaining action recovery",()=>{
  it("ordered trusted literals survive lost host cache and resume only third without reattesting JSON", async()=>{
    const f=fixture();
    const plan=stampLiteralPlanOrigin({...f.prepared.plan,sourceDependencySheets:[],
      actions:[111,222,333].map((value,index)=>({type:"writeValues" as const,sheet:"Target",range:`A${index+1}`,values:[[value]]}))},"literal_range_edit");
    const prepared=prepareWpsPlan(f.application,[1,2],plan);
    expect(wpsResumeContext(prepared)).toEqual({version:2,hostVersion:"12.0",sources:[]});
    expect(wpsResumeContext(prepareWpsPlan(f.application,[1,2],JSON.parse(JSON.stringify(plan))))).toBeNull();
    f.state.descriptor=ledger.wpsRecoveryDescriptor(prepared);
    await expect(executeWpsPlan(f.application,prepared,undefined,undefined,async(phase,c)=>{
      if(phase==="after"&&c.actionIndex===1) throw new Error("lost literal receipt");
    },operationId,{taskId:"task-1",...recoveryChannel})).rejects.toThrow("lost literal receipt");
    expect([1,2,3].map(i=>f.cells.get(`Target!A${i}`))).toEqual([111,222,"old3"]);
    f.state.nextActionIndex=2; // The existing in-flight reconciliation proves action 1, without writing.
    f.storage.clear(); f.writes.length=0;
    const reopened={...f.application};
    expect((await ensureWpsRecoveryMaterial(reopened,f.state)).kind).toBe("package");
    wire(f.state);
    const result=await resumeWpsExecution(reopened,f.state,()=>[1,2]);
    expect(result.state).toBe("verified");
    expect(f.writes).toEqual(["Target!A3"]);
    expect(vi.mocked(ledger.startWpsAction).mock.calls.map(call=>call[1].actionIndex)).toEqual([2]);
    expect([1,2,3].map(i=>f.cells.get(`Target!A${i}`))).toEqual([111,222,333]);
    await undoWpsExecution(reopened,operationId);
    expect([1,2,3].map(i=>f.cells.get(`Target!A${i}`))).toEqual(["old1","old2","old3"]);
  });
  /*
   * 2026-10-27 真实 WPS 验收的缺陷回归：完整退出 WPS 再打开，`PluginStorage` 里读不到
   * 原执行记录，于是"已核对第二步、仍可继续第三步"在真机上断链。
   *
   * 为什么这个用例**不能**用宿主临时文件（`Env.GetTempPath` + `FileSystem`）来替：
   * 那是另一套存储方案，不是本产品要落地的东西；把替身写成那样，等于用被测结构自己
   * 编一份"证据"。这里换成真实架构：**换掉宿主**（新 `WpsApplication` 对象 + 空
   * `PluginStorage`），**保留本机后端那一份材料**（后端是另一个进程，WPS 重启与它无关），
   * 且材料只能以文本跨边界（读回要重新解析、逐字节比对）。
   */
  it("full process restart loses PluginStorage but resumes only third from durable material", async () => {
    const f = fixture();
    await expect(executeWpsPlan(f.application, f.prepared, undefined, undefined, async (phase, checkpoint) => {
      if (phase === "after" && checkpoint.actionIndex === 1) throw new Error("lost receipt");
    }, operationId, {taskId: "task-1", ...recoveryChannel})).rejects.toThrow("lost receipt");
    // 写前落盘确实发生过：材料在后端那一份里，且正文是按协议校验过的文本。
    expect(backendRecovery().ids()).toContain(operationId);
    f.state.nextActionIndex = 2;
    // WPS 进程退出：宿主缓存没了，应用对象也换了新的。后端材料**不动**。
    f.storage.clear();
    const reopened = {...f.application};
    expect(wpsUndoRecordFor(reopened, operationId)).toBeNull();
    f.writes.length = 0;
    wire(f.state);
    // 核对材料：从后端读回 → 校验 → 水合回宿主缓存。这一步之后既有判据原样工作。
    const material = await ensureWpsRecoveryMaterial(reopened, f.state);
    expect(material.kind).toBe("package");
    expect(canResumeWpsExecution(reopened, f.state, [1, 2])).toBe(true);
    const result = await resumeWpsExecution(reopened, f.state, () => [1, 2]);
    expect(f.writes).toEqual(["Target!A3"]);
    expect(ledger.startWpsAction).toHaveBeenCalledTimes(1);
    expect(vi.mocked(ledger.startWpsAction).mock.calls[0][1].actionIndex).toBe(2);
    expect(result.state).toBe("verified");
    // 执行成功**不等于**任务结束：材料必须还在，否则接下来"确认/放弃"就断了。
    expect(backendRecovery().ids()).toContain(operationId);
    await undoWpsExecution(reopened, operationId);
    expect([1, 2, 3].map(i => f.cells.get(`Target!A${i}`))).toEqual(["old1", "old2", "old3"]);
    // 终态回收是唯一合法的删除时机（面板撤销收尾处调用）。
    retireWpsRecoveryPackage(operationId);
    await Promise.resolve();
    expect(backendRecovery().ids()).toEqual([]);
  });

  it("撤销成功后 record 变 restored，续跑在拿到执行权前就被 state 闸拒绝、零写入", async () => {
    /*
     * 「撤销后不可能再算 resume」的**唯一**判据不是材料回收（`retireWpsRecoveryPackage`），
     * 而是 `prepareWpsResume` 里的 `record.state` fail-closed：撤销把宿主缓存里这条记录
     * 的 state 改成 `restored`，而可续跑集合只有 `ready / applying / partial_write_recoverable`。
     * 材料回收是终态清理，且因 `ensureWpsRecoveryMaterial` 宿主缓存优先（`wpsUndoRecordFor`
     * 还能按 executionId 找到这条 restored 记录），它**不能**单独保证阻断 —— 真正兜底的是
     * 下面这条 state 闸。迟到响应 / 旧回调 / 连点都从同一个 `prepareWpsResume` 入口被挡。
     */
    const f = await interrupted(2);
    wire(f.state);
    expect(canResumeWpsExecution(f.application, f.state, [1, 2])).toBe(true);
    // 撤销：目标回到执行前，宿主缓存里这条记录的 state 被改成 restored。
    await undoWpsExecution(f.application, operationId);
    expect([1, 2, 3].map(i => f.cells.get(`Target!A${i}`))).toEqual(["old1", "old2", "old3"]);
    const record = wpsUndoRecordFor(f.application, operationId)!;
    expect(record.state).toBe("restored");
    // 材料层面：宿主缓存优先会把它读成"可用"（host），所以材料回收不是这里的判据。
    expect((await ensureWpsRecoveryMaterial(f.application, f.state)).kind).toBe("host");
    // 真正阻断：`prepareWpsResume` 因 restored 不在可续跑状态集合抛错。
    const reason = wpsResumeUnavailableReason(f.application, f.state, [1, 2]);
    expect(reason).toContain("本机续跑材料缺失、过期、被替换或仍需核对");
    expect(canResumeWpsExecution(f.application, f.state, [1, 2])).toBe(false);
    // 旧回调 / 迟到响应走到执行器同样被挡在 claim 之前：撤销本身回退写了两个目标格，
    // 基线取撤销之后，续跑不得再新增任何写入。
    const writesAfterUndo = [...f.writes];
    await expect(resumeWpsExecution(f.application, f.state, () => [1, 2])).rejects.toThrow();
    expect(ledger.claimWpsExecutionResume).not.toHaveBeenCalled();
    expect(f.writes).toEqual(writesAfterUndo);
  });

  it("撤销后跨重启：宿主缓存清空、材料删失败仍在，disposition 终态仍挡住续跑", async () => {
    /*
     * 用户第 3 点的跨重启形态：撤销成功后 `record.state="restored"` 只落在**宿主缓存**，
     * `retireWpsRecoveryPackage` 删后端材料是 best-effort。清空缓存重开后：
     * - 宿主缓存没了 → `wpsUndoRecordFor` 读不到 restored 记录；
     * - 后端材料若删失败还在 → `ensureWpsRecoveryMaterial` 会**水合**，按账本游标把
     *   state 推回 `partial_write_recoverable`（`nextActionIndex=2 < 3`）；
     * - `prepareWpsResume` 的可续跑集合含 `partial_write_recoverable` ⇒ 旧代码会放行续跑。
     *
     * 现在靠**账本 disposition 投影**（`undoneAt`）在 `wpsResumeUnavailableReason` 与
     * `resumeWpsExecution` 入口各挡一道，不依赖材料删没删干净、也不依赖宿主缓存里
     * 有没有那条 `restored` 记录。
     */
    const f = await interrupted(2);
    wire(f.state);
    // 撤销成功并留痕（undoneAt 落账）；后端材料删除失败，仍留在盘上。
    await undoWpsExecution(f.application, operationId);
    f.state.disposition = { undoneAt: 1_790_000_700_000, abandonedAt: null };
    // 清空宿主缓存（重开）：record.state="restored" 随之消失，材料却还在后端。
    f.storage.clear();
    const reopened = { ...f.application };
    // 材料确实还在（删失败）：水合会推回 partial_write_recoverable，但 disposition 终态先拦。
    expect((await ensureWpsRecoveryMaterial(reopened, f.state)).kind).toBe("package");
    // 终态判据：canResume 为 false，原因是"已经撤销过"，不是"材料缺失"。
    const reason = wpsResumeUnavailableReason(reopened, f.state, [1, 2]);
    expect(reason).toContain("已经撤销过");
    expect(canResumeWpsExecution(reopened, f.state, [1, 2])).toBe(false);
    // 执行器入口同样被挡在 claim 之前：旧回调 / 迟到响应也续不了。
    const writesBefore = [...f.writes];
    await expect(resumeWpsExecution(reopened, f.state, () => [1, 2])).rejects.toThrow("已经撤销过");
    expect(ledger.claimWpsExecutionResume).not.toHaveBeenCalled();
    expect(f.writes).toEqual(writesBefore);
  });

  it("撤销后跨重启：水合读 disposition 终态，把记录推回 restored 而非 partial_write_recoverable，撤销入口不重新开放", async () => {
    /*
     * 第 3 点的另一半：续跑被 `wpsResumeUnavailableReason` 挡住还不够。撤销入口的判定
     * 走 `latestWpsUndo`（只认 `verified / partial_write_recoverable / undo_conflict`）。
     * 若水合按账本游标把一次已撤销的执行推回 `partial_write_recoverable`，`latestWpsUndo`
     * 会重新把它交出来 —— 备份抽屉与结果卡的「撤销本次执行」按钮就重新出现，让用户
     * 对一张已经回退掉的表再撤一次。所以水合必须读 disposition：任一终态都水合成
     * `restored`（`latestWpsUndo` 过滤它、`prepareWpsResume` 也不认它），两头一起堵。
     */
    const f = await interrupted(2);
    wire(f.state);
    await undoWpsExecution(f.application, operationId);
    f.state.disposition = { undoneAt: 1_790_000_700_000, abandonedAt: null };
    // 清空宿主缓存（重开），后端材料删失败仍在。
    f.storage.clear();
    const reopened = { ...f.application };
    // 水合仍返回 package（材料在），但推回宿主缓存的 record.state 必须是 restored。
    const material = await ensureWpsRecoveryMaterial(reopened, f.state);
    expect(material.kind).toBe("package");
    const hydrated = wpsUndoRecordFor(reopened, operationId)!;
    expect(hydrated.state).toBe("restored");
    // 撤销入口判定：`latestWpsUndo` 过滤 restored ⇒ 返回 null ⇒ 按钮不重新出现。
    expect(latestWpsUndo(reopened)).toBeNull();
    // 续跑同样被挡（record.state=restored 不在可续跑集合）。
    expect(canResumeWpsExecution(reopened, f.state, [1, 2])).toBe(false);
  });

  it("留痕失败 + 材料删失败 + 跨重启：旧材料水合后，撤销过的执行续跑与重复撤销都零写入", async () => {
    /*
     * 用户点名的组合风险：本轮的 disposition 闸**依赖上报成功**。这里刻意让上报失败
     * （`state.disposition` 保持 undefined）且材料删除失败（不调 retire、后端材料仍在），
     * 再清空宿主缓存重开、从旧材料水合 —— 于是 disposition 闸不生效，水合只能按账本游标
     * 把 state 推回 `partial_write_recoverable`（如实：无法证明已撤销）。
     *
     * 此时唯一的防线是**执行入口自身**读宿主现场：撤销把目标恢复成了写前值，而账本登记的是
     * 写后指纹 —— 续跑与撤销都必须在这一层被拦住。判据落在**真实宿主写入计数**上，
     * 不看卡片文案、也不只看 record.state。
     */
    const f = await interrupted(2);
    wire(f.state);
    await undoWpsExecution(f.application, operationId);
    // 撤销确实发生：目标回到写前值 —— 这就是后面"现场与旧执行证据不一致"的根据。
    expect([1, 2, 3].map(i => f.cells.get(`Target!A${i}`))).toEqual(["old1", "old2", "old3"]);
    // 双失败现场：留痕没落库（disposition 仍是 undefined）、材料没删（仍在后端）。
    expect(f.state.disposition).toBeUndefined();
    expect(backendRecovery().ids()).toContain(operationId);
    // 完整退出宿主：宿主缓存清空，`record.state="restored"` 随之消失。
    f.storage.clear();
    const reopened = { ...f.application };
    expect(wpsUndoRecordFor(reopened, operationId)).toBeNull();
    // 从这一刻起，任何宿主写入都算"重新写入"，必须为 0。
    f.writes.length = 0;
    // 从旧材料水合：disposition 读不到 ⇒ 按游标推回 partial_write_recoverable（不是 restored）。
    expect((await ensureWpsRecoveryMaterial(reopened, f.state)).kind).toBe("package");
    const hydrated = wpsUndoRecordFor(reopened, operationId)!;
    expect(hydrated.state).toBe("partial_write_recoverable");
    expect(hydrated.actions[0].appliedRows).toBeGreaterThan(0);
    expect(hydrated.actions[0].afterFingerprint).toBeTruthy();
    /*
     * 卡面层（第 4 点）：证据不足时**保持阻断并说明原因**，不推断撤销成功。
     * 现场（写前值）与账本登记（写后指纹）矛盾 ⇒ 只读核对判 conflict ⇒ 结论 blocked，
     * 不给任何写入/补记按钮，也不出现"可继续"。
     */
    const inspection = inspectWpsRecoveryState(reopened, f.state, [1, 2]);
    expect(inspection).toBe("conflict");
    const verdict = wpsReconcileVerdict({
      state: f.state, identityMatch: true, localInspection: inspection,
      resumeUnavailableReason: null, resumeAvailable: false
    });
    expect(verdict.kind).toBe("blocked");
    /*
     * 如实呈现：留痕失败时水合推回 `partial_write_recoverable`，`latestWpsUndo` **确实**
     * 会把这条记录交出来 —— 也就是"撤销入口重新出现了"。安全性不能建立在"按钮不出现"上，
     * 只能建立在"点下去写不进去"上（下面 ② 用真实写入计数证明）。
     */
    expect(latestWpsUndo(reopened)?.executionId).toBe(operationId);
    /*
     * ① 续跑 —— 并且故意用**撤销前捕获的旧 state** 直接调执行器（旧回调入口，绕过卡面）。
     *    `prepareWpsResume` 读宿主现场：已完成动作是写前值 ≠ 登记的写后指纹 ⇒ 在 claim
     *    之前就抛错。断言顺序刻意"先写计数、再 claim 次数"：万一这层被拆掉，先暴露的是
     *    **越过了执行边界**（拿到了执行权），而不是一句文案不匹配。
     */
    const resumeOutcome = await resumeWpsExecution(reopened, f.state, () => [1, 2])
      .then(() => "resolved", (error) => error instanceof Error ? error.message : String(error));
    expect(f.writes).toEqual([]);
    expect(ledger.claimWpsExecutionResume).not.toHaveBeenCalled();
    expect(resumeOutcome).toContain("已完成或未完成目标已变化");
    /*
     * ② 重复撤销：水合记录的 appliedRows>0、afterFingerprint 是写后，现场是写前 ⇒
     *    `undoWpsExecution` 在**任何写入之前**判冲突停下（只改记录状态，不碰单元格）。
     */
    await expect(undoWpsExecution(reopened, operationId)).rejects.toThrow("在执行后又被修改");
    expect(f.writes).toEqual([]);
    // 原表保持撤销后的现场：既没被"重新写入"改回写后值，也没被二次恢复覆盖。
    expect([1, 2, 3].map(i => f.cells.get(`Target!A${i}`))).toEqual(["old1", "old2", "old3"]);
  });

  it("核对通过到写入之间现场被改：执行入口在写入前按已完成动作的写后现场复核拦住（零写入）", async () => {
    /*
     * 用户第 3 点的"检查至写入之间的状态变化"：`resumeWpsExecution` 在拿到执行权
     * （`await claim`）之后会**再核一次** `prepareWpsResume`，那次读的是当时的宿主现场。
     * 这里在 claim 期间把已完成动作的区域改掉，模拟"核对时还没变、写入前已经变了"。
     * 判据同样落在实际写入计数上：不写、也不进入任何动作的 start/submit。
     */
    const f = await interrupted(2);
    // claim 是真实异步点：在它返回之前篡改已完成动作的区域。
    wire(f.state, () => { f.cells.set("Target!A1", "tampered"); });
    f.writes.length = 0;
    await expect(resumeWpsExecution(f.application, f.state, () => [1, 2]))
      .rejects.toThrow("已完成或未完成目标已变化");
    expect(f.writes).toEqual([]);
    expect(ledger.startWpsAction).not.toHaveBeenCalled();
  });

  it("动作之间现场被改：后续动作写入前按已完成动作的写后现场复核拦住（零写入）", async () => {
    /*
     * `prepareWpsResume` 只在续跑开头核一次。真正在**每个动作写入之前**核"更早已完成动作
     * 的写后现场"的是 `executeWpsPlan` 里的 `assertPriorWritesCurrent`。这里在第一个恢复
     * 动作的写前检查点（真实 await 点）篡改一个更早完成的区域，模拟"检查之后、写入之前"。
     * 判据先看**写计数**：拆掉这层会真的把剩余动作写进原表（不是文案变化）。
     */
    const f = await interrupted(2);
    wire(f.state, undefined, () => { f.cells.set("Target!A1", "tampered-between"); });
    f.writes.length = 0;
    const outcome = await resumeWpsExecution(f.application, f.state, () => [1, 2])
      .then(() => "resolved", (error) => error instanceof Error ? error.message : String(error));
    expect(f.writes).toEqual([]);
    expect(outcome).toContain("等待期间已执行区域发生外部变化");
  });

  it("放弃变更后 disposition.abandonedAt 同样挡住续跑", async () => {
    const f = await interrupted(2);
    wire(f.state);
    f.state.disposition = { undoneAt: null, abandonedAt: 1_790_000_700_000 };
    const reason = wpsResumeUnavailableReason(f.application, f.state, [1, 2]);
    expect(reason).toContain("放弃变更");
    expect(canResumeWpsExecution(f.application, f.state, [1, 2])).toBe(false);
    const writesBefore = [...f.writes];
    await expect(resumeWpsExecution(f.application, f.state, () => [1, 2])).rejects.toThrow("放弃变更");
    expect(ledger.claimWpsExecutionResume).not.toHaveBeenCalled();
    expect(f.writes).toEqual(writesBefore);
  });

  it("a failed write-before-persist stops with zero host writes and keeps the reason", async () => {
    /*
     * 落盘失败必须**零宿主写入**：写了一半却没有材料的现场，比不写更糟 ——
     * 用户会照着界面去手动改原表，把真正的现场彻底搞乱。
     *
     * 这条同时钉住"归一"位置：抛错必须落到执行器自己的 catch 里，
     * 记录变成"未写入（restored）"且带着**原始**原因，不能被后续状态覆盖成一句"已恢复"。
     */
    const f = fixture();
    await expect(executeWpsPlan(f.application, f.prepared, undefined, undefined, undefined, operationId,
      {taskId: "task-1", persistMaterial: async () => { throw new Error("磁盘写不进去"); }})).rejects.toThrow("磁盘写不进去");
    expect(f.writes).toEqual([]);
    expect([1, 2, 3].map(i => f.cells.get(`Target!A${i}`))).toEqual(["old1", "old2", "old3"]);
    const record = wpsUndoRecordFor(f.application, operationId)!;
    expect(record.state).toBe("restored");
    expect(record.error).toContain("磁盘写不进去");
    expect(record.actions.every(action => action.appliedRows === 0)).toBe(true);
    // 材料没存下 ⇒ 也没资格说自己"可续跑"。
    expect(backendRecovery().ids()).toEqual([]);
  });

  it("material written by a newer format version is refused instead of half-read", async () => {
    const f = fixture();
    await expect(executeWpsPlan(f.application, f.prepared, undefined, undefined, async (phase, checkpoint) => {
      if (phase === "after" && checkpoint.actionIndex === 1) throw new Error("lost receipt");
    }, operationId, {taskId: "task-1", ...recoveryChannel})).rejects.toThrow("lost receipt");
    f.state.nextActionIndex = 2;
    f.storage.clear();
    const reopened = {...f.application};
    const stored = backendRecovery().load(operationId)!;
    const material = await ensureWpsRecoveryMaterial(reopened, f.state, {
      // 只把**文本**换成"更高版本"的材料：读回仍走解析，不许按老字段硬猜。
      load: async () => ({...stored, packageJson: JSON.stringify({...JSON.parse(stored.packageJson), schemaVersion: 2})})
    });
    expect(material.kind).toBe("unavailable");
    expect(canResumeWpsExecution(reopened, f.state, [1, 2])).toBe(false);
    f.writes.length = 0;
    await expect(resumeWpsExecution(reopened, f.state, () => [1, 2])).rejects.toThrow();
    expect(f.writes).toEqual([]);
  });

  /*
   * 缺陷回归（2026-10-27 第二轮）：持久包有效期 7 天，`prepareWpsResume` 的综合条件里
   * 却还留着一句无条件的 `age > wpsResumeMaxAgeSeconds*1000`（30 分钟）。于是"材料在盘上、
   * 也还在保留期内"，超过半小时照样不能续跑 —— 用户看到的就是"完整退出后过一会儿就再也
   * 接不上"。这里用**可控时钟**推进到 31 分钟，不许用真实等待。
   */
  it("a durable package older than the 30-minute cache window still resumes only the third action", async () => {
    const f = fixture();
    const startedAt = new Date();
    await expect(executeWpsPlan(f.application, f.prepared, undefined, undefined, async (phase, checkpoint) => {
      if (phase === "after" && checkpoint.actionIndex === 1) throw new Error("lost receipt");
    }, operationId, {taskId: "task-1", ...recoveryChannel})).rejects.toThrow("lost receipt");
    f.state.nextActionIndex = 2;
    f.storage.clear();
    const reopened = {...f.application};
    f.writes.length = 0;
    wire(f.state);
    expect((await ensureWpsRecoveryMaterial(reopened, f.state)).kind).toBe("package");
    vi.useFakeTimers({toFake: ["Date"]});
    try {
      // 31 分钟：早已超过宿主缓存的 30 分钟，但距离 7 天的保留期还很远。
      vi.setSystemTime(new Date(startedAt.getTime() + 31 * 60 * 1000));
      expect(canResumeWpsExecution(reopened, f.state, [1, 2])).toBe(true);
      const result = await resumeWpsExecution(reopened, f.state, () => [1, 2]);
      // 宿主写调用记录是唯一能证明"前两步没被重写"的东西：最终值相等不算。
      expect(f.writes).toEqual(["Target!A3"]);
      expect(vi.mocked(ledger.startWpsAction).mock.calls.map(call => call[1].actionIndex)).toEqual([2]);
      expect(result.state).toBe("verified");
    } finally {
      vi.useRealTimers();
    }
  });

  it("durable expiry boundary equals the server's, and an expired package writes nothing", async () => {
    const f = fixture();
    await expect(executeWpsPlan(f.application, f.prepared, undefined, undefined, async (phase, checkpoint) => {
      if (phase === "after" && checkpoint.actionIndex === 1) throw new Error("lost receipt");
    }, operationId, {taskId: "task-1", ...recoveryChannel})).rejects.toThrow("lost receipt");
    f.state.nextActionIndex = 2;
    f.storage.clear();
    const reopened = {...f.application};
    expect((await ensureWpsRecoveryMaterial(reopened, f.state)).kind).toBe("package");
    const record = wpsUndoRecordFor(reopened, operationId)!;
    const deadline = Date.parse(record.resumeExpiresAt!);
    expect(Number.isFinite(deadline)).toBe(true);
    // 三档：到期前一毫秒仍可续跑；**恰好到期**就算过期（与服务端 `deadline <= now` 同口径）；
    // 到期后当然过期。不许留"差一点点也算有效"的灰区。
    expect(wpsResumeWindow(record, deadline - 1)).toMatchObject({ok: true, persistent: true});
    expect(wpsResumeWindow(record, deadline)).toEqual({ok: false, reason: "本机恢复包已超过保留期。"});
    expect(wpsResumeWindow(record, deadline + 1)).toEqual({ok: false, reason: "本机恢复包已超过保留期。"});
    expect(() => prepareWpsResume(reopened, f.state, [1, 2], deadline - 1)).not.toThrow();
    expect(() => prepareWpsResume(reopened, f.state, [1, 2], deadline)).toThrow("超过保留期");
    wire(f.state);
    f.writes.length = 0;
    vi.useFakeTimers({toFake: ["Date"]});
    try {
      vi.setSystemTime(new Date(deadline));
      await expect(resumeWpsExecution(reopened, f.state, () => [1, 2])).rejects.toThrow("超过保留期");
    } finally {
      vi.useRealTimers();
    }
    expect(f.writes).toEqual([]);
    expect(ledger.startWpsAction).not.toHaveBeenCalled();
  });

  it("a corrupt durable expiry is refused instead of falling back to the cache window", async () => {
    const f = await interrupted();
    const record = wpsUndoRecordFor(f.application, operationId)!;
    // "字段在，但读不出时间"：不许退回 30 分钟去放行，也不许当成永久有效。
    expect(wpsResumeWindow({...record, resumeExpiresAt: "not-a-time"} as typeof record, Date.now()))
      .toEqual({ok: false, reason: "本机恢复包的保留期限字段无效，不能续跑。"});
    f.storage.set("excel-bro.wps.undo.v1", JSON.stringify([{...record, resumeExpiresAt: "not-a-time"}]));
    expect(canResumeWpsExecution(f.application, f.state, [1, 2])).toBe(false);
    await expect(resumeWpsExecution(f.application, f.state, () => [1, 2])).rejects.toThrow("保留期限字段无效");
    expect(f.writes).toEqual([]);
  });

  it("a record without a durable expiry keeps the old cache window, and invalid times never pass", async () => {
    const f = await interrupted();
    // 清掉持久材料 ⇒ 只剩旧宿主缓存，走的就是历史那套 30 分钟规则。
    backendRecovery().clear();
    expect((await ensureWpsRecoveryMaterial(f.application, f.state)).kind).toBe("host");
    const record = wpsUndoRecordFor(f.application, operationId)!;
    expect(record.resumeExpiresAt).toBeUndefined();
    const createdAt = Date.parse(record.createdAt);
    const window = capabilities.taskRecovery.wpsResumeMaxAgeSeconds * 1000;
    expect(wpsResumeWindow(record, createdAt + window - 1)).toMatchObject({ok: true, persistent: false});
    expect(wpsResumeWindow(record, createdAt + window)).toMatchObject({ok: false});
    expect(wpsResumeWindow(record, createdAt + window + 1)).toMatchObject({ok: false});
    // 创建时间无效 / 指向未来：任何期限都不能因此"还没到"。
    expect(wpsResumeWindow({...record, createdAt: "2000-13-40"} as typeof record, createdAt)).toMatchObject({ok: false});
    expect(wpsResumeWindow({...record, createdAt: new Date(createdAt + 86_400_000).toISOString()} as typeof record, createdAt))
      .toEqual({ok: false, reason: "本机恢复记录已过期或创建时间无效。"});
    expect(canResumeWpsExecution(f.application, f.state, [1, 2])).toBe(true);
    expect(() => prepareWpsResume(f.application, f.state, [1, 2], createdAt + window)).toThrow();
  });

  it("durable material that is gone is reported as missing, never faked", async () => {
    const f = fixture();
    await expect(executeWpsPlan(f.application, f.prepared, undefined, undefined, async (phase, checkpoint) => {
      if (phase === "after" && checkpoint.actionIndex === 1) throw new Error("lost receipt");
    }, operationId, {taskId: "task-1", ...recoveryChannel})).rejects.toThrow("lost receipt");
    f.state.nextActionIndex = 2;
    f.storage.clear();
    const reopened = {...f.application};
    backendRecovery().clear(); // 后端那一份也没了（换了机器 / 被清理）
    f.writes.length = 0;
    const material = await ensureWpsRecoveryMaterial(reopened, f.state);
    expect(material.kind).toBe("none");
    expect(canResumeWpsExecution(reopened, f.state, [1, 2])).toBe(false);
    await expect(resumeWpsExecution(reopened, f.state, () => [1, 2])).rejects.toThrow();
    expect(f.writes).toEqual([]);
  });

  it("expired / corrupt / identity-mismatched material stops before any write", async () => {
    const f = fixture();
    await expect(executeWpsPlan(f.application, f.prepared, undefined, undefined, async (phase, checkpoint) => {
      if (phase === "after" && checkpoint.actionIndex === 1) throw new Error("lost receipt");
    }, operationId, {taskId: "task-1", ...recoveryChannel})).rejects.toThrow("lost receipt");
    f.state.nextActionIndex = 2;
    f.storage.clear();
    const reopened = {...f.application};
    for (const damage of ["expired", "corrupt"] as const) {
      const copy = fixture();
      resetBackendRecovery();
      await expect(executeWpsPlan(copy.application, copy.prepared, undefined, undefined, async (phase, checkpoint) => {
        if (phase === "after" && checkpoint.actionIndex === 1) throw new Error("lost receipt");
      }, operationId, {taskId: "task-1", ...recoveryChannel})).rejects.toThrow("lost receipt");
      copy.storage.clear();
      const host = {...copy.application};
      if (damage === "expired") backendRecovery().forceExpire(operationId);
      else backendRecovery().corrupt(operationId, '{"schemaVersion":1}');
      copy.writes.length = 0;
      const material = await ensureWpsRecoveryMaterial(host, copy.state);
      expect(material.kind).toBe("unavailable");
      expect(canResumeWpsExecution(host, copy.state, [1, 2])).toBe(false);
      await expect(resumeWpsExecution(host, copy.state, () => [1, 2])).rejects.toThrow();
      expect(copy.writes).toEqual([]);
    }
    void reopened;
  });

  it("real acceptance fault leaves second in flight, fresh host checks it and resumes only third",async()=>{
    const f=fixture(), m=new Map<string,string>();
    vi.stubGlobal("localStorage",{getItem:(k:string)=>m.get(k)??null,setItem:(k:string,v:string)=>m.set(k,v),removeItem:(k:string)=>m.delete(k)});
    try {
      armRecoveryProbe(true,f.state.workbookIdentity);
      bindRecoveryProbe(true,f.state.workbookIdentity,operationId,f.prepared.plan.actions,true);
      const receipts:number[]=[]; f.state.nextActionIndex=0;
      await expect(executeWpsPlan(f.application,f.prepared,undefined,undefined,async(phase,c)=>{
        if(phase==="before") {f.state.inFlightActionIndex=c.actionIndex;return;}
        interruptRecoveryProbe(true,operationId,c.actionIndex);
        receipts.push(c.actionIndex); f.state.nextActionIndex=c.actionIndex+1;f.state.inFlightActionIndex=null;
      },operationId,{taskId:"task-1",...recoveryChannel})).rejects.toThrow("第二步已写入并回读");
      expect(receipts).toEqual([0]);expect(f.state.nextActionIndex).toBe(1);expect(f.state.inFlightActionIndex).toBe(1);
      expect([1,2,3].map(i=>f.cells.get(`Target!A${i}`))).toEqual([10,20,"old3"]);
      const reopened={...f.application};
      expect(inspectWpsRecoveryState(reopened,f.state,[1,2])).toBe("completed");
      // API transition only after the production native inspection, same as the panel's reconciliation endpoint.
      f.state.nextActionIndex=2;f.state.inFlightActionIndex=null;wire(f.state);
      const before=[...f.writes];await resumeWpsExecution(reopened,f.state,()=>[1,2]);
      expect(f.writes.slice(before.length)).toEqual(["Target!A3"]);
      expect(vi.mocked(ledger.startWpsAction).mock.calls.map(c=>c[1].actionIndex)).toEqual([2]);
      expect(vi.mocked(ledger.completeWpsAction).mock.calls.map(c=>c[1].actionIndex)).toEqual([2]);
    } finally {vi.unstubAllGlobals();}
  });

  it("three value copies use the production lowering, execute once and undo all three",async()=>{
    const f=fixture();
    [111,222,333].forEach((v,i)=>f.cells.set(`Source!A${i+1}`,v));
    const prepared=prepareWpsPlan(f.application,[1,2],{...f.prepared.plan,sourceDependencySheets:undefined,
      actions:[1,2,3].map(r=>({type:"copyRange",sheet:"Target",sourceSheet:"Source",sourceRange:`A${r}`,
        targetRange:`A${r}`,copyType:"values",skipBlanks:false,transpose:false}))});
    await executeWpsPlan(f.application,prepared,undefined,undefined,undefined,operationId,{taskId:"task-1",...recoveryChannel});
    expect(f.writes).toEqual(["Target!A1","Target!A2","Target!A3"]);
    expect([1,2,3].map(r=>f.cells.get(`Target!A${r}`))).toEqual([111,222,333]);
    await undoWpsExecution(f.application,operationId);
    expect([1,2,3].map(r=>f.cells.get(`Target!A${r}`))).toEqual(["old1","old2","old3"]);
    expect([1,2,3].map(r=>f.cells.get(`Source!A${r}`))).toEqual([111,222,333]);
  });

  it("dependent sequential copies remain blocked before any writes",()=>{
    const f=fixture();
    expect(()=>prepareWpsPlan(f.application,[1,2],{...f.prepared.plan,actions:[
      {type:"copyRange",sheet:"Target",sourceSheet:"Source",sourceRange:"A1",targetRange:"A1",copyType:"values",skipBlanks:false,transpose:false},
      {type:"copyRange",sheet:"Target",sourceSheet:"Target",sourceRange:"A1",targetRange:"A2",copyType:"values",skipBlanks:false,transpose:false}
    ]})).toThrow("复制来源必须");
    expect(f.writes).toEqual([]);
  });
  it("independent copy batch with formula sources remains blocked before writes",()=>{
    const f=fixture(); f.cells.set("Source!A1","=1+1");
    expect(()=>prepareWpsPlan(f.application,[1,2],{...f.prepared.plan,actions:[1,2].map(r=>({
      type:"copyRange",sheet:"Target",sourceSheet:"Source",sourceRange:"A1",targetRange:`A${r}`,copyType:"values",skipBlanks:false,transpose:false
    }))})).toThrow("公式依赖");
    expect(f.writes).toEqual([]);
  });
  it.each([1,2])("after %i completed actions, writes only the suffix and undo restores the entire original plan",async next=>{
    const f=await interrupted(next); wire(f.state);
    // Pane recreation: use a fresh application object with the same host and PluginStorage.
    const reopened={...f.application};
    expect(canResumeWpsExecution(reopened,f.state,[2,1])).toBe(true);
    const verdict=wpsReconcileVerdict({state:f.state,identityMatch:true,localInspection:"not_started",resumeAvailable:true});
    expect(wpsReconcileActions(verdict)).toContain("resume");
    const record=await resumeWpsExecution(reopened,f.state,()=>[1,2]);
    expect(f.writes).toEqual(Array.from({length:3-next},(_,i)=>`Target!A${next+i+1}`));
    expect(record.executionId).toBe(operationId); expect(record.state).toBe("verified");expect(record.resume).toBeUndefined();
    expect(ledger.startWpsAction).toHaveBeenCalledTimes(3-next);expect(ledger.finishWpsExecution).toHaveBeenCalledTimes(1);
    expect(reportWpsHostWrites).toHaveBeenCalledTimes(1);
    const report = vi.mocked(reportWpsHostWrites).mock.calls[0];
    expect(report[0]).toBe(operationId);
    expect(report[1].entries.map(entry => entry.actionIndex)).toEqual(Array.from({length: 3-next}, (_, i) => next+i));
    for (const entry of report[1].entries) {
      expect(entry).toMatchObject({outcome: "written", readback: "verified"});
      expect(entry.hostWriteCalls).toBeGreaterThan(0);
    }

    await undoWpsExecution(reopened,operationId);
    expect([1,2,3].map(i=>f.cells.get(`Target!A${i}`))).toEqual(["old1","old2","old3"]);
  });
  it("SQLite sorted object keys preserve original evidence and only resume the third action", async () => {
    const f = await interrupted();
    const sorted = (value: unknown): unknown => Array.isArray(value) ? value.map(sorted) :
      value && typeof value === "object" ? Object.fromEntries(Object.entries(value).sort(([a],[b]) => a.localeCompare(b)).map(([k,v]) => [k, sorted(v)])) : value;
    f.state.descriptor = sorted(f.state.descriptor) as typeof f.state.descriptor;
    wire(f.state);
    expect(canResumeWpsExecution(f.application, f.state, [1,2])).toBe(true);
    await resumeWpsExecution(f.application, f.state, () => [1,2]);
    expect(f.writes).toEqual(["Target!A3"]);
    expect(ledger.startWpsAction).toHaveBeenCalledTimes(1);
  });
  it.each(["source","completed target","remaining target","authorization","expired","task","matrix","old record","inflight"])("refuses %s before claiming or writing",async issue=>{
    const f=await interrupted(); let indexes=[1,2];
    const records=JSON.parse(String(f.storage.get("excel-bro.wps.undo.v1")));
    if(issue==="source")f.cells.set("Source!A1","changed");
    if(issue==="completed target")f.cells.set("Target!A1","changed");
    if(issue==="remaining target")f.cells.set("Target!A3","changed");
    if(issue==="authorization")indexes=[2];
    if(issue==="expired")records[0].createdAt="2000-01-01T00:00:00Z";
    if(issue==="task")f.state.taskId="other";
    if(issue==="matrix")records[0].actions[2].values=[[999]];
    if(issue==="old record")delete records[0].resume;
    if(issue==="inflight")f.state.inFlightActionIndex=2;
    f.storage.set("excel-bro.wps.undo.v1",JSON.stringify(records));
    expect(canResumeWpsExecution(f.application,f.state,indexes)).toBe(false);
    await expect(resumeWpsExecution(f.application,f.state,()=>indexes)).rejects.toThrow();
    expect(ledger.claimWpsExecutionResume).not.toHaveBeenCalled();expect(f.writes).toEqual([]);
  });
  it.each([
    ["claim","source"], ["checkpoint","source"],
    ["claim","authorization"], ["checkpoint","authorization"]
  ])("rechecks after awaited %s when %s changes",async (point, issue)=>{
    const f=await interrupted(); let indexes=[1,2];
    const change=()=>{
      if(issue==="source") f.cells.set("Source!A1","changed during await");
      else indexes=[2];
    };
    wire(f.state,point==="claim"?change:undefined,point==="checkpoint"?change:undefined);
    await expect(resumeWpsExecution(f.application,f.state,()=>indexes)).rejects.toThrow();
    expect(f.writes).toEqual([]);expect(ledger.suspendWpsExecution).toHaveBeenCalledTimes(1);
  });
  it("lost claim / concurrent owner never writes",async()=>{
    const f=await interrupted();vi.mocked(ledger.claimWpsExecutionResume).mockRejectedValueOnce(new Error("409 owner active"));
    await expect(resumeWpsExecution(f.application,f.state,()=>[1,2])).rejects.toThrow("409");expect(f.writes).toEqual([]);
  });
  it("production copyRange expansion supplies complete evidence without invented dependencies",()=>{
    const f=fixture();
    const prepared=prepareWpsPlan(f.application,[1,2],{...f.prepared.plan,sourceDependencySheets:undefined,
      actions:[{type:"copyRange",sourceSheet:"Source",sourceRange:"A1:A1",sheet:"Target",targetRange:"A3",copyType:"values",skipBlanks:false,transpose:false}]});
    const descriptor=ledger.wpsRecoveryDescriptor(prepared);
    expect(descriptor.resumeContext?.sources.map(source=>source.sheetName)).toEqual(["Source"]);
    expect(descriptor.actions.map(action=>action.type)).toEqual(["writeValues"]);
  });
  it("cannot smuggle raw values in resume metadata; old descriptors still accepted",()=>{
    const f=fixture(),d=f.state.descriptor;
    expect(JSON.stringify(d)).not.toContain("private-source");expect(d.resumeContext).toBeTruthy();
    expect(()=>assertWpsRecoveryDescriptor(d)).not.toThrow();
    expect(()=>assertWpsRecoveryDescriptor({actions:d.actions})).not.toThrow();
    expect(()=>assertWpsRecoveryDescriptor({...d,resumeContext:{...d.resumeContext,values:[["private"]]}})).toThrow();
  });
  it("preserves original age and rollback contents on a second interruption",async()=>{
    const f=await interrupted(1),before=wpsUndoRecordFor(f.application,operationId)!;wire(f.state);
    vi.mocked(ledger.completeWpsAction).mockRejectedValueOnce(new Error("lost again"));
    await expect(resumeWpsExecution(f.application,f.state,()=>[1,2])).rejects.toThrow("lost again");
    const after=wpsUndoRecordFor(f.application,operationId)!;
    expect(after.createdAt).toBe(before.createdAt);expect(after.actions.map(a=>a.beforeValues)).toEqual(before.actions.map(a=>a.beforeValues));
    expect(f.writes).toEqual(["Target!A2"]);expect(after.resume).toBeTruthy();
    expect(()=>prepareWpsResume(f.application,{...f.state,nextActionIndex:2},[1,2])).not.toThrow();
  });
});

/*
 * 2026-10-28 真机：卡片只说"工作簿、宿主版本或授权已变化，不能续跑"。
 * 这三个条件要改的地方完全不同（打开原簿 / 收回或补上勾选 / 换回原版本），
 * 合并成一句时用户很容易读成"材料丢了"，去重开材料而不是改正授权。
 *
 * 下面三条各自只破坏**一个**条件，断言文案点名的那一类 —— 谁都不许互相冒充，
 * 也不许把"授权不同"说成"材料缺失"。判据本身没有放宽：三条仍然全部必须成立。
 */
describe("恢复续跑判据：三类不匹配各自点名", () => {
  async function interruptedWithExtraSheet(next = 2) {
    const f = fixture({extraSheet: true});
    await expect(executeWpsPlan(f.application, f.prepared, undefined, undefined, async (phase, checkpoint) => {
      if (phase === "after" && checkpoint.actionIndex === next - 1) throw new Error("lost receipt");
    }, operationId, {taskId: "task-1", ...recoveryChannel})).rejects.toThrow("lost receipt");
    f.state.nextActionIndex = next;
    f.writes.length = 0;
    return f;
  }

  it("授权集合与原登记不同时说授权（几张 vs 几张），不说材料", async () => {
    const f = await interruptedWithExtraSheet();
    const reason = wpsResumeUnavailableReason(f.application, f.state, [1, 2, 3]);
    expect(reason).toContain("当前授权 3 张表，原执行登记 2 张表");
    expect(reason).not.toContain("恢复记录");
    expect(reason).not.toContain("材料");
    // 收回多勾的那张就回到"可以继续"：判据没有因为多一张表而变得更严，也没有放宽。
    expect(wpsResumeUnavailableReason(f.application, f.state, [1, 2])).toBeNull();
  });

  it("宿主版本变化时说版本，不说授权或材料", async () => {
    const f = await interruptedWithExtraSheet();
    const other: typeof f.application = {...f.application, Version: "12.1"};
    const reason = wpsResumeUnavailableReason(other, f.state, [1, 2]);
    expect(reason).toContain("宿主版本已变化");
    expect(reason).toContain("登记 12.0");
    expect(reason).not.toContain("授权");
  });

  it("换了一个工作簿时说身份，不说材料；同名不同路径也算换", async () => {
    const f = await interruptedWithExtraSheet();
    const other: typeof f.application = {...f.application,
      ActiveWorkbook: {...f.application.ActiveWorkbook!, FullName: "C://Data//other.xlsx"}};
    const reason = wpsResumeUnavailableReason(other, f.state, [1, 2]);
    expect(reason).toContain("身份不一致");
    expect(reason).not.toContain("材料");
    expect(reason).not.toContain("授权");
  });
});

/*
 * U1a：公式写入（`writeFormulas`）的精确续跑。
 *
 * 这一组守三件事：
 * 1. 公式能不能进续跑协议（以前被 `type === "writeValues"` 拦在门外，公式计划一中断就只能重来）；
 * 2. **公式文本与计算值必须区分** —— 两条公式可以算出同一个数，凭显示值相同认定"已写入"
 *    会把公式写成常量、也会把别人的改动当成自己写的；
 * 3. 边界不放宽：来源变了、引用了授权外的表、记录里的目标矩阵退化成常量，一律停。
 */
describe("公式写入的精确续跑（U1a）", () => {
  type AnalysisPlanLike = ReturnType<typeof fixture>["prepared"]["plan"];

  /** 三步公式计划：每步写一格，公式引用 Source!A1（跨表依赖）。 */
  function formulaPlan(f: ReturnType<typeof fixture>, formulas = [1, 2, 3].map(i => `=Source!A1+${i}`)): AnalysisPlanLike {
    return {
      ...f.prepared.plan,
      sourceDependencySheets: ["Source"],
      actions: formulas.map((formula, index) => ({
        type: "writeFormulas" as const,
        sheet: "Target",
        range: `A${index + 1}:A${index + 1}`,
        formulas: [[formula]]
      }))
    };
  }

  function prepareFormula(f: ReturnType<typeof fixture>, formulas?: string[]) {
    const prepared = prepareWpsPlan(f.application, [1, 2], formulaPlan(f, formulas),
      [{ sheetName: "Source", range: "A1:A1", values: [[10]] }]);
    f.state.descriptor = ledger.wpsRecoveryDescriptor(prepared);
    return prepared;
  }

  async function interruptAfterFirstFormula(f: ReturnType<typeof fixture>, prepared: ReturnType<typeof prepareFormula>) {
    await expect(executeWpsPlan(f.application, prepared, undefined, undefined, async (phase, checkpoint) => {
      if (phase === "after" && checkpoint.actionIndex === 1) throw new Error("lost formula receipt");
    }, operationId, { taskId: "task-1", ...recoveryChannel })).rejects.toThrow("lost formula receipt");
    f.state.nextActionIndex = 2; // 生产环境由核对证明前两步已落地。
    f.writes.length = 0;
  }

  it("公式计划能登记精确续跑材料，描述符里的动作类型仍是公式", () => {
    const f = fixture();
    f.cells.set("Source!A1", 10);
    const prepared = prepareFormula(f);
    expect(wpsResumeContext(prepared)).toMatchObject({ version: 1 });
    expect(ledger.wpsRecoveryDescriptor(prepared).actions.map(action => action.type))
      .toEqual(["writeFormulas", "writeFormulas", "writeFormulas"]);
  });

  it("中断在第二步之后：续跑只写第三格，已写的公式不重写", async () => {
    const f = fixture();
    f.cells.set("Source!A1", 10);
    const prepared = prepareFormula(f);
    await interruptAfterFirstFormula(f, prepared);
    // 前两格已经是公式现场（不是常量）：写进去的是文本，读回来也是文本。
    expect(f.cells.get("Target!A1")).toEqual({ formula: "=Source!A1+1", value: null });
    expect(f.cells.get("Target!A3")).toBe("old3");

    wire(f.state);
    await resumeWpsExecution(f.application, f.state, () => [1, 2]);

    expect(f.writes).toEqual(["Target!A3"]);
    expect(vi.mocked(ledger.startWpsAction).mock.calls.map(call => call[1].actionIndex)).toEqual([2]);
    expect(f.cells.get("Target!A3")).toEqual({ formula: "=Source!A1+3", value: null });
  });

  it("**计算值一样、公式不一样**的现场不算已写入，也不给完成结论", async () => {
    const f = fixture();
    f.cells.set("Source!A1", 10);
    const prepared = prepareFormula(f);
    await executeWpsPlan(f.application, prepared, undefined, undefined, undefined, operationId,
      { taskId: "task-1", ...recoveryChannel });
    // 第一步的公式是 `=Source!A1+1`（算 11）。这里换成 `=5+6`：**算出来也是 11**。
    f.cells.set("Target!A1", { formula: "=5+6", value: 11 });
    const inspection = inspectWpsUndoRecord(f.application, operationId);
    expect(inspection.kind).not.toBe("completed");
    expect(inspection.kind).toBe("unknown");
  });

  it("记录里的写后矩阵退化成计算值：拒绝续跑（数对也不行）", async () => {
    const f = fixture();
    f.cells.set("Source!A1", 10);
    const prepared = prepareFormula(f);
    await interruptAfterFirstFormula(f, prepared);
    /*
     * 模拟"整条材料被投影成值矩阵"：本机记录与账本描述**一起**换成一格常量 11
     * （正是 `=Source!A1+1` 的计算值）。指纹自洽，所以前面的形状/指纹闸都会放行 ——
     * 此时唯一能挡住它的就是"目标矩阵必须是公式文本"这条。
     */
    const record = wpsUndoRecordFor(f.application, operationId)!;
    record.actions[2].values = [[11]] as never;
    adoptWpsUndoRecord(f.application, record);
    f.state.descriptor.actions[2].expectedFingerprint = matrixFingerprintV2([[11]], [[11]]);
    expect(() => prepareWpsResume(f.application, f.state, [1, 2])).toThrow("不是公式文本");
  });

  it("公式读了没有读取记录的表：**登记阶段**就不给精确续跑资格", () => {
    const f = fixture({ extraSheet: true });
    f.cells.set("Source!A1", 10);
    f.cells.set("Extra!A1", 1);
    const prepared = prepareFormula(f, ["=Source!A1+1", "=Source!A1+2", "=Extra!A1+3"]);
    /*
     * `Extra` 既不在可信记账（`sourceDependencySheets`，只有 Source）里，也没有
     * 读过的区域 —— 于是没有"它当时是什么值"的基线。按产品规则（公式恢复要求登记的
     * 来源证据仍然一致），这种情况**不能**声称能精确续跑：连材料都不登记，
     * 而不是等到续跑时才拦。
     */
    expect(wpsResumeContext(prepared)).toBeNull();
  });

  /*
   * 覆盖缺口（2026-10-30 本轮）：证据以前只按**表名**放行 —— 读取记录覆盖
   * `Source!A1:A3`，公式却引用 `Source!B10`，登记照样通过，改掉 B10 也拦不住。
   * 下面三条把"证据必须真的盖住公式读的那块地方"钉住：
   *  - 覆盖不到 ⇒ 改 B10 必须拒绝且零写入；
   *  - 覆盖得到 ⇒ 来源没变时只跑剩下的动作；
   *  - 有效区域变大（原来为空的一格被填上）同样算来源变化。
   */
  it("公式引用了读取记录没盖住的那一格：改动它之后拒绝续跑，一格都不写", async () => {
    const f = fixture();
    [1, 2, 3].forEach(row => f.cells.set(`Source!A${row}`, row));
    f.cells.set("Source!B10", 5);
    // 读取记录只覆盖 A1:A3（模型当时只读了这三格），公式读的却是 B10。
    const prepared = prepareWpsPlan(f.application, [1, 2], formulaPlan(f, [
      "=Source!B10+1", "=Source!B10+2", "=Source!B10+3"
    ]), [{ sheetName: "Source", range: "A1:A3", values: [[1], [2], [3]] }]);
    f.state.descriptor = ledger.wpsRecoveryDescriptor(prepared);
    await interruptAfterFirstFormula(f, prepared);
    f.cells.set("Source!B10", 99);

    // 拒绝点必须在拿到执行权之前：`prepareWpsResume` 就是这个入口。
    await expect(resumeWpsExecution(f.application, f.state, () => [1, 2]))
      .rejects.toThrow("来源已变化");
    expect(f.writes).toEqual([]);
    expect(f.cells.get("Target!A3")).toBe("old3");
  });

  it("依赖表整块有效区域就是证据：来源没变时只执行剩下的那一步", async () => {
    const f = fixture();
    [1, 2, 3].forEach(row => f.cells.set(`Source!A${row}`, row));
    f.cells.set("Source!B10", 5);
    const prepared = prepareWpsPlan(f.application, [1, 2], formulaPlan(f, [
      "=Source!B10+1", "=Source!B10+2", "=Source!B10+3"
    ]), [{ sheetName: "Source", range: "A1:A3", values: [[1], [2], [3]] }]);
    f.state.descriptor = ledger.wpsRecoveryDescriptor(prepared);
    // 登记的证据里必须包含一块盖住 B10 的区域（= 依赖表当时的整张有效区域）。
    expect(wpsResumeContext(prepared)!.sources.some(source =>
      source.sheetName === "Source" && source.range === "A1:B10")).toBe(true);
    await interruptAfterFirstFormula(f, prepared);

    wire(f.state);
    await resumeWpsExecution(f.application, f.state, () => [1, 2]);

    expect(f.writes).toEqual(["Target!A3"]);
    expect(f.cells.get("Target!A3")).toEqual({ formula: "=Source!B10+3", value: null });
  });

  it("公式引用的那一格被清空（有效区域缩回登记范围）：仍然拒绝续跑", async () => {
    const f = fixture();
    [1, 2, 3].forEach(row => f.cells.set(`Source!A${row}`, row));
    f.cells.set("Source!B10", 5);
    const prepared = prepareWpsPlan(f.application, [1, 2], formulaPlan(f, [
      "=Source!B10+1", "=Source!B10+2", "=Source!B10+3"
    ]), [{ sheetName: "Source", range: "A1:A3", values: [[1], [2], [3]] }]);
    f.state.descriptor = ledger.wpsRecoveryDescriptor(prepared);
    await interruptAfterFirstFormula(f, prepared);
    /*
     * 清空 B10：有效区域缩回 `A1:A3`，恰好落在登记的那块矩形里 ——
     * 只看"有效区域还在不在证据范围内"会放行，**只有**登记时那一块
     * `A1:B10` 的值指纹能发现它变了（5 → 空）。
     */
    f.cells.delete("Source!B10");

    await expect(resumeWpsExecution(f.application, f.state, () => [1, 2]))
      .rejects.toThrow("来源已变化");
    expect(f.writes).toEqual([]);
  });

  it("预览后依赖表变了：公式照常写下去（恢复基线不拦正常执行）", async () => {
    const f = fixture();
    [1, 2, 3].forEach(row => f.cells.set(`Source!A${row}`, row));
    f.cells.set("Source!B10", 5);
    const prepared = prepareWpsPlan(f.application, [1, 2], formulaPlan(f, [
      "=Source!B10+1", "=Source!B10+2", "=Source!B10+3"
    ]), [{ sheetName: "Source", range: "A1:A3", values: [[1], [2], [3]] }]);
    /*
     * 预览之后、执行之前改掉 B10：它在恢复基线（`Source!A1:B10`）里，但不在模型
     * 读过的 `A1:A3` 里。公式写下去的是**文本**，宿主按那一刻的源值自己算 ——
     * 产品规则针对恢复，不改变正常公式的自动重算，所以这里必须照常写。
     */
    f.cells.set("Source!B10", 99);

    await executeWpsPlan(f.application, prepared, undefined, undefined, undefined, operationId,
      { taskId: "task-1", ...recoveryChannel });

    expect(f.cells.get("Target!A1")).toEqual({ formula: "=Source!B10+1", value: null });
  });

  it("依赖表原来为空的一格被填上（有效区域变大）：同样按来源已变化拒绝", async () => {
    const f = fixture();
    [2, 3].forEach(row => f.cells.set(`Source!A${row}`, row));
    f.cells.set("Source!A1", 10); // 与 `prepareFormula` 登记的来源快照一致
    // 公式只读 A1；登记证据覆盖 A1:A3，A1 的值一个字都没改。
    const prepared = prepareFormula(f);
    await interruptAfterFirstFormula(f, prepared);
    // 变化发生在被登记区域**之外**：有效区域从 A1:A3 长到 A1:B10。
    f.cells.set("Source!B10", 5);

    await expect(resumeWpsExecution(f.application, f.state, () => [1, 2]))
      .rejects.toThrow("来源已变化");
    expect(f.writes).toEqual([]);
  });

  it("依赖表有效区域大到无法登记证据：不登记精确续跑，普通公式执行照常", async () => {
    const f = fixture();
    f.cells.set("Source!A1", 10);
    // `Source!ZZ500` 让包围盒变成 A1:ZZ500 = 351,000 格，超过 wpsResumeMaxSourceCells。
    f.cells.set("Source!ZZ500", 1);
    const prepared = prepareFormula(f);
    expect(wpsResumeContext(prepared)).toBeNull();
    // 拒绝的只是"精确恢复"，不是写入本身：公式照常写进去、宿主自己算。
    await executeWpsPlan(f.application, prepared, undefined, undefined, undefined, operationId,
      { taskId: "task-1", ...recoveryChannel });
    expect(f.cells.get("Target!A1")).toEqual({ formula: "=Source!A1+1", value: null });
  });

  it("动态引用的公式：不登记精确续跑材料（解析不出读哪张表，就没有证据口径）", () => {
    const f = fixture({ extraSheet: true });
    f.cells.set("Source!A1", 10);
    const prepared = prepareFormula(f, [
      "=Source!A1+1",
      "=Source!A1+2",
      '=INDIRECT("Source!A1")+3'
    ]);
    // 同一份计划只差这一条公式：`INDIRECT` 解析不出依赖表 ⇒ 没有可比对的证据。
    expect(wpsResumeContext(prepared)).toBeNull();
  });

  it("公式引用了授权之外的表：续跑时拒绝，不放大写入边界", async () => {
    const f = fixture({ extraSheet: true });
    f.cells.set("Source!A1", 10);
    f.cells.set("Extra!A1", 1);
    // 登记时只读 Source（可续跑）。续跑前把材料里第三条公式换成引用 Extra ——
    // 那是本次授权之外的依赖，必须拒绝，而不是照着它写出去。
    const prepared = prepareFormula(f);
    await interruptAfterFirstFormula(f, prepared);
    const record = wpsUndoRecordFor(f.application, operationId)!;
    record.actions[2].values = [["=Extra!A1+3"]];
    adoptWpsUndoRecord(f.application, record);
    // 材料自洽（写后矩阵与登记指纹一致），所以形状/指纹闸都会放行 ——
    // 此时唯一还站着的就是"公式依赖必须仍在授权内"这一条。
    f.state.descriptor.actions[2].expectedFingerprint = matrixFingerprintV2(
      [["=Extra!A1+3"]],
      [["=Extra!A1+3"]]
    );
    expect(() => prepareWpsResume(f.application, f.state, [1, 2])).toThrow("授权之外");
  });

  it("来源表变了：公式计划同样拒绝续跑（来源校验没有为公式放宽）", async () => {
    const f = fixture();
    f.cells.set("Source!A1", 10);
    const prepared = prepareFormula(f);
    await interruptAfterFirstFormula(f, prepared);
    f.cells.set("Source!A1", 99);
    expect(() => prepareWpsResume(f.application, f.state, [1, 2])).toThrow("来源已变化");
  });

  it("核对与续跑资格判断只读：一个单元格都不写", async () => {
    const f = fixture();
    f.cells.set("Source!A1", 10);
    const prepared = prepareFormula(f);
    await interruptAfterFirstFormula(f, prepared);
    inspectWpsUndoRecord(f.application, operationId);
    wpsResumeUnavailableReason(f.application, f.state, [1, 2]);
    expect(f.writes).toEqual([]);
  });

  it("放弃执行把原公式与计算值一起恢复（不是拿值覆盖回去）", async () => {
    // `=1+1` 在替身里登记的"宿主重算结果"是 2：撤销写回公式后，宿主会重算成 2。
    const f = fixture({ formulaValues: { "=1+1": 2 } });
    f.cells.set("Source!A1", 10);
    f.cells.set("Target!A1", { formula: "=1+1", value: 2 });
    const prepared = prepareFormula(f);
    await executeWpsPlan(f.application, prepared, undefined, undefined, undefined, operationId,
      { taskId: "task-1", ...recoveryChannel });
    expect(f.cells.get("Target!A1")).toEqual({ formula: "=Source!A1+1", value: null });
    await undoWpsExecution(f.application, operationId);
    expect(f.cells.get("Target!A1")).toEqual({ formula: "=1+1", value: 2 });
    expect(f.cells.get("Target!A2")).toBe("old2");
  });
});

/*
 * 覆盖证据的**持久化往返**（2026-10-30 第二十三轮核实，不加功能）。
 *
 * 上一组验的是"判据对不对"，这一组只验**判据的输入是否真的跨进程活着**：
 * 登记时补采的矩形与两个指纹 → 白名单投影 → 恢复包（落盘）→ 换宿主、清空 PluginStorage →
 * 只从材料水合 → 续跑。`formulaSourceCoverage` / `coverageOnly` 都**不持久化**，
 * 所以这些用例断言的是"材料里那六个字段 + 计划里的公式文本"够不够重建判据。
 *
 * 三条负向用例的判别点是同一个：**任何"现读一次当基线"的补采都会让它们变绿**，
 * 而它们必须红（拒绝 + 零写入）。
 */
describe("公式覆盖证据的持久化往返（重启后只用恢复包）", () => {
  const COVERED_FORMULAS = ["=Source!B10+1", "=Source!B10+2", "=Source!B10+3"];

  function coveredPlan(f: ReturnType<typeof fixture>) {
    return {
      ...f.prepared.plan,
      sourceDependencySheets: ["Source"],
      actions: COVERED_FORMULAS.map((formula, index) => ({
        type: "writeFormulas" as const,
        sheet: "Target",
        range: `A${index + 1}:A${index + 1}`,
        formulas: [[formula]]
      }))
    };
  }

  /** 登记并执行到第二个动作之后中断：恢复包在这一刻已经落盘。 */
  async function registered(f: ReturnType<typeof fixture>) {
    [1, 2, 3].forEach(row => f.cells.set(`Source!A${row}`, row));
    f.cells.set("Source!B10", 5);
    const prepared = prepareWpsPlan(f.application, [1, 2], coveredPlan(f),
      [{ sheetName: "Source", range: "A1:A3", values: [[1], [2], [3]] }]);
    f.state.descriptor = ledger.wpsRecoveryDescriptor(prepared);
    await expect(executeWpsPlan(f.application, prepared, undefined, undefined, async (phase, checkpoint) => {
      if (phase === "after" && checkpoint.actionIndex === 1) throw new Error("lost coverage receipt");
    }, operationId, { taskId: "task-1", ...recoveryChannel })).rejects.toThrow("lost coverage receipt");
    f.state.nextActionIndex = 2;
    f.writes.length = 0;
    expect(backendRecovery().ids()).toContain(operationId);
    return prepared;
  }

  /** 丢掉宿主缓存与账号内存里的一切（含原 `prepared`），只从恢复包水合。 */
  async function reopen(f: ReturnType<typeof fixture>) {
    f.storage.clear();
    const reopened = { ...f.application };
    expect(wpsUndoRecordFor(reopened, operationId)).toBeNull();
    const material = await ensureWpsRecoveryMaterial(reopened, f.state);
    return { reopened, material };
  }

  it("重启后只用恢复包：覆盖矩形从材料重建，来源没变时只写剩余动作", async () => {
    const f = fixture();
    const prepared = await registered(f);
    // 内存里的两块元数据都不持久化：覆盖记账是映射表，`coverageOnly` 只在 `prepared` 上。
    expect(prepared.formulaSourceCoverage).toEqual({ Source: "A1:B10" });
    expect(prepared.sources.find(source => source.range === "A1:B10")!.coverageOnly).toBe(true);

    const { reopened, material } = await reopen(f);
    expect(material.kind).toBe("package");
    if (material.kind !== "package")
      throw new Error(material.kind === "unavailable" ? material.detail : `材料状态 ${material.kind}`);
    /*
     * 材料里能重建判据的**只有** `sourceEvidence.sources`：矩形 + 行列数 + 两个指纹。
     * 六个字段就是协议的全部（`hasOnlyKeys` 会拒掉多余字段），`coverageOnly` 与
     * `formulaSourceCoverage` 都不在其中 —— 覆盖判据靠"该表有哪块矩形 + 计划里的公式文本"重建。
     */
    const sources = material.record.resume!.context.sources;
    expect(sources.map(source => `${source.sheetName}!${source.range}`)).toEqual(["Source!A1:A3", "Source!A1:B10"]);
    expect(Object.keys(sources[1]).sort()).toEqual(
      ["columnCount", "previewFingerprint", "range", "rowCount", "sheetName", "valueFingerprint"]);

    wire(f.state);
    await resumeWpsExecution(reopened, f.state, () => [1, 2]);

    expect(f.writes).toEqual(["Target!A3"]);
    expect(f.cells.get("Target!A3")).toEqual({ formula: "=Source!B10+3", value: null });
  });

  it("重启后 B10 改值：拒绝且零写入（判据来自材料里的指纹，不是现读）", async () => {
    const f = fixture();
    await registered(f);
    const { reopened } = await reopen(f);
    f.cells.set("Source!B10", 99);

    await expect(resumeWpsExecution(reopened, f.state, () => [1, 2])).rejects.toThrow("来源已变化");
    expect(f.writes).toEqual([]);
    expect(f.cells.get("Target!A3")).toBe("old3");
  });

  it("重启后 B10 被清空（有效区域缩回）：拒绝且零写入", async () => {
    const f = fixture();
    await registered(f);
    const { reopened } = await reopen(f);
    // 有效区域缩回 `A1:A3`，仍被登记矩形包含 ⇒ 区域核对放行，只有**持久化的整块指纹**能发现它变了。
    f.cells.delete("Source!B10");

    await expect(resumeWpsExecution(reopened, f.state, () => [1, 2])).rejects.toThrow("来源已变化");
    expect(f.writes).toEqual([]);
  });

  it("重启后在登记区域之外新增值：拒绝且零写入", async () => {
    const f = fixture();
    await registered(f);
    const { reopened } = await reopen(f);
    // 新增的 C1 让有效区域从 A1:B10 长到 A1:C10 —— 登记矩形盖不住了。
    f.cells.set("Source!C1", 7);

    await expect(resumeWpsExecution(reopened, f.state, () => [1, 2])).rejects.toThrow("来源已变化");
    expect(f.writes).toEqual([]);
  });

  it("旧材料只有 A1:A3 证据、现场已有 B10：拒绝，不拿现读补成旧基线", async () => {
    const f = fixture();
    await registered(f);
    const current = wpsUndoRecordFor(f.application, operationId)!;
    // 模拟"本轮之前登记的"材料：来源证据只有模型读过的那块 A1:A3。
    const legacyContext = {
      ...current.resume!.context,
      sources: current.resume!.context.sources.filter(source => source.range === "A1:A3")
    };
    const legacyPackage = buildWpsRecoveryPackage({ ...current, resume: { ...current.resume!, context: legacyContext } });
    if (!legacyPackage) throw new Error("旧材料构建失败");
    // 同一 operationId 只允许一份材料：先回收再按旧口径落盘（正文仍走协议校验）。
    backendRecovery().remove(operationId);
    backendRecovery().save(operationId, JSON.stringify(legacyPackage));
    // 服务端那一份也退回旧口径，否则会先撞"本机与服务端不一致"，撞不到本轮要验的覆盖判据。
    f.state.descriptor = { ...f.state.descriptor, resumeContext: legacyContext };

    const { reopened, material } = await reopen(f);
    expect(material.kind).toBe("package");
    /*
     * 现场是 A1:B10（B10 有值），材料只盖到 A1:A3。唯一能让它通过的写法是
     * "现读一次 B10 当作当时的基线" —— 正是被禁止的。所以必须拒绝，且理由要说
     * 到"有效区域超出登记证据"，而不是笼统的"材料不可用"。
     */
    await expect(resumeWpsExecution(reopened, f.state, () => [1, 2]))
      .rejects.toThrow(/有效区域已超出登记的来源证据/);
    expect(f.writes).toEqual([]);
    expect(f.cells.get("Target!A3")).toBe("old3");
  });
});

/*
 * copyRange 的能力证据（2026-10-30 先确认再决定要不要写实现）。
 *
 * 结论先行：`copyRange` 在准备阶段被展开成 `writeValues` + 来源快照，形状本来就满足
 * 续跑门槛，**这条用例只是把"能续"钉住**。谁以后把展开改成别的形状，这里会先红。
 */
describe("copyRange 的精确续跑（能力证据）", () => {
  it("三步复制中断在第二步之后：续跑只写第三格", async () => {
    const f = fixture();
    [111, 222, 333].forEach((value, index) => f.cells.set(`Source!A${index + 1}`, value));
    const prepared = prepareWpsPlan(f.application, [1, 2], {
      ...f.prepared.plan,
      sourceDependencySheets: ["Source"],
      actions: [1, 2, 3].map(row => ({
        type: "copyRange" as const,
        sheet: "Target",
        sourceSheet: "Source",
        sourceRange: `A${row}`,
        targetRange: `A${row}`,
        copyType: "values" as const,
        skipBlanks: false,
        transpose: false
      }))
    });
    // 展开之后是值动作 —— 这正是它能进续跑协议的原因。
    expect(ledger.wpsRecoveryDescriptor(prepared).actions.map(action => action.type))
      .toEqual(["writeValues", "writeValues", "writeValues"]);
    expect(wpsResumeContext(prepared)).not.toBeNull();
    f.state.descriptor = ledger.wpsRecoveryDescriptor(prepared);

    await expect(executeWpsPlan(f.application, prepared, undefined, undefined, async (phase, checkpoint) => {
      if (phase === "after" && checkpoint.actionIndex === 1) throw new Error("lost copy receipt");
    }, operationId, { taskId: "task-1", ...recoveryChannel })).rejects.toThrow("lost copy receipt");
    f.state.nextActionIndex = 2;
    f.writes.length = 0;

    wire(f.state);
    await resumeWpsExecution(f.application, f.state, () => [1, 2]);

    expect(f.writes).toEqual(["Target!A3"]);
    expect([1, 2, 3].map(row => f.cells.get(`Target!A${row}`))).toEqual([111, 222, 333]);
  });
});
