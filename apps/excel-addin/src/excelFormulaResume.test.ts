import { beforeEach, describe, expect, it } from "vitest";
import {
  PlanExecutionError,
  captureWorkbookSourceFingerprint,
  currentOfficeWorkbookIdentity,
  executePlan,
  officeRecoverySupport,
  reconcileOfficeRangeAction,
  undoExecution
} from "./excel";
import type { AnalysisPlan, ExcelAction, PlanExecutionResult } from "./contracts";
import {
  createOfficeTestHost,
  installOfficeTestHost,
  type TestOfficeHost
} from "./officeTestHost";
import { FORMULA_SOURCE_CHANGED_MESSAGE } from "@excel-bro/core/formulaWriteback";

/**
 * U1a：`writeFormulas` 的精确恢复（Excel 端）。
 *
 * 判据只有一条，但必须**分三种现场**说清楚：
 *  - 写前（`not_applied`）：现场还是写前那条公式 ⇒ 可以写；
 *  - 写后（`applied`）：现场就是目标公式文本 ⇒ 补记回执，**不重写**；
 *  - 其它（`unknown`）：只写了一部分、或被换成另一条公式 ⇒ **停下来核对**。
 *
 * 「不能仅凭计算值相同认定已写过」是本文件的核心否定判据：`=数据!A1*2` 与 `=2*2`
 * 在真实 Excel 里都算出 4，但它们是两个不同的写入目标。判据比的是**公式文本**
 * （`Excel.Range.formulas`，不是 `values`），所以两条公式换成任意一条都能区分。
 *
 * 证据边界：**替身宿主证据，不是真实 Excel 实机证据**。替身不计算公式（公式格读回来
 * 就是公式文本，`values` 也是文本），因此"算出同一个数"这件事只在注释里表述，
 * 真正钉住的是"按文本比、不按结果比"这条判据本身。
 */

const SOURCE = "数据";
const TARGET = "结果";
const OPERATION = "excel-exec-formula-0001";

function plan(actions: ExcelAction[]): AnalysisPlan {
  return {
    id: "plan-formula-resume",
    title: "公式",
    summary: "公式",
    assumptions: [],
    warnings: [],
    actions,
    sourceDependencySheets: [SOURCE]
  };
}

/** 三个公式动作分别写 A1/A2/A3，目标文本互不相同，便于按格断言"写了哪几格"。 */
function formulaPlan(): AnalysisPlan {
  return plan([
    { type: "writeFormulas", sheet: TARGET, range: "A1", formulas: [[`=${SOURCE}!A1*2`]] },
    { type: "writeFormulas", sheet: TARGET, range: "A2", formulas: [[`=${SOURCE}!A1*3`]] },
    { type: "writeFormulas", sheet: TARGET, range: "A3", formulas: [[`=${SOURCE}!A1*4`]] }
  ]);
}

let host: TestOfficeHost;

beforeEach(() => {
  host = createOfficeTestHost({
    url: "https://contoso.sharepoint.com/formula.xlsx",
    activeSheet: SOURCE,
    selectedRange: "A1",
    sheets: [
      { name: SOURCE, values: [[2]] },
      { name: TARGET, values: [["占位"]] }
    ]
  });
  installOfficeTestHost(host);
});

function writtenAddresses(): string[] {
  return host.writes().map((write) => `${write.sheet}!${write.address}`);
}

describe("计算公式动作的精确续跑（U1a）", () => {
  it("中断在第二步回执：续跑只写第三步，前两步一格都不重写", async () => {
    // 登记时钉住的来源证据（= 数据 的当前值）。续跑会重新读同一张表比对，
    // 本用例不改源值，所以基线一直成立。
    const registered: AnalysisPlan = {
      ...formulaPlan(),
      sourceDependencyFingerprint: await captureWorkbookSourceFingerprint([SOURCE])
    };
    let failure: PlanExecutionError | null = null;
    try {
      await executePlan(registered, {
        authorizedSheetNames: [SOURCE, TARGET],
        operationId: OPERATION,
        recovery: {
          startActionIndex: 0,
          onBeforeAction: async () => undefined,
          onActionCompleted: async ({ actionIndex }) => {
            if (actionIndex === 1) throw new Error("模拟提交第二步回执时断电");
          }
        }
      });
    } catch (error) {
      failure = error as PlanExecutionError;
    }
    expect(failure).toBeInstanceOf(PlanExecutionError);
    const checkpoint = failure!.undoSnapshot!;
    // 现场事实：第二条已经写进宿主，但回执没落地 —— 这正是"在途不明"的来源。
    expect(checkpoint.uncertainActionIndexes).toEqual([1]);
    expect(host.cell(TARGET, "A1")).toBe(`=${SOURCE}!A1*2`);
    expect(host.cell(TARGET, "A2")).toBe(`=${SOURCE}!A1*3`);
    expect(host.cell(TARGET, "A3")).toBe("");

    host.resetWrites();
    const result = await executePlan(registered, {
      authorizedSheetNames: [SOURCE, TARGET],
      operationId: OPERATION,
      recovery: {
        startActionIndex: 2,
        undoSnapshot: checkpoint,
        onBeforeAction: async () => undefined,
        onActionCompleted: async () => undefined
      }
    });

    // 只碰第三步：写前快照读取不算写，读回核对也不算写。
    expect(writtenAddresses()).toEqual(["结果!A3"]);
    expect(host.cell(TARGET, "A3")).toBe(`=${SOURCE}!A1*4`);
    // 三步的公式文本都在现场，独立回读逐条通过。
    expect(result.verification.checks.length).toBeGreaterThanOrEqual(3);
    expect(result.verification.checks.every((check) => check.passed)).toBe(true);
    /*
     * 这里 status 是 `executed_unverified`，**只因为本测试手写的 `formulaPlan()`
     * 没有带 `acceptanceCriteria`**。`verificationGaps` 只认计划声明的验收条件
     * （`plan.acceptanceCriteria` + `dynamicCriteria`），没有它就判不出"已核实"。
     *
     * 真实产品路径要拿到 `verified` 得**三条同时成立**（2026-10-30 收紧表述，
     * 原来写"真实路径不会走到这里"是不带条件的断言，太满）：
     * ①服务端为每个 `writeFormulas` 自动补 `formulasEqual`
     *   （`models.py` 的 `add_deterministic_acceptance_criteria`，见
     *   `server/tests/test_models.py::test_write_formulas_receives_formulas_equal_acceptance`）；
     * ②恢复**读路径**把它带回给客户端 —— 续跑读的是
     *   `GET /api/office/executions/{id}` 返回的 `state.plan`，不是本端副本
     *   （`server/tests/test_office_resume_acceptance.py` 钉住这一环）；
     * ③当前现场回读全部通过。
     *
     * 任一条不成立（判据缺失 / 读响应没带 / 回读不符 / 证据不足）就还是
     * `executed_unverified`，并被 `ResultReview.adoptCompleted`（`verified !== true` 即抛）
     * 挡在审阅之前。所以别把"计划里有 acceptanceCriteria"当成"验证成功"。
     *
     * 本测试钉的是"没带验收条件时如实记成未核实、不冒充已核实"这条兜底 ——
     * `executed_unverified` 不是"恢复的默认结果"，而是"判据缺失时的诚实降级"。
     * 断言它不是为了庆祝"未核实"，而是防止将来有人把缺判据悄悄改成 `verified`。
     */
    expect(result.verification.status).toBe("executed_unverified");
  });

  it("核对不重写：现场已是目标公式 ⇒ applied，且一个单元格都没写", async () => {
    const target = formulaPlan();
    host.setCell(TARGET, "A1", `=${SOURCE}!A1`, { asFormula: true });
    let failure: PlanExecutionError | null = null;
    try {
      await executePlan(target, {
        authorizedSheetNames: [SOURCE, TARGET],
        operationId: OPERATION,
        recovery: {
          startActionIndex: 0,
          onBeforeAction: async () => undefined,
          onActionCompleted: async ({ actionIndex }) => {
            if (actionIndex === 0) throw new Error("模拟提交回执时断电");
          }
        }
      });
    } catch (error) {
      failure = error as PlanExecutionError;
    }
    const before = failure!.undoSnapshot!.ranges[0];
    expect(before.formulas).toEqual([[`=${SOURCE}!A1`]]);
    host.resetWrites();

    const outcome = await reconcileOfficeRangeAction({
      plan: target,
      operationId: OPERATION,
      workbookIdentity: currentOfficeWorkbookIdentity()!,
      actionIndex: 0,
      before
    });
    expect(outcome.outcome).toBe("applied");
    // 补记回执只追加恢复点，不回到宿主再写一次。
    expect(outcome.checkpoint?.ranges.map((range) => range.actionIndex)).toEqual([0]);
    expect(host.writes()).toEqual([]);
  });

  it("算出同一个数的另一条公式不算已写入：文本不同 ⇒ unknown，而不是 applied", async () => {
    const target = formulaPlan();
    host.setCell(TARGET, "A1", `=${SOURCE}!A1`, { asFormula: true });
    let failure: PlanExecutionError | null = null;
    try {
      await executePlan(target, {
        authorizedSheetNames: [SOURCE, TARGET],
        operationId: OPERATION,
        recovery: {
          startActionIndex: 0,
          onBeforeAction: async () => undefined,
          onActionCompleted: async ({ actionIndex }) => {
            if (actionIndex === 0) throw new Error("模拟提交回执时断电");
          }
        }
      });
    } catch (error) {
      failure = error as PlanExecutionError;
    }
    const before = failure!.undoSnapshot!.ranges[0];
    host.resetWrites();

    // `=2*2` 与目标 `=数据!A1*2`（数据!A1 = 2）在真实 Excel 里都算出 4。
    // 只看结果会误判成"已写入"；判据比的是公式文本，所以这里必须是 unknown。
    host.setCell(TARGET, "A1", "=2*2", { asFormula: true });
    const unknown = await reconcileOfficeRangeAction({
      plan: target,
      operationId: OPERATION,
      workbookIdentity: currentOfficeWorkbookIdentity()!,
      actionIndex: 0,
      before
    });
    expect(unknown.outcome).toBe("unknown");
    expect(unknown.checkpoint).toBeUndefined();

    // 回到写前那条公式 ⇒ 明确"没写过"，可以让续跑重新写这一步。
    host.setCell(TARGET, "A1", `=${SOURCE}!A1`, { asFormula: true });
    const notApplied = await reconcileOfficeRangeAction({
      plan: target,
      operationId: OPERATION,
      workbookIdentity: currentOfficeWorkbookIdentity()!,
      actionIndex: 0,
      before
    });
    expect(notApplied.outcome).toBe("not_applied");
    expect(notApplied.checkpoint).toBeUndefined();
    expect(host.writes()).toEqual([]);
  });

  it("放弃本次执行：把原公式与格式一起写回，不留目标公式", async () => {
    const target = formulaPlan();
    host.setCell(TARGET, "A1", `=${SOURCE}!A1*10`, { asFormula: true });
    host.setCellStyle(TARGET, "A1", { fontBold: true });
    host.setCellDisplay(TARGET, "A1", { format: "0.00" });

    const result = await executePlan(target, {
      authorizedSheetNames: [SOURCE, TARGET],
      operationId: OPERATION
    });
    expect(host.cell(TARGET, "A1")).toBe(`=${SOURCE}!A1*2`);

    await undoExecution(result.undoSnapshot!);
    expect(host.cell(TARGET, "A1")).toBe(`=${SOURCE}!A1*10`);
    // 值/公式动作不动格式，撤销也只写回公式那一层。
    expect(host.cellStyle(TARGET, "A1")?.fontBold).toBe(true);
    expect(host.cellNumberFormat(TARGET, "A1")).toBe("0.00");
  });

  it("R1C1 铺满公式不在跨重启恢复白名单里（连效果核对都不给）", () => {
    const support = officeRecoverySupport(
      plan([
        {
          type: "writeFormulas",
          sheet: TARGET,
          range: "A1:A2",
          formulas: [[`=${SOURCE}!A1`], [`=${SOURCE}!A2`]],
          formulaR1C1: "=R[1]C[1]"
        }
      ])
    );
    expect(support).toEqual({
      supported: false,
      reason: "第 1 步 R1C1 铺满公式暂不支持跨重启效果核对"
    });
  });

  /*
   * **来源值变了，公式计划怎么办** —— 产品规则（2026-10-30 拍板）：
   * 公式精确恢复要求登记的来源证据仍然一致；来源值变化时 Excel / WPS 都**停止续跑**
   * 并明确提示，不自动重新规划。规则只针对恢复，不改变正常公式的自动重算。
   *
   * 判据读的是**公式真正依赖的那几张表的值**（`=数据!A1*3` ⇒ 数据），不是本次授权
   * 范围的整张快照 —— 拿无关快照比对既会在无关表被改时误拒，也证明不了公式的输入没变。
   */
  it("续跑时来源值变了：拒绝续跑，一格都不写，并且说的是「来源已变化」", async () => {
    const evidence = await captureWorkbookSourceFingerprint([SOURCE]);
    const anchored: AnalysisPlan = {
      ...formulaPlan(),
      sourceDependencyFingerprint: evidence
    };
    let checkpoint: PlanExecutionResult["undoSnapshot"] | null = null;
    try {
      await executePlan(anchored, {
        authorizedSheetNames: [SOURCE, TARGET],
        operationId: OPERATION,
        recovery: {
          startActionIndex: 0,
          onBeforeAction: async () => undefined,
          onActionCompleted: async ({ actionIndex }) => {
            if (actionIndex === 0) throw new Error("模拟提交第一步回执时断电");
          }
        }
      });
    } catch (error) {
      checkpoint = (error as PlanExecutionError).undoSnapshot ?? null;
    }
    expect(checkpoint).not.toBeNull();
    expect(host.cell(TARGET, "A1")).toBe(`=${SOURCE}!A1*2`);

    // 中断期间外部把源值从 2 改成 99。公式文本一个字符都没变 ——
    // 这正是"文本相同不能替代来源校验"的现场：计划里还是 `=数据!A1*3`。
    host.setCell(SOURCE, "A1", 99);
    host.resetWrites();
    expect(anchored.actions).toEqual(formulaPlan().actions);

    await expect(
      executePlan(anchored, {
        authorizedSheetNames: [SOURCE, TARGET],
        operationId: OPERATION,
        recovery: {
          startActionIndex: 1,
          undoSnapshot: checkpoint!,
          onBeforeAction: async () => undefined,
          onActionCompleted: async () => undefined
        }
      })
    ).rejects.toThrow(FORMULA_SOURCE_CHANGED_MESSAGE);
    // 零写入：连"先把没跑过的写了再说"都不允许。
    expect(host.writes()).toEqual([]);
    expect(host.cell(TARGET, "A2")).toBe("");
    expect(host.cell(TARGET, "A3")).toBe("");
  });

  it("来源没变：照常续跑，只写没跑过的那几步", async () => {
    const anchored: AnalysisPlan = {
      ...formulaPlan(),
      sourceDependencyFingerprint: await captureWorkbookSourceFingerprint([SOURCE])
    };
    let checkpoint: PlanExecutionResult["undoSnapshot"] | null = null;
    try {
      await executePlan(anchored, {
        authorizedSheetNames: [SOURCE, TARGET],
        operationId: OPERATION,
        recovery: {
          startActionIndex: 0,
          onBeforeAction: async () => undefined,
          onActionCompleted: async ({ actionIndex }) => {
            if (actionIndex === 0) throw new Error("模拟提交第一步回执时断电");
          }
        }
      });
    } catch (error) {
      checkpoint = (error as PlanExecutionError).undoSnapshot ?? null;
    }
    host.resetWrites();

    const result = await executePlan(anchored, {
      authorizedSheetNames: [SOURCE, TARGET],
      operationId: OPERATION,
      recovery: {
        startActionIndex: 1,
        undoSnapshot: checkpoint!,
        onBeforeAction: async () => undefined,
        onActionCompleted: async () => undefined
      }
    });

    expect(writtenAddresses()).toEqual([`${TARGET}!A2`, `${TARGET}!A3`]);
    expect(host.cell(TARGET, "A2")).toBe(`=${SOURCE}!A1*3`);
    expect(result.verification.checks.every((check) => check.passed)).toBe(true);
  });

  it("续跑被拒之后，放弃仍能把原值恢复回来（拒绝不等于卡死）", async () => {
    host.setCell(TARGET, "A1", `=${SOURCE}!A1*10`, { asFormula: true });
    const anchored: AnalysisPlan = {
      ...formulaPlan(),
      sourceDependencyFingerprint: await captureWorkbookSourceFingerprint([SOURCE])
    };
    let failure: PlanExecutionError | null = null;
    try {
      await executePlan(anchored, {
        authorizedSheetNames: [SOURCE, TARGET],
        operationId: OPERATION,
        recovery: {
          startActionIndex: 0,
          onBeforeAction: async () => undefined,
          onActionCompleted: async ({ actionIndex }) => {
            if (actionIndex === 0) throw new Error("模拟提交第一步回执时断电");
          }
        }
      });
    } catch (error) {
      failure = error as PlanExecutionError;
    }
    const checkpoint = failure!.undoSnapshot!;
    expect(host.cell(TARGET, "A1")).toBe(`=${SOURCE}!A1*2`);

    host.setCell(SOURCE, "A1", 99);
    await expect(
      executePlan(anchored, {
        authorizedSheetNames: [SOURCE, TARGET],
        operationId: OPERATION,
        recovery: {
          startActionIndex: 1,
          undoSnapshot: checkpoint,
          onBeforeAction: async () => undefined,
          onActionCompleted: async () => undefined
        }
      })
    ).rejects.toThrow(FORMULA_SOURCE_CHANGED_MESSAGE);

    // 拒绝只挡"继续写"，不挡"退回去"：撤销点仍在，原公式照旧写回。
    await undoExecution(checkpoint);
    expect(host.cell(TARGET, "A1")).toBe(`=${SOURCE}!A1*10`);
    expect(host.cell(TARGET, "A2")).toBe("");
  });

  it("没有登记来源证据的公式计划：拒绝续跑，不拿现读补基线", async () => {
    let checkpoint: PlanExecutionResult["undoSnapshot"] | null = null;
    try {
      await executePlan(formulaPlan(), {
        authorizedSheetNames: [SOURCE, TARGET],
        operationId: OPERATION,
        recovery: {
          startActionIndex: 0,
          onBeforeAction: async () => undefined,
          onActionCompleted: async ({ actionIndex }) => {
            if (actionIndex === 0) throw new Error("模拟提交第一步回执时断电");
          }
        }
      });
    } catch (error) {
      checkpoint = (error as PlanExecutionError).undoSnapshot ?? null;
    }
    host.resetWrites();
    await expect(
      executePlan(formulaPlan(), {
        authorizedSheetNames: [SOURCE, TARGET],
        operationId: OPERATION,
        recovery: {
          startActionIndex: 1,
          undoSnapshot: checkpoint!,
          onBeforeAction: async () => undefined,
          onActionCompleted: async () => undefined
        }
      })
    ).rejects.toThrow("没有登记公式来源的证据");
    expect(host.writes()).toEqual([]);
  });

  /*
   * 同一条闸门的**负向锚点**：结构闸与值指纹闸读的是不同的事实。
   * 少了它，"放行"与"根本没检查"就分不开 —— 这里读到的是"表不在了"（结构），
   * 上面那条读到的是"表还在、值变了"（值）。
   */
  it("续跑时点名的来源表已不在工作簿里：拒绝执行，一格都不写", async () => {
    const anchored: AnalysisPlan = {
      ...formulaPlan(),
      sourceFingerprint: await captureWorkbookSourceFingerprint([SOURCE]),
      sourceFingerprintSheets: [SOURCE, "已删除表"]
    };
    host.resetWrites();
    await expect(
      executePlan(anchored, {
        // 授权照登记时的原样带上（续跑时授权来自服务端记录），表本身已经不在了。
        authorizedSheetNames: [SOURCE, TARGET, "已删除表"],
        operationId: OPERATION
      })
    ).rejects.toThrow("计划所需的来源已发生变化");
    expect(host.writes()).toHaveLength(0);
  });
});

/**
 * 登记闸：**公式依赖的那几张表，能不能被登记的证据覆盖**。
 *
 * 这一步在"写第一个格之前"发生（`officeRecoverySupport` → 服务端登记）。它判的是
 * "有没有资格登记精确恢复"，不是"这次能不能写" —— 判成不支持时计划照样执行，
 * 只是不声称能跨重启续跑。两者不许混着说。
 */
describe("公式计划的恢复登记闸（officeRecoverySupport）", () => {
  const support = (actions: ExcelAction[], dependency?: string[]) =>
    officeRecoverySupport({
      id: "plan-formula-support",
      title: "公式",
      summary: "公式",
      assumptions: [],
      warnings: [],
      actions,
      ...(dependency === undefined ? {} : { sourceDependencySheets: dependency })
    });

  it("跨表公式 + 记账覆盖：可以登记，并把要钉指纹的表交出来", () => {
    expect(
      support(
        [{ type: "writeFormulas", sheet: TARGET, range: "A1", formulas: [[`=${SOURCE}!A1*2`]] }],
        [SOURCE]
      )
    ).toEqual({ supported: true, sourceEvidenceSheets: [SOURCE] });
  });

  it("同表公式（=A1*2）：不登记，理由说清是「分不开本次写入与外部修改」", () => {
    const verdict = support([
      { type: "writeFormulas", sheet: TARGET, range: "A1", formulas: [["=A1*2"]] }
    ], [SOURCE]);
    // 先看理由：理由里带着"为什么不可分离"，只断言 supported=false 会把原因丢掉。
    expect(verdict.reason ?? "").toContain("写入目标表自身");
    expect(verdict.supported).toBe(false);
  });

  it("同表公式只是不能**精确恢复**：普通执行照常把它写进去", async () => {
    host.resetWrites();
    const sameSheet: AnalysisPlan = plan([
      { type: "writeFormulas", sheet: TARGET, range: "B1", formulas: [["=A1*2"]] }
    ]);
    // 这里刻意用 `executePlan`（不带 recovery）：`officeRecoverySupport` 只回答
    // "能不能跨重启续跑"，不回答"能不能写"。两者混着说，就会把一条恢复能力的缺失
    // 读成产品不能写同表公式。
    await executePlan(sameSheet, {
      authorizedSheetNames: [SOURCE, TARGET],
      operationId: OPERATION
    });
    expect(writtenAddresses()).toEqual([`${TARGET}!B1`]);
    expect(host.cell(TARGET, "B1")).toBe("=A1*2");
  });

  it("公式读了记账里没有的表：不登记，点名是哪张表", () => {
    const verdict = support([
      { type: "writeFormulas", sheet: TARGET, range: "A1", formulas: [["=参数!A1*2"]] }
    ], [SOURCE]);
    expect(verdict.reason ?? "").toContain("参数");
    expect(verdict.reason ?? "").toContain("没有本机读取记录");
    expect(verdict.supported).toBe(false);
  });

  it("没有可信记账（缺失）：不登记，不拿指纹范围冒充依赖", () => {
    const verdict = support([
      { type: "writeFormulas", sheet: TARGET, range: "A1", formulas: [[`=${SOURCE}!A1*2`]] }
    ]);
    expect(verdict.reason ?? "").toContain("可信记账");
    expect(verdict.supported).toBe(false);
  });

  it("动态引用：不登记（解析不出读哪张表，也就没法定证据口径）", () => {
    const verdict = support([
      { type: "writeFormulas", sheet: TARGET, range: "A1", formulas: [['=INDIRECT("数据!A1")']] }
    ], [SOURCE]);
    expect(verdict.reason ?? "").toContain("动态引用");
    expect(verdict.supported).toBe(false);
  });
});
