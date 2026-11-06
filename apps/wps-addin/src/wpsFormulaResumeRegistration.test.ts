/**
 * 四表公式计划的**恢复登记载荷**（前端实际提交的那一份）。
 *
 * 存在理由（2026-11-01 真机）：用户的真实场景是
 *
 * - 来源表 `来源!A1:B5`（只读）；
 * - 三个独立动作：`结果1!B2 =来源!B2*1`、`结果2!B2 =来源!B4*2`、`结果3!B2 =来源!B5*3`；
 * - 目标原先为空。
 *
 * 计划生成正常，写入前登记被服务端拒掉，界面显示「无法持久登记本次写入现场…」。
 * 根因在服务端（`_validate_wps_recovery_descriptor` 把 `all(type == "writeValues")`
 * 折进一个 `valid` 布尔，凡是带来源证据的公式动作一律被拒），**不在**这里：端上早就
 * 把 `writeFormulas` 算作可带来源证据的动作（`isResumableValueAction`）。
 *
 * 这个文件钉住"前端实际提交的载荷长什么样"，与
 * `server/tests/test_wps_formula_resume.py`（同一份形状、服务端一侧）成对：
 * 两端各断言同一份形状，任何一端单独改动都会有一侧变红。
 */

import {describe, expect, it} from "vitest";
import type {AnalysisPlan} from "../../excel-addin/src/contracts";
import {assertWpsRecoveryDescriptor} from "../../excel-addin/src/contracts";
import {prepareWpsPlan, executeWpsPlan, wpsResumeContextOutcome} from "./wpsPlanExecutor";
import {wpsRecoveryDescriptor} from "./wpsExecutionLedger";
import {fixture} from "./wpsResume.testHost";

const SHEETS = ["来源", "结果1", "结果2", "结果3"] as const;

/** 与真机同形状的计划：三个独立动作，来源只读，目标原为空。 */
const FOUR_SHEET_FORMULA_PLAN: AnalysisPlan = {
  id: "plan-four-sheet-formulas",
  title: "三张结果表分别引用来源的不同行",
  summary: "结果1!B2 =来源!B2*1；结果2!B2 =来源!B4*2；结果3!B2 =来源!B5*3",
  assumptions: [],
  warnings: [],
  sourceDependencySheets: ["来源"],
  actions: [
    {type: "writeFormulas", sheet: "结果1", range: "B2:B2", formulas: [["=来源!B2*1"]]},
    {type: "writeFormulas", sheet: "结果2", range: "B2:B2", formulas: [["=来源!B4*2"]]},
    {type: "writeFormulas", sheet: "结果3", range: "B2:B2", formulas: [["=来源!B5*3"]]}
  ],
  acceptanceCriteria: []
};

/** 来源 A1:B5 的现场值（两列五行），也是本轮可信读取记下的那份矩阵。 */
const SOURCE_MATRIX = [[10, 100], [20, 200], [30, 300], [40, 400], [50, 500]];

function prepare() {
  const f = fixture({sheetNames: [...SHEETS]});
  // 来源 A1:B5 摆上值（覆盖替身预置的 A1 种子格）；三个目标格保持为空 —— 真机里它们原本是空的。
  SOURCE_MATRIX.forEach((row, rowIndex) => {
    row.forEach((value, columnIndex) => {
      f.cells.set(`来源!${columnIndex === 0 ? "A" : "B"}${rowIndex + 1}`, value);
    });
  });
  for (const name of ["结果1", "结果2", "结果3"]) f.cells.set(`${name}!B2`, null);
  const prepared = prepareWpsPlan(f.application, [1, 2, 3, 4], FOUR_SHEET_FORMULA_PLAN,
    [{sheetName: "来源", range: "A1:B5", values: SOURCE_MATRIX}]);
  return {f, prepared, descriptor: wpsRecoveryDescriptor(prepared)};
}

describe("四表公式计划的恢复登记载荷", () => {
  it("来源证据盖住整张来源表，动作类型仍是公式", () => {
    const {descriptor} = prepare();
    /*
     * 这两条正是服务端判据要读的东西：`resumeContext.version === 1` 且来源非空。
     * 服务端此前在这份载荷上只看到"动作不是 writeValues"就整份拒掉。
     */
    expect(descriptor.resumeContext).toMatchObject({version: 1});
    expect(descriptor.resumeContext!.sources.map(source => ({
      sheetName: source.sheetName,
      range: source.range,
      rowCount: source.rowCount,
      columnCount: source.columnCount
    }))).toEqual([{sheetName: "来源", range: "A1:B5", rowCount: 5, columnCount: 2}]);
    // 指纹必须真的算过（`v2:<长度>:<8 位>`），不是占位串。
    for (const source of descriptor.resumeContext!.sources) {
      expect(source.previewFingerprint).toMatch(/^v2:\d+:[0-9a-f]{8}$/);
      expect(source.valueFingerprint).toMatch(/^v2:\d+:[0-9a-f]{8}$/);
    }

    expect(descriptor.actions.map(action => ({type: action.type, sheet: action.sheet, range: action.range})))
      .toEqual([
        {type: "writeFormulas", sheet: "结果1", range: "B2:B2"},
        {type: "writeFormulas", sheet: "结果2", range: "B2:B2"},
        {type: "writeFormulas", sheet: "结果3", range: "B2:B2"}
      ]);
    // 三个目标互不相同、都不是来源表 —— 服务端"来源即目标"那条在这里根本不成立。
    expect(new Set(descriptor.actions.map(action => action.sheet)).size).toBe(3);
    expect(descriptor.actions.some(action => action.sheet === "来源")).toBe(false);
  });

  it("载荷里没有公式文本，也没有单元格值矩阵", () => {
    const {descriptor} = prepare();
    const serialized = JSON.stringify(descriptor);
    expect(serialized).not.toContain("=来源!");
    for (const forbidden of ["\"values\"", "\"formulas\"", "\"preview\"", "\"cells\""]) {
      expect(serialized).not.toContain(forbidden);
    }
    // 锚点：来源证据确实在载荷里（否则上面"没有值"的断言可能只是因为它什么都没带）。
    expect(serialized).toContain("valueFingerprint");
  });

  it("这份载荷通过端上的出站契约闸", () => {
    const {descriptor} = prepare();
    // 端上自己认为它可以出站；服务端此前却不认 —— 这是判据漂移，不是载荷可疑。
    expect(() => assertWpsRecoveryDescriptor(descriptor)).not.toThrow();
  });

  it("登记成功后：前两步写入并回读，第二步回执前中断，第三步一格没写", async () => {
    /*
     * 真机要求的验证顺序：登记成功 → 前两步写入并回读 → 第二步回执前中断 → 第三步未写。
     *
     * 用**同一份四表夹具**（不是另造一个更简单的形状）：三个动作分属三张不同的表，
     * 所以"第三步未写"必须是"`结果3!B2` 仍然是空的"，不可能被前两步捎带写成。
     *
     * 中断点选在 `after`（回读之后、回执之前）：`onActionCheckpoint` 抛错会中止后续
     * 动作 —— 拿不到持久回执就不继续写。
     */
    const {f, prepared} = prepare();
    const started: number[] = [];
    const completed: number[] = [];
    const persisted: string[] = [];
    await expect(executeWpsPlan(f.application, prepared, undefined, undefined,
      async (phase, checkpoint) => {
        if (phase === "before") started.push(checkpoint.actionIndex);
        if (phase === "after") {
          completed.push(checkpoint.actionIndex);
          if (checkpoint.actionIndex === 1) throw new Error("第二步回执前中断");
        }
      }, "wps-execution-four-sheet", {
      taskId: "task-formula",
      /*
       * 本用例的对象是"写入与中断的边界"，不是恢复包的落盘耐久性（那条由
       * `wpsRecoveryPackage` / `wpsExecutionResume` 的用例覆盖）。执行器要求
       * 可精确续跑的执行必须**先**有落盘通道，所以这里给一个记账用的替身，
       * 只证明"写前确实被要求落盘过一次"。
       */
      persistMaterial: async (record) => {
        persisted.push(record.executionId);
      }
    })).rejects.toThrow("第二步回执前中断");

    // 前两步真的进了宿主并回读过；第三步连开始都没开始。
    expect(started).toEqual([0, 1]);
    expect(completed).toEqual([0, 1]);
    // 恢复材料在**第一次宿主写入之前**就已经被要求落盘过（顺序保证，不是事后补存）。
    expect(persisted).toHaveLength(1);
    expect(f.writes).toEqual(["结果1!B2", "结果2!B2"]);
    // 写进去的是**公式文本**，不是计算值（替身不猜计算结果，落的是 null）。
    expect(f.cells.get("结果1!B2")).toEqual({formula: "=来源!B2*1", value: null});
    expect(f.cells.get("结果2!B2")).toEqual({formula: "=来源!B4*2", value: null});
    expect(f.cells.get("结果3!B2")).toBeNull();
  });

  it("计划漏声明 sourceDependencySheets：仍登记精确恢复（依赖以公式解析为准，不静默降级）", () => {
    /*
     * 2026-11-01 真机（WPS 公式 B 首轮，`wps-execution-4998b190…`）：计划的
     * `sourceDependencySheets` 为空，判据读声明直接返回 null —— 登记不带
     * `resumeContext`、恢复包不生成，恢复卡只能说"没有恢复材料"，事后还查不出原因。
     *
     * 但 `prepareWpsPlan` 已经从 `=来源!B2*1` 解析出 `来源`，也把 `来源!A1:B5` 的
     * 证据采进了 `sources`。**声明缺失不等于没有依赖**：解析结果就是同一份事实。
     */
    const f = fixture({sheetNames: [...SHEETS]});
    SOURCE_MATRIX.forEach((row, rowIndex) => {
      row.forEach((value, columnIndex) => {
        f.cells.set(`来源!${columnIndex === 0 ? "A" : "B"}${rowIndex + 1}`, value);
      });
    });
    for (const name of ["结果1", "结果2", "结果3"]) f.cells.set(`${name}!B2`, null);
    const plan = {
      ...FOUR_SHEET_FORMULA_PLAN,
      sourceDependencySheets: undefined as unknown as string[]
    };
    const prepared = prepareWpsPlan(f.application, [1, 2, 3, 4], plan,
      [{sheetName: "来源", range: "A1:B5", values: SOURCE_MATRIX}]);
    // 判据的正面锚点：证据确实在（没有它，"能登记"可能只是压根没读来源）。
    expect(prepared.sources.map(s => s.sheetName)).toContain("来源");
    // 不再因为"声明为空"而静默丢失精确恢复资格。
    const descriptor = wpsRecoveryDescriptor(prepared);
    expect(descriptor.resumeContext).toMatchObject({version: 1});
    expect(descriptor.resumeContext!.sources.map(s => `${s.sheetName}!${s.range}`))
      .toEqual(["来源!A1:B5"]);
  });

  it("wpsResumeContextOutcome：给出「为什么不能精确恢复」的同一份原因", () => {
    /*
     * 用一个**允许写、但不能精确恢复**的结构：第三条公式引用了写入目标表自身的单元格
     * （`=B1*3`，未限定表名 ⇒ `formulaDependencies.local`）。`formulaRequirement.ok`
     * 为 false，旧实现这里只返回 null；新实现必须带上"写入目标表自身"这条原因，而不是吞掉。
     *
     * 换成同表引用而不是 `INDIRECT`：动态引用现在在 `prepareWpsPlan` 就被**拒写**
     * （授权闸无法判定依赖），根本走不到"求原因"这一步 —— 两条边界不能互相顶替。
     */
    const f = fixture({sheetNames: [...SHEETS]});
    SOURCE_MATRIX.forEach((row, rowIndex) => {
      row.forEach((value, columnIndex) => {
        f.cells.set(`来源!${columnIndex === 0 ? "A" : "B"}${rowIndex + 1}`, value);
      });
    });
    for (const name of ["结果1", "结果2", "结果3"]) f.cells.set(`${name}!B2`, null);
    const planLocal: AnalysisPlan = {
      ...FOUR_SHEET_FORMULA_PLAN,
      actions: [
        {type: "writeFormulas", sheet: "结果1", range: "B2", formulas: [["=来源!B2*1"]]},
        {type: "writeFormulas", sheet: "结果2", range: "B2", formulas: [["=来源!B4*2"]]},
        {type: "writeFormulas", sheet: "结果3", range: "B2", formulas: [["=B1*3"]]}
      ]
    };
    const prepared = prepareWpsPlan(f.application, [1, 2, 3, 4], planLocal,
      [{sheetName: "来源", range: "A1:B5", values: SOURCE_MATRIX}]);
    const outcome = wpsResumeContextOutcome(prepared);
    expect(outcome.context).toBeNull();
    expect(outcome.reason).toContain("写入目标表自身");
  });

  it("声明非空但漏掉另一张实际依赖表：证据照采，但不登记精确恢复（不静默放行）", () => {
    /*
     * 权威来源是**公式文本的解析结果**，`sourceDependencySheets` 只是声明。
     * 声明漏了一张实际依赖表时两种错法都要被挡住：
     *  - 只按声明采证据 ⇒ `辅助` 没有基线，续跑时改它拦不住（漏采）；
     *  - 只按声明判依赖 ⇒ `辅助` 根本不在判据里，等于静默放行。
     * 所以断言分三处：证据覆盖两张表（没漏采）、登记被拒（没放行）、原因点名漏的那张。
     */
    const f = fixture({sheetNames: ["来源", "辅助", "结果1", "结果2", "结果3"]});
    for (const name of ["来源", "辅助"]) {
      SOURCE_MATRIX.forEach((row, rowIndex) => {
        row.forEach((value, columnIndex) => {
          f.cells.set(`${name}!${columnIndex === 0 ? "A" : "B"}${rowIndex + 1}`, value);
        });
      });
    }
    for (const name of ["结果1", "结果2", "结果3"]) f.cells.set(`${name}!B2`, null);
    const plan: AnalysisPlan = {
      ...FOUR_SHEET_FORMULA_PLAN,
      id: "plan-declaration-misses-a-dependency",
      // 只声明了「来源」，第三条公式却读「辅助」。
      sourceDependencySheets: ["来源"],
      actions: [
        {type: "writeFormulas", sheet: "结果1", range: "B2", formulas: [["=来源!B2*1"]]},
        {type: "writeFormulas", sheet: "结果2", range: "B2", formulas: [["=辅助!B2*2"]]},
        {type: "writeFormulas", sheet: "结果3", range: "B2", formulas: [["=来源!B5*3"]]}
      ]
    };
    const prepared = prepareWpsPlan(f.application, [1, 2, 3, 4, 5], plan);

    // ① 没漏采：两张实际依赖表都补了整块有效区域的证据。
    const covered = prepared.sources.map(source => `${source.sheetName}!${source.range}`).sort();
    expect(covered).toEqual(["来源!A1:B5", "辅助!A1:B5"].sort());
    expect(Object.keys(prepared.formulaSourceCoverage ?? {}).sort())
      .toEqual(["来源", "辅助"].sort());

    // ② 没静默放行：不登记精确恢复，且原因点名漏声明的那张表。
    const outcome = wpsResumeContextOutcome(prepared);
    expect(outcome.context).toBeNull();
    expect(outcome.reason).toContain("辅助");
    expect(outcome.reason).toContain("没有声明");
    expect(wpsRecoveryDescriptor(prepared).resumeContext ?? null).toBeNull();

    // ③ 登记被拒只影响恢复资格，不影响"这一次准备"：一格都没写。
    expect(f.writes).toEqual([]);
  });

  it("实际公式引用未授权的表：写入前拒绝整份计划，零写入", () => {
    /*
     * 授权边界是用户勾选的那几张表。`=其它!A1` 会把没勾选的表的值算进已授权目标格，
     * 所以必须在**写入前**拒绝 —— 只"不登记精确恢复"是恢复口径，挡不住这次读。
     * 正向锚点：同一份计划把那张表也勾上就能通过，证明拒绝点确实在授权而不是公式形状。
     */
    const f = fixture({sheetNames: ["来源", "结果1", "结果2", "结果3", "其它"]});
    SOURCE_MATRIX.forEach((row, rowIndex) => {
      row.forEach((value, columnIndex) => {
        f.cells.set(`来源!${columnIndex === 0 ? "A" : "B"}${rowIndex + 1}`, value);
      });
    });
    for (const name of ["结果1", "结果2", "结果3"]) f.cells.set(`${name}!B2`, null);
    const plan: AnalysisPlan = {
      ...FOUR_SHEET_FORMULA_PLAN,
      id: "plan-unauthorized-dependency",
      actions: [
        {type: "writeFormulas", sheet: "结果1", range: "B2", formulas: [["=来源!B2*1"]]},
        {type: "writeFormulas", sheet: "结果2", range: "B2", formulas: [["=其它!B2*2"]]},
        {type: "writeFormulas", sheet: "结果3", range: "B2", formulas: [["=来源!B5*3"]]}
      ]
    };
    // 只勾前四张：`其它` 存在但未授权。
    expect(() => prepareWpsPlan(f.application, [1, 2, 3, 4], plan))
      .toThrow("未授权的工作表「其它」");
    expect(f.writes).toEqual([]);

    // 正向锚点：把 `其它` 一起勾上，同一份计划就能准备（拒绝并非来自公式形状）。
    const prepared = prepareWpsPlan(f.application, [1, 2, 3, 4, 5], plan);
    expect(prepared.actions).toHaveLength(3);
  });

  it("动态引用解析不出依赖：写入前拒绝整份计划，零写入（与 Excel 同一条边界）", () => {
    /*
     * `INDIRECT` / `OFFSET` 的目标是**运行时算出来的**，静态解析不出它读哪张表。
     * 这意味着授权闸在这条公式上**无法判定**，而"无法判定"不等于"通过"：
     * 放行等于把"读哪张表"的决定权交给公式自己，用户勾的边界形同建议。
     *
     * 所以边界与 Excel 端 `validatePlan` 的 `references.dynamic` 一致：**首写前拒绝**。
     * 曾经这里只取消精确恢复资格、普通执行照写 —— 那是把恢复口径当成了授权口径。
     *
     * 本用例不启用验收探针、不带恢复通道：拒绝不依赖探针是否 arm。
     * 正向锚点：把动态引用换成能解析的 `=来源!B5*3`，同一份计划就能准备并通过。
     */
    const f = fixture({sheetNames: [...SHEETS]});
    SOURCE_MATRIX.forEach((row, rowIndex) => {
      row.forEach((value, columnIndex) => {
        f.cells.set(`来源!${columnIndex === 0 ? "A" : "B"}${rowIndex + 1}`, value);
      });
    });
    for (const name of ["结果1", "结果2", "结果3"]) f.cells.set(`${name}!B2`, null);
    const plan: AnalysisPlan = {
      ...FOUR_SHEET_FORMULA_PLAN,
      id: "plan-dynamic-dependency",
      actions: [
        {type: "writeFormulas", sheet: "结果1", range: "B2", formulas: [["=来源!B2*1"]]},
        {type: "writeFormulas", sheet: "结果2", range: "B2", formulas: [["=来源!B4*2"]]},
        {type: "writeFormulas", sheet: "结果3", range: "B2", formulas: [["=INDIRECT(\"来源!B5\")*3"]]}
      ]
    };
    expect(() => prepareWpsPlan(f.application, [1, 2, 3, 4], plan)).toThrow("动态引用");
    // 零写入：拒绝在任何宿主读取/写入之前，三格仍是空的。
    expect(f.writes).toEqual([]);
    expect(["结果1", "结果2", "结果3"].map(name => f.cells.get(`${name}!B2`))).toEqual([null, null, null]);

    // 正向锚点：只把第三条换成静态可解析的引用，同一份计划就能准备。
    const staticPlan: AnalysisPlan = {
      ...plan,
      actions: [
        {type: "writeFormulas", sheet: "结果1", range: "B2", formulas: [["=来源!B2*1"]]},
        {type: "writeFormulas", sheet: "结果2", range: "B2", formulas: [["=来源!B4*2"]]},
        {type: "writeFormulas", sheet: "结果3", range: "B2", formulas: [["=来源!B5*3"]]}
      ]
    };
    const prepared = prepareWpsPlan(f.application, [1, 2, 3, 4], staticPlan);
    expect(prepared.actions).toHaveLength(3);
    expect(wpsResumeContextOutcome(prepared).context).not.toBeNull();
  });
});
