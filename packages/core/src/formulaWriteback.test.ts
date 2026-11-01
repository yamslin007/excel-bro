import { describe, expect, it } from "vitest";
import {
  FORMULA_SOURCE_CHANGED_MESSAGE,
  FORMULA_SOURCE_EVIDENCE_MISSING_MESSAGE,
  FORMULA_SOURCE_UNVERIFIABLE_MESSAGE,
  formulaCellSite,
  formulaDependencies,
  formulaMatrixSite,
  formulaResumeGate,
  formulaSheetReferences,
  formulaSourceEvidenceRequirement,
  formulaTextEqual,
  formulaWriteMatrixIssue,
  isFormulaText,
  verifyFormulaResumeSources
} from "./formulaWriteback";

/**
 * 这组用例守的是一条口径：**公式的身份是它的文本，不是它算出来的数。**
 *
 * 所以最关键的几格都用"计算值相同、文本不同"来构造：如果哪天有人把判据改成
 * 拿显示值比较，这些用例会红。
 */
describe("公式文本身份", () => {
  it("算得出同一个数的两条公式不是同一条", () => {
    expect(formulaTextEqual("=1+1", "=2")).toBe(false);
    expect(formulaTextEqual("=A1+B1", "=B1+A1")).toBe(false);
  });

  it("相对引用与绝对引用是两个写入目标，不做归一化", () => {
    expect(formulaTextEqual("=A1*2", "=$A$1*2")).toBe(false);
    expect(formulaTextEqual("=SUM(A1:A3)", "=SUM(B1:B3)")).toBe(false);
  });

  it("公式文本逐字符比较：不改大小写、不 trim", () => {
    expect(formulaTextEqual("=sum(A1:A3)", "=SUM(A1:A3)")).toBe(false);
    expect(formulaTextEqual("=A1", " =A1")).toBe(false);
    expect(formulaTextEqual("=A1", "=A1")).toBe(true);
  });

  it("常量永远不是公式，哪怕长得像", () => {
    expect(isFormulaText(2)).toBe(false);
    expect(isFormulaText("2")).toBe(false);
    expect(isFormulaText("=2")).toBe(true);
    expect(formulaTextEqual(2, "=2")).toBe(false);
  });
});

describe("写后目标矩阵必须是公式文本", () => {
  it("退化成常量就拒：算出来的数一样也不行", () => {
    // `=1+1` 的计算值是 2；把矩阵换成 [[2]] 时，值一样、写入语义完全不同。
    const issue = formulaWriteMatrixIssue([[2]]);
    expect(issue).toEqual({ row: 0, column: 0, kind: "not_text" });
    expect(formulaWriteMatrixIssue([["2"]])).toEqual({ row: 0, column: 0, kind: "not_formula" });
  });

  it("点名到具体格子，不笼统说'格式不对'", () => {
    expect(formulaWriteMatrixIssue([["=A1", null]])).toEqual({ row: 0, column: 1, kind: "not_text" });
    expect(formulaWriteMatrixIssue([["=A1"], ["=A2"]])).toBeNull();
  });
});

describe("公式引用了哪些表", () => {
  it("普通表名、带引号的表名都认得", () => {
    expect(formulaSheetReferences([["=Source!A1*2"]]).sheets).toEqual(["Source"]);
    expect(formulaSheetReferences([["='My Sheet'!A1+Sheet2!B2"]]).sheets).toEqual(["My Sheet", "Sheet2"]);
  });

  it("字符串字面量里的 ! 不算引用", () => {
    expect(formulaSheetReferences([['="Source!A1"']]).sheets).toEqual([]);
  });

  it("外部工作簿引用原样保留，交给调用方拒绝", () => {
    expect(formulaSheetReferences([["=[Book1.xlsx]Sheet1!A1"]]).sheets).toEqual(["[Book1.xlsx]Sheet1"]);
  });

  it("动态引用解析不出来就说不知道，不假装没依赖", () => {
    const indirect = formulaSheetReferences([["=INDIRECT(\"A1\")"]]);
    expect(indirect.dynamic).toBe(true);
    expect(formulaSheetReferences([["=OFFSET(A1,1,0)"]]).dynamic).toBe(true);
  });

  it("非公式格不贡献引用", () => {
    expect(formulaSheetReferences([[null, "Source!A1" as unknown as string]]).sheets).toEqual([]);
  });
});

describe("一格现场：写前 / 写后 / 判不出来 / 都不是", () => {
  it("写前", () => {
    expect(formulaCellSite("=old", "=A1*2", "=old")).toBe("before");
  });

  it("写后", () => {
    expect(formulaCellSite("=old", "=A1*2", "=A1*2")).toBe("written");
  });

  it("被改成第三条公式 ⇒ 都不是，停下来核对", () => {
    expect(formulaCellSite("=old", "=A1*2", "=A1*3")).toBe("unknown");
  });

  it("**显示值是我们要的数、但公式不是** ⇒ 不算写入", () => {
    // 该格计算值可能正好等于 `=A1*2` 的结果，但现场已经不是公式了（被粘成常量）。
    expect(formulaCellSite("=old", "=A1*2", null)).toBe("unknown");
  });

  it("写前写后文本一样时不许报'已写入'", () => {
    // 把 `=A1` 再写成 `=A1`：文本无法证明写过。
    expect(formulaCellSite("=A1", "=A1", "=A1")).toBe("indeterminate");
    expect(formulaCellSite("=A1", "=A1", "=B1")).toBe("unknown");
  });
});

describe("整块现场", () => {
  const before = [["=old1"], ["=old2"], ["=old3"]];
  const written = [["=A1*2"], ["=A2*2"], ["=A3*2"]];

  it("全写前 / 全写后", () => {
    expect(formulaMatrixSite(before, written, before).kind).toBe("before");
    expect(formulaMatrixSite(before, written, written).kind).toBe("written");
  });

  it("写了一半就是在途不明，不是'已写入'", () => {
    const site = formulaMatrixSite(before, written, [["=A1*2"], ["=old2"], ["=old3"]]);
    expect(site.kind).toBe("partial");
    expect(site.detail).toEqual({ row: 1, column: 0, site: "before" });
  });

  it("有一格被人改过 ⇒ unknown 优先，别猜", () => {
    const site = formulaMatrixSite(before, written, [["=A1*2"], ["=A2*2"], ["=手改"]]);
    expect(site.kind).toBe("unknown");
    expect(site.detail).toEqual({ row: 2, column: 0, site: "unknown" });
  });

  it("形状对不上就是 unknown，不谈写没写", () => {
    expect(formulaMatrixSite(before, written, [["=A1*2"]]).kind).toBe("unknown");
    expect(formulaMatrixSite([], [], []).kind).toBe("unknown");
  });

  it("多列区域逐格归类", () => {
    const site = formulaMatrixSite([["=a", "=b"]], [["=A1", "=B1"]], [["=A1", "=B1"]]);
    expect(site.kind).toBe("written");
    expect(site.cells).toEqual([["written", "written"]]);
  });
});

describe("续跑闸", () => {
  const before = [["=old"]];
  const written = [["=A1*2"]];

  it("写前才写、写后就跳过、其余停下来核对", () => {
    expect(formulaResumeGate(formulaMatrixSite(before, written, before))).toBe("write");
    expect(formulaResumeGate(formulaMatrixSite(before, written, written))).toBe("skip");
    expect(formulaResumeGate(formulaMatrixSite(before, written, [["=手改"]]))).toBe("reconcile");
    expect(formulaResumeGate(formulaMatrixSite(before, written, [["=old"], ["=x"]]))).toBe("reconcile");
  });
});

/**
 * 依赖面：恢复时"来源证据必须覆盖哪些表"。
 *
 * 2026-10-30 的产品规则：公式精确恢复要求登记的来源证据仍然一致；来源值变了就停止
 * 续跑。这条规则能落地的前提是"依赖看得见" —— 所以这里既钉跨表引用，也钉**未限定
 * 表名的本表引用**（`=A1*2`）与**定义名称**（解析不出 ⇒ 依赖不确定）。漏掉后者，
 * 一条同表公式就能绕开来源校验。
 */
describe("公式依赖面：跨表、同表、动态、外部与名称", () => {
  it("显式跨表引用登记成依赖表", () => {
    expect(formulaDependencies([["=数据!A1*2", "='汇总 表'!B2"]])).toEqual({
      sheets: ["数据", "汇总 表"],
      local: false,
      dynamic: false,
      external: false,
      unresolved: false
    });
    // 区域引用同样带表名：`SUM(数据!A:A)` 不是本表引用。
    expect(formulaDependencies([["=SUM(数据!A:A)"]])).toEqual({
      sheets: ["数据"],
      local: false,
      dynamic: false,
      external: false,
      unresolved: false
    });
  });

  it("未限定表名的单元格/整列引用算**本表**依赖", () => {
    for (const formula of ["=A1*2", "=$A$1*2", "=SUM(A1:B2)", "=SUM(A:A)", "=SUM(3:5)"]) {
      const deps = formulaDependencies([[formula]]);
      expect({ formula, local: deps.local }).toEqual({ formula, local: true });
      expect(deps.unresolved).toBe(false);
    }
  });

  it("函数名、布尔字面量、科学计数不被误当成引用", () => {
    expect(formulaDependencies([["=LOG10(2)"]])).toEqual({
      sheets: [], local: false, dynamic: false, external: false, unresolved: false
    });
    expect(formulaDependencies([["=TRUE"]])).toEqual({
      sheets: [], local: false, dynamic: false, external: false, unresolved: false
    });
    expect(formulaDependencies([["=1E5*2"]])).toEqual({
      sheets: [], local: false, dynamic: false, external: false, unresolved: false
    });
    // `LOG10(A1)` 只有一个未限定引用 A1，不该再额外报"认不出的名字"。
    expect(formulaDependencies([["=LOG10(A1)"]]).unresolved).toBe(false);
  });

  it("双引号字面量里的 `!` 不是引用，但字面量外面的 A1 仍是本表依赖", () => {
    expect(formulaDependencies([['="数据!A1"&A1']])).toEqual({
      sheets: [], local: true, dynamic: false, external: false, unresolved: false
    });
  });

  it("外部工作簿与动态引用各归各档（都不进依赖表）", () => {
    expect(formulaDependencies([["=[Book1.xlsx]Sheet1!A1"]]).external).toBe(true);
    expect(formulaDependencies([["=[Book1.xlsx]Sheet1!A1"]]).sheets).toEqual([]);
    expect(formulaDependencies([["=INDIRECT(\"数据!A1\")"]]).dynamic).toBe(true);
    expect(formulaDependencies([["=OFFSET(数据!A1,1,0)"]]).dynamic).toBe(true);
  });

  it("认不出的裸标识符（定义名称）标成 unresolved，不当成没有依赖", () => {
    expect(formulaDependencies([["=税率*2"]]).unresolved).toBe(true);
    expect(formulaDependencies([["=MyRate*2"]]).unresolved).toBe(true);
  });
});

describe("公式计划的来源证据要求（续跑放行判据）", () => {
  const plan = (actions: unknown[]) => ({ actions }) as never;

  it("跨表公式：放行，并给出必须被证据覆盖的表", () => {
    expect(
      formulaSourceEvidenceRequirement(
        plan([
          { type: "writeFormulas", sheet: "结果", formulas: [["=数据!A1*2"]] }
        ])
      )
    ).toEqual({ ok: true, sheets: ["数据"] });
  });

  it("公式不读任何来源（=1+1）：放行且不需要证据", () => {
    expect(
      formulaSourceEvidenceRequirement(
        plan([{ type: "writeFormulas", sheet: "结果", formulas: [["=1+1"]] }])
      )
    ).toEqual({ ok: true, sheets: [] });
  });

  it("同表引用：拒绝，理由是「没法把本次写入与外部修改分开」", () => {
    const verdict = formulaSourceEvidenceRequirement(
      plan([{ type: "writeFormulas", sheet: "结果", formulas: [["=A1*2"]] }])
    );
    expect(verdict.ok).toBe(false);
    expect(verdict.ok ? "" : verdict.reason).toContain("写入目标表自身");
  });

  it("引用的表也是本次写入目标：拒绝", () => {
    const verdict = formulaSourceEvidenceRequirement(
      plan([
        { type: "writeFormulas", sheet: "结果", formulas: [["=明细!A1"]] },
        { type: "writeValues", sheet: "明细", range: "A1", values: [[1]] }
      ])
    );
    expect(verdict.ok).toBe(false);
    expect(verdict.ok ? "" : verdict.reason).toContain("也是本次写入目标");
  });

  it("动态 / 外部 / 名称：分别给出各自的拒绝理由", () => {
    const reason = (formulas: string[][]) =>
      formulaSourceEvidenceRequirement(
        plan([{ type: "writeFormulas", sheet: "结果", formulas }])
      );
    const dynamic = reason([["=INDIRECT(\"数据!A1\")"]]);
    expect(dynamic.ok ? "" : dynamic.reason).toContain("动态引用");
    const external = reason([["=[Book1.xlsx]Sheet1!A1"]]);
    expect(external.ok ? "" : external.reason).toContain("外部工作簿");
    const named = reason([["=税率*2"]]);
    expect(named.ok ? "" : named.reason).toContain("名称引用");
  });

  it("只看 writeFormulas：值动作不产生来源证据要求", () => {
    expect(
      formulaSourceEvidenceRequirement(
        plan([{ type: "writeValues", sheet: "结果", range: "A1", values: [[1]] }])
      )
    ).toEqual({ ok: true, sheets: [] });
  });
});

/**
 * 续跑前的来源证据判定 —— 产品规则（2026-10-30）：**登记的来源证据仍然一致**才允许续跑。
 *
 * 这组用例守三条，缺一条就退化成"看一眼公式文本就放行"：
 * ① 比对的是公式真正依赖的那几张表的**值**（不是授权范围的整张快照）；
 * ② 公式文本一字未变、来源值变了 ⇒ 必须拒绝（文本相同**不能**替代来源校验）；
 * ③ 没有证据 / 读不到 ⇒ 拒绝，不是"当作没变"。
 */
describe("续跑前的来源证据（值指纹）", () => {
  const plan = (actions: unknown[], fingerprint?: string) =>
    ({
      actions,
      ...(fingerprint === undefined ? {} : { sourceDependencyFingerprint: fingerprint })
    }) as never;

  const crossSheet = (fingerprint?: string) =>
    plan(
      [{ type: "writeFormulas", sheet: "结果", formulas: [["=数据!A1*2", "=汇总!B2"]] }],
      fingerprint
    );

  it("来源没变：放行，且读取的就是公式依赖的那两张表", async () => {
    const read: string[][] = [];
    const verdict = await verifyFormulaResumeSources(crossSheet("fnv:aa"), async (sheets) => {
      read.push([...sheets]);
      return "fnv:aa";
    });
    expect(verdict).toEqual({ verdict: "ok", sheets: ["数据", "汇总"] });
    // 读数**只**覆盖依赖表：授权范围里其它表不参与比对（否则无关改动会误拒）。
    expect(read).toEqual([["数据", "汇总"]]);
  });

  it("公式文本没变、来源值变了：拒绝，并给出「来源已变化」", async () => {
    const verdict = await verifyFormulaResumeSources(crossSheet("fnv:aa"), async () => "fnv:bb");
    expect(verdict).toEqual({ verdict: "refused", reason: FORMULA_SOURCE_CHANGED_MESSAGE });
  });

  it("没登记指纹：拒绝（不拿现读补一个基线）", async () => {
    let read = 0;
    const verdict = await verifyFormulaResumeSources(crossSheet(), async () => {
      read += 1;
      return "fnv:aa";
    });
    expect(verdict).toEqual({
      verdict: "refused",
      reason: FORMULA_SOURCE_EVIDENCE_MISSING_MESSAGE
    });
    // 没有可比对的基线就不该去读 —— 读了也只能"证明现在是什么"。
    expect(read).toBe(0);
  });

  it("读不出（抛错或空串）：拒绝，不说成「没变」", async () => {
    const thrown = await verifyFormulaResumeSources(crossSheet("fnv:aa"), async () => {
      throw new Error("宿主读不到");
    });
    expect(thrown).toEqual({ verdict: "refused", reason: FORMULA_SOURCE_UNVERIFIABLE_MESSAGE });
    const empty = await verifyFormulaResumeSources(crossSheet("fnv:aa"), async () => "");
    expect(empty).toEqual({ verdict: "refused", reason: FORMULA_SOURCE_UNVERIFIABLE_MESSAGE });
  });

  it("公式不读任何来源（=1+1）：不适用，且一次宿主读取都不发生", async () => {
    let read = 0;
    const verdict = await verifyFormulaResumeSources(
      plan([{ type: "writeFormulas", sheet: "结果", formulas: [["=1+1"]] }], "fnv:aa"),
      async () => {
        read += 1;
        return "fnv:aa";
      }
    );
    expect(verdict).toEqual({ verdict: "not-applicable" });
    expect(read).toBe(0);
  });

  it("同表引用（=A1*2）即使有指纹也拒绝：依赖面不可分离", async () => {
    const verdict = await verifyFormulaResumeSources(
      plan([{ type: "writeFormulas", sheet: "结果", formulas: [["=A1*2"]] }], "fnv:aa"),
      async () => "fnv:aa"
    );
    expect(verdict.verdict).toBe("refused");
    expect(verdict.verdict === "refused" ? verdict.reason : "").toContain("写入目标表自身");
  });
});
