/**
 * 列映射的**语义边界**锁（与共享夹具互补，不重复它）。
 *
 * 共享夹具 `config/column_mapping_cases.json` 钉的是"两份镜像实现是否逐格一致"。
 * 本文件钉的是另一件事：**哪些输入必须拒绝、必须澄清，而不是被静默兜底**。
 * 这两类断言不能塞进共享夹具，因为夹具只能表达"期望值 / 期望错误"两态，
 * 而这里要钉的是**入口处的判断顺序**（先辨空白、再辨显式、再辨未知）与
 * **不允许出现的兜底结果**。
 *
 * 三条被钉住的语义：
 *
 * 1. **空白 ≠ 未知类别**。空白是"源本来没有值"，走 blankLabel；
 *    未知是"源有值但映射表没覆盖"，必须拒绝。两者不得互相冒充 ——
 *    若实现把未知值兜成 blankLabel，用户会在单元格层面看到"未分类"，
 *    而真实缺口被吞掉。
 * 2. **"无问题" 这类显式映射项不是空白**。它出现在 mapping 里就说明
 *    调用方已经为它指定了目标（比如 "/"），这与"真空白写什么"是两个问题。
 *    把显式项当空白、或把空白当显式项，都会改变输出。
 * 3. **显式空串映射项优先于空白策略**，且只对**空串**生效，
 *    不波及 null 与纯空白字符串 —— 三者在"是不是空白"上一致，
 *    但在"有没有被显式覆盖"上不同。
 *
 * 文件名刻意**不含**业务词：`test_no_business_hardcoding.py` 会扫生产代码，
 * 而这里出现的取样值（列名、标签）属于测试数据，不构成对实现的约束。
 */

import { describe, expect, it } from "vitest";

import {
  buildColumnMappingActions,
  NeedsClarificationError,
  UnmappedCategoryError,
  type ColumnMappingAction,
  type ColumnMappingInput,
  type ColumnMappingPlanAction
} from "./columnMapping";

/**
 * 取第一段写值动作的值矩阵。
 *
 * `ColumnMappingPlanAction` 是 `writeValues | clearRange` 的联合，而 `clearRange`
 * 只清内容、不带 `values`。所以"断言矩阵"必须按 `type` 先收窄 —— 直接 `actions[0].values`
 * 在类型层就不成立。清空动作有没有写对，由断言**整条动作**（`type`/`range`/`applyTo`）
 * 的用例负责，不靠从它身上读一个不存在的字段。
 */
function writeMatrix(actions: readonly ColumnMappingPlanAction[]): string[][] {
  const action = actions.find(
    (candidate): candidate is ColumnMappingAction =>
      candidate.type === "writeValues"
  );
  if (!action) throw new Error("预期至少有一个 writeValues 动作，实际一个都没有");
  return action.values;
}

/** 只关心"不许出现什么"，输入刻意保持与业务无关的通用标签。 */
function run(
  sourceValues: (string | number | boolean | null)[][],
  mapping: Record<string, string>,
  blankLabel: string,
  /** 多值语义参数，按平铺字段透传（与协议、夹具同形）。 */
  extra: Partial<
    Pick<
      ColumnMappingInput,
      "delimiters" | "valueRank" | "undeclaredSeparators"
    >
  > = {}
) {
  return buildColumnMappingActions({
    sourceSheet: "表一",
    targetSheet: "表一",
    sourceRange: `A1:A${sourceValues.length}`,
    targetRange: `C1:C${sourceValues.length}`,
    sourceValues,
    mapping,
    blankLabel,
    ...extra
  });
}

describe("空白与未知类别必须分开处理", () => {
  it("真空白走 blankLabel，不抛错", () => {
    const actions = run([[null]], { 甲: "一" }, "空");
    expect(writeMatrix(actions)).toEqual([["空"]]);
  });

  it("无值但有映射表仍未覆盖的值：拒绝，而不是兜成 blankLabel", () => {
    // 若实现把"未命中"兜成 blankLabel，这条会静默通过 —— 正是要钉住的失败模式。
    expect(() => run([["乙"]], { 甲: "一" }, "空")).toThrow(
      UnmappedCategoryError
    );
  });

  it("空白与未知混在同一列：报错时只列未知值，不把空白也算进缺口", () => {
    let thrown: unknown;
    try {
      run([[null], [" "], ["乙"]], { 甲: "一" }, "空");
    } catch (error) {
      thrown = error;
    }
    expect(thrown).toBeInstanceOf(UnmappedCategoryError);
    // 空格是空白（`isBlank` 对纯空白字符串为真），不应被当成需要补全的类别。
    expect((thrown as UnmappedCategoryError).values).toEqual(["乙"]);
  });
});

describe("显式映射项与空白策略的边界", () => {
  it("映射表里写了非空标签的显式项，就按它输出，不因它像空白而降级", () => {
    // "无问题" 是**非空**来源值；实现不得因为它的语义像"没有内容"而改走空白。
    const actions = run([["无问题"]], { 无问题: "/" }, "未填");
    expect(writeMatrix(actions)).toEqual([["/"]]);
  });

  it("真空白在有显式项存在时仍走 blankLabel，不借用显式项的目标", () => {
    const actions = run([[null], ["无问题"]], { 无问题: "/" }, "未填");
    expect(writeMatrix(actions)).toEqual([["未填"], ["/"]]);
  });

  it("显式空串映射项只覆盖空串，不覆盖 null 与纯空白字符串", () => {
    const actions = run([[""], [null], ["  "]], { "": "空串" }, "真空格");
    expect(writeMatrix(actions)).toEqual([["空串"], ["真空格"], ["真空格"]]);
  });

  it("映射表外的显式项不会被兜底：多个未知值一次列全", () => {
    let thrown: unknown;
    try {
      run([["丙"], ["丁"], ["丙"]], { 甲: "一" }, "空");
    } catch (error) {
      thrown = error;
    }
    expect(thrown).toBeInstanceOf(UnmappedCategoryError);
    expect((thrown as UnmappedCategoryError).values).toEqual(["丙", "丁"]);
  });
});

describe("标签值本身是空白符号时不被误解为空值", () => {
  it("目标标签为符号（如 / 或 —）时原样写入", () => {
    const actions = run([["甲"], ["乙"]], { 甲: "/", 乙: "—" }, "空");
    expect(writeMatrix(actions)).toEqual([["/"], ["—"]]);
  });
});

/**
 * 多值部分的语义边界。核心是**两类归因不得互相冒充**：
 *
 * - 已按声明的分隔符拆开、拆出的项有未覆盖 → `UnmappedCategoryError`，
 *   补救动作只有「补类别映射」一条；
 * - 整格没命中、又像是用了别的分隔符写法 → `NeedsClarificationError`，
 *   补救动作有两个可能（补分隔符 / 补映射），必须让用户选。
 *
 * 把两者合成一个错误，会让"该让用户选"被误当成"补映射表就行"，
 * 用户会去改一个没问题的设置，而真正的缺口一直留着。
 */
describe("多值拆分后的两类归因必须分开", () => {
  it("已声明分隔符对上了：归因是类别缺口，缺失项按拆开后的项报", () => {
    let thrown: unknown;
    try {
      run(
        [["甲 丙"]],
        { 甲: "壹" },
        "空",
        { delimiters: [" "], valueRank: [{ value: "甲", rank: 1 }] }
      );
    } catch (error) {
      thrown = error;
    }
    expect(thrown).toBeInstanceOf(UnmappedCategoryError);
    // 用户要补的是「丙」，不是整格原文「甲 丙」。
    expect((thrown as UnmappedCategoryError).values).toEqual(["丙"]);
  });

  it("拆开的项全部未覆盖时，归因仍是类别缺口且列全", () => {
    let thrown: unknown;
    try {
      run(
        [["丙 丁"]],
        { 甲: "壹" },
        "空",
        { delimiters: [" "], valueRank: [{ value: "甲", rank: 1 }] }
      );
    } catch (error) {
      thrown = error;
    }
    expect(thrown).toBeInstanceOf(UnmappedCategoryError);
    expect((thrown as UnmappedCategoryError).values).toEqual(["丙", "丁"]);
  });

  it("整格含未声明的分隔符写法：归因不确定，要用户选", () => {
    let thrown: unknown;
    try {
      run(
        [["甲/乙"], ["甲"]],
        { 甲: "壹", 乙: "贰" },
        "空",
        {
          delimiters: [" "],
          valueRank: [
            { value: "甲", rank: 1 },
            { value: "乙", rank: 0 }
          ],
          undeclaredSeparators: ["/"]
        }
      );
    } catch (error) {
      thrown = error;
    }
    expect(thrown).toBeInstanceOf(NeedsClarificationError);
    expect((thrown as NeedsClarificationError).cells).toEqual(["甲/乙"]);
  });

  it("看不出任何分隔形式时归因确定：就是缺这个类别", () => {
    let thrown: unknown;
    try {
      run(
        [["丙"]],
        { 甲: "壹" },
        "空",
        { delimiters: [" "], valueRank: [{ value: "甲", rank: 1 }] }
      );
    } catch (error) {
      thrown = error;
    }
    // 与上一条成对：同一入口只差"格子里有没有像分隔符的写法"，
    // 归因却不同，输出也必须不同。
    expect(thrown).toBeInstanceOf(UnmappedCategoryError);
    expect((thrown as UnmappedCategoryError).values).toEqual(["丙"]);
  });

  it("整格能精确命中时就不拆：整格精确优先于拆分", () => {
    const actions = run(
      [["甲 乙"]],
      { "甲 乙": "合", 甲: "壹", 乙: "贰" },
      "空",
      {
        delimiters: [" "],
        valueRank: [
          { value: "甲", rank: 1 },
          { value: "乙", rank: 0 }
        ]
      }
    );
    expect(writeMatrix(actions)).toEqual([["合"]]);
  });

  it("拆分后按规则给出的等级取最高级", () => {
    const actions = run(
      [["乙 甲"], ["甲"]],
      { 甲: "壹", 乙: "贰" },
      "空",
      {
        delimiters: [" "],
        valueRank: [
          { value: "甲", rank: 1 },
          { value: "乙", rank: 0 }
        ]
      }
    );
    expect(writeMatrix(actions)).toEqual([["壹"], ["壹"]]);
  });

  it("多候选并存时按出现位置切分，不按声明序", () => {
    const actions = run(
      [["甲-乙|丙"]],
      { 甲: "壹", "乙|丙": "合" },
      "空",
      { delimiters: ["|", "-"], valueRank: [{ value: "甲", rank: 1 }] }
    );
    expect(writeMatrix(actions)).toEqual([["壹"]]);
  });
});

describe("澄清内容不得把源值带回调用方", () => {
  it("归因不确定时，错误文案里不含整格原文", () => {
    // 这个 message 会经 `ColumnMappingClarification.message` 上到界面/协议；
    // 带上原文就等于把本端数据顺着"澄清"这条通道送出去。要展示原文的
    // 界面本来就能读自己的源值，不需要（也不该）由实现代转。
    let thrown: unknown;
    try {
      run([["甲/乙"]], { 甲: "一" }, "空", {
        delimiters: [" "],
        undeclaredSeparators: ["/"],
        valueRank: [{ value: "甲", rank: 1 }]
      });
    } catch (error) {
      thrown = error;
    }
    expect(thrown).toBeInstanceOf(NeedsClarificationError);
    expect((thrown as Error).message).not.toContain("甲/乙");
    // 但"有几格处理不了"这个计数必须留在结构化字段里。
    expect((thrown as NeedsClarificationError).cells).toEqual(["甲/乙"]);
  });

  it("类别缺口报的是拆开的项，不是整格原文", () => {
    let thrown: unknown;
    try {
      run([["丙 丁"]], { 甲: "一" }, "空", {
        delimiters: [" "],
        valueRank: [{ value: "甲", rank: 1 }]
      });
    } catch (error) {
      thrown = error;
    }
    expect(thrown).toBeInstanceOf(UnmappedCategoryError);
    // 用户要补的是"丙""丁"两条映射，给他"丙 丁"等于让他自己去拆。
    expect((thrown as UnmappedCategoryError).values).toEqual(["丙", "丁"]);
  });
});
