import { describe, expect, it } from "vitest";
import { messageRowPresentation, type MessageRowLike } from "./messageRows";

/** 单步消息理解的过程日志（`isCompactMessageProcess` 判为紧凑的那种）。 */
const compactProcess = {
  outcome: "completed",
  steps: [{ id: "message-intent", status: "completed", phase: "intent" }]
};

/** 三步执行的过程日志（不是紧凑的那种）。 */
const multiStepProcess = {
  outcome: "completed",
  steps: [
    { id: "writing:0", status: "completed", phase: "write" },
    { id: "writing:1", status: "completed", phase: "write" },
    { id: "writing:2", status: "completed", phase: "write" }
  ]
};

const user = (text: string): MessageRowLike => ({ role: "user", text });
const assistant = (patch: Partial<MessageRowLike> = {}): MessageRowLike => ({
  role: "assistant",
  ...patch
});

const authors = (rows: readonly MessageRowLike[]) =>
  messageRowPresentation(rows).map((row) => row.author);
const starts = (rows: readonly MessageRowLike[]) =>
  messageRowPresentation(rows).map((row) => row.groupStart);

describe("一段回复只署一次名", () => {
  it("过程行 + 结论行（同一出处）：只有过程行署名", () => {
    expect(
      authors([
        user("把这些公式写进去"),
        assistant({ processLog: multiStepProcess }),
        assistant({ text: "已执行 3 步并通过独立验证。" })
      ])
    ).toEqual([false, true, false]);
  });

  it("组首没有徽标、后续行才标出「本地工具」：不再印第二次", () => {
    // 组首那颗「Excel Bro」已经把这一段认下了；再印一次是两颗一模一样的抬头。
    expect(
      authors([
        user("把这些公式写进去"),
        assistant({ processLog: multiStepProcess }),
        assistant({ text: "已执行 3 步并通过独立验证。" }),
        assistant({ text: "共 12 行", badge: "本地工具" })
      ])
    ).toEqual([false, true, false, false]);
  });

  it("组首有徽标、后续行没有：也不重印", () => {
    expect(
      authors([
        user("算一下"),
        assistant({ processLog: multiStepProcess, badge: "Agent" }),
        assistant({ text: "结论" })
      ])
    ).toEqual([false, true, false]);
  });

  it("同一组的行不是组首行，组与组之间的行才是", () => {
    expect(
      starts([
        user("第一问"),
        assistant({ processLog: multiStepProcess }),
        assistant({ text: "结论" }),
        user("第二问"),
        assistant({ text: "另一段结论" })
      ])
    ).toEqual([true, true, false, true, true]);
  });
});

describe("不变量 A：换了来源徽标就重新署一次", () => {
  it("过程行是 Agent、下一行是需求确认 ⇒ 两行都署名", () => {
    // 反向锚点：把 `badge !== groupBadge` 去掉，「需求确认」这个抬头就没了，
    // 用户看到一张没有出处的问卷。
    expect(
      authors([
        user("算一下"),
        assistant({ processLog: multiStepProcess, badge: "Agent" }),
        assistant({ text: "怎样分级？", badge: "需求确认" })
      ])
    ).toEqual([false, true, true]);
  });

  it("徽标相同不算换事，仍然只署一次", () => {
    expect(
      authors([
        user("查一下"),
        assistant({ processLog: multiStepProcess, badge: "Agent" }),
        assistant({ text: "共 12 行", badge: "Agent" })
      ])
    ).toEqual([false, true, false]);
  });
});

describe("不变量 B：单步消息理解把名字让给结论行", () => {
  it("过程行没有内容、下一行是结论 ⇒ 名字落在结论行上", () => {
    expect(
      authors([
        user("谢谢帮助"),
        assistant({ text: "", processLog: compactProcess }),
        assistant({ text: "不客气，随时告诉我。" })
      ])
    ).toEqual([false, false, true]);
  });

  it("多步过程行不适用这条让位（它自己有内容要看）", () => {
    expect(
      authors([
        user("分析当前表"),
        assistant({ text: "", processLog: multiStepProcess }),
        assistant({ text: "结论" })
      ])
    ).toEqual([false, true, false]);
  });

  it("让位的目标行被跳过时不把名字让给看不见的一行", () => {
    expect(
      authors([
        user("谢谢帮助"),
        assistant({ text: "", processLog: compactProcess }),
        assistant({ text: "不客气，随时告诉我。", skipped: true }),
        assistant({ text: "还有什么要做的？" })
      ])
    ).toEqual([false, false, false, true]);
  });
});

describe("不变量 C：换回复段必须重新署名", () => {
  it("继续开新段，两段之间没有用户消息也不合并", () => {
    // 反向锚点：只按"上一条是不是助手"分组（不认段身份）时，这条会得到
    // [false, true, false]（新段第一行被当成续行），用户看到一段凭空冒出来的进度。
    expect(
      authors([
        user("跑三步公式"),
        assistant({ processLog: multiStepProcess, segmentId: "seg-a" }),
        assistant({ processLog: multiStepProcess, segmentId: "seg-b" })
      ])
    ).toEqual([false, true, true]);
  });

  it("同段的后续行仍是续行", () => {
    expect(
      authors([
        user("跑三步公式"),
        assistant({ processLog: multiStepProcess, segmentId: "seg-a" }),
        assistant({ text: "结论", segmentId: "seg-a" })
      ])
    ).toEqual([false, true, false]);
  });

  it("老消息没有段身份时按同一组处理，不把历史拆碎", () => {
    expect(
      authors([
        user("老对话"),
        assistant({ processLog: multiStepProcess }),
        assistant({ text: "老结论" })
      ])
    ).toEqual([false, true, false]);
  });
});

describe("分组边界", () => {
  it("用户消息永远不署名，但一定是组首行", () => {
    expect(authors([user("你好"), user("再问一句")])).toEqual([false, false]);
    expect(starts([user("你好"), assistant({ text: "在" })])).toEqual([true, true]);
  });

  it("跳过的行不参与分组，也不占作者名", () => {
    const rows: MessageRowLike[] = [
      user("算一下"),
      assistant({ text: "先看范围", skipped: true }),
      assistant({ text: "范围已确认" })
    ];
    expect(authors(rows)).toEqual([false, false, true]);
    expect(starts(rows)).toEqual([true, false, true]);
  });

  it("空数组不炸，长度与入参一一对应", () => {
    expect(messageRowPresentation([])).toEqual([]);
    expect(authors([user("你好")])).toHaveLength(1);
  });
});
