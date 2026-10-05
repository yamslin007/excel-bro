import { describe, expect, it } from "vitest";

import { wpsClarificationVerdict } from "./wpsClarificationVerdict";

/**
 * 三种结果必须**互不冒充**：
 *
 * - `"same"`：能**正面证明**仍是原授权对象 → 继续回答。
 * - `"changed"`：有**正面证据**说明授权对象或表头变了 → 旧卡失效。
 * - `"unknown"`：**证明不了**（读到失败、或对象引用换了但无法区分
 *   "同一对象的新包装"与"真的换了对象"）→ 既不放行，也**不谎称"数据已变化"**。
 *
 * 之前的实现把后两种压成一个布尔值、统一报"数据范围已经变化" —— 那是假话：
 * 用户只是改了数据格又撤回，范围根本没变，只是我们证明不了。
 */

const identity = (combined: string) => combined;

function verdict(input: {
  identity?: { recorded: string | null; current: string | null };
  header?: { recorded: string | null; current: string | null };
  objectIdentity: "same" | "different" | "unknown" | null;
}) {
  return wpsClarificationVerdict({
    identity: input.identity ?? { recorded: identity("bk|1:S"), current: identity("bk|1:S") },
    header: input.header ?? { recorded: "h1", current: "h1" },
    objectIdentity: input.objectIdentity
  });
}

describe("澄清范围判定：same / changed / unknown", () => {
  it("数据格编辑后撤回：三条判据都能证明相同，必须放行继续回答", () => {
    expect(verdict({ objectIdentity: "same" })).toBe("same");
  });

  it("宿主对同一张表返回了新包装对象：证明不了相同，但不是 changed", () => {
    // 这是本次真机问题的核心：引用不相等只说明"证明不了"，不说明"变了"。
    expect(verdict({ objectIdentity: "unknown" })).toBe("unknown");
  });

  it("同名异路径工作簿：身份指纹有变化证据，判 changed", () => {
    expect(
      verdict({
        identity: { recorded: identity("bk|1:S"), current: identity("bk|1:S") },
        objectIdentity: "different"
      })
    ).toBe("changed");
  });

  it("同索引、同名、同表头重建的工作表：证明不了身份，不得谎称数据已变化", () => {
    expect(verdict({ objectIdentity: "unknown" })).toBe("unknown");
  });

  it("表头确实被改写（且读到了）：判 changed", () => {
    expect(verdict({ header: { recorded: "h1", current: "h2" }, objectIdentity: "same" })).toBe(
      "changed"
    );
  });

  it("表头读取失败：算 unknown，不算 changed —— 读不到不等于变了", () => {
    expect(verdict({ header: { recorded: "h1", current: null }, objectIdentity: "unknown" })).toBe(
      "unknown"
    );
    // 反过来：对象身份能被**正面证明**相同时，表头这次读不到只是这条附加判据
    // 没跑成，不构成"变化证据"，也不该因此把用户挡回去。
    expect(verdict({ header: { recorded: "h1", current: null }, objectIdentity: "same" })).toBe(
      "same"
    );
  });

  it("身份指纹读取失败：算 unknown，不算 changed", () => {
    expect(
      verdict({ identity: { recorded: identity("bk|1:S"), current: null }, objectIdentity: "unknown" })
    ).toBe("unknown");
  });

  it("面板重开后拿不到内存对象身份：算 unknown", () => {
    expect(verdict({ objectIdentity: null })).toBe("unknown");
  });

  it("变化证据优先于对象身份相同：不得因为引用相等就放行一张换了表头的表", () => {
    expect(verdict({ header: { recorded: "h1", current: "h2" }, objectIdentity: "same" })).toBe(
      "changed"
    );
  });

  it("旧卡没有指纹（无法比对）时不能算相同，也不能算变化", () => {
    expect(
      verdict({
        identity: { recorded: null, current: identity("bk|1:S") },
        header: { recorded: null, current: "h1" },
        objectIdentity: "unknown"
      })
    ).toBe("unknown");
  });

  it("只有名称/表头相同不足以放行", () => {
    // 指纹一致 + 表头一致，但对象身份证明不了 → 仍然不得放行。
    expect(verdict({ objectIdentity: "unknown" })).not.toBe("same");
  });
});
