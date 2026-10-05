import { describe, expect, it } from "vitest";
import type { ColumnMappingRule } from "./contracts";
import {
  addExactMapping,
  addSplitMapping,
  suggestMappingDelimiter,
  uncoveredSplitItems
} from "./localMappingResolution";
import { buildColumnMappingActions } from "@excel-bro/core";
import { normalizeStoredMessage } from "./conversation";
import type { ChatMessage } from "./types/chat";

const rule: ColumnMappingRule = {
  kind: "columnMapping",
  sourceSheet: "样本",
  targetSheet: "样本",
  sourceRange: "C2:C3",
  targetRange: "D2:D3",
  mapping: [
    { source: "甲", target: "低" },
    { source: "乙", target: "高" }
  ],
  blankLabel: "空"
};

describe("本地映射结构化补全（Excel/WPS 共用）", () => {
  it("重开时使仅存于内存的规则草稿失效，不留一张可误点的澄清卡", () => {
    const pending: ChatMessage = {
      id: "local-1",
      role: "assistant",
      text: "",
      createdAt: "2026-10-01T00:00:00.000Z",
      clarification: {
        id: "gap-1",
        localMapping: true,
        summary: "本地映射需要你决定",
        question: "有未覆盖的来源值",
        reason: "源值只在本机展示",
        scopeLabel: "样本!C2:C3",
        options: [],
        originalPrompt: "按类别映射",
        scopeFingerprint: "scope",
        round: 0,
        status: "pending"
      }
    };
    expect(normalizeStoredMessage(pending).clarification?.status).toBe("invalidated");
    expect(JSON.stringify(pending)).not.toContain("甲，乙");
  });

  it("只把用户明确确认的整格映射补进草稿，不把复合值猜成空白", () => {
    const patched = addExactMapping(rule, ["甲，乙"], ["高"]);
    expect(patched.mapping.at(-1)).toEqual({ source: "甲，乙", target: "高" });
    expect(rule.mapping).toHaveLength(2);
    expect(() => addExactMapping(rule, ["甲，乙"], [""])).toThrow();
    expect(() => addExactMapping(rule, ["甲，乙"], ["=1+1"])).toThrow();
  });

  it("候选分隔符只供确认；拆分后仍须显式声明优先顺序", () => {
    expect(suggestMappingDelimiter(rule, ["甲，乙"])).toBe("，");
    expect(uncoveredSplitItems(rule, ["甲，乙"], "，")).toEqual([]);
    expect(() => addSplitMapping(rule, {
      delimiter: "，", missingItems: [], targets: [], priorityTargets: []
    })).toThrow(/排序/);
    const patched = addSplitMapping(rule, {
      delimiter: "，", missingItems: [], targets: [], priorityTargets: ["高", "低"]
    });
    expect(patched.delimiters).toEqual(["，"]);
    expect(patched.valueRank).toEqual([
      { value: "甲", rank: 1 }, { value: "乙", rank: 2 }
    ]);
    const actions = buildColumnMappingActions({
      sourceSheet: "样本", targetSheet: "样本", sourceRange: "C2:C3", targetRange: "D2:D3",
      sourceValues: [["甲，乙"], ["甲"]],
      mapping: Object.fromEntries(patched.mapping.map((pair) => [pair.source, pair.target])),
      blankLabel: patched.blankLabel,
      delimiters: patched.delimiters,
      valueRank: patched.valueRank
    });
    expect(actions[0]).toMatchObject({ type: "writeValues", values: [["高"], ["低"]] });
  });

  it("新增拆分项后仍须覆盖全部目标并确认高低，不靠映射声明顺序", () => {
    expect(uncoveredSplitItems(rule, ["甲；丙"], "；")).toEqual(["丙"]);
    const patched = addSplitMapping(rule, {
      delimiter: "；", missingItems: ["丙"], targets: ["中"],
      priorityTargets: ["高", "中", "低"]
    });
    expect(patched.mapping.at(-1)).toEqual({ source: "丙", target: "中" });
    expect(patched.valueRank?.map((entry) => entry.rank)).toEqual([1, 3, 2]);
    expect(() => addSplitMapping(rule, {
      delimiter: "；", missingItems: ["丙"], targets: ["中"], priorityTargets: ["高", "低"]
    })).toThrow(/遗漏或重复/);
  });
});
