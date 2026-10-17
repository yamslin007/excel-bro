import { describe, expect, it } from "vitest";
import type { FormatActionLike, FormatPropertyKey, FormatSnapshot } from "./formatActions";
import {
  FORMAT_IMPACT_UNRESOLVED_TEXT,
  FORMAT_VALUE_NOT_CAPTURED_TEXT,
  buildFormatChangeReview,
  borderDescription,
  formatActionEdits,
  formatActionImpact,
  formatChangeSummary,
  formatImpactNotice,
  formatPlanPreview,
  formatPlanPropertyGroups,
  formatValueText,
  isFormatActionType
} from "./formatActions";

function edits(action: FormatActionLike) {
  return formatActionEdits(action).map((edit) => `${edit.property}=${String(edit.value)}`);
}

function snapshot(
  properties: Partial<Record<FormatPropertyKey, string | number | boolean | null>>,
  captured = true
): FormatSnapshot {
  return { properties, captured };
}

describe("格式动作 · 只列显式声明要改的属性", () => {
  it("setFont 五项各自可选，缺省项不出现在改动清单里", () => {
    // `bold:false` 是"取消加粗"，缺省是"不要动" —— 两者都不能漏，也不能混。
    expect(edits({ type: "setFont", sheet: "S", range: "A1", bold: false })).toEqual([
      "fontBold=false"
    ]);
    expect(
      edits({ type: "setFont", sheet: "S", range: "A1", bold: true, italic: true })
    ).toEqual(["fontBold=true", "fontItalic=true"]);
    expect(
      edits({
        type: "setFont",
        sheet: "S",
        range: "A1",
        name: "微软雅黑",
        size: 14,
        color: "#112233"
      })
    ).toEqual(["fontName=微软雅黑", "fontSize=14", "fontColor=#112233"]);
    // 一个属性都没给 ⇒ 空清单，调用方看得见"这个动作什么也不做"。
    expect(edits({ type: "setFont", sheet: "S", range: "A1" })).toEqual([]);
  });

  it("setBorders 每条边各出一行，缺省线型/颜色/线宽按协议默认值补齐", () => {
    expect(
      edits({ type: "setBorders", sheet: "S", range: "A1:B2", sides: ["top", "bottom"] })
    ).toEqual([
      `borderTop=${borderDescription({})}`,
      `borderBottom=${borderDescription({})}`
    ]);
    expect(borderDescription({ style: "dash", color: "#FF0000", weight: "medium" })).toBe(
      "dash medium #FF0000"
    );
    // 未知边名不兜底成某条真实边：宁可少一行，也不能画用户没要求的边。
    expect(
      edits({ type: "setBorders", sheet: "S", range: "A1", sides: ["diagonal"] })
    ).toEqual([]);
  });

  it("setAlignment / resizeRange 同样只列显式项", () => {
    expect(
      edits({ type: "setAlignment", sheet: "S", range: "A1", wrapText: false })
    ).toEqual(["wrapText=false"]);
    expect(
      edits({ type: "setAlignment", sheet: "S", range: "A1", horizontal: "center" })
    ).toEqual(["alignmentHorizontal=center"]);
    expect(edits({ type: "setAlignment", sheet: "S", range: "A1" })).toEqual([]);
    expect(edits({ type: "resizeRange", sheet: "S", range: "A1", rowHeight: 24 })).toEqual([
      "rowHeight=24"
    ]);
    expect(edits({ type: "resizeRange", sheet: "S", range: "A1" })).toEqual([]);
  });

  it("autofit 会动行高与列宽两项，即使它没有固定值", () => {
    expect(edits({ type: "autofit", sheet: "S", range: "A1" })).toEqual([
      "rowHeight=null",
      "columnWidth=null"
    ]);
  });

  it("动作类型白名单认得本轮七种，别的一律不认", () => {
    for (const type of [
      "setFill",
      "setFont",
      "setNumberFormat",
      "setBorders",
      "setAlignment",
      "resizeRange",
      "autofit"
    ]) {
      expect(isFormatActionType(type)).toBe(true);
    }
    expect(isFormatActionType("writeValues")).toBe(false);
    expect(isFormatActionType("setFont ")).toBe(false);
  });
});

describe("格式动作 · 真实影响范围", () => {
  it("行高落在多行区域上会声明整行，而不是那一小块", () => {
    const impact = formatActionImpact({ type: "resizeRange", sheet: "S", range: "B3:D5", rowHeight: 24 });
    expect(impact).toEqual({ kind: "rows", sheet: "S", range: "B3:D5", rows: "3–5" });
    expect(formatImpactNotice(impact!)).toContain("整行");
    expect(formatImpactNotice(impact!)).toContain("3–5");
  });

  it("列宽落在多列区域上会声明整列", () => {
    const impact = formatActionImpact({
      type: "resizeRange",
      sheet: "S",
      range: "B3:D5",
      columnWidth: 18
    });
    expect(impact).toEqual({ kind: "columns", sheet: "S", range: "B3:D5", columns: "B–D" });
    expect(formatImpactNotice(impact!)).toContain("整列");
  });

  it("同时改行高与列宽才说两者都受影响", () => {
    const impact = formatActionImpact({
      type: "resizeRange",
      sheet: "S",
      range: "B3:D5",
      rowHeight: 24,
      columnWidth: 18
    });
    expect(impact?.kind).toBe("rowsAndColumns");
  });

  it("resizeRange 一个维度都没给时不说影响整行/整列", () => {
    // 这里是"计划坏掉了"的上游信号；影响范围退回区域本身，不能凭空声称整行。
    const impact = formatActionImpact({ type: "resizeRange", sheet: "S", range: "B3:D5" });
    expect(impact).toEqual({ kind: "range", sheet: "S", range: "B3:D5" });
  });

  it("单行/单列区域不写成 3–3", () => {
    expect(formatActionImpact({ type: "resizeRange", sheet: "S", range: "C7:D9", rowHeight: 20 })).toEqual({
      kind: "rows",
      sheet: "S",
      range: "C7:D9",
      rows: "7–9"
    });
    const single = formatActionImpact({ type: "autofit", sheet: "S", range: "C7" });
    expect(single).toEqual({
      kind: "rowsAndColumns",
      sheet: "S",
      range: "C7",
      rows: "7",
      columns: "C"
    });
  });

  it("区域不是明确 A1 时返回 null，调用方不得退化成『只影响这块』", () => {
    expect(formatActionImpact({ type: "resizeRange", sheet: "S", range: "整列", rowHeight: 20 })).toBeNull();
    expect(formatActionImpact({ type: "resizeRange", sheet: "S", range: "D9:B3", rowHeight: 20 })).toBeNull();
  });

  it("非行高列宽动作仍然只影响区域本身", () => {
    expect(formatActionImpact({ type: "setFill", sheet: "S", range: "A1:B2" })).toEqual({
      kind: "range",
      sheet: "S",
      range: "A1:B2"
    });
  });
});

describe("格式动作 · 属性值展示", () => {
  it("布尔显示是/否，行高字号带磅、列宽带字符宽", () => {
    expect(formatValueText("fontBold", true)).toBe("是");
    expect(formatValueText("wrapText", false)).toBe("否");
    expect(formatValueText("rowHeight", 24)).toBe("24 磅");
    expect(formatValueText("fontSize", 14)).toBe("14 磅");
    expect(formatValueText("columnWidth", 18)).toBe("18 字符宽");
    expect(formatValueText("numberFormat", "#,##0.00")).toBe("#,##0.00");
  });

  it("空字符串与 null 不是一回事", () => {
    expect(formatValueText("fontName", "")).toBe("（空字符串）");
    expect(formatValueText("rowHeight", null)).toBe("（按内容自适应）");
  });
});

describe("格式动作 · 修改前后对照", () => {
  const properties: FormatPropertyKey[] = ["fontBold", "fillColor", "rowHeight"];

  it("只比对计划声明要改的属性", () => {
    const review = buildFormatChangeReview({
      sheet: "S",
      range: "A1:B2",
      status: "applied",
      properties: ["fontBold"],
      before: snapshot({ fontBold: false, fillColor: "#ffffff", rowHeight: 20 }),
      after: snapshot({ fontBold: true, fillColor: "#ff0000", rowHeight: 40 })
    });
    // 填充色与行高虽然前后都不同，但这次没打算改它们，不许出现在对照里。
    expect(review.rows.map((row) => row.property)).toEqual(["fontBold"]);
    expect(review.changedCount).toBe(1);
  });

  it("两侧都采到且真不一样才算改动", () => {
    const review = buildFormatChangeReview({
      sheet: "S",
      range: "A1",
      status: "applied",
      properties,
      before: snapshot({ fontBold: false, fillColor: "#FFFFFF", rowHeight: 20 }),
      after: snapshot({ fontBold: true, fillColor: "#ffffff", rowHeight: 20 })
    });
    const changed = review.rows.filter((row) => row.changed).map((row) => row.property);
    // 颜色大小写不同不算改动 —— 宿主回读常改大小写，报成改动会让用户以为写没生效。
    expect(changed).toEqual(["fontBold"]);
  });

  it("任一侧未采集就不算改动，并如实标注", () => {
    const review = buildFormatChangeReview({
      sheet: "S",
      range: "A1",
      status: "executed_unverified",
      properties: ["fontName"],
      before: snapshot({ fontName: "Arial" }, false),
      after: snapshot({ fontName: "微软雅黑" })
    });
    expect(review.rows[0].beforeCaptured).toBe(false);
    expect(review.rows[0].beforeText).toBe(FORMAT_VALUE_NOT_CAPTURED_TEXT);
    expect(review.rows[0].changed).toBe(false);
    expect(review.hasUncaptured).toBe(true);
  });

  it("缺键也算未采集，不得被当成默认值参与比较", () => {
    const review = buildFormatChangeReview({
      sheet: "S",
      range: "A1",
      status: "applied",
      properties: ["fontItalic"],
      before: snapshot({}),
      after: snapshot({ fontItalic: false })
    });
    expect(review.rows[0].beforeCaptured).toBe(false);
    expect(review.rows[0].changed).toBe(false);
  });

  it("摘要不带任何属性值", () => {
    const review = buildFormatChangeReview({
      sheet: "S",
      range: "A1",
      status: "applied",
      properties: ["fontName"],
      before: snapshot({ fontName: "Arial" }),
      after: snapshot({ fontName: "微软雅黑" })
    });
    expect(formatChangeSummary(review)).toEqual({
      sheet: "S",
      range: "A1",
      status: "applied",
      changedCount: 1,
      propertyCount: 1,
      hasUncaptured: false
    });
    expect(JSON.stringify(formatChangeSummary(review))).not.toContain("微软雅黑");
  });
});

describe("格式动作 · 计划级属性归并", () => {
  it("同区域多个动作的属性合并，顺序稳定且不重复", () => {
    const groups = formatPlanPropertyGroups([
      { type: "setFont", sheet: "S", range: "A1:C1", bold: true, size: 14 },
      { type: "setFill", sheet: "S", range: "A1:C1", color: "#DDEEFF" },
      { type: "setAlignment", sheet: "S", range: "A1:C1", horizontal: "center" }
    ]);
    expect(groups).toHaveLength(1);
    expect(groups[0].properties).toEqual([
      "fontSize",
      "fontBold",
      "fillColor",
      "alignmentHorizontal"
    ]);
  });

  it("同区域影响范围取更宽的那个", () => {
    const groups = formatPlanPropertyGroups([
      { type: "setFill", sheet: "S", range: "A2:D2", color: "#DDEEFF" },
      { type: "resizeRange", sheet: "S", range: "A2:D2", rowHeight: 26 }
    ]);
    // 只要区域里有一个动作影响整行，整条区域的预览就必须说影响整行。
    expect(groups[0].impact?.kind).toBe("rows");
  });

  it("不同区域分成不同组", () => {
    const groups = formatPlanPropertyGroups([
      { type: "setFont", sheet: "S", range: "A1:C1", bold: true },
      { type: "setFont", sheet: "S", range: "A2:C9", size: 11 },
      { type: "setFont", sheet: "T", range: "A2:C9", size: 11 }
    ]);
    expect(groups.map((group) => `${group.sheet}!${group.range}`)).toEqual([
      "S!A1:C1",
      "S!A2:C9",
      "T!A2:C9"
    ]);
  });
});

describe("格式动作 · 确认前预览", () => {
  it("列区域内的格式改动，不重复念区域", () => {
    const preview = formatPlanPreview([
      { type: "setFont", sheet: "S", range: "A1:C1", bold: true, name: "微软雅黑" },
      { type: "setFill", sheet: "S", range: "A1:C1", color: "#DDEEFF" },
      { type: "setNumberFormat", sheet: "S", range: "B2:B9", formatCode: "#,##0.00" }
    ]);
    expect(preview).toEqual([
      {
        key: "S!A1:C1",
        sheet: "S",
        range: "A1:C1",
        changes: ["字体 → 微软雅黑", "加粗 → 是", "填充色 → #DDEEFF"],
        impactNotice: null
      },
      {
        key: "S!B2:B9",
        sheet: "S",
        range: "B2:B9",
        changes: ["数字格式 → #,##0.00"],
        impactNotice: null
      }
    ]);
  });

  it("行高/列宽必须带出真实影响范围", () => {
    const preview = formatPlanPreview([
      { type: "resizeRange", sheet: "S", range: "A1:C10", rowHeight: 24, columnWidth: 18 }
    ]);
    expect(preview[0].changes).toEqual(["行高 → 24 磅", "列宽 → 18 字符宽"]);
    expect(preview[0].impactNotice).toContain("整行");
    expect(preview[0].impactNotice).toContain("整列");
  });

  it("区域解析不出来时也给警告，不说『只影响这一小块』", () => {
    const preview = formatPlanPreview([
      { type: "resizeRange", sheet: "S", range: "整列", rowHeight: 24 }
    ]);
    expect(preview[0].impactNotice).toBe(FORMAT_IMPACT_UNRESOLVED_TEXT);
  });

  it("同区域同属性被两个动作声明时只留最后一条", () => {
    const preview = formatPlanPreview([
      { type: "setFont", sheet: "S", range: "A1", size: 12 },
      { type: "setFont", sheet: "S", range: "A1", size: 14 }
    ]);
    expect(preview[0].changes).toEqual(["字号 → 14 磅"]);
  });

  it("没有格式动作时返回空列表", () => {
    expect(formatPlanPreview([])).toEqual([]);
  });
});

describe("格式动作 · 验收场景（把选中区域整理成报表）", () => {
  /**
   * 用户给的通用任务：设置表头字体和填充、金额格式、边框、对齐和列宽，
   * 不改变数据与公式，先预览、确认后应用。
   *
   * 这里把它的**语义**钉成用例：改哪些属性、哪些动作会溢出区域。
   * 真实执行与回读在两端各自的端到端用例里，不在这里冒充。
   */
  const acceptancePlan: FormatActionLike[] = [
    {
      type: "setFont",
      sheet: "明细",
      range: "A1:D1",
      bold: true,
      name: "微软雅黑",
      size: 12
    },
    { type: "setFill", sheet: "明细", range: "A1:D1", color: "#DDEEFF" },
    {
      type: "setAlignment",
      sheet: "明细",
      range: "A1:D1",
      horizontal: "center",
      vertical: "center",
      wrapText: true
    },
    {
      type: "setBorders",
      sheet: "明细",
      range: "A1:D10",
      sides: [
        "top",
        "bottom",
        "left",
        "right",
        "insideHorizontal",
        "insideVertical"
      ],
      style: "continuous",
      color: "#000000",
      weight: "thin"
    },
    { type: "setNumberFormat", sheet: "明细", range: "C2:C10", formatCode: "#,##0.00" },
    { type: "resizeRange", sheet: "明细", range: "A1:D10", columnWidth: 16 }
  ];

  it("表头字体+填充+对齐只影响表头那一行，不声称影响整行整列", () => {
    const preview = formatPlanPreview(acceptancePlan);
    const header = preview.find((entry) => entry.range === "A1:D1")!;
    expect(header.changes).toEqual([
      "字体 → 微软雅黑",
      "字号 → 12 磅",
      "加粗 → 是",
      "填充色 → #DDEEFF",
      "水平对齐 → center",
      "垂直对齐 → center",
      "自动换行 → 是"
    ]);
    // 行高列宽没被碰过 ⇒ 不许冒出"影响整行/整列"。
    expect(header.impactNotice).toBeNull();
  });

  it("金额格式单独成组，只落在金额列上", () => {
    const preview = formatPlanPreview(acceptancePlan);
    const amount = preview.find((entry) => entry.range === "C2:C10")!;
    expect(amount.changes).toEqual(["数字格式 → #,##0.00"]);
    expect(amount.impactNotice).toBeNull();
  });

  it("列宽必须声明整列，且不得顺带声称行高也变了", () => {
    const preview = formatPlanPreview(acceptancePlan);
    const body = preview.find((entry) => entry.range === "A1:D10")!;
    expect(body.changes).toEqual([
      "上边框 → continuous thin #000000",
      "下边框 → continuous thin #000000",
      "左边框 → continuous thin #000000",
      "右边框 → continuous thin #000000",
      "内部横线 → continuous thin #000000",
      "内部竖线 → continuous thin #000000",
      "列宽 → 16 字符宽"
    ]);
    expect(body.impactNotice).toContain("整列");
    expect(body.impactNotice).not.toContain("整行");
  });

  it("预览是只读的：不改动计划对象本身", () => {
    // 预览若会写回计划（顺手补默认值、填空白），"预览不能提前修改原表"
    // 这条就只剩 UI 层面的承诺了 —— 数据模型早就被改过。
    const frozen = JSON.parse(JSON.stringify(acceptancePlan));
    const before = JSON.stringify(acceptancePlan);
    formatPlanPreview(acceptancePlan);
    formatPlanPropertyGroups(acceptancePlan);
    for (const action of acceptancePlan) formatActionEdits(action);
    expect(JSON.stringify(acceptancePlan)).toBe(before);
    expect(before).toBe(JSON.stringify(frozen));
  });
});
