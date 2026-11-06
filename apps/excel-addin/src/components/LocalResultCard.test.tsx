// @vitest-environment jsdom
/**
 * 本地结果卡。
 *
 * 关键语义：卡片拿到的**只有指针**（引用 + 结构，没有一行值）；值必须由本机
 * 回环端点按引用取回。所以要证的不是"能渲染表格"，而是：
 * - 指针上确实不带值（隐私面）；
 * - 取回失败时**如实**说不可用，而不是渲染一张空表冒充成功；
 * - `load` 每渲染换一次身份也不会陷入"取一次 → 再渲染 → 再取"的循环。
 */

import { act, cleanup, render, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";

import type { LocalResultCard as LocalResultPointer, SdkBridgeResultResponse } from "../contracts";
import { LocalResultCard } from "./LocalResultCard";
import { executeQueryTableData } from "../dataTools";

afterEach(cleanup);

const REF = "res_" + "a".repeat(32);

function pointer(overrides: Partial<LocalResultPointer> = {}): LocalResultPointer {
  return {
    resultRef: REF,
    title: "按类别汇总",
    headers: ["类别", "数量"],
    rowCount: 2,
    columnCount: 2,
    sourceSheets: ["Sheet1", "Sheet2"],
    complete: true,
    ...overrides
  };
}

function payload(overrides: Partial<SdkBridgeResultResponse> = {}): SdkBridgeResultResponse {
  return {
    resultRef: REF,
    title: "按类别汇总",
    headers: ["类别", "数量"],
    rows: [
      ["甲", 2],
      ["乙", 3]
    ],
    sourceSheets: ["Sheet1", "Sheet2"],
    warnings: [],
    complete: true,
    calculation: "按类别计数",
    scannedRows: 9,
    ...overrides
  };
}

describe("LocalResultCard", () => {
  it("共享结果卡展示真实执行器计算的两表总数", async () => {
    const result = executeQueryTableData({
      id: "total-card", tool: "query_table", arguments: {
        mode: "aggregate", groupBy: ["类别"], combine: { mode: "union" },
        metrics: [{ operation: "countRows", outputName: "数量" }]
      }
    }, [
      { name: "Sheet1", values: [["类别"], ["甲"], ["乙"]] },
      { name: "Sheet2", values: [["类别"], ["甲"], ["乙"], ["乙"]] }
    ]);
    const ui = render(<LocalResultCard card={pointer()} load={async () => payload(result)} />);
    await waitFor(() => expect(ui.getByText(/筛选后总记录数：5/)).toBeTruthy());
  });
  it("指针本身不含任何值，只展示结构与来源", () => {
    const load = vi.fn(() => new Promise<SdkBridgeResultResponse | null>(() => {}));
    const ui = render(<LocalResultCard card={pointer()} load={load} />);
    const text = ui.container.textContent ?? "";
    expect(text).toContain("按类别汇总");
    expect(text).toContain("2 行 × 2 列");
    expect(text).toContain("Sheet1、Sheet2");
    // 指针上唯一的"值"是表头；数据行一个都没出现。
    expect(text).not.toContain("甲");
  });

  it("挂载即按引用取值，取回后渲染真实表格", async () => {
    const load = vi.fn(async () => payload());
    const ui = render(<LocalResultCard card={pointer()} load={load} />);
    await waitFor(() => expect(ui.getByRole("table")).toBeTruthy());
    expect(load).toHaveBeenCalledWith(REF);
    const cells = Array.from(ui.getByRole("table").querySelectorAll("td")).map(
      (cell) => cell.textContent
    );
    expect(cells).toEqual(["甲", "2", "乙", "3"]);
    expect(ui.container.textContent).toContain("按类别计数");
  });

  it("取不到值时如实说不可用，不渲染空表冒充成功", async () => {
    const load = vi.fn(async () => null);
    const ui = render(<LocalResultCard card={pointer()} load={load} />);
    await waitFor(() => expect(ui.getByText(/结果已不可用/)).toBeTruthy());
    expect(ui.queryByRole("table")).toBeNull();
  });

  it("取值抛异常同样落到不可用（断线不把界面打崩）", async () => {
    const load = vi.fn(async () => {
      throw new Error("network down");
    });
    const ui = render(<LocalResultCard card={pointer()} load={load} />);
    await waitFor(() => expect(ui.getByText(/结果已不可用/)).toBeTruthy());
    expect(ui.container.textContent).not.toContain("network down");
  });

  /*
   * 提示语必须**按真实原因**写。
   *
   * 2026-11-01 真机：用户只要求查「来源」表，结果卡却列出四张来源表、三张各报
   * 「缺少列2」，提示语却是「结果不完整（扫描已达上限）」—— 三种完全不同的原因
   * 共用一句最吓人的话，用户自然当成读表读崩了。下面三条分别钉住三种事实。
   */
  it("缺字段：说是哪张表缺哪一列，不说成扫描达上限", async () => {
    const ui = render(
      <LocalResultCard
        card={pointer({
          complete: false,
          sourceSheets: ["来源", "结果1"],
          incomplete: [
            { kind: "missing_field", sheet: "结果1", fields: ["列2"] }
          ]
        })}
        load={vi.fn(async () => payload({ complete: false }))}
      />
    );
    await waitFor(() =>
      expect(ui.getByText(/「结果1」没有字段：列2/)).toBeTruthy()
    );
    expect(ui.container.textContent).not.toContain("扫描");
  });

  it("展示截断：只说行数，不说成结果不完整", async () => {
    const ui = render(
      <LocalResultCard
        card={pointer({
          truncated: true,
          truncation: { shownRows: 20, totalRows: 137, limit: 20 }
        })}
        load={vi.fn(async () => payload())}
      />
    );
    await waitFor(() =>
      expect(ui.getByText(/仅显示前 20 行，共 137 行/)).toBeTruthy()
    );
    expect(ui.container.textContent).not.toContain("结果不完整");
  });

  it("说了不完整却没有任何原因：如实说没记录到，不替它编一个", async () => {
    const ui = render(
      <LocalResultCard
        card={pointer({ complete: false })}
        load={vi.fn(async () => payload({ complete: false }))}
      />
    );
    await waitFor(() =>
      expect(ui.getByText(/没有记录到具体原因/)).toBeTruthy()
    );
    expect(ui.container.textContent).not.toContain("扫描已达上限");
  });

  /*
   * 告警不能静默丢弃。
   *
   * "有 2 条记录未匹配""存在重复键，已按 first 取值"这类信息决定了用户能否
   * 信任这份结果。它们只存在于 `warnings` 里 —— 界面不渲染，用户就只看到一张
   * 看起来正常的表，误以为数据是完整的。
   */
  it("结果卡渲染告警，不静默丢弃未匹配明细与重复键提示", async () => {
    const warnings = [
      "有 2 条记录在字典表中未匹配，已列入未匹配明细。",
      "检测到重复键「甲」，已按 first 取值。"
    ];
    const ui = render(
      <LocalResultCard
        card={pointer()}
        load={vi.fn(async () => payload({ warnings }))}
      />
    );
    await waitFor(() => expect(ui.getByText(/未匹配明细/)).toBeTruthy());
    expect(ui.getByText(/重复键/)).toBeTruthy();
    // 告警必须以可被辅助技术识别的列表呈现，而不是一条拼接文本。
    expect(ui.container.querySelectorAll(".local-result-card-warnings li")).toHaveLength(2);
  });

  it("没有告警时不渲染空告警区（不制造噪声）", async () => {
    const ui = render(
      <LocalResultCard card={pointer()} load={vi.fn(async () => payload())} />
    );
    await waitFor(() => expect(ui.getByRole("table")).toBeTruthy());
    expect(ui.container.querySelector(".local-result-card-warnings")).toBeNull();
  });

  it("`load` 每渲染换身份也只取一次（不会取值循环）", async () => {
    const calls: string[] = [];
    const makeLoader = () => async (resultRef: string) => {
      calls.push(resultRef);
      return payload();
    };
    const ui = render(<LocalResultCard card={pointer()} load={makeLoader()} />);
    await waitFor(() => expect(ui.getByRole("table")).toBeTruthy());
    ui.rerender(<LocalResultCard card={pointer()} load={makeLoader()} />);
    ui.rerender(<LocalResultCard card={pointer()} load={makeLoader()} />);
    await act(async () => {
      await Promise.resolve();
    });
    expect(calls).toEqual([REF]);
  });
});
