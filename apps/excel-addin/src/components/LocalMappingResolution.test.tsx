// @vitest-environment jsdom
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { ColumnMappingRule } from "../contracts";
import { LocalMappingResolution } from "./LocalMappingResolution";

afterEach(cleanup);

const rule: ColumnMappingRule = {
  kind: "columnMapping",
  sourceSheet: "样本",
  targetSheet: "样本",
  sourceRange: "C2:C2",
  targetRange: "D2:D2",
  mapping: [
    { source: "甲", target: "低" },
    { source: "乙", target: "高" }
  ],
  blankLabel: "空"
};

describe("本地映射补全卡 · 两宿主共用", () => {
  it("切到拆分规则仍留在卡上；未显式确认优先级不能生成预览", () => {
    const onResolve = vi.fn().mockResolvedValue(undefined);
    const onCancel = vi.fn();
    render(<LocalMappingResolution
      rule={rule} values={["甲，乙"]} allowExactMapping busy={false}
      onResolve={onResolve} onCancel={onCancel}
    />);
    fireEvent.click(screen.getByRole("button", { name: "改为说明拆分规则" }));
    expect(onCancel).not.toHaveBeenCalled();
    expect(screen.getByLabelText("拆分符")).toHaveProperty("value", "，");
    expect(screen.getByRole("button", { name: "按补充映射重新生成预览" })).toHaveProperty("disabled", true);
    fireEvent.click(screen.getAllByRole("button", { name: "上移" })[1]);
    fireEvent.click(screen.getByLabelText("我已核对上述优先顺序（最上方优先）"));
    fireEvent.click(screen.getByRole("button", { name: "按补充映射重新生成预览" }));
    expect(onResolve).toHaveBeenCalledTimes(1);
    const updated = onResolve.mock.calls[0][0] as ColumnMappingRule;
    expect(updated.delimiters).toEqual(["，"]);
    expect(updated.valueRank).toEqual([
      { value: "甲", rank: 1 },
      { value: "乙", rank: 2 }
    ]);
  });

  it("整格映射必须由用户填写，组件不把未知值默认为空白标签", () => {
    const onResolve = vi.fn().mockResolvedValue(undefined);
    render(<LocalMappingResolution
      rule={rule} values={["甲，乙"]} allowExactMapping busy={false}
      onResolve={onResolve} onCancel={() => undefined}
    />);
    fireEvent.click(screen.getByRole("button", { name: "按补充映射重新生成预览" }));
    expect(onResolve).not.toHaveBeenCalled();
    expect(screen.getByRole("alert").textContent).toContain("目标值不能为空");
    fireEvent.change(screen.getByLabelText("来源值 甲，乙 的目标值"), { target: { value: "高" } });
    fireEvent.click(screen.getByRole("button", { name: "按补充映射重新生成预览" }));
    expect(onResolve).toHaveBeenCalledTimes(1);
    expect((onResolve.mock.calls[0][0] as ColumnMappingRule).mapping.at(-1)).toEqual({
      source: "甲，乙", target: "高"
    });
  });

  it("空白卡只给四个空白选项，不叠一张补映射表", () => {
    const onResolve = vi.fn().mockResolvedValue(undefined);
    render(<LocalMappingResolution
      rule={rule} values={[]} allowExactMapping={false} busy={false}
      reason="blank" blankCount={2}
      onResolve={onResolve} onCancel={() => undefined}
    />);
    // 事实来自真实读取：卡上说的是"读到 2 个空白格"，不是一句泛泛的报错。
    expect(screen.getByText("2", { exact: false })).toBeTruthy();
    expect(document.body.textContent).toContain("个空白格");
    // 空白这一支不复用类别/分隔符的选项：没有补映射入口。
    expect(screen.queryByLabelText("拆分符")).toBeNull();
    expect(screen.queryByRole("button", { name: "按补充映射重新生成预览" })).toBeNull();
    expect(screen.getAllByRole("button", { name: "暂不修改" })).toHaveLength(1);
  });

  it("四个空白选项各自落成不同的结构化策略，不互相冒充", () => {
    const onResolve = vi.fn().mockResolvedValue(undefined);
    render(<LocalMappingResolution
      rule={rule} values={[]} allowExactMapping={false} busy={false}
      reason="blank" blankCount={1}
      onResolve={onResolve} onCancel={() => undefined}
    />);
    const submit = () => screen.getByRole("button", { name: "按选定的空白处理重新生成预览" });

    // 默认就是"保留目标原值"：不写入，不是读回原值再写回。
    fireEvent.click(submit());
    expect((onResolve.mock.calls[0][0] as ColumnMappingRule).blankPolicy).toEqual({ kind: "keep" });
    expect((onResolve.mock.calls[0][0] as ColumnMappingRule).blankLabel).toBeUndefined();

    fireEvent.click(screen.getByLabelText("清空目标内容 —— 清成真空单元格，不是写一个空字符串"));
    fireEvent.click(submit());
    expect((onResolve.mock.calls[1][0] as ColumnMappingRule).blankPolicy).toEqual({ kind: "clear" });

    fireEvent.click(screen.getByLabelText("写入空字符串 —— 明确写一个空字符串常量，与「清空」不同"));
    fireEvent.click(submit());
    expect((onResolve.mock.calls[2][0] as ColumnMappingRule).blankPolicy).toEqual({ kind: "value", value: "" });

    fireEvent.click(screen.getByLabelText("写入下面填的指定值"));
    // 选了"写指定值"却没填值：不替他挑默认标签，按钮保持不可用。
    expect(submit()).toHaveProperty("disabled", true);
    expect(onResolve).toHaveBeenCalledTimes(3);
    fireEvent.change(screen.getByLabelText("空白格写入的值"), { target: { value: "未分类" } });
    fireEvent.click(submit());
    expect((onResolve.mock.calls[3][0] as ColumnMappingRule).blankPolicy).toEqual({ kind: "value", value: "未分类" });
  });

  it("宿主回调抛错时不把可能含源值的异常原文展示出来", async () => {
    const secret = "仅在源格中的异常内容-5H9T";
    render(<LocalMappingResolution
      rule={rule} values={["甲，乙"]} allowExactMapping busy={false}
      onResolve={vi.fn().mockRejectedValue(new Error(secret))}
      onCancel={() => undefined}
    />);
    fireEvent.change(screen.getByLabelText("来源值 甲，乙 的目标值"), { target: { value: "高" } });
    fireEvent.click(screen.getByRole("button", { name: "按补充映射重新生成预览" }));
    await waitFor(() => expect(screen.getByRole("alert").textContent).toContain("原表未写入"));
    expect(document.body.textContent).not.toContain(secret);
  });
});
