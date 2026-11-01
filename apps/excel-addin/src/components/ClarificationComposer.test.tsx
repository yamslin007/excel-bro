// @vitest-environment jsdom
import { cleanup, fireEvent, render } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { ClarificationComposer } from "./ClarificationComposer";

afterEach(cleanup);

const fields = [
  { id: "scope", label: "时间范围" },
  { id: "unit", label: "统计单位" }
];
const options = [
  { id: "0", label: "按会员等级", resolution: "按会员等级汇总" },
  { id: "1", label: "按门店", resolution: "按门店汇总" }
];

function mount(overrides: Partial<Parameters<typeof ClarificationComposer>[0]> = {}) {
  const props = {
    question: "请填写两个信息",
    options: [] as typeof options,
    selected: [] as string[],
    multiple: false,
    fields,
    fieldValues: {} as Record<string, string>,
    freeText: "",
    busy: false,
    onToggle: vi.fn(),
    onFieldChange: vi.fn(),
    ...overrides
  };
  return { ...render(<ClarificationComposer {...props} />), props };
}

describe("多字段澄清：必填、拼法、禁用", () => {
  /*
   * 这三条从已删除的 `ClarificationFields.test.tsx` 迁过来：
   * 旧组件被 `ClarificationComposer` 完全替代之后，它验的是"字段不能提交一半"、
   * "答案要带标签"、"busy 时不许再改"。行为没变，只是换了一个界面 —— 断言必须跟着搬家，
   * 不能随组件一起删掉。
   */
  it("必填没填完 ⇒ 预览行说明还差哪几项，不把半张表单当答案", () => {
    const { container } = mount({ fieldValues: { scope: "本月" } });
    const preview = container.querySelector(".clarification-composer-preview")!;
    expect(preview.getAttribute("data-ready")).toBe("false");
    expect(preview.textContent).toContain("统计单位");
    expect(preview.textContent).not.toContain("将发送");
  });

  it("全部填好 ⇒ 预览是「标签：值」逐行拼出来的答案", () => {
    const { container } = mount({ fieldValues: { scope: "本月", unit: "件" } });
    const preview = container.querySelector(".clarification-composer-preview")!;
    expect(preview.getAttribute("data-ready")).toBe("true");
    // 换行在预览里折成 " / "，但顺序与标签必须保持。
    expect(preview.textContent).toBe("将发送：时间范围：本月 / 统计单位：件");
  });

  it("busy 时输入框与选项都禁用（不给「还能再改一点」的错觉）", () => {
    const { container } = mount({ busy: true, options });
    const areas = [...container.querySelectorAll("textarea")];
    expect(areas).toHaveLength(2);
    expect(areas.every((input) => input.disabled)).toBe(true);
    const buttons = [...container.querySelectorAll("button")];
    expect(buttons.length).toBeGreaterThan(0);
    expect(buttons.every((button) => button.disabled)).toBe(true);
  });
});

describe("选项：点选只选中，不在这里提交", () => {
  it("点一下只报出选中的 id，界面里没有第二颗发送按钮", () => {
    const { container, props } = mount({ options, selected: [] });
    const first = container.querySelector<HTMLButtonElement>('[aria-pressed="false"]')!;
    fireEvent.click(first);
    expect(props.onToggle).toHaveBeenCalledWith("0");
    /*
     * 这一条是这次交互改造的核心：回答区**不许**自带发送按钮。
     * 多一颗就等于同一个动作两个入口，而且它不受输入框状态约束。
     */
    const labels = [...container.querySelectorAll("button")].map((button) => button.textContent);
    expect(labels).not.toContain("发送");
    expect(labels).not.toContain("发送回答");
  });

  it("多选时同一项可以再点一次取消（选中态由调用方给，组件只报告点击）", () => {
    const { container, props } = mount({ options, selected: ["0"], multiple: true });
    const chosen = container.querySelector('[aria-pressed="true"]')!;
    expect(chosen.getAttribute("data-selected")).toBe("true");
    fireEvent.click(chosen);
    expect(props.onToggle).toHaveBeenCalledWith("0");
  });

  it("选项 + 自由补充：预览里两句话都在（补充不许被丢掉）", () => {
    /*
     * 必须显式 `fields: []`：提交判据的优先级是 字段 > 选项 > 自由输入，
     * 默认 props 带着两个字段会把这条用例压成"还有 2 项没有填"。
     */
    const { container } = mount({ fields: [], options, selected: ["1"], freeText: "只看华东" });
    expect(container.querySelector(".clarification-composer-preview")!.textContent)
      .toBe("将发送：按门店 / 补充：只看华东");
  });

  it("框选按钮组通过标签区分单选与多选，不伪装为原生单选框", () => {
    const single = mount({ options, selected: [] });
    expect(single.container.querySelector('[role="group"][aria-label="单选"]')).not.toBeNull();
    expect(single.container.textContent).not.toContain("请选择一项，或写下你的想法。");
    cleanup();
    const multi = mount({ options, selected: [], multiple: true });
    expect(multi.container.querySelector('[role="group"][aria-label="可多选"]')).not.toBeNull();
  });
});
