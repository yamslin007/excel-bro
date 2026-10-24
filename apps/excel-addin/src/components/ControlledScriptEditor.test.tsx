// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { ControlledScriptEditor } from "./ControlledScriptEditor";
import PlanEditorDialog from "./PlanEditorDialog";
import { validatePlan } from "../api";

vi.mock("../api", () => ({ validatePlan: vi.fn(async plan => plan), chatModifyPlan: vi.fn() }));
vi.mock("../utils/imageStorage", () => ({ loadThemeSettings: vi.fn(async () => null) }));
const original = {id:"old",title:"旧计划",summary:"",assumptions:[],warnings:[],actions:[{type:"setFont",sheet:"One",range:"A1",bold:true}]};
afterEach(() => { cleanup(); vi.unstubAllGlobals(); vi.clearAllMocks(); localStorage.clear(); });

describe("shared controlled script entry", () => {
  it("applies a fresh multi-sheet plan without carrying old results or invoking validation/execution", () => {
    const onApply = vi.fn();
    render(<ControlledScriptEditor draft={JSON.stringify({...original, localResultRef:"stale", acceptanceCriteria:[{}]})} getAuthorizedSheets={() => ["One","Two"]} onApply={onApply} />);
    fireEvent.click(screen.getByText("填入多表示例"));
    fireEvent.click(screen.getByText("生成待核对计划"));
    const plan = JSON.parse(onApply.mock.calls[0][0]);
    expect(plan.actions.map((action: {sheet:string}) => action.sheet)).toEqual(["One","Two"]);
    expect(plan.id).not.toBe("old");
    expect(plan.localResultRef).toBeUndefined();
    expect(plan.acceptanceCriteria).toBeUndefined();
    expect(validatePlan).not.toHaveBeenCalled();
  });
  it("rechecks selected sheets on generation and leaves the original draft intact on refusal", () => {
    const onApply = vi.fn(); let authorized = ["One"];
    render(<ControlledScriptEditor draft={JSON.stringify(original)} getAuthorizedSheets={() => authorized} onApply={onApply} />);
    fireEvent.click(screen.getByText("从当前计划生成代码"));
    authorized = ["Two"];
    fireEvent.click(screen.getByText("生成待核对计划"));
    expect(screen.getByRole("alert").textContent).toContain("未勾选");
    expect(onApply).not.toHaveBeenCalled();
  });
  it("Excel dialog routes the generated plan through existing server validation on Save", async () => {
    const send = vi.fn();
    vi.stubGlobal("Office", {onReady: (fn: () => void) => fn(), context: {ui:{messageParent:send}}});
    vi.spyOn(window, "close").mockImplementation(() => {});
    localStorage.setItem("EB_PLAN_EDITOR_PAYLOAD", JSON.stringify({type:"plan-editor-init",payload:{
      plan:original, target:{kind:"message",messageId:"m1"}, authorizedSheets:["One","Two"],
      workbook:{name:"Test.xlsx",activeWorksheet:"One",worksheets:[{name:"One"},{name:"Two"},{name:"Secret"}]}
    }}));
    render(<PlanEditorDialog />);
    await waitFor(() => expect(screen.getByText("填入多表示例")).toBeTruthy());
    fireEvent.click(screen.getByText("填入多表示例"));
    fireEvent.click(screen.getByText("生成待核对计划"));
    expect(validatePlan).not.toHaveBeenCalled();
    fireEvent.click(screen.getByText("保存", {selector:"button"}));
    await waitFor(() => expect(validatePlan).toHaveBeenCalledOnce());
    const [plan, selected] = vi.mocked(validatePlan).mock.calls[0];
    expect(selected).toEqual(["One","Two"]);
    expect(plan.actions.map(action => action.sheet)).toEqual(selected);
    await waitFor(() => expect(send.mock.calls.some(([message]) => JSON.parse(message).type === "plan-editor-save")).toBe(true));
  });
});
