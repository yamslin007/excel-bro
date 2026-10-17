// @vitest-environment jsdom
import { act, cleanup, renderHook, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { useScopeAdmission } from "./useScopeAdmission";
import { admissionItem, type ScopeAdmissionReport } from "@excel-bro/core/scopeAdmission";
import capabilities from "../../../../config/capabilities.json";

afterEach(() => { cleanup(); vi.useRealTimers(); });
const report = (id: string, fingerprint = "one"): ScopeAdmissionReport => ({ items: [admissionItem(id)], fingerprint });
describe("selection eligibility shared by both hosts", () => {
  it("late results cannot enable another selection", async () => {
    let resolve!: (value: ScopeAdmissionReport) => void;
    const inspect = vi.fn(() => new Promise<ScopeAdmissionReport>((done) => { resolve = done; }));
    const h = renderHook(({ id }) => useScopeAdmission({ enabled: true, scopeKey: "file", ids: [id], inspect }), { initialProps: { id: "a" } });
    const first = resolve;
    h.rerender({ id: "b" });
    expect(h.result.current.canCommit).toBe(false);
    await act(async () => { first(report("a")); });
    expect(h.result.current.canCommit).toBe(false);
    await act(async () => { resolve(report("b")); });
    expect(h.result.current.canCommit).toBe(true);
    expect(h.result.current.item("a")).toBeUndefined();
  });
  it("commit rechecks file identity and refuses a replaced file", async () => {
    const inspect = vi.fn(async () => report("a"));
    const h = renderHook(() => useScopeAdmission({ enabled: true, scopeKey: "file", ids: ["a"], inspect }));
    await waitFor(() => expect(h.result.current.canCommit).toBe(true));
    inspect.mockResolvedValue(report("a", "replacement"));
    await act(async () => { expect(await h.result.current.verify()).toBeNull(); });
    expect(h.result.current.item("a")?.reason).toBe("target_changed");
    expect(h.result.current.canCommit).toBe(false);
  });
  it.each([
    { items: [] }, { items: [admissionItem("foreign")] },
    { items: [{ id: "a", status: "ready", reason: "forged" }] },
    { items: [{ id: "a", status: "ready", reason: "read_only" }] }
  ])("rejects missing, foreign or forged results %#", async (bad) => {
    const h = renderHook(() => useScopeAdmission({ enabled: true, scopeKey: "file", ids: ["a"], inspect: async () => bad as ScopeAdmissionReport }));
    await waitFor(() => expect(h.result.current.checking).toBe(false));
    expect(h.result.current.canCommit).toBe(false);
    expect(h.result.current.item("a")?.reason).toBe("identity_unverified");
  });
  it("write restrictions are explicit and still allow analysis selection", async () => {
    const h = renderHook(() => useScopeAdmission({ enabled: true, scopeKey: "file", ids: ["a"],
      inspect: async () => ({ items: [admissionItem("a", "protected")] }) }));
    await waitFor(() => expect(h.result.current.canCommit).toBe(true));
    expect(h.result.current.item("a")?.status).toBe("read_only");
  });
  it("timeout fails closed and late completion cannot restore readiness", async () => {
    vi.useFakeTimers();
    let resolve!: (value: ScopeAdmissionReport) => void;
    const h = renderHook(() => useScopeAdmission({ enabled: true, scopeKey: "file", ids: ["a"],
      inspect: () => new Promise((done) => { resolve = done; }) }));
    await act(async () => { await vi.advanceTimersByTimeAsync(capabilities.scopeAdmission.timeoutMs); });
    expect(h.result.current.canCommit).toBe(false);
    await act(async () => { resolve(report("a")); });
    expect(h.result.current.canCommit).toBe(false);
  });
});
