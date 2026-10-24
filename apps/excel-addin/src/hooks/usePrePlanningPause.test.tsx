// @vitest-environment jsdom
import { act, renderHook } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";
import { usePrePlanningPause } from "./usePrePlanningPause";

const target = { fileKey: "book:1", name: "test.xlsx", selected: ["s1"],
  sheets: [{ id: "s1", name: "Sheet1" }] };

describe("pre-planning request retention", () => {
  function setup() {
    const hook = renderHook(({ context }) => usePrePlanningPause<{ prompt: string }>(context),
      { initialProps: { context: "chat-1:manual" } });
    let id = "";
    act(() => { id = hook.result.current.remember({ messageId: "message-1", conversationId: "chat-1",
      target, mode: "manual", data: { prompt: "original request" } }); });
    return { ...hook, id };
  }
  it.each(["file", "scope", "mode", "unreadable"] as const)("rejects %s changes", async (change) => {
    const { result, id } = setup();
    const run = vi.fn();
    await act(async () => result.current.restart(id, "chat-1", change === "mode" ? "auto" : "manual",
      async () => change === "unreadable" ? null : { ...target,
        fileKey: change === "file" ? "book:2" : target.fileKey,
        selected: change === "scope" ? ["s1", "s2"] : target.selected }, run));
    expect(run).not.toHaveBeenCalled();
    expect(result.current.entries[0].error).toContain("无法核实");
  });
  it("claims once while verification is pending and consumes before dispatch", async () => {
    const { result, id } = setup();
    let release!: (value: typeof target) => void;
    const gate = new Promise<typeof target>((resolve) => { release = resolve; });
    const run = vi.fn(async () => {});
    const first = result.current.restart(id, "chat-1", "manual", () => gate, run);
    await result.current.restart(id, "chat-1", "manual", async () => target, run);
    expect(run).not.toHaveBeenCalled();
    await act(async () => { release(target); await first; });
    expect(run).toHaveBeenCalledExactlyOnceWith({ prompt: "original request" });
    expect(result.current.entries).toHaveLength(0);
  });
  it("old render callbacks read live candidates during verification and nested dispatch", async () => {
    const { result, id } = setup();
    const previousRender = result.current;
    let release!: (value: typeof target) => void;
    const gate = new Promise<typeof target>((resolve) => { release = resolve; });
    const run = vi.fn(async () => {
      expect(previousRender.entries).toHaveLength(1);
      expect(previousRender.pendingEntries()).toHaveLength(0);
    });
    const first = previousRender.restart(id, "chat-1", "manual", () => gate, run);
    expect(previousRender.pendingEntries()).toHaveLength(0);
    await act(async () => { release(target); await first; });
    expect(run).toHaveBeenCalledOnce();
  });
  it("cannot dispatch into another conversation after an in-flight verification", async () => {
    const { result, id, rerender } = setup();
    let release!: (value: typeof target) => void;
    const gate = new Promise<typeof target>((resolve) => { release = resolve; });
    const run = vi.fn();
    const first = result.current.restart(id, "chat-1", "manual", () => gate, run);
    rerender({ context: "chat-2:manual" });
    await act(async () => { release(target); await first; });
    await result.current.restart(id, "chat-2", "manual", async () => target, run);
    expect(run).not.toHaveBeenCalled();
    expect(result.current.entries).toHaveLength(1);
  });
  it("closing the pane while verifying cannot restart the request", async () => {
    const { result, id, unmount } = setup();
    let release!: (value: typeof target) => void;
    const gate = new Promise<typeof target>((resolve) => { release = resolve; });
    const run = vi.fn();
    const first = result.current.restart(id, "chat-1", "manual", () => gate, run);
    unmount();
    release(target);
    await first;
    expect(run).not.toHaveBeenCalled();
  });
});
