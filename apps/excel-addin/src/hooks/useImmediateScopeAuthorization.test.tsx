// @vitest-environment jsdom
import { act, renderHook, waitFor } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";
import { useImmediateScopeAuthorization } from "./useImmediateScopeAuthorization";

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

describe("checkbox authorization", () => {
  it("requires a picker gesture; catalog readiness alone never grants access", async () => {
    const apply = vi.fn(async () => true);
    const { result, rerender } = renderHook(({ ready }) => useImmediateScopeAuthorization({ ready, apply }), { initialProps: { ready: false } });
    rerender({ ready: true });
    expect(apply).not.toHaveBeenCalled();
    act(() => result.current.request());
    await waitFor(() => expect(result.current.pending).toBe(false));
    expect(apply).toHaveBeenCalledOnce();
  });

  it("serializes rapid choices, invalidates the in-flight choice and applies only the latest", async () => {
    const first = deferred<boolean>();
    const current: Array<() => boolean> = [];
    const apply = vi.fn((isCurrent: () => boolean) => {
      current.push(isCurrent);
      return current.length === 1 ? first.promise : Promise.resolve(true);
    });
    const { result } = renderHook(() => useImmediateScopeAuthorization({ ready: true, apply }));
    act(() => result.current.request());
    expect(current[0]()).toBe(true);
    act(() => { result.current.request(); result.current.request(); });
    expect(result.current.isPending()).toBe(true);
    expect(current[0]()).toBe(false);
    expect(apply).toHaveBeenCalledOnce();
    await act(async () => first.resolve(true));
    await waitFor(() => expect(result.current.pending).toBe(false));
    expect(apply).toHaveBeenCalledTimes(2);
    expect(current[1]()).toBe(true);
    // Another gesture after completion must still render (generation != render counter).
    act(() => result.current.request());
    await waitFor(() => expect(apply).toHaveBeenCalledTimes(3));
  });

  it("keeps a failed selection blocked without an automatic retry, and permits an explicit retry", async () => {
    const apply = vi.fn().mockResolvedValueOnce(false).mockResolvedValue(true);
    const { result, rerender } = renderHook(() => useImmediateScopeAuthorization({ ready: true, apply }));
    await act(async () => result.current.request());
    rerender();
    expect(apply).toHaveBeenCalledOnce();
    expect(result.current.pending).toBe(true);
    await act(async () => result.current.request());
    expect(result.current.pending).toBe(false);
    expect(apply).toHaveBeenCalledTimes(2);
  });

  it("an unmounted selector cannot publish a late verification", async () => {
    const pending = deferred<boolean>();
    let current!: () => boolean;
    const { result, unmount } = renderHook(() => useImmediateScopeAuthorization({ ready: true, apply: (isCurrent) => { current = isCurrent; return pending.promise; } }));
    act(() => result.current.request());
    unmount();
    expect(current()).toBe(false);
    pending.resolve(true);
  });
});
