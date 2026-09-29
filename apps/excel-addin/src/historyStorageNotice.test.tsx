// @vitest-environment jsdom
// 本文件断言**用户实际看到的文案**，而不是 hook 的布尔值。
//
// 起因：`legacyCleanupNotice` 只有一个布尔值，而它背后其实有四种真实形态。早先的
// 实现在 v4 写入失败时**无条件**置 true，于是"写失败 + 本机压根没有 v3 旧键"这条
// 也会渲染出"旧版本记录待清理 / 历史已经以脱敏形式保存" —— 两句话都不成立：
// 没有旧键可清理，本次历史也从没保存过。
//
// 布尔值测不出这种谎报（它确实是 `true`），只有把渲染结果摆到 DOM 上读文案才行。
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";
import HistoryStorageNotice, {
  historyStorageNoticeCopy,
  isHistoryStorageNoticeVisible,
  type HistoryStorageStatus
} from "./components/HistoryStorageNotice";
import {
  resolveHistoryStorageStatus,
  useConversation
} from "./hooks/useConversation";
import {
  CHAT_STORAGE_KEY,
  LEGACY_CHAT_STORAGE_KEY,
  MODEL_STORAGE_KEY
} from "./conversation";

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT =
  true;

let mounted: Array<{ container: HTMLElement; unmount: () => void }> = [];

async function mountStatus(status: HistoryStorageStatus) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  await act(async () => {
    root.render(<HistoryStorageNotice status={status} />);
  });
  const entry = {
    container,
    unmount: () => {
      act(() => {
        root.unmount();
      });
      container.remove();
    }
  };
  mounted.push(entry);
  return entry;
}

/** 挂载真实 hook，并返回读取当前存储状态的工具。 */
async function mountHook(userRender?: (element: React.ReactElement) => void) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  let current: ReturnType<typeof useConversation> | null = null;
  function Probe() {
    current = useConversation();
    // 让真实 hook 的存储状态也经过一次 DOM 渲染 —— 断言的对象是"用户看到的"。
    return <HistoryStorageNotice status={current.historyStorageStatus} />;
  }
  const element = <Probe />;
  await act(async () => {
    if (userRender) userRender(element);
    else root.render(element);
  });
  await act(async () => {
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
  });
  const entry = {
    container,
    get hook(): ReturnType<typeof useConversation> {
      if (!current) throw new Error("hook 尚未挂载");
      return current;
    },
    unmount: () => {
      act(() => {
        root.unmount();
      });
      container.remove();
    }
  };
  mounted.push(entry);
  return entry;
}

afterEach(() => {
  mounted.forEach((entry) => entry.unmount());
  mounted = [];
  document.body.innerHTML = "";
  localStorage.clear();
});

/** 落盘持续失败、但 v3 键**允许**删除（互补于"删不掉、写得进"那组）。 */
function blockV4WritesOnly() {
  const originalSetItem = Storage.prototype.setItem;
  const spy = vi
    .spyOn(Storage.prototype, "setItem")
    .mockImplementation(function (this: Storage, key: string, value: string) {
      if (key === CHAT_STORAGE_KEY) {
        throw new Error("QuotaExceededError: 空间已满");
      }
      originalSetItem.call(this, key, value);
    });
  return { restore: () => spy.mockRestore() };
}

describe("history storage notice copy", () => {
  it("maps each of the four real outcomes to a truthful sentence", () => {
    // 1) 写失败 + 旧键存在：说"本次没保存"，且**明说旧记录被保留、未清理**。
    const failedWithLegacy = historyStorageNoticeCopy(
      "save-failed-legacy-present"
    );
    expect(failedWithLegacy).not.toBeNull();
    expect(failedWithLegacy?.heading).toContain("没能保存");
    expect(failedWithLegacy?.detail).toContain("旧版本");
    expect(failedWithLegacy?.detail).toContain("保留");
    // 不得声称 v4 已保存。
    expect(failedWithLegacy?.detail).not.toContain("已经保存到本机");

    // 2) 写失败 + 无旧键：说"未保存"，且**不得**出现"待清理"口径（无键可清）。
    const failedNoLegacy = historyStorageNoticeCopy("save-failed-no-legacy");
    expect(failedNoLegacy).not.toBeNull();
    expect(failedNoLegacy?.heading).toContain("未能保存");
    expect(failedNoLegacy?.detail).toContain("没有旧版本记录");
    expect(JSON.stringify(failedNoLegacy)).not.toContain("待清理");

    // 3) 写成功 + 删不掉：可以说明"脱敏后的 v4 已保存"，但必须说旧键仍待清理。
    const savedPending = historyStorageNoticeCopy("saved-legacy-pending");
    expect(savedPending).not.toBeNull();
    expect(savedPending?.heading).toContain("待清理");
    expect(savedPending?.detail).toContain("已经保存到本机");
    expect(savedPending?.detail).toContain("旧版本");
    // 不得把它说成"保存失败"。
    expect(JSON.stringify(savedPending)).not.toContain("没能保存");

    // 4) 全部成功：不渲染任何警告。
    expect(historyStorageNoticeCopy("saved-clean")).toBeNull();
    expect(isHistoryStorageNoticeVisible("saved-clean")).toBe(false);
    expect(isHistoryStorageNoticeVisible("saved-legacy-pending")).toBe(true);
  });

  it("derives the status from this attempt's write outcome and legacy presence", () => {
    expect(resolveHistoryStorageStatus(true, false)).toBe("saved-clean");
    expect(resolveHistoryStorageStatus(true, true)).toBe("saved-legacy-pending");
    expect(resolveHistoryStorageStatus(false, true)).toBe(
      "save-failed-legacy-present"
    );
    expect(resolveHistoryStorageStatus(false, false)).toBe(
      "save-failed-no-legacy"
    );
  });
});

describe("rendered notice actually says the right thing", () => {
  it("renders nothing at all when both the write and the cleanup succeeded", async () => {
    const view = await mountStatus("saved-clean");
    expect(view.container.querySelector(".storage-write-warning")).toBeNull();
    expect((view.container.textContent ?? "").trim()).toBe("");
  });

  it("renders the pending-cleanup wording only when there is a legacy key to clean", async () => {
    const pending = await mountStatus("saved-legacy-pending");
    const text = pending.container.textContent ?? "";
    expect(text).toContain("旧版本记录待清理");
    expect(text).toContain("已经保存到本机");
    // 仍然是同一个可被 CSS 命中的告警块。
    expect(pending.container.querySelector(".storage-write-warning")).not.toBeNull();
    expect(
      pending.container.querySelector(".storage-write-warning")?.getAttribute("role")
    ).toBe("alert");
  });

  it("never shows the legacy-cleanup wording when there is no legacy key", async () => {
    const view = await mountStatus("save-failed-no-legacy");
    const text = view.container.textContent ?? "";
    expect(text).toContain("未能保存");
    // 关键：没有旧键就不能说"待清理"、更不能说"历史已经以脱敏形式保存"。
    expect(text).not.toContain("旧版本记录待清理");
    expect(text).not.toContain("已经保存");
  });
});

describe("through the real hook", () => {
  const SECRET = "sk-live-notice-state-9c31";

  it("says the save failed (and the legacy record is kept) when the v4 write keeps failing", async () => {
    localStorage.setItem(
      LEGACY_CHAT_STORAGE_KEY,
      JSON.stringify([
        {
          id: "legacy-fail",
          role: "user",
          text: `用 ${SECRET} 汇总这张表`,
          createdAt: "2026-09-26T00:00:00.000Z"
        }
      ])
    );
    localStorage.setItem(MODEL_STORAGE_KEY, "connection:first");

    const fault = blockV4WritesOnly();
    const view = await mountHook();
    try {
      const text = view.container.textContent ?? "";
      // 事实一：本次保存失败。
      expect(view.hook.historyStorageStatus).toBe("save-failed-legacy-present");
      expect(text).toContain("没能保存");
      // 事实二：旧记录仍在（不能被说成"已清理"）。
      expect(text).toContain("保留");
      // 不得谎称 v4 已保存。
      expect(text).not.toContain("已经保存到本机");
      // 旧键确实还在。
      expect(localStorage.getItem(LEGACY_CHAT_STORAGE_KEY)).not.toBeNull();
      // 密钥不泄漏到界面。
      expect(text).not.toContain(SECRET);
    } finally {
      view.unmount();
      fault.restore();
    }
  });

  it("says the history was not saved — not that a legacy record awaits cleanup — when there is no v3 key", async () => {
    // 没有任何 v3 旧键，v4 又写不进去：这是最容易谎报的组合。
    localStorage.removeItem(LEGACY_CHAT_STORAGE_KEY);

    const fault = blockV4WritesOnly();
    const view = await mountHook();
    try {
      const text = view.container.textContent ?? "";
      expect(view.hook.historyStorageStatus).toBe("save-failed-no-legacy");
      expect(text).toContain("未能保存");
      // 核心反向断言：不能出现旧键清理口径。
      expect(text).not.toContain("旧版本记录待清理");
      expect(text).not.toContain("已经保存");
      expect(localStorage.getItem(CHAT_STORAGE_KEY)).toBeNull();
    } finally {
      view.unmount();
      fault.restore();
    }
  });

  it("says the sanitized v4 is saved but the old key still needs cleanup when only removal fails", async () => {
    localStorage.setItem(
      LEGACY_CHAT_STORAGE_KEY,
      JSON.stringify([
        {
          id: "legacy-removal",
          role: "user",
          text: `用 ${SECRET} 汇总这张表`,
          createdAt: "2026-09-26T00:00:00.000Z"
        }
      ])
    );

    const originalRemoveItem = Storage.prototype.removeItem;
    const removeSpy = vi
      .spyOn(Storage.prototype, "removeItem")
      .mockImplementation(function (this: Storage, key: string) {
        if (key === LEGACY_CHAT_STORAGE_KEY) {
          throw new Error(`removeItem 被拒绝 ${SECRET}`);
        }
        originalRemoveItem.call(this, key);
      });

    const view = await mountHook();
    try {
      const text = view.container.textContent ?? "";
      expect(view.hook.historyStorageStatus).toBe("saved-legacy-pending");
      expect(text).toContain("旧版本记录待清理");
      expect(text).toContain("已经保存到本机");
      // 不得说成保存失败。
      expect(text).not.toContain("没能保存");
      // 异常原文（含密钥）不得泄漏，v4 不得含密钥。
      expect(text).not.toContain(SECRET);
      expect(localStorage.getItem(CHAT_STORAGE_KEY) ?? "").not.toContain(SECRET);
    } finally {
      view.unmount();
      removeSpy.mockRestore();
    }
  });

  it("drops every warning once both the write and the cleanup succeed", async () => {
    localStorage.setItem(
      LEGACY_CHAT_STORAGE_KEY,
      JSON.stringify([
        {
          id: "legacy-happy",
          role: "user",
          text: "帮我汇总这张表",
          createdAt: "2026-09-26T00:00:00.000Z"
        }
      ])
    );

    const view = await mountHook();
    try {
      expect(view.hook.historyStorageStatus).toBe("saved-clean");
      expect(view.container.querySelector(".storage-write-warning")).toBeNull();
      expect((view.container.textContent ?? "").trim()).toBe("");
      expect(localStorage.getItem(LEGACY_CHAT_STORAGE_KEY)).toBeNull();
      expect(localStorage.getItem(CHAT_STORAGE_KEY)).not.toBeNull();
    } finally {
      view.unmount();
    }
  });
});
