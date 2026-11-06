// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createFocusEnvelope } from "@excel-bro/core/focusProtocol";
import FocusWorkspace from "./FocusWorkspace";

(
  globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }
).IS_REACT_ACT_ENVIRONMENT = true;

/**
 * Excel 专注窗口子窗口的握手编排回归。
 *
 * 2026-11-03 真机暴露：子窗口停在「正在与任务窗格握手」，永远不进入内容。根因是
 * 子窗口用**自己发消息的序号**（发完 hello 后 seqRef=1）去当父窗格消息的 `minSeq`
 * 下限，而父窗格第一条快照 seq 也是 1 ⇒ `1 <= 1` 被误判 stale-seq 拒收。
 * 父/子各维护独立序号空间，子窗口必须用「已收到的父消息序号」做 minSeq（初始 -1）。
 */

type ReceiveHandler = (event: { message: string }) => void;

interface OfficeStub {
  onReady: (cb: () => void) => Promise<unknown>;
  AsyncResultStatus: { Succeeded: string; Failed: string };
  EventType: { DialogParentMessageReceived: string };
  context: {
    ui: {
      addHandlerAsync: (type: string, handler: ReceiveHandler, cb: (r: { status: string }) => void) => void;
      messageParent: (body: string) => void;
    };
  };
}

const SESSION = "0123456789abcdef0123456789abcdef";
const WORKBOOK = { name: "预算.xlsx", identity: "excel-path:c:\\data\\预算.xlsx" };

function makeSnapshot() {
  return {
    initialView: "conversation" as const,
    activeConversationId: "c1",
    conversations: [
      {
        id: "c1",
        title: "整理销售表",
        updatedAt: "2026-11-03T08:00:00.000Z",
        messages: [{ id: "m1", role: "user" as const, text: "把空行删掉", createdAt: "2026-11-03T08:00:00.000Z" }]
      }
    ],
    tools: [],
    workbook: WORKBOOK
  };
}

/**
 * 一个模拟"对话框子窗口"的 Office 替身：`messageParent` 把消息交给测试侧，测试侧再借
 * `addHandlerAsync` 注册的回调把父窗格消息"投递"回子窗口，模拟真实的 DialogApi 往返。
 */
function makeChildOffice() {
  let receive: ReceiveHandler | null = null;
  const parentMessages: string[] = [];

  const office: OfficeStub = {
    onReady: (cb) => {
      cb();
      return Promise.resolve({ host: "Excel", platform: "PC" });
    },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    EventType: { DialogParentMessageReceived: "dialogParentMessageReceived" },
    context: {
      ui: {
        addHandlerAsync: (_type, handler, cb) => {
          receive = handler;
          cb({ status: "succeeded" });
        },
        messageParent: (body) => {
          parentMessages.push(body);
        }
      }
    }
  };

  return {
    office,
    sentMessages: () => parentMessages,
    /** 父窗格向子窗口投递一条消息（等价于 dialog.messageChild）。 */
    deliverToChild: (body: string) => {
      receive?.({ message: body });
    }
  };
}

let officeLike: ReturnType<typeof makeChildOffice>;

beforeEach(() => {
  officeLike = makeChildOffice();
  vi.stubGlobal("Office", officeLike.office);
});

afterEach(() => {
  vi.unstubAllGlobals();
});

async function mount() {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  await act(async () => {
    root.render(<FocusWorkspace />);
  });
  return {
    container,
    unmount: () => {
      act(() => root.unmount());
      container.remove();
    }
  };
}

describe("Excel 专注窗口子窗口握手", () => {
  it("父窗格回的第一条快照（seq=1）不会被误判过期：子窗口进入内容态", async () => {
    // 让子窗口从 URL 拿到会话标识。
    window.history.replaceState(null, "", `?session=${SESSION}`);

    const { container, unmount } = await mount();

    // 子窗口应发出 hello（可能重发，但至少一次）。
    expect(officeLike.sentMessages().length).toBeGreaterThan(0);
    const hello = JSON.parse(officeLike.sentMessages()[0]);
    expect(hello.kind).toBe("hello");
    expect(hello.seq).toBe(1); // 子窗口发 hello 后自己的序号已经推进到 1。

    // 父窗格回一条 seq=1 的快照 —— 这正是修复前被误判 stale-seq 的那条。
    officeLike.deliverToChild(
      JSON.stringify(
        createFocusEnvelope({
          kind: "snapshot",
          sessionId: SESSION,
          seq: 1,
          sentAt: Date.now(),
          workbook: WORKBOOK,
          snapshot: makeSnapshot()
        })
      )
    );

    await act(async () => {
      await new Promise<void>((resolve) => setTimeout(resolve, 0));
    });

    // 进入内容态：出现「已同步」徽标与对话标题，而不是停在「正在握手」。
    expect(container.textContent).toContain("已同步");
    expect(container.textContent).toContain("整理销售表");
    expect(container.textContent).not.toContain("正在与任务窗格握手");

    unmount();
  });

  it("仍拒绝序号真正倒退的父窗格消息（独立序号空间的负向锚点）", async () => {
    window.history.replaceState(null, "", `?session=${SESSION}`);
    const { container, unmount } = await mount();

    // 先收到 seq=5 的快照，进入内容态。
    officeLike.deliverToChild(
      JSON.stringify(
        createFocusEnvelope({
          kind: "snapshot",
          sessionId: SESSION,
          seq: 5,
          sentAt: Date.now(),
          workbook: WORKBOOK,
          snapshot: makeSnapshot()
        })
      )
    );
    await act(async () => {
      await new Promise<void>((resolve) => setTimeout(resolve, 0));
    });
    expect(container.textContent).toContain("已同步");

    // 再来一条 seq=3 的快照（真正的倒退），应被拒收，界面不回退、也不报错。
    officeLike.deliverToChild(
      JSON.stringify(
        createFocusEnvelope({
          kind: "snapshot",
          sessionId: SESSION,
          seq: 3,
          sentAt: Date.now(),
          workbook: WORKBOOK,
          snapshot: makeSnapshot()
        })
      )
    );
    await act(async () => {
      await new Promise<void>((resolve) => setTimeout(resolve, 0));
    });
    expect(container.textContent).toContain("已同步");

    unmount();
  });
});
