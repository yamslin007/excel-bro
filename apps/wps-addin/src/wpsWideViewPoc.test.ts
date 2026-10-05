import { afterEach, describe, expect, it, vi } from "vitest";
import type { WpsApplication } from "./wpsHost";
import {
  probeWideViewDialog,
  WIDE_VIEW_TARGET_HEIGHT,
  WIDE_VIEW_TARGET_WIDTH
} from "./wpsWideViewPoc";

function hostWith(extra: Partial<WpsApplication>): WpsApplication {
  return {
    Version: "12.0",
    PluginStorage: { getItem: () => undefined, setItem: () => undefined },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined,
    ...extra
  };
}

afterEach(() => {
  vi.restoreAllMocks();
});

describe("WPS 独立宽窗口能力探针", () => {
  it("宿主没有 ShowDialog/CreateWebDialog 时如实报 missing，不回退到任何替代方案", () => {
    const probe = probeWideViewDialog(hostWith({}));
    expect(probe.status).toBe("missing");
    expect(probe.presentApis).toEqual([]);
    expect(probe.openedWith).toBeNull();
    // 失败必须带原因，否则"记录 API/版本/现象"这条要求落不了地。
    expect(probe.error).toContain("ShowDialog");
  });

  it("ShowDialog 返回 true 只说明请求被接受，按逻辑像素传参且不伪造关闭结果", () => {
    const showDialog = vi.fn(() => true);
    const probe = probeWideViewDialog(hostWith({ ShowDialog: showDialog }));

    expect(probe.status).toBe("requested");
    expect(probe.openedWith).toBe("ShowDialog");
    expect(probe.closed).toBeNull();
    expect(probe.dialogMembers).toEqual([]);

    const [, title, width, height, modal] = showDialog.mock.calls[0] as unknown as [
      string,
      string,
      number,
      number,
      boolean
    ];
    expect(title).toContain("变更审阅");
    expect(modal).toBe(false);
    // 官方签名使用逻辑像素；不能再乘 devicePixelRatio。
    expect(width).toBe(WIDE_VIEW_TARGET_WIDTH);
    expect(height).toBe(WIDE_VIEW_TARGET_HEIGHT);
  });

  it("ShowDialog 抛错时状态是 error 而不是 available，且保留原始错误文本", () => {
    const probe = probeWideViewDialog(
      hostWith({
        ShowDialog: () => {
          throw new Error("no such method");
        }
      })
    );
    expect(probe.status).toBe("error");
    expect(probe.openedWith).toBeNull();
    expect(probe.error).toContain("no such method");
  });

  it("旧宿主若返回句柄，也不得未经观察就自动关闭", () => {
    const probe = probeWideViewDialog(hostWith({ ShowDialog: () => ({ ID: 3 }) }));
    expect(probe.status).toBe("requested");
    expect(probe.closed).toBeNull();
    expect(probe.dialogMembers).toContain("ID");
  });

  it("ShowDialog 返回 false 时再试 CreateWebDialog，并显式显示窗口", () => {
    const createWebDialog = vi.fn(() => ({ ID: 9, Visible: false, Close: () => undefined }));
    const probe = probeWideViewDialog(
      hostWith({
        ShowDialog: () => false,
        CreateWebDialog: createWebDialog
      })
    );
    expect(probe.status).toBe("requested");
    expect(probe.openedWith).toBe("CreateWebDialog");
    expect(createWebDialog).toHaveBeenCalledTimes(1);
    expect(createWebDialog.mock.results[0].value.Visible).toBe(true);
    expect(probe.closed).toBeNull();
  });

  it("CreateWebDialog 没有可显示句柄时不谎报窗口已打开", () => {
    const probe = probeWideViewDialog(hostWith({ CreateWebDialog: () => true }));
    expect(probe.status).toBe("error");
    expect(probe.error).toContain("句柄");
  });

  it("报告宿主版本，便于失败时定位是哪个版本缺这个 API", () => {
    const probe = probeWideViewDialog(hostWith({ Version: "12.1.0.28488" }));
    expect(probe.hostVersion).toBe("12.1.0.28488");
  });

  it("通信探针只在 URL 带随机校验码，不带工作簿或单元格内容", () => {
    const showDialog = vi.fn((..._args: unknown[]) => true);
    probeWideViewDialog(hostWith({ ShowDialog: showDialog }), "safe-nonce");
    const url = new URL(String(showDialog.mock.calls[0][0]));
    expect(url.searchParams.get("probeNonce")).toBe("safe-nonce");
    expect(url.searchParams.size).toBe(1);
  });
});
