// @vitest-environment jsdom
import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  shouldShowWorkDirectoryGuide,
  WorkDirectoryGuide
} from "./WorkDirectoryGuide";

afterEach(() => cleanup());

describe("WorkDirectoryGuide", () => {
  it("appears only after earlier onboarding and before a directory is configured", () => {
    expect(
      shouldShowWorkDirectoryGuide({
        ready: true,
        previousStepComplete: true,
        folderConfigured: false,
        dismissed: false,
        blocked: false
      })
    ).toBe(true);
    expect(
      shouldShowWorkDirectoryGuide({
        ready: true,
        previousStepComplete: true,
        folderConfigured: true,
        dismissed: false,
        blocked: false
      })
    ).toBe(false);
    expect(
      shouldShowWorkDirectoryGuide({
        ready: true,
        previousStepComplete: true,
        folderConfigured: false,
        dismissed: true,
        blocked: false
      })
    ).toBe(false);
  });

  it("explains authorization and delayed backup creation", () => {
    render(
      <WorkDirectoryGuide onDismiss={vi.fn()} onChoose={vi.fn()} />
    );

    expect(
      screen.getByRole("region", { name: "首次工作目录设置引导" })
    ).toBeTruthy();
    expect(screen.getByText(/只有你随后勾选的工作表会被授权/)).toBeTruthy();
    expect(screen.getByText(/恢复点只会在发生写入后创建/)).toBeTruthy();
  });

  it("supports skipping and choosing without hidden side effects", () => {
    const onDismiss = vi.fn();
    const onChoose = vi.fn();
    render(
      <WorkDirectoryGuide onDismiss={onDismiss} onChoose={onChoose} />
    );

    fireEvent.click(screen.getByRole("button", { name: "以后再设置" }));
    fireEvent.click(screen.getByRole("button", { name: "选择工作目录" }));

    expect(onDismiss).toHaveBeenCalledOnce();
    expect(onChoose).toHaveBeenCalledOnce();
  });

  it("disables both actions while the folder picker is busy", () => {
    render(
      <WorkDirectoryGuide busy onDismiss={vi.fn()} onChoose={vi.fn()} />
    );

    expect(
      (screen.getByRole("button", {
        name: "以后再设置"
      }) as HTMLButtonElement).disabled
    ).toBe(true);
    expect(
      (screen.getByRole("button", {
        name: "选择工作目录"
      }) as HTMLButtonElement).disabled
    ).toBe(true);
  });
});
