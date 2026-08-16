// @vitest-environment jsdom
import { describe, it, expect } from "vitest";
import { renderHook, act } from "@testing-library/react";
import type { SlashCommand } from "../SlashCommandAutocomplete";
import {
  filterSlashCommands,
  useSlashCommands
} from "./useSlashCommands";

const commands: SlashCommand[] = [
  {
    value: "function",
    label: "函数",
    description: "生成 Excel 公式"
  },
  {
    value: "model",
    label: "模型",
    description: "切换模型"
  }
];

describe("useSlashCommands", () => {
  it("should initialize with autocomplete closed and command mode", () => {
    const { result } = renderHook(() => useSlashCommands());

    expect(result.current.showAutocomplete).toBe(false);
    expect(result.current.filter).toBe("");
    expect(result.current.mode).toBe("command");
  });

  it("should hide autocomplete when no slash exists", () => {
    const { result } = renderHook(() => useSlashCommands());

    act(() => {
      expect(result.current.detectSlashCommand("hello", 5)).toBe(false);
    });

    expect(result.current.showAutocomplete).toBe(false);
  });

  it("should detect slash at line start and set filter", () => {
    const { result } = renderHook(() => useSlashCommands());

    act(() => {
      expect(result.current.detectSlashCommand("/fun", 4)).toBe(true);
    });

    expect(result.current.showAutocomplete).toBe(true);
    expect(result.current.filter).toBe("fun");
  });

  it("should reject slash preceded by non-space", () => {
    const { result } = renderHook(() => useSlashCommands());

    act(() => {
      expect(result.current.detectSlashCommand("a/fun", 5)).toBe(false);
    });

    expect(result.current.showAutocomplete).toBe(false);
  });

  it("should reject space after slash in command mode", () => {
    const { result } = renderHook(() => useSlashCommands());

    act(() => {
      expect(result.current.detectSlashCommand("/fun x", 5)).toBe(false);
    });

    expect(result.current.showAutocomplete).toBe(false);
  });

  it("should allow space after slash in model mode", () => {
    const { result } = renderHook(() => useSlashCommands());

    act(() => result.current.enterModelMode());
    act(() => {
      expect(result.current.detectSlashCommand("/model x", 8)).toBe(true);
    });

    expect(result.current.showAutocomplete).toBe(true);
    expect(result.current.filter).toBe("model x");
  });

  it("should close autocomplete and reset mode", () => {
    const { result } = renderHook(() => useSlashCommands());

    act(() => {
      result.current.detectSlashCommand("/fun", 4);
      result.current.enterModelMode();
      result.current.closeAutocomplete();
    });

    expect(result.current.showAutocomplete).toBe(false);
    expect(result.current.filter).toBe("");
    expect(result.current.mode).toBe("command");
  });

  it("should reset to initial state", () => {
    const { result } = renderHook(() => useSlashCommands());

    act(() => {
      result.current.detectSlashCommand("/fun", 4);
      result.current.reset();
    });

    expect(result.current.showAutocomplete).toBe(false);
    expect(result.current.filter).toBe("");
    expect(result.current.mode).toBe("command");
  });

  it("should filter commands by value, label and description", () => {
    expect(filterSlashCommands(commands, "")).toEqual(commands);
    expect(filterSlashCommands(commands, "fun")).toHaveLength(1);
    expect(filterSlashCommands(commands, "函")).toHaveLength(1);
    expect(filterSlashCommands(commands, "公式")).toHaveLength(1);
    expect(filterSlashCommands(commands, "missing")).toHaveLength(0);
  });
});
