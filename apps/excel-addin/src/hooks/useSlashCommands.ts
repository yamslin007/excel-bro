import { useState, useCallback } from "react";
import type { SlashCommand } from "../SlashCommandAutocomplete";

/**
 * 斜杠命令状态类型
 */
export type SlashMode = "command" | "model";

/**
 * 斜杠命令管理 Hook
 *
 * 职责：
 * - 管理斜杠命令自动补全状态
 * - 检测用户输入中的斜杠命令
 * - 处理命令模式切换（一级命令 vs 模型选择）
 * - 过滤和匹配命令
 */
export function useSlashCommands() {
  const [showAutocomplete, setShowAutocomplete] = useState(false);
  const [filter, setFilter] = useState("");
  const [mode, setMode] = useState<SlashMode>("command");

  /**
   * 检测输入文本中的斜杠命令
   * 返回：是否应该显示自动补全
   */
  const detectSlashCommand = useCallback(
    (text: string, cursorPosition: number): boolean => {
      // 只在光标位置检测
      const textBeforeCursor = text.substring(0, cursorPosition);

      // 找到最后一个 /
      const lastSlashIndex = textBeforeCursor.lastIndexOf("/");

      if (lastSlashIndex === -1) {
        setShowAutocomplete(false);
        return false;
      }

      // / 必须在行首或前面是空格
      const charBeforeSlash =
        lastSlashIndex > 0
          ? textBeforeCursor[lastSlashIndex - 1]
          : undefined;

      if (charBeforeSlash !== undefined && charBeforeSlash !== " ") {
        setShowAutocomplete(false);
        return false;
      }

      // 提取 / 后面的文本作为过滤词
      const afterSlash = textBeforeCursor.substring(lastSlashIndex + 1);

      // / 后面不能有空格（除非是在选择模型的二级菜单中）
      if (mode === "command" && afterSlash.includes(" ")) {
        setShowAutocomplete(false);
        return false;
      }

      setFilter(afterSlash);
      setShowAutocomplete(true);
      return true;
    },
    [mode]
  );

  /**
   * 关闭自动补全
   */
  const closeAutocomplete = useCallback(() => {
    setShowAutocomplete(false);
    setFilter("");
    setMode("command");
  }, []);

  /**
   * 进入模型选择模式
   */
  const enterModelMode = useCallback(() => {
    setMode("model");
    setFilter("");
  }, []);

  /**
   * 返回命令模式
   */
  const exitModelMode = useCallback(() => {
    setMode("command");
    setFilter("");
  }, []);

  /**
   * 重置到初始状态
   */
  const reset = useCallback(() => {
    setShowAutocomplete(false);
    setFilter("");
    setMode("command");
  }, []);

  return {
    // 状态
    showAutocomplete,
    filter,
    mode,

    // 操作
    detectSlashCommand,
    closeAutocomplete,
    enterModelMode,
    exitModelMode,
    reset,

    // 内部状态设置（给特殊场景使用）
    setShowAutocomplete,
    setFilter,
    setMode,
  };
}

/**
 * 根据过滤词匹配斜杠命令
 */
export function filterSlashCommands(
  commands: SlashCommand[],
  filter: string
): SlashCommand[] {
  if (!filter) {
    return commands;
  }

  const lowerFilter = filter.toLowerCase();

  return commands.filter(
    (cmd) =>
      cmd.value.toLowerCase().includes(lowerFilter) ||
      (cmd.label && cmd.label.toLowerCase().includes(lowerFilter)) ||
      cmd.description.toLowerCase().includes(lowerFilter)
  );
}
