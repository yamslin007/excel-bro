import { useEffect, useRef, useState } from "react";

export interface SlashCommand {
  // value：选中时回传给 onSelect 的标识（命令名或 modelId）。
  value: string;
  // label：显示名（不含斜杠前缀）；缺省时用 value。
  label?: string;
  description: string;
  // showSlashPrefix：是否在名字前显示 "/"，命令项为 true，模型项为 false。
  showSlashPrefix?: boolean;
  // active：是否为当前生效项（如当前选中的模型），显示勾选。
  active?: boolean;
  disabled?: boolean;
}

interface SlashCommandAutocompleteProps {
  visible: boolean;
  commands: SlashCommand[];
  onSelect: (value: string) => void;
  filter?: string;
  // title：菜单顶部标题，如"选择模型"；缺省不显示。
  title?: string;
}

export function SlashCommandAutocomplete({
  visible,
  commands,
  onSelect,
  filter = "",
  title
}: SlashCommandAutocompleteProps) {
  const [selectedIndex, setSelectedIndex] = useState(0);
  const listRef = useRef<HTMLDivElement>(null);

  // 过滤（按显示名或描述匹配）
  const filteredCommands = commands.filter((cmd) => {
    const name = cmd.label ?? cmd.value;
    const keyword = filter.toLowerCase();
    return (
      name.toLowerCase().includes(keyword) ||
      cmd.description.toLowerCase().includes(keyword)
    );
  });

  useEffect(() => {
    setSelectedIndex(0);
  }, [filter, visible]);

  useEffect(() => {
    if (listRef.current && visible) {
      const selected = listRef.current.querySelector(
        `[data-index="${selectedIndex}"]`
      ) as HTMLElement;
      if (selected) {
        selected.scrollIntoView({ block: "nearest" });
      }
    }
  }, [selectedIndex, visible]);

  useEffect(() => {
    if (!visible) return;

    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === "ArrowDown") {
        e.preventDefault();
        setSelectedIndex((prev) =>
          prev < filteredCommands.length - 1 ? prev + 1 : prev
        );
      } else if (e.key === "ArrowUp") {
        e.preventDefault();
        setSelectedIndex((prev) => (prev > 0 ? prev - 1 : prev));
      } else if (e.key === "Enter" || e.key === "Tab") {
        const target = filteredCommands[selectedIndex];
        if (target && !target.disabled) {
          e.preventDefault();
          onSelect(target.value);
        }
      }
    };

    window.addEventListener("keydown", handleKeyDown, { capture: true });
    return () => {
      window.removeEventListener("keydown", handleKeyDown, { capture: true });
    };
  }, [visible, selectedIndex, filteredCommands, onSelect]);

  if (!visible || filteredCommands.length === 0) {
    return null;
  }

  return (
    <div ref={listRef} className="slash-command-autocomplete">
      {title && <div className="slash-command-title">{title}</div>}
      {filteredCommands.map((cmd, index) => {
        const name = cmd.label ?? cmd.value;
        return (
          <div
            key={cmd.value}
            data-index={index}
            className={`slash-command-item ${
              index === selectedIndex ? "selected" : ""
            } ${cmd.disabled ? "disabled" : ""}`}
            onMouseDown={(e) => {
              // onMouseDown 抢在 textarea 的 onBlur 之前触发选中。
              e.preventDefault();
              if (!cmd.disabled) onSelect(cmd.value);
            }}
            onMouseEnter={() => setSelectedIndex(index)}
          >
            <div className="slash-command-name">
              {cmd.showSlashPrefix === false ? name : `/${name}`}
              {cmd.active && <span className="slash-command-check">✓</span>}
            </div>
            <div className="slash-command-description">{cmd.description}</div>
          </div>
        );
      })}
    </div>
  );
}
