/**
 * 专注窗口的展示数据结构。
 *
 * 类型本身已移到共享层（`@excel-bro/core/focusProtocol`），这里只保留 Excel 端的
 * 转出点，避免 App.tsx / FocusWorkspace.tsx 各自从不同位置取类型。
 */
export type {
  FocusMessage,
  FocusConversation,
  FocusTool,
  FocusWorkbookRef,
  FocusSnapshot
} from "@excel-bro/core/focusProtocol";
