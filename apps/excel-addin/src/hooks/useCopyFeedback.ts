import { useCallback, useEffect, useRef, useState } from "react";
import { copyTextToClipboard } from "../utils";

interface UseCopyFeedbackOptions {
  onMessageCopyError?: () => void;
  onFormulaCopyError?: (messageId: string) => void;
}

/**
 * 复制反馈 Hook
 *
 * 职责：
 * - 管理消息和公式预览的“已复制”反馈状态
 * - 管理两个反馈自动清除定时器
 * - 在卸载时清理定时器，避免更新已卸载组件
 */
export function useCopyFeedback({
  onMessageCopyError,
  onFormulaCopyError
}: UseCopyFeedbackOptions) {
  const [copiedMessageId, setCopiedMessageId] = useState<string | null>(null);
  const [copiedFunctionPreviewId, setCopiedFunctionPreviewId] = useState<
    string | null
  >(null);
  const mountedRef = useRef(true);
  const copyFeedbackTimerRef = useRef<number | null>(null);
  const functionCopyTimerRef = useRef<number | null>(null);

  useEffect(
    () => {
      // 必须在 setup 里**重新武装**：`main.tsx` 用 `React.StrictMode` 包 `<App />`，
      // 开发构建下 React 会重放 effect（setup → cleanup → setup）。
      // 只在 cleanup 里置 false、不在 setup 里置回 true 的话，重放之后这个开关
      // 会**永久**停在 false，之后每次点复制都在写入"已复制"状态之前被它拦掉 ——
      // 界面表现就是"点了复制毫无反应"。
      mountedRef.current = true;
      return () => {
        mountedRef.current = false;
        if (copyFeedbackTimerRef.current !== null) {
          window.clearTimeout(copyFeedbackTimerRef.current);
        }
        if (functionCopyTimerRef.current !== null) {
          window.clearTimeout(functionCopyTimerRef.current);
        }
      };
    },
    []
  );

  const copyMessageText = useCallback(
    async (messageId: string, text: string) => {
      if (!text.trim()) return;
      const copied = await copyTextToClipboard(text);
      if (!mountedRef.current) return;
      if (!copied) {
        onMessageCopyError?.();
        return;
      }

      setCopiedMessageId(messageId);
      if (copyFeedbackTimerRef.current !== null) {
        window.clearTimeout(copyFeedbackTimerRef.current);
      }
      copyFeedbackTimerRef.current = window.setTimeout(() => {
        setCopiedMessageId((current) =>
          current === messageId ? null : current
        );
        copyFeedbackTimerRef.current = null;
      }, 1600);
    },
    [onMessageCopyError]
  );

  const copyFunctionFormula = useCallback(
    async (messageId: string, formula: string) => {
      const copied = await copyTextToClipboard(formula);
      if (!mountedRef.current) return;
      if (!copied) {
        onFormulaCopyError?.(messageId);
        return;
      }

      setCopiedFunctionPreviewId(messageId);
      if (functionCopyTimerRef.current !== null) {
        window.clearTimeout(functionCopyTimerRef.current);
      }
      functionCopyTimerRef.current = window.setTimeout(() => {
        setCopiedFunctionPreviewId((current) =>
          current === messageId ? null : current
        );
        functionCopyTimerRef.current = null;
      }, 1500);
    },
    [onFormulaCopyError]
  );

  return {
    copiedMessageId,
    copiedFunctionPreviewId,
    copyMessageText,
    copyFunctionFormula
  };
}
