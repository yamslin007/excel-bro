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
  const copyFeedbackTimerRef = useRef<number | null>(null);
  const functionCopyTimerRef = useRef<number | null>(null);

  useEffect(
    () => () => {
      if (copyFeedbackTimerRef.current !== null) {
        window.clearTimeout(copyFeedbackTimerRef.current);
      }
      if (functionCopyTimerRef.current !== null) {
        window.clearTimeout(functionCopyTimerRef.current);
      }
    },
    []
  );

  const copyMessageText = useCallback(
    async (messageId: string, text: string) => {
      if (!text.trim()) return;
      const copied = await copyTextToClipboard(text);
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
