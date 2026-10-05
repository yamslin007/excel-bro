type MessageCopyButtonProps = {
  copied: boolean;
  onClick: () => void;
};

/** Shared copy affordance for Excel and WPS conversation messages. */
export default function MessageCopyButton({ copied, onClick }: MessageCopyButtonProps) {
  return (
    <button
      className={`message-copy-button${copied ? " copied" : ""}`}
      type="button"
      title={copied ? "已复制" : "复制这条消息"}
      aria-label={copied ? "消息已复制" : "复制这条消息"}
      onClick={onClick}
    >
      {copied ? (
        <>
          <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
            <path d="m5 12 4 4L19 6" />
          </svg>
          <span>已复制</span>
        </>
      ) : (
        <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
          <rect x="8" y="8" width="11" height="11" rx="2" />
          <path d="M16 8V6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h2" />
        </svg>
      )}
    </button>
  );
}
