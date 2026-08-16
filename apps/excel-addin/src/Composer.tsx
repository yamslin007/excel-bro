// 输入区：范围切换按钮 + 消息输入框 + 图片附件 + 发送/停止。
// 纯展示组件，JSX 从 App.tsx 逐字搬移；状态与回调由 props 传入。
import type {
  ChangeEvent,
  ClipboardEvent,
  Dispatch,
  DragEvent,
  KeyboardEvent,
  PointerEvent,
  RefObject,
  SetStateAction
} from "react";
import {
  SlashCommandAutocomplete,
  type SlashCommand
} from "./SlashCommandAutocomplete";
import type { PendingImage } from "./imageAttachments";
import type {
  ExecutionUndoSnapshot,
  FolderCatalog,
  WorkbookSnapshot
} from "./contracts";
import type { SourceMode, WorkbookScopeMode } from "./App";

interface ComposerProps {
  contextOpen: boolean;
  setContextOpen: Dispatch<SetStateAction<boolean>>;
  setSheetSearch: Dispatch<SetStateAction<string>>;
  workbook: WorkbookSnapshot | null;
  sourceMode: SourceMode;
  folderCatalog: FolderCatalog | null;
  folderSheetKeys: string[];
  workbookScopeMode: WorkbookScopeMode;
  workbookDataPeriod: string | null;
  selectedSheetNames: string[];
  draggingImage: boolean;
  setDraggingImage: Dispatch<SetStateAction<boolean>>;
  handleImageDrop: (event: DragEvent<HTMLDivElement>) => void;
  setComposerHeight: Dispatch<SetStateAction<number | null>>;
  startComposerResize: (event: PointerEvent<HTMLDivElement>) => void;
  moveComposerResize: (event: PointerEvent<HTMLDivElement>) => void;
  finishComposerResize: (event: PointerEvent<HTMLDivElement>) => void;
  resizeComposerWithKeyboard: (event: KeyboardEvent<HTMLDivElement>) => void;
  showSlashAutocomplete: boolean;
  setShowSlashAutocomplete: Dispatch<SetStateAction<boolean>>;
  slashMode: "command" | "model";
  slashCommands: SlashCommand[];
  slashModelCommands: SlashCommand[];
  handleSlashCommandSelect: (value: string) => void;
  slashFilter: string;
  composerInputRef: RefObject<HTMLTextAreaElement | null>;
  prompt: string;
  handleComposerChange: (event: ChangeEvent<HTMLTextAreaElement>) => void;
  handleComposerKeyDown: (event: KeyboardEvent<HTMLTextAreaElement>) => void;
  handleImagePaste: (event: ClipboardEvent<HTMLTextAreaElement>) => void;
  pendingImages: PendingImage[];
  setPendingImages: Dispatch<SetStateAction<PendingImage[]>>;
  setImageError: Dispatch<SetStateAction<string>>;
  imageError: string;
  imageInputRef: RefObject<HTMLInputElement | null>;
  addImageFiles: (files: File[]) => Promise<void>;
  supportsVision: boolean;
  busy: boolean;
  lastUndoSnapshot: ExecutionUndoSnapshot | null;
  undoLastExecution: () => Promise<void>;
  steerTurn: (text: string) => void;
  stopTurn: () => void;
  sendMessage: () => Promise<void>;
  replyTarget: { messageId: string; originalPrompt: string } | null;
  clearReplyTarget: () => void;
}

export default function Composer({
  contextOpen,
  setContextOpen,
  setSheetSearch,
  workbook,
  sourceMode,
  folderCatalog,
  folderSheetKeys,
  workbookScopeMode,
  workbookDataPeriod,
  selectedSheetNames,
  draggingImage,
  setDraggingImage,
  handleImageDrop,
  setComposerHeight,
  startComposerResize,
  moveComposerResize,
  finishComposerResize,
  resizeComposerWithKeyboard,
  showSlashAutocomplete,
  setShowSlashAutocomplete,
  slashMode,
  slashCommands,
  slashModelCommands,
  handleSlashCommandSelect,
  slashFilter,
  composerInputRef,
  prompt,
  handleComposerChange,
  handleComposerKeyDown,
  handleImagePaste,
  pendingImages,
  setPendingImages,
  setImageError,
  imageError,
  imageInputRef,
  addImageFiles,
  supportsVision,
  busy,
  lastUndoSnapshot,
  undoLastExecution,
  steerTurn,
  stopTurn,
  sendMessage,
  replyTarget,
  clearReplyTarget
}: ComposerProps) {
  return (
    <>
      <button
        className={`scope-trigger ${contextOpen ? "open" : ""}`}
        onClick={() => {
          setSheetSearch("");
          setContextOpen((value) => !value);
        }}
        aria-expanded={contextOpen}
      >
        <i aria-hidden="true">▦</i>
        <span>
          {workbook && sourceMode === "workbook" && (
            <>
              <small title={workbook.name}>{workbook.name}</small>
              <span className="scope-sep" aria-hidden="true">
                |
              </span>
            </>
          )}
          <strong>
            {!workbook
              ? "正在读取工作簿"
              : sourceMode === "folder"
                ? folderCatalog
                  ? `文件夹 · 已选 ${folderSheetKeys.length} 个工作表`
                  : "选择文件夹"
                : workbookScopeMode === "auto"
                  ? `${
                      workbookDataPeriod
                        ? `${workbookDataPeriod} · `
                        : ""
                    }当前表 ${workbook.activeWorksheet}`
                  : `${
                      workbookDataPeriod
                        ? `${workbookDataPeriod} · `
                        : ""
                    }已固定 ${selectedSheetNames.length} 个工作表`}
          </strong>
        </span>
        <b>{contextOpen ? "⌄" : "⌃"}</b>
      </button>

      <div
        className={`composer-input-region ${
          draggingImage ? "dragging-image" : ""
        }`}
        onDragEnter={(event) => {
          event.preventDefault();
          setDraggingImage(true);
        }}
        onDragOver={(event) => event.preventDefault()}
        onDragLeave={(event) => {
          if (!event.currentTarget.contains(event.relatedTarget as Node)) {
            setDraggingImage(false);
          }
        }}
        onDrop={handleImageDrop}
      >
        <div
          className="composer-resize-handle"
          role="separator"
          aria-label="上下拖动调整输入框高度，双击恢复自动高度"
          aria-orientation="horizontal"
          tabIndex={0}
          onPointerDown={startComposerResize}
          onPointerMove={moveComposerResize}
          onPointerUp={finishComposerResize}
          onPointerCancel={finishComposerResize}
          onDoubleClick={() => setComposerHeight(null)}
          onKeyDown={resizeComposerWithKeyboard}
        >
          <i />
        </div>
        <div className="composer-box">
          <SlashCommandAutocomplete
            visible={showSlashAutocomplete}
            commands={slashMode === "model" ? slashModelCommands : slashCommands}
            onSelect={handleSlashCommandSelect}
            filter={slashFilter}
            title={slashMode === "model" ? "选择模型" : undefined}
          />
          {replyTarget && (
            <div className="composer-reply-chip">
              <span title={replyTarget.originalPrompt}>
                正在回复：{replyTarget.originalPrompt}
              </span>
              <button
                type="button"
                aria-label="取消回复"
                title="取消回复"
                onClick={clearReplyTarget}
              >
                ✕
              </button>
            </div>
          )}
          <textarea
            ref={composerInputRef}
            aria-label="给 Excel Bro 发消息"
            value={prompt}
            onChange={handleComposerChange}
            onKeyDown={handleComposerKeyDown}
            onPaste={handleImagePaste}
            onBlur={() => setShowSlashAutocomplete(false)}
            placeholder={
              workbook
                ? "描述你想查询、分析或修改的内容…"
                : "可以先输入需求，工作簿读取完成后即可发送…"
            }
            rows={1}
          />
          {pendingImages.length > 0 && (
            <div className="composer-attachments">
              {pendingImages.map((image) => (
                <figure key={image.id}>
                  <img src={image.previewUrl} alt={image.name} />
                  <figcaption title={image.name}>{image.name}</figcaption>
                  <button
                    onClick={() => {
                      setPendingImages((current) =>
                        current.filter((item) => item.id !== image.id)
                      );
                      setImageError("");
                    }}
                    aria-label={`移除图片 ${image.name}`}
                  >
                    ×
                  </button>
                </figure>
              ))}
            </div>
          )}
          {imageError && (
            <div className="composer-image-error">{imageError}</div>
          )}
          <div className="composer-toolbar">
            <input
              ref={imageInputRef}
              className="image-file-input"
              type="file"
              accept="image/png,image/jpeg,image/webp"
              multiple
              onChange={(event) => {
                const files = [...(event.target.files ?? [])];
                event.target.value = "";
                void addImageFiles(files);
              }}
            />
            <div className="composer-tools-left">
              <button
                className="attach-image-button"
                disabled={!supportsVision || busy}
                onClick={() => imageInputRef.current?.click()}
                title={
                  supportsVision
                    ? "添加图片，也可以直接粘贴或拖入截图"
                    : "当前模型不支持图片"
                }
                aria-label="添加图片"
              >
                <svg
                  viewBox="0 0 24 24"
                  aria-hidden="true"
                  focusable="false"
                >
                  <rect x="3.5" y="4.5" width="17" height="15" rx="2.5" />
                  <circle cx="9" cy="9.5" r="1.5" />
                  <path d="m5.5 17 4.2-4.4 3.1 3 2.2-2.2 3.5 3.6" />
                </svg>
                <span>图片</span>
              </button>
              {lastUndoSnapshot && (
                <button
                  className="attach-image-button"
                  disabled={busy}
                  onClick={() => void undoLastExecution()}
                  title="撤销上一次 Excel Bro 执行"
                >
                  ↶ <span>撤销</span>
                </button>
              )}
            </div>
            {busy ? (
              // 运行中：有字=转向（带话打断重跑），无字=纯停止。
              <button
                className={`send-button ${
                  prompt.trim() ? "is-steer" : "is-stop"
                }`}
                onClick={() =>
                  prompt.trim() ? steerTurn(prompt) : stopTurn()
                }
                aria-label={prompt.trim() ? "打断并补充" : "停止"}
                title={
                  prompt.trim()
                    ? "打断当前处理，带上这句话重新开始"
                    : "停止当前处理"
                }
              >
                {prompt.trim() ? "↑" : "■"}
              </button>
            ) : (
              <button
                className="send-button"
                disabled={
                  !workbook ||
                  (!prompt.trim() && pendingImages.length === 0) ||
                  (pendingImages.length > 0 && !supportsVision)
                }
                onClick={() => void sendMessage()}
                aria-label="发送"
              >
                ↑
              </button>
            )}
          </div>
        </div>
      </div>
      <span>
        {busy
          ? "运行中：输入新内容按 Enter 可打断并转向，留空点 ■ 停止"
          : "Enter 发送 · Shift + Enter 换行 · 可粘贴截图 · 写入操作会先预览"}
      </span>
    </>
  );
}
