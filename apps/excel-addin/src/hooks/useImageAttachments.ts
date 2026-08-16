import { useState, useCallback, type DragEvent } from "react";
import {
  MAX_IMAGE_ATTACHMENTS,
  prepareImageFile,
  type PendingImage
} from "../imageAttachments";

/**
 * 图片附件管理 Hook
 *
 * 职责：
 * - 管理待发送的图片列表
 * - 处理图片拖拽上传
 * - 图片验证和错误处理
 * - 图片大小限制和格式检查
 */
export function useImageAttachments() {
  const [pendingImages, setPendingImages] = useState<PendingImage[]>([]);
  const [imageError, setImageError] = useState("");
  const [draggingImage, setDraggingImage] = useState(false);

  /**
   * 添加图片到待发送列表
   * 执行验证、压缩、转换为 base64
   */
  const addImage = useCallback(
    async (file: File) => {
      setImageError("");

      // 检查数量限制
      if (pendingImages.length >= MAX_IMAGE_ATTACHMENTS) {
        setImageError(`最多只能上传 ${MAX_IMAGE_ATTACHMENTS} 张图片`);
        return;
      }

      try {
        const prepared = await prepareImageFile(file);
        setPendingImages((prev) => [...prev, prepared]);
      } catch (error) {
        const message =
          error instanceof Error ? error.message : "图片处理失败";
        setImageError(message);
      }
    },
    [pendingImages.length]
  );

  /**
   * 移除指定索引的图片
   */
  const removeImage = useCallback((index: number) => {
    setPendingImages((prev) => prev.filter((_, i) => i !== index));
    setImageError("");
  }, []);

  /**
   * 清空所有待发送图片
   */
  const clearImages = useCallback(() => {
    setPendingImages([]);
    setImageError("");
  }, []);

  /**
   * 处理文件拖拽放下
   */
  const handleDrop = useCallback(
    async (e: DragEvent<HTMLTextAreaElement>) => {
      e.preventDefault();
      e.stopPropagation();
      setDraggingImage(false);

      const files = Array.from(e.dataTransfer.files);
      const imageFiles = files.filter((file) =>
        file.type.startsWith("image/")
      );

      if (imageFiles.length === 0) {
        setImageError("请拖入图片文件");
        return;
      }

      // 按顺序添加图片（受数量限制）
      for (const file of imageFiles) {
        if (pendingImages.length >= MAX_IMAGE_ATTACHMENTS) {
          setImageError(`最多只能上传 ${MAX_IMAGE_ATTACHMENTS} 张图片`);
          break;
        }
        await addImage(file);
      }
    },
    [pendingImages.length, addImage]
  );

  /**
   * 处理拖拽进入区域
   */
  const handleDragEnter = useCallback((e: DragEvent<HTMLTextAreaElement>) => {
    e.preventDefault();
    e.stopPropagation();

    // 检查是否包含文件
    if (e.dataTransfer.types.includes("Files")) {
      setDraggingImage(true);
    }
  }, []);

  /**
   * 处理拖拽离开区域
   */
  const handleDragLeave = useCallback((e: DragEvent<HTMLTextAreaElement>) => {
    e.preventDefault();
    e.stopPropagation();

    // 只有真正离开 textarea 时才取消高亮
    // relatedTarget 为 null 表示离开了文档
    if (!e.currentTarget.contains(e.relatedTarget as Node)) {
      setDraggingImage(false);
    }
  }, []);

  /**
   * 处理拖拽悬停
   */
  const handleDragOver = useCallback((e: DragEvent<HTMLTextAreaElement>) => {
    e.preventDefault();
    e.stopPropagation();
  }, []);

  return {
    // 状态
    pendingImages,
    imageError,
    draggingImage,

    // 操作
    addImage,
    removeImage,
    clearImages,

    // 拖拽事件处理
    handleDrop,
    handleDragEnter,
    handleDragLeave,
    handleDragOver,

    // 内部状态设置（给子组件使用）
    setImageError,
    setDraggingImage,
  };
}
