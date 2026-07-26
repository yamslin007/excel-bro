import type { ImageAttachment } from "./contracts";
import capabilities from "../../../config/capabilities.json";

export const MAX_IMAGE_ATTACHMENTS = capabilities.images.maxAttachments;
export const MAX_IMAGE_BYTES = capabilities.images.maxBytes;
const MAX_SOURCE_IMAGE_BYTES = capabilities.images.maxSourceBytes;
const MAX_IMAGE_DIMENSION = capabilities.images.maxDimension;

export interface PendingImage extends ImageAttachment {
  id: string;
  previewUrl: string;
}

function readAsDataUrl(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(new Error("读取图片失败"));
    reader.readAsDataURL(blob);
  });
}

function loadImage(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("无法解析这张图片"));
    image.src = url;
  });
}

function canvasBlob(
  canvas: HTMLCanvasElement,
  mediaType: "image/png" | "image/jpeg",
  quality?: number
): Promise<Blob> {
  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (blob) => (blob ? resolve(blob) : reject(new Error("压缩图片失败"))),
      mediaType,
      quality
    );
  });
}

function attachmentFromDataUrl(
  fileName: string,
  dataUrl: string
): PendingImage {
  const match = /^data:(image\/(?:png|jpeg|webp));base64,(.+)$/s.exec(dataUrl);
  if (!match) throw new Error("图片编码格式不受支持");
  return {
    id:
      globalThis.crypto?.randomUUID?.() ??
      `${Date.now()}-${Math.random().toString(16).slice(2)}`,
    name: fileName || "截图",
    mediaType: match[1] as ImageAttachment["mediaType"],
    data: match[2],
    previewUrl: dataUrl
  };
}

export async function prepareImageFile(file: File): Promise<PendingImage> {
  if (!["image/png", "image/jpeg", "image/webp"].includes(file.type)) {
    throw new Error("仅支持 PNG、JPEG 和 WebP 图片");
  }
  if (file.size > MAX_SOURCE_IMAGE_BYTES) {
    throw new Error("原始图片不能超过 12 MB");
  }

  const originalUrl = URL.createObjectURL(file);
  try {
    const image = await loadImage(originalUrl);
    const largestDimension = Math.max(image.naturalWidth, image.naturalHeight);
    if (
      largestDimension <= MAX_IMAGE_DIMENSION &&
      file.size <= MAX_IMAGE_BYTES
    ) {
      return attachmentFromDataUrl(file.name, await readAsDataUrl(file));
    }

    const scale = Math.min(1, MAX_IMAGE_DIMENSION / largestDimension);
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(image.naturalWidth * scale));
    canvas.height = Math.max(1, Math.round(image.naturalHeight * scale));
    const context = canvas.getContext("2d");
    if (!context) throw new Error("当前环境无法压缩图片");
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.drawImage(image, 0, 0, canvas.width, canvas.height);

    const compressed = await canvasBlob(canvas, "image/jpeg", 0.84);
    if (compressed.size > MAX_IMAGE_BYTES) {
      throw new Error("压缩后的图片仍超过 4 MB，请裁剪后重试");
    }
    return attachmentFromDataUrl(
      file.name.replace(/\.[^.]+$/, "") + ".jpg",
      await readAsDataUrl(compressed)
    );
  } finally {
    URL.revokeObjectURL(originalUrl);
  }
}
