const MAX_IMAGE_SIZE = 10 * 1024 * 1024;

export const ACCEPTED_IMAGE_TYPES = ["image/jpeg", "image/png", "image/webp"];

export function validateImageFile(file: File): void {
  if (file.size > MAX_IMAGE_SIZE) {
    throw new Error("图片文件过大，请选择小于 10MB 的图片");
  }
  if (!ACCEPTED_IMAGE_TYPES.includes(file.type)) {
    throw new Error("请选择 JPG、PNG 或 WebP 图片");
  }
}

export function convertToJPEG(file: File, quality = 0.9): Promise<Blob> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = () => {
      const image = new Image();
      image.onload = () => {
        const canvas = document.createElement("canvas");
        canvas.width = image.width;
        canvas.height = image.height;
        const context = canvas.getContext("2d");
        if (!context) {
          reject(new Error("无法创建图片处理画布"));
          return;
        }
        context.drawImage(image, 0, 0);
        canvas.toBlob(
          (blob) => {
            if (blob) {
              resolve(blob);
            } else {
              reject(new Error("图片处理失败"));
            }
          },
          "image/jpeg",
          quality
        );
      };
      image.onerror = () => reject(new Error("图片加载失败"));
      image.src = String(reader.result ?? "");
    };

    reader.onerror = () => reject(new Error("文件读取失败"));
    reader.readAsDataURL(file);
  });
}
