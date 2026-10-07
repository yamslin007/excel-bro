import { useState, useRef, useCallback, useEffect, type ClipboardEvent, type DragEvent, type SetStateAction } from "react";
import { MAX_IMAGE_ATTACHMENTS, prepareImageFile, type PendingImage } from "../imageAttachments";

/** Shared by Excel and WPS. Bytes stay in composer memory until explicit send. */
export function useImageAttachments() {
  const [pendingImages, updateImages] = useState<PendingImage[]>([]);
  const imagesRef = useRef<PendingImage[]>([]);
  const generation = useRef(0);
  const reservations = useRef(0);
  const nextOrder = useRef(0);
  const imageOrders = useRef(new Map<string, number>());
  const [preparingImages, setPreparingImages] = useState(false);
  const [imageError, setImageError] = useState("");
  const [draggingImage, setDraggingImage] = useState(false);
  const setPendingImages = useCallback((action: SetStateAction<PendingImage[]>) => {
    generation.current += 1;
    reservations.current = 0;
    setPreparingImages(false);
    const next = typeof action === "function" ? action(imagesRef.current) : action;
    imagesRef.current = next.slice(0, MAX_IMAGE_ATTACHMENTS);
    imageOrders.current = new Map(imagesRef.current.map((image, index) => [image.id, index]));
    nextOrder.current = imagesRef.current.length;
    updateImages(imagesRef.current);
  }, []);
  useEffect(() => () => { generation.current += 1; }, []);

  const addImage = useCallback(async (file: File) => {
    if (imagesRef.current.length + reservations.current >= MAX_IMAGE_ATTACHMENTS) {
      setImageError(`最多只能上传 ${MAX_IMAGE_ATTACHMENTS} 张图片`);
      return;
    }
    const currentGeneration = generation.current;
    const order = nextOrder.current++;
    reservations.current += 1;
    setPreparingImages(true);
    setImageError("");
    try {
      const prepared = await prepareImageFile(file);
      if (generation.current !== currentGeneration) return;
      imageOrders.current.set(prepared.id, order);
      imagesRef.current = [...imagesRef.current, prepared].sort((a, b) =>
        (imageOrders.current.get(a.id) ?? 0) - (imageOrders.current.get(b.id) ?? 0));
      updateImages(imagesRef.current);
    } catch (error) {
      if (generation.current === currentGeneration) {
        setImageError(error instanceof Error ? error.message : "图片处理失败");
      }
    } finally {
      if (generation.current === currentGeneration) {
        reservations.current -= 1;
        setPreparingImages(reservations.current > 0);
      }
    }
  }, []);
  const addImageFiles = useCallback(async (files: File[]) => {
    await Promise.all(files.map(addImage));
  }, [addImage]);
  const removeImage = useCallback((index: number) => {
    setPendingImages((current) => current.filter((_, i) => i !== index));
    setImageError("");
  }, [setPendingImages]);
  const clearImages = useCallback(() => {
    setPendingImages([]);
    setImageError("");
  }, [setPendingImages]);
  const handleDrop = useCallback(async (event: DragEvent<HTMLElement>) => {
    event.preventDefault(); event.stopPropagation(); setDraggingImage(false);
    const files = Array.from(event.dataTransfer.files).filter((file) => file.type.startsWith("image/"));
    if (!files.length) { setImageError("请拖入图片文件"); return; }
    await addImageFiles(files);
  }, [addImageFiles]);
  const handleImagePaste = useCallback((event: ClipboardEvent<HTMLTextAreaElement>) => {
    const files = Array.from(event.clipboardData.items)
      .filter((item) => item.kind === "file" && item.type.startsWith("image/"))
      .map((item) => item.getAsFile()).filter((file): file is File => file !== null);
    if (!files.length) return;
    event.preventDefault();
    void addImageFiles(files);
  }, [addImageFiles]);
  const handleDragEnter = useCallback((event: DragEvent<HTMLElement>) => {
    event.preventDefault(); event.stopPropagation();
    if (event.dataTransfer.types.includes("Files")) setDraggingImage(true);
  }, []);
  const handleDragLeave = useCallback((event: DragEvent<HTMLElement>) => {
    event.preventDefault(); event.stopPropagation();
    const target = event.relatedTarget instanceof Node ? event.relatedTarget : null;
    if (!event.currentTarget.contains(target)) setDraggingImage(false);
  }, []);
  const handleDragOver = useCallback((event: DragEvent<HTMLElement>) => {
    event.preventDefault(); event.stopPropagation();
  }, []);
  return { pendingImages, preparingImages, imageError, draggingImage, addImage, addImageFiles,
    removeImage, clearImages, handleDrop, handleImagePaste, handleDragEnter, handleDragLeave,
    handleDragOver, setPendingImages, setImageError, setDraggingImage };
}
