// @vitest-environment jsdom
import { describe, it, expect, vi, beforeEach } from "vitest";
import { renderHook, act } from "@testing-library/react";
import type { DragEvent } from "react";
import {
  MAX_IMAGE_ATTACHMENTS,
  prepareImageFile,
  type PendingImage
} from "../imageAttachments";
import { useImageAttachments } from "./useImageAttachments";

vi.mock("../imageAttachments", () => ({
  MAX_IMAGE_ATTACHMENTS: 2,
  prepareImageFile: vi.fn()
}));

const prepareImageFileMock = vi.mocked(prepareImageFile);

function imageFile(name = "photo.png"): File {
  return new File(["image"], name, { type: "image/png" });
}

function pendingImage(name: string): PendingImage {
  return {
    id: `id-${name}`,
    name,
    mediaType: "image/png",
    data: "base64",
    previewUrl: `blob:${name}`
  };
}

function dropEvent(files: File[], types = ["Files"]) {
  return {
    preventDefault: vi.fn(),
    stopPropagation: vi.fn(),
    dataTransfer: {
      files,
      types
    }
  } as unknown as DragEvent<HTMLTextAreaElement>;
}

describe("useImageAttachments", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    prepareImageFileMock.mockImplementation(async (file) =>
      pendingImage(file.name)
    );
  });

  it("should initialize with empty image state", () => {
    const { result } = renderHook(() => useImageAttachments());

    expect(result.current.pendingImages).toEqual([]);
    expect(result.current.imageError).toBe("");
    expect(result.current.draggingImage).toBe(false);
  });

  it("should add a prepared image", async () => {
    const { result } = renderHook(() => useImageAttachments());

    await act(async () => {
      await result.current.addImage(imageFile("a.png"));
    });

    expect(prepareImageFileMock).toHaveBeenCalledTimes(1);
    expect(result.current.pendingImages).toHaveLength(1);
    expect(result.current.pendingImages[0].name).toBe("a.png");
  });

  it("should set an error when image preparation fails", async () => {
    prepareImageFileMock.mockRejectedValue(new Error("图片太大"));
    const { result } = renderHook(() => useImageAttachments());

    await act(async () => {
      await result.current.addImage(imageFile("bad.png"));
    });

    expect(result.current.pendingImages).toEqual([]);
    expect(result.current.imageError).toBe("图片太大");
  });

  it("should reject images beyond the limit", async () => {
    const { result } = renderHook(() => useImageAttachments());

    await act(async () => {
      await result.current.addImage(imageFile("a.png"));
    });
    await act(async () => {
      await result.current.addImage(imageFile("b.png"));
    });
    await act(async () => {
      await result.current.addImage(imageFile("c.png"));
    });

    expect(result.current.pendingImages).toHaveLength(MAX_IMAGE_ATTACHMENTS);
    expect(result.current.imageError).toContain("最多只能上传");
  });

  it("should add accepted files on drop and report excess", async () => {
    const { result } = renderHook(() => useImageAttachments());

    await act(async () => {
      await result.current.handleDrop(
        dropEvent([
          imageFile("a.png"),
          imageFile("b.png"),
          imageFile("c.png")
        ])
      );
    });

    expect(prepareImageFileMock).toHaveBeenCalledTimes(2);
    expect(result.current.pendingImages).toHaveLength(2);
    expect(result.current.imageError).toContain("最多只能上传");
  });

  it("should set an error when dropping non-image files", async () => {
    const { result } = renderHook(() => useImageAttachments());
    const event = dropEvent([new File(["x"], "notes.txt", { type: "text/plain" })], [
      "Files"
    ]);

    await act(async () => {
      await result.current.handleDrop(event);
    });

    expect(result.current.pendingImages).toEqual([]);
    expect(result.current.imageError).toBe("请拖入图片文件");
  });

  it("should mark dragging state on file drag enter", () => {
    const { result } = renderHook(() => useImageAttachments());
    const event = {
      preventDefault: vi.fn(),
      stopPropagation: vi.fn(),
      dataTransfer: {
        types: ["Files"],
        files: []
      }
    } as unknown as DragEvent<HTMLTextAreaElement>;

    act(() => result.current.handleDragEnter(event));

    expect(result.current.draggingImage).toBe(true);
    expect(event.preventDefault).toHaveBeenCalled();
    expect(event.stopPropagation).toHaveBeenCalled();
  });

  it("should not mark dragging when drop types do not include Files", () => {
    const { result } = renderHook(() => useImageAttachments());
    const event = {
      preventDefault: vi.fn(),
      stopPropagation: vi.fn(),
      dataTransfer: {
        types: ["text/plain"],
        files: []
      }
    } as unknown as DragEvent<HTMLTextAreaElement>;

    act(() => result.current.handleDragEnter(event));

    expect(result.current.draggingImage).toBe(false);
  });

  it("should remove an image by index", async () => {
    const { result } = renderHook(() => useImageAttachments());

    await act(async () => {
      await result.current.addImage(imageFile("a.png"));
      await result.current.addImage(imageFile("b.png"));
    });

    act(() => result.current.removeImage(0));

    expect(result.current.pendingImages).toHaveLength(1);
    expect(result.current.pendingImages[0].name).toBe("b.png");
  });

  it("should clear all images", async () => {
    const { result } = renderHook(() => useImageAttachments());

    await act(async () => {
      await result.current.addImage(imageFile("a.png"));
    });
    act(() => result.current.clearImages());

    expect(result.current.pendingImages).toEqual([]);
  });
});
