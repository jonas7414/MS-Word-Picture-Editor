import type { LoadedImage } from "../editor/types";

export const maxImageBytes = 10 * 1024 * 1024;

const supportedTypes = new Set(["image/png", "image/jpeg", "image/webp"]);

export type ImageValidationResult =
  | { ok: true }
  | { ok: false; message: string };

export function validateImageFile(file: File): ImageValidationResult {
  if (!supportedTypes.has(file.type)) {
    return { ok: false, message: "僅支援 PNG、JPG、JPEG、WEBP 圖片。" };
  }

  if (file.size > maxImageBytes) {
    return { ok: false, message: "圖片不可超過 10 MB。" };
  }

  return { ok: true };
}

export async function loadImageFile(file: File): Promise<LoadedImage> {
  const validation = validateImageFile(file);
  if (!validation.ok) {
    throw new Error(validation.message);
  }

  const src = await readFileAsDataUrl(file);
  const dimensions = await readImageDimensions(src);

  return {
    src,
    name: file.name,
    width: dimensions.width,
    height: dimensions.height
  };
}

function readFileAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(new Error("無法讀取圖片。"));
    reader.readAsDataURL(file);
  });
}

export function readImageDimensions(
  src: string
): Promise<{ width: number; height: number }> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () =>
      resolve({
        width: image.naturalWidth,
        height: image.naturalHeight
      });
    image.onerror = () => reject(new Error("無法載入圖片。"));
    image.src = src;
  });
}
