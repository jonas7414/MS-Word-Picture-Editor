import type { Annotation, LoadedImage } from "../editor/types";

const arrowHeadLength = 14;
const textPaddingX = 4;
const textPaddingY = 5;

export async function exportAnnotatedPng({
  image,
  annotations
}: {
  image: LoadedImage;
  annotations: Annotation[];
}): Promise<string> {
  const bitmap = await loadBitmap(image.src);
  const canvas = document.createElement("canvas");
  canvas.width = image.width;
  canvas.height = image.height;

  const context = canvas.getContext("2d");
  if (!context) {
    throw new Error("無法建立匯出畫布。");
  }

  context.drawImage(bitmap, 0, 0, image.width, image.height);
  annotations.forEach((annotation) => drawAnnotation(context, annotation));

  return canvas.toDataURL("image/png");
}

function loadBitmap(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("無法載入匯出圖片。"));
    image.src = src;
  });
}

function drawAnnotation(
  context: CanvasRenderingContext2D,
  annotation: Annotation
) {
  context.save();
  context.strokeStyle = annotation.stroke;
  context.lineWidth = annotation.strokeWidth;
  applyStrokeStyle(context, annotation.strokeStyle);

  if (annotation.type === "text") {
    drawText(context, annotation);
    context.restore();
    return;
  }

  if (annotation.type === "line" || annotation.type === "arrow") {
    drawLine(context, annotation.x1, annotation.y1, annotation.x2, annotation.y2);
    if (annotation.type === "arrow") {
      drawArrowHead(
        context,
        annotation.x1,
        annotation.y1,
        annotation.x2,
        annotation.y2
      );
    }
    context.restore();
    return;
  }

  if (annotation.type === "rectangle") {
    context.fillStyle = annotation.fill;
    if (annotation.fill !== "transparent") {
      context.fillRect(
        annotation.x,
        annotation.y,
        annotation.width,
        annotation.height
      );
    }
    context.strokeRect(
      annotation.x,
      annotation.y,
      annotation.width,
      annotation.height
    );
    context.restore();
    return;
  }

  if (annotation.type === "ellipse") {
    context.beginPath();
    context.ellipse(
      annotation.x + annotation.width / 2,
      annotation.y + annotation.height / 2,
      annotation.width / 2,
      annotation.height / 2,
      0,
      0,
      Math.PI * 2
    );
    context.fillStyle = annotation.fill;
    if (annotation.fill !== "transparent") {
      context.fill();
    }
    context.stroke();
    context.restore();
    return;
  }

  if (annotation.type === "callout") {
    drawLine(
      context,
      annotation.x,
      annotation.y,
      annotation.targetX,
      annotation.targetY
    );
    drawArrowHead(
      context,
      annotation.x,
      annotation.y,
      annotation.targetX,
      annotation.targetY
    );
    drawText(context, annotation);
  }
  context.restore();
}

function drawText(
  context: CanvasRenderingContext2D,
  annotation: Extract<Annotation, { type: "text" | "callout" }>
) {
  context.font = `${annotation.fontSize}px sans-serif`;
  context.fillStyle = annotation.background;
  context.fillRect(
    annotation.x - textPaddingX,
    annotation.y - annotation.fontSize,
    annotation.text.length * annotation.fontSize,
    annotation.fontSize + textPaddingY * 2
  );
  context.fillStyle = annotation.color;
  context.fillText(annotation.text, annotation.x, annotation.y);
}

function drawLine(
  context: CanvasRenderingContext2D,
  x1: number,
  y1: number,
  x2: number,
  y2: number
) {
  context.beginPath();
  context.moveTo(x1, y1);
  context.lineTo(x2, y2);
  context.stroke();
}

function applyStrokeStyle(
  context: CanvasRenderingContext2D,
  strokeStyle: Annotation["strokeStyle"]
) {
  context.setLineDash(strokeStyle === "dashed" ? [10, 7] : []);
}

function drawArrowHead(
  context: CanvasRenderingContext2D,
  x1: number,
  y1: number,
  x2: number,
  y2: number
) {
  const angle = Math.atan2(y2 - y1, x2 - x1);

  context.beginPath();
  context.moveTo(x2, y2);
  context.lineTo(
    x2 - arrowHeadLength * Math.cos(angle - Math.PI / 6),
    y2 - arrowHeadLength * Math.sin(angle - Math.PI / 6)
  );
  context.lineTo(
    x2 - arrowHeadLength * Math.cos(angle + Math.PI / 6),
    y2 - arrowHeadLength * Math.sin(angle + Math.PI / 6)
  );
  context.closePath();
  context.fillStyle = String(context.strokeStyle);
  context.fill();
}
