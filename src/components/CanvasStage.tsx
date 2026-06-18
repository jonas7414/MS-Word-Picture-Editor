import type { PointerEvent } from "react";
import { useRef, useState } from "react";
import type { Annotation, LoadedImage } from "../editor/types";
import { AnnotationSvg, type ResizeHandle } from "../render/annotationSvg";

export function CanvasStage({
  image,
  annotations,
  selectedId,
  onAddAt,
  onMove,
  onSelect,
  onUpdate
}: {
  image: LoadedImage | null;
  annotations: Annotation[];
  selectedId: string | null;
  onAddAt: (x: number, y: number) => void;
  onSelect: (id: string | null) => void;
  onUpdate: (id: string, patch: Partial<Annotation>) => void;
  onMove: (id: string, delta: { x: number; y: number }) => void;
}) {
  const canvasRef = useRef<SVGSVGElement | null>(null);
  const [dragging, setDragging] = useState<{
    id: string;
    mode: "move" | "resize";
    handle?: ResizeHandle;
    lastPoint: { x: number; y: number };
  } | null>(null);

  if (!image) {
    return (
      <section aria-label="尚未載入圖片" className="empty-state">
        <p>上傳圖片後即可加入文字、箭頭與基礎圖案。</p>
      </section>
    );
  }

  function handlePointerDown(event: PointerEvent<SVGSVGElement>) {
    if (!image) return;

    const point = toImagePoint(event);

    onSelect(null);
    onAddAt(point.x, point.y);
  }

  function toImagePoint(event: PointerEvent<SVGElement>) {
    if (!image) return { x: 0, y: 0 };

    const rect = (canvasRef.current ?? event.currentTarget).getBoundingClientRect();
    const width = rect.width || image.width;
    const height = rect.height || image.height;
    const scale = Math.min(width / image.width, height / image.height) || 1;
    const renderedWidth = image.width * scale;
    const renderedHeight = image.height * scale;
    const offsetX = (width - renderedWidth) / 2;
    const offsetY = (height - renderedHeight) / 2;
    const clientX = Number.isFinite(event.clientX) ? event.clientX : rect.left;
    const clientY = Number.isFinite(event.clientY) ? event.clientY : rect.top;
    const imageX = (clientX - rect.left - offsetX) / scale;
    const imageY = (clientY - rect.top - offsetY) / scale;

    return {
      x: Math.round(clamp(imageX, 0, image.width)),
      y: Math.round(clamp(imageY, 0, image.height))
    };
  }

  function clamp(value: number, min: number, max: number) {
    return Math.min(max, Math.max(min, value));
  }

  function handlePointerMove(event: PointerEvent<SVGSVGElement>) {
    if (!dragging || !image) return;

    const annotation = annotations.find((item) => item.id === dragging.id);
    if (!annotation) return;

    const point = toImagePoint(event);

    if (dragging.mode === "move") {
      onMove(annotation.id, {
        x: point.x - dragging.lastPoint.x,
        y: point.y - dragging.lastPoint.y
      });
      setDragging({ ...dragging, lastPoint: point });
      return;
    }

    if (
      dragging.handle === "shape-bottom-right" &&
      (annotation.type === "rectangle" || annotation.type === "ellipse")
    ) {
      onUpdate(annotation.id, {
        width: Math.max(12, point.x - annotation.x),
        height: Math.max(12, point.y - annotation.y)
      } as Partial<Annotation>);
      return;
    }

    if (
      dragging.handle === "line-end" &&
      (annotation.type === "line" || annotation.type === "arrow")
    ) {
      onUpdate(annotation.id, {
        x2: point.x,
        y2: point.y
      } as Partial<Annotation>);
      return;
    }

    if (dragging.handle === "callout-target" && annotation.type === "callout") {
      onUpdate(annotation.id, {
        targetX: point.x,
        targetY: point.y
      } as Partial<Annotation>);
    }
  }

  function handleResizeStart(id: string, handle: ResizeHandle) {
    onSelect(id);
    setDragging({ id, mode: "resize", handle, lastPoint: { x: 0, y: 0 } });
  }

  function handleMoveStart(
    id: string,
    event: PointerEvent<SVGElement>
  ) {
    const point = toImagePoint(event);
    onSelect(id);
    setDragging({
      id,
      mode: "move",
      lastPoint: point
    });
  }

  return (
    <section className="canvas-wrap">
      <svg
        aria-label="圖片標註畫布"
        className="canvas-stage"
        ref={canvasRef}
        viewBox={`0 0 ${image.width} ${image.height}`}
        onPointerDown={handlePointerDown}
        onPointerMove={handlePointerMove}
        onPointerUp={() => setDragging(null)}
        onPointerLeave={() => setDragging(null)}
      >
        <image height={image.height} href={image.src} width={image.width} />
        <AnnotationSvg
          annotations={annotations}
          selectedId={selectedId}
          onSelect={onSelect}
          onMoveStart={handleMoveStart}
          onResizeStart={handleResizeStart}
        />
      </svg>
    </section>
  );
}
