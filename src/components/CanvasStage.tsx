import type { PointerEvent } from "react";
import { useRef, useState } from "react";
import {
  defaultCalloutTargetOffset,
  defaultCircleSize,
  defaultFill,
  defaultFontSize,
  defaultLineLength,
  defaultRectangleSize,
  defaultStroke,
  defaultStrokeWidth,
  defaultTextBackground,
  defaultTextColor
} from "../editor/defaults";
import type { Annotation, LoadedImage, Tool } from "../editor/types";
import { AnnotationSvg, type ResizeHandle } from "../render/annotationSvg";

export function CanvasStage({
  image,
  annotations,
  activeTool,
  selectedId,
  onAddAt,
  onMove,
  onSelect,
  onUpdate
}: {
  image: LoadedImage | null;
  annotations: Annotation[];
  activeTool: Tool;
  selectedId: string | null;
  onAddAt: (x: number, y: number) => void;
  onSelect: (id: string | null) => void;
  onUpdate: (id: string, patch: Partial<Annotation>) => void;
  onMove: (id: string, delta: { x: number; y: number }) => void;
}) {
  const canvasRef = useRef<SVGSVGElement | null>(null);
  const [previewPoint, setPreviewPoint] = useState<{ x: number; y: number } | null>(
    null
  );
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

    setPreviewPoint(null);
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
    if (!image) return;

    const point = toImagePoint(event);

    if (!dragging) {
      setPreviewPoint(activeTool === "select" ? null : point);
      return;
    }

    const annotation = annotations.find((item) => item.id === dragging.id);
    if (!annotation) return;

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
    setPreviewPoint(null);
    onSelect(id);
    setDragging({ id, mode: "resize", handle, lastPoint: { x: 0, y: 0 } });
  }

  function handleMoveStart(
    id: string,
    event: PointerEvent<SVGElement>
  ) {
    const point = toImagePoint(event);
    setPreviewPoint(null);
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
        className={`canvas-stage tool-cursor-${activeTool}`}
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
        <ToolPreview activeTool={activeTool} point={previewPoint} />
      </svg>
    </section>
  );
}

function ToolPreview({
  activeTool,
  point
}: {
  activeTool: Tool;
  point: { x: number; y: number } | null;
}) {
  if (!point || activeTool === "select") return null;

  const common = {
    fill: defaultFill,
    opacity: 0.72,
    pointerEvents: "none" as const,
    stroke: defaultStroke,
    strokeWidth: defaultStrokeWidth
  };

  if (activeTool === "text") {
    const text = "零件名稱";
    return (
      <g aria-hidden="true" className="tool-preview" pointerEvents="none">
        <rect
          fill={defaultTextBackground}
          height={defaultFontSize + 10}
          opacity="0.72"
          rx="3"
          stroke={defaultStroke}
          strokeWidth={defaultStrokeWidth}
          width={text.length * defaultFontSize}
          x={point.x - 4}
          y={point.y - defaultFontSize}
        />
        <text
          dominantBaseline="alphabetic"
          fill={defaultTextColor}
          fontSize={defaultFontSize}
          opacity="0.72"
          x={point.x}
          y={point.y}
        >
          {text}
        </text>
      </g>
    );
  }

  if (activeTool === "line" || activeTool === "arrow") {
    return (
      <line
        {...common}
        aria-hidden="true"
        className="tool-preview"
        markerEnd={activeTool === "arrow" ? "url(#arrow-head)" : undefined}
        strokeLinecap="round"
        x1={point.x}
        x2={point.x + defaultLineLength}
        y1={point.y}
        y2={point.y}
      />
    );
  }

  if (activeTool === "callout") {
    const text = "零件名稱";
    return (
      <g aria-hidden="true" className="tool-preview" pointerEvents="none">
        <line
          {...common}
          markerEnd="url(#arrow-head)"
          strokeLinecap="round"
          x1={point.x}
          x2={point.x + defaultCalloutTargetOffset.x}
          y1={point.y}
          y2={point.y + defaultCalloutTargetOffset.y}
        />
        <rect
          fill={defaultTextBackground}
          height={defaultFontSize + 10}
          opacity="0.72"
          rx="3"
          stroke={defaultStroke}
          strokeWidth={defaultStrokeWidth}
          width={text.length * defaultFontSize}
          x={point.x - 4}
          y={point.y - defaultFontSize}
        />
        <text
          dominantBaseline="alphabetic"
          fill={defaultTextColor}
          fontSize={defaultFontSize}
          opacity="0.72"
          x={point.x}
          y={point.y}
        >
          {text}
        </text>
      </g>
    );
  }

  if (activeTool === "ellipse") {
    return (
      <ellipse
        {...common}
        aria-hidden="true"
        className="tool-preview"
        cx={point.x + defaultCircleSize.width / 2}
        cy={point.y + defaultCircleSize.height / 2}
        rx={defaultCircleSize.width / 2}
        ry={defaultCircleSize.height / 2}
      />
    );
  }

  return (
    <rect
      {...common}
      aria-hidden="true"
      className="tool-preview"
      height={defaultRectangleSize.height}
      strokeDasharray={activeTool === "dashedRectangle" ? "10 7" : undefined}
      width={defaultRectangleSize.width}
      x={point.x}
      y={point.y}
    />
  );
}
