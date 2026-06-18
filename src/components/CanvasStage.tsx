import type { PointerEvent, WheelEvent } from "react";
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
  zoom,
  pan,
  selectedId,
  onAddAt,
  onPan,
  onZoom,
  onMove,
  onSelect,
  onUpdate
}: {
  image: LoadedImage | null;
  annotations: Annotation[];
  activeTool: Tool;
  zoom: number;
  pan: { x: number; y: number };
  selectedId: string | null;
  onAddAt: (x: number, y: number) => void;
  onPan: (delta: { x: number; y: number }) => void;
  onZoom: (zoom: number, anchor: { x: number; y: number }) => void;
  onSelect: (id: string | null) => void;
  onUpdate: (id: string, patch: Partial<Annotation>) => void;
  onMove: (id: string, delta: { x: number; y: number }) => void;
}) {
  const canvasRef = useRef<SVGSVGElement | null>(null);
  const [previewPoint, setPreviewPoint] = useState<{ x: number; y: number } | null>(
    null
  );
  const [dragging, setDragging] = useState<{
    id?: string;
    mode: "move" | "resize" | "pan";
    handle?: ResizeHandle;
    lastClient?: { x: number; y: number };
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

    if (activeTool === "select" && isCanvasBackgroundTarget(event.target)) {
      setPreviewPoint(null);
      onSelect(null);
      setDragging({
        mode: "pan",
        lastClient: { x: event.clientX, y: event.clientY },
        lastPoint: point
      });
      return;
    }

    setPreviewPoint(null);
    onSelect(null);
    onAddAt(point.x, point.y);
  }

  function toImagePoint(event: PointerEvent<SVGElement>) {
    return clientPointToImagePoint(event.clientX, event.clientY, event.currentTarget);
  }

  function clientPointToImagePoint(
    clientXValue: number,
    clientYValue: number,
    fallbackTarget: Element
  ) {
    if (!image) return { x: 0, y: 0 };

    const viewBox = getViewBox();
    const rect = (canvasRef.current ?? fallbackTarget).getBoundingClientRect();
    const width = rect.width || image.width;
    const height = rect.height || image.height;
    const scale = Math.min(width / viewBox.width, height / viewBox.height) || 1;
    const renderedWidth = viewBox.width * scale;
    const renderedHeight = viewBox.height * scale;
    const offsetX = (width - renderedWidth) / 2;
    const offsetY = (height - renderedHeight) / 2;
    const clientX = Number.isFinite(clientXValue) ? clientXValue : rect.left;
    const clientY = Number.isFinite(clientYValue) ? clientYValue : rect.top;
    const imageX = viewBox.x + (clientX - rect.left - offsetX) / scale;
    const imageY = viewBox.y + (clientY - rect.top - offsetY) / scale;

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

    if (dragging.mode === "pan") {
      const lastClient = dragging.lastClient ?? {
        x: event.clientX,
        y: event.clientY
      };
      const imageDelta = clientDeltaToImageDelta({
        x: event.clientX - lastClient.x,
        y: event.clientY - lastClient.y
      });
      onPan({ x: -imageDelta.x, y: -imageDelta.y });
      setDragging({
        ...dragging,
        lastClient: { x: event.clientX, y: event.clientY },
        lastPoint: point
      });
      return;
    }

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

  function handleWheel(event: WheelEvent<SVGSVGElement>) {
    if (!image || !event.ctrlKey) return;

    event.preventDefault();
    const anchor = clientPointToImagePoint(
      event.clientX,
      event.clientY,
      event.currentTarget
    );
    const zoomFactor = event.deltaY < 0 ? 1.15 : 1 / 1.15;
    onZoom(zoom * zoomFactor, anchor);
  }

  function clientDeltaToImageDelta(delta: { x: number; y: number }) {
    if (!image) return { x: 0, y: 0 };

    const rect = canvasRef.current?.getBoundingClientRect();
    const viewBox = getViewBox();
    const width = rect?.width || image.width;
    const height = rect?.height || image.height;
    const scale = Math.min(width / viewBox.width, height / viewBox.height) || 1;

    return {
      x: delta.x / scale,
      y: delta.y / scale
    };
  }

  function isCanvasBackgroundTarget(target: EventTarget) {
    return (
      target === canvasRef.current ||
      (target instanceof Element &&
        target.tagName.toLowerCase() === "image" &&
        target.parentElement === canvasRef.current)
    );
  }

  function handleResizeStart(id: string, handle: ResizeHandle) {
    setPreviewPoint(null);
    onSelect(id);
    setDragging({ id, mode: "resize", handle, lastPoint: { x: 0, y: 0 } });
  }

  function getViewBox() {
    if (!image) return { x: 0, y: 0, width: 1, height: 1 };

    const viewWidth = image.width / zoom;
    const viewHeight = image.height / zoom;
    return {
      x: clamp(pan.x, 0, Math.max(0, image.width - viewWidth)),
      y: clamp(pan.y, 0, Math.max(0, image.height - viewHeight)),
      width: viewWidth,
      height: viewHeight
    };
  }

  const viewBox = getViewBox();

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
        viewBox={`${viewBox.x} ${viewBox.y} ${viewBox.width} ${viewBox.height}`}
        onPointerDown={handlePointerDown}
        onPointerMove={handlePointerMove}
        onPointerUp={() => setDragging(null)}
        onPointerLeave={() => setDragging(null)}
        onWheel={handleWheel}
      >
        <image height={image.height} href={image.src} width={image.width} />
        <AnnotationSvg
          annotations={annotations}
          selectedId={selectedId}
          onSelect={onSelect}
          onMoveStart={handleMoveStart}
          onResizeStart={handleResizeStart}
        />
        <ToolPreview activeTool={activeTool} point={previewPoint} zoom={zoom} />
      </svg>
    </section>
  );
}

function ToolPreview({
  activeTool,
  point,
  zoom
}: {
  activeTool: Tool;
  point: { x: number; y: number } | null;
  zoom: number;
}) {
  if (!point || activeTool === "select") return null;
  const scale = 1 / zoom;

  const common = {
    fill: defaultFill,
    opacity: 0.72,
    pointerEvents: "none" as const,
    stroke: defaultStroke,
    strokeWidth: Math.max(1, defaultStrokeWidth * scale)
  };

  if (activeTool === "text") {
    const text = "零件名稱";
    return (
      <g aria-hidden="true" className="tool-preview" pointerEvents="none">
        <rect
          fill={defaultTextBackground}
          height={(defaultFontSize + 10) * scale}
          opacity="0.72"
          rx="3"
          stroke={defaultStroke}
          strokeWidth={Math.max(1, defaultStrokeWidth * scale)}
          width={text.length * defaultFontSize * scale}
          x={point.x - 4 * scale}
          y={point.y - defaultFontSize * scale}
        />
        <text
          dominantBaseline="alphabetic"
          fill={defaultTextColor}
          fontSize={Math.max(6, defaultFontSize * scale)}
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
        x2={point.x + defaultLineLength * scale}
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
          x2={point.x + defaultCalloutTargetOffset.x * scale}
          y1={point.y}
          y2={point.y + defaultCalloutTargetOffset.y * scale}
        />
        <rect
          fill={defaultTextBackground}
          height={(defaultFontSize + 10) * scale}
          opacity="0.72"
          rx="3"
          stroke={defaultStroke}
          strokeWidth={Math.max(1, defaultStrokeWidth * scale)}
          width={text.length * defaultFontSize * scale}
          x={point.x - 4 * scale}
          y={point.y - defaultFontSize * scale}
        />
        <text
          dominantBaseline="alphabetic"
          fill={defaultTextColor}
          fontSize={Math.max(6, defaultFontSize * scale)}
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
        cx={point.x + (defaultCircleSize.width * scale) / 2}
        cy={point.y + (defaultCircleSize.height * scale) / 2}
        rx={(defaultCircleSize.width * scale) / 2}
        ry={(defaultCircleSize.height * scale) / 2}
      />
    );
  }

  return (
    <rect
      {...common}
      aria-hidden="true"
      className="tool-preview"
      height={defaultRectangleSize.height * scale}
      strokeDasharray={activeTool === "dashedRectangle" ? "10 7" : undefined}
      width={defaultRectangleSize.width * scale}
      x={point.x}
      y={point.y}
    />
  );
}
