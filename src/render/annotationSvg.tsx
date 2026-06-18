import type { PointerEvent, PointerEventHandler } from "react";
import type { Annotation } from "../editor/types";

type Props = {
  annotations: Annotation[];
  selectedId: string | null;
  onSelect: (id: string) => void;
  onMoveStart: (id: string, event: PointerEvent<SVGGElement>) => void;
  onResizeStart: (id: string, handle: ResizeHandle) => void;
};

export type ResizeHandle = "shape-bottom-right" | "line-end" | "callout-target";

type AnnotationNodeProps = {
  annotation: Annotation;
  selected: boolean;
  onSelect: (id: string) => void;
  onMoveStart: (id: string, event: PointerEvent<SVGGElement>) => void;
  onResizeStart: (id: string, handle: ResizeHandle) => void;
};

const selectedStroke = "#2563eb";
const textPaddingX = 4;
const textPaddingY = 5;
const dashedStroke = "10 7";

export function AnnotationSvg({
  annotations,
  selectedId,
  onSelect,
  onMoveStart,
  onResizeStart
}: Props) {
  return (
    <g>
      <defs>
        <marker
          id="arrow-head"
          markerHeight="10"
          markerUnits="strokeWidth"
          markerWidth="10"
          orient="auto"
          refX="8"
          refY="3"
        >
          <path d="M0,0 L0,6 L9,3 z" fill="#dc2626" />
        </marker>
      </defs>
      {annotations.map((annotation) => (
        <AnnotationNode
          key={annotation.id}
          annotation={annotation}
          selected={annotation.id === selectedId}
          onSelect={onSelect}
          onMoveStart={onMoveStart}
          onResizeStart={onResizeStart}
        />
      ))}
    </g>
  );
}

function AnnotationNode({
  annotation,
  selected,
  onSelect,
  onMoveStart,
  onResizeStart
}: AnnotationNodeProps) {
  const handlePointerDown: PointerEventHandler<SVGGElement> = (event) => {
    event.stopPropagation();
    onSelect(annotation.id);
    onMoveStart(annotation.id, event);
  };

  const common = {
    onPointerDown: handlePointerDown,
    role: "button",
    style: { cursor: selected ? "move" : "pointer" }
  };

  if (annotation.type === "text") {
    return (
      <g {...common} aria-label={`選取標註 ${annotation.text}`}>
        <TextBox annotation={annotation} selected={selected} />
      </g>
    );
  }

  if (annotation.type === "line" || annotation.type === "arrow") {
    return (
      <g {...common} aria-label={`選取標註 ${annotation.type}`}>
        <line
          data-testid={`${annotation.type}-${annotation.id}`}
          markerEnd={annotation.type === "arrow" ? "url(#arrow-head)" : undefined}
          stroke={annotation.stroke}
          strokeDasharray={annotation.strokeStyle === "dashed" ? dashedStroke : undefined}
          strokeLinecap="round"
          strokeWidth={annotation.strokeWidth}
          x1={annotation.x1}
          x2={annotation.x2}
          y1={annotation.y1}
          y2={annotation.y2}
        />
        {selected ? (
          <circle
            cx={annotation.x2}
            cy={annotation.y2}
            data-testid={`selected-outline-${annotation.id}`}
            fill={selectedStroke}
            onPointerDown={(event) => {
              event.stopPropagation();
              onResizeStart(annotation.id, "line-end");
            }}
            r="5"
            style={{ cursor: "nwse-resize" }}
          />
        ) : null}
      </g>
    );
  }

  if (annotation.type === "callout") {
    return (
      <g {...common} aria-label={`選取標註 ${annotation.text}`}>
        <line
          markerEnd="url(#arrow-head)"
          stroke={annotation.stroke}
          strokeDasharray={annotation.strokeStyle === "dashed" ? dashedStroke : undefined}
          strokeLinecap="round"
          strokeWidth={annotation.strokeWidth}
          x1={annotation.x}
          x2={annotation.targetX}
          y1={annotation.y}
          y2={annotation.targetY}
        />
        <TextBox annotation={annotation} selected={selected} />
        {selected ? (
          <circle
            cx={annotation.targetX}
            cy={annotation.targetY}
            data-testid={`resize-handle-${annotation.id}`}
            fill={selectedStroke}
            onPointerDown={(event) => {
              event.stopPropagation();
              onResizeStart(annotation.id, "callout-target");
            }}
            r="5"
            style={{ cursor: "crosshair" }}
          />
        ) : null}
      </g>
    );
  }

  if (annotation.type === "rectangle" || annotation.type === "ellipse") {
    const outline = selected ? (
      <rect
        data-testid={`selected-outline-${annotation.id}`}
        fill="transparent"
        height={annotation.height + 8}
        stroke={selectedStroke}
        strokeDasharray="4 4"
        width={annotation.width + 8}
        x={annotation.x - 4}
        y={annotation.y - 4}
      />
    ) : null;
    const resizeHandle = selected ? (
      <rect
        data-testid={`resize-handle-${annotation.id}`}
        fill={selectedStroke}
        height="9"
        onPointerDown={(event) => {
          event.stopPropagation();
          onResizeStart(annotation.id, "shape-bottom-right");
        }}
        rx="2"
        style={{ cursor: "nwse-resize" }}
        width="9"
        x={annotation.x + annotation.width - 4.5}
        y={annotation.y + annotation.height - 4.5}
      />
    ) : null;

    if (annotation.type === "rectangle") {
      return (
        <g {...common} aria-label={`選取標註 ${annotation.type}`}>
          <rect
            fill={annotation.fill}
            height={annotation.height}
            stroke={annotation.stroke}
            strokeDasharray={annotation.strokeStyle === "dashed" ? dashedStroke : undefined}
            strokeWidth={annotation.strokeWidth}
            width={annotation.width}
            x={annotation.x}
            y={annotation.y}
          />
          {outline}
          {resizeHandle}
        </g>
      );
    }

    return (
      <g {...common} aria-label={`選取標註 ${annotation.type}`}>
        <ellipse
          cx={annotation.x + annotation.width / 2}
          cy={annotation.y + annotation.height / 2}
          fill={annotation.fill}
          rx={annotation.width / 2}
          ry={annotation.height / 2}
          stroke={annotation.stroke}
          strokeDasharray={annotation.strokeStyle === "dashed" ? dashedStroke : undefined}
          strokeWidth={annotation.strokeWidth}
        />
        {outline}
        {resizeHandle}
      </g>
    );
  }

  return null;
}

function TextBox({
  annotation,
  selected
}: {
  annotation: Extract<Annotation, { type: "text" | "callout" }>;
  selected: boolean;
}) {
  return (
    <>
      <rect
        data-testid={selected ? `selected-outline-${annotation.id}` : undefined}
        fill={annotation.background}
        height={annotation.fontSize + textPaddingY * 2}
        rx="3"
        stroke={selected ? selectedStroke : annotation.stroke}
        width={annotation.text.length * annotation.fontSize}
        x={annotation.x - textPaddingX}
        y={annotation.y - annotation.fontSize}
      />
      <text
        dominantBaseline="alphabetic"
        fill={annotation.color}
        fontSize={annotation.fontSize}
        x={annotation.x}
        y={annotation.y}
      >
        {annotation.text}
      </text>
    </>
  );
}
