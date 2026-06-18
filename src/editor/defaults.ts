import type {
  Annotation,
  CalloutAnnotation,
  LineAnnotation,
  Point,
  ShapeAnnotation,
  TextAnnotation,
  Tool
} from "./types";

export const defaultStroke = "#dc2626";
export const defaultFill = "transparent";
export const defaultTextColor = "#111827";
export const defaultTextBackground = "#ffffff";
export const defaultStrokeWidth = 3;
export const defaultFontSize = 18;
export const defaultStrokeStyle = "solid" as const;
export const defaultLineLength = 96;
export const defaultCalloutTargetOffset = { x: 96, y: 40 };
export const defaultRectangleSize = { width: 120, height: 72 };
export const defaultCircleSize = { width: 96, height: 96 };

let nextId = 1;

export function resetAnnotationIdsForTests() {
  nextId = 1;
}

function createId() {
  const id = `annotation-${nextId}`;
  nextId += 1;
  return id;
}

export function createAnnotation(tool: "text", point: Point): TextAnnotation;
export function createAnnotation(
  tool: "line" | "arrow",
  point: Point
): LineAnnotation;
export function createAnnotation(
  tool: "rectangle" | "dashedRectangle" | "ellipse",
  point: Point
): ShapeAnnotation;
export function createAnnotation(
  tool: "callout",
  point: Point
): CalloutAnnotation;
export function createAnnotation(
  tool: Exclude<Tool, "select">,
  point: Point
): Annotation {
  const base = {
    id: createId(),
    stroke: defaultStroke,
    fill: defaultFill,
    strokeWidth: defaultStrokeWidth,
    strokeStyle: defaultStrokeStyle
  };

  if (tool === "text") {
    return {
      ...base,
      type: "text",
      x: point.x,
      y: point.y,
      text: "零件名稱",
      fontSize: defaultFontSize,
      color: defaultTextColor,
      background: defaultTextBackground
    };
  }

  if (tool === "line" || tool === "arrow") {
    return {
      ...base,
      type: tool,
      x1: point.x,
      y1: point.y,
      x2: point.x + defaultLineLength,
      y2: point.y
    };
  }

  if (tool === "callout") {
    return {
      ...base,
      type: "callout",
      x: point.x,
      y: point.y,
      text: "零件名稱",
      fontSize: defaultFontSize,
      color: defaultTextColor,
      background: defaultTextBackground,
      targetX: point.x + defaultCalloutTargetOffset.x,
      targetY: point.y + defaultCalloutTargetOffset.y
    };
  }

  return {
    ...base,
    type: tool === "dashedRectangle" ? "rectangle" : tool,
    strokeStyle: tool === "dashedRectangle" ? "dashed" : base.strokeStyle,
    x: point.x,
    y: point.y,
    width:
      tool === "ellipse" ? defaultCircleSize.width : defaultRectangleSize.width,
    height:
      tool === "ellipse" ? defaultCircleSize.height : defaultRectangleSize.height
  };
}
