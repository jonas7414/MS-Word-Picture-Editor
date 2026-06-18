import type { Annotation, Point } from "./types";

export function moveAnnotation(annotation: Annotation, delta: Point): Annotation {
  switch (annotation.type) {
    case "line":
    case "arrow":
      return {
        ...annotation,
        x1: annotation.x1 + delta.x,
        y1: annotation.y1 + delta.y,
        x2: annotation.x2 + delta.x,
        y2: annotation.y2 + delta.y
      };
    case "callout":
      return {
        ...annotation,
        x: annotation.x + delta.x,
        y: annotation.y + delta.y,
        targetX: annotation.targetX + delta.x,
        targetY: annotation.targetY + delta.y
      };
    case "text":
    case "rectangle":
    case "ellipse":
      return {
        ...annotation,
        x: annotation.x + delta.x,
        y: annotation.y + delta.y
      };
  }
}

export function resizeShape(
  annotation: Annotation,
  width: number,
  height: number
): Annotation {
  if (annotation.type !== "rectangle" && annotation.type !== "ellipse") {
    return annotation;
  }

  return {
    ...annotation,
    width: Math.max(12, width),
    height: Math.max(12, height)
  };
}
