import { describe, expect, it } from "vitest";
import { moveAnnotation, resizeShape } from "./geometry";
import type { Annotation } from "./types";

describe("geometry helpers", () => {
  it("moves line endpoints together", () => {
    const line: Annotation = {
      id: "line-1",
      type: "line",
      x1: 10,
      y1: 20,
      x2: 30,
      y2: 40,
      stroke: "#dc2626",
      fill: "transparent",
      strokeWidth: 3
    };

    expect(moveAnnotation(line, { x: 5, y: -10 })).toMatchObject({
      x1: 15,
      y1: 10,
      x2: 35,
      y2: 30
    });
  });

  it("moves callout label and target together", () => {
    const callout: Annotation = {
      id: "callout-1",
      type: "callout",
      x: 10,
      y: 20,
      text: "零件名稱",
      fontSize: 18,
      color: "#111827",
      background: "#ffffff",
      targetX: 90,
      targetY: 60,
      stroke: "#dc2626",
      fill: "transparent",
      strokeWidth: 3
    };

    expect(moveAnnotation(callout, { x: -2, y: 6 })).toMatchObject({
      x: 8,
      y: 26,
      targetX: 88,
      targetY: 66
    });
  });

  it("resizes rectangles with minimum dimensions", () => {
    const rectangle: Annotation = {
      id: "box-1",
      type: "rectangle",
      x: 10,
      y: 20,
      width: 80,
      height: 40,
      stroke: "#dc2626",
      fill: "transparent",
      strokeWidth: 3
    };

    expect(resizeShape(rectangle, 3, 4)).toMatchObject({
      width: 12,
      height: 12
    });
  });
});
