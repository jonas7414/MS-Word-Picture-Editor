import { fireEvent, render, screen } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";
import type { Annotation } from "../editor/types";
import { AnnotationSvg } from "./annotationSvg";

const annotations: Annotation[] = [
  {
    id: "text-1",
    type: "text",
    x: 12,
    y: 24,
    text: "馬達",
    fontSize: 18,
    color: "#111827",
    background: "#ffffff",
    stroke: "#dc2626",
    fill: "transparent",
    strokeWidth: 3
  },
  {
    id: "arrow-1",
    type: "arrow",
    x1: 50,
    y1: 60,
    x2: 140,
    y2: 90,
    stroke: "#dc2626",
    fill: "transparent",
    strokeWidth: 3
  }
];

describe("AnnotationSvg", () => {
  it("renders text and arrow annotations and reports selection", () => {
    const onSelect = vi.fn();
    const onMoveStart = vi.fn();
    const onResizeStart = vi.fn();

    render(
      <svg>
        <AnnotationSvg
          annotations={annotations}
          selectedId="text-1"
          onSelect={onSelect}
          onMoveStart={onMoveStart}
          onResizeStart={onResizeStart}
        />
      </svg>
    );

    expect(screen.getByText("馬達")).toBeInTheDocument();
    fireEvent.pointerDown(screen.getByLabelText("選取標註 馬達"));
    expect(onSelect).toHaveBeenCalledWith("text-1");
    expect(screen.getByTestId("selected-outline-text-1")).toBeInTheDocument();
    expect(screen.getByTestId("arrow-arrow-1")).toBeInTheDocument();
  });
});
