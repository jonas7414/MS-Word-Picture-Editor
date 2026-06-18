import { afterEach, describe, expect, it, vi } from "vitest";
import type { Annotation, LoadedImage } from "../editor/types";
import { exportAnnotatedPng } from "./exportPng";

describe("exportAnnotatedPng", () => {
  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("returns a PNG data URL and draws text and rectangle annotations", async () => {
    const drawImage = vi.fn();
    const fillText = vi.fn();
    const strokeRect = vi.fn();
    const context = {
      drawImage,
      fillText,
      strokeRect,
      beginPath: vi.fn(),
      closePath: vi.fn(),
      ellipse: vi.fn(),
      fill: vi.fn(),
      fillRect: vi.fn(),
      lineTo: vi.fn(),
      moveTo: vi.fn(),
      restore: vi.fn(),
      save: vi.fn(),
      setLineDash: vi.fn(),
      stroke: vi.fn(),
      set fillStyle(_value: string) {},
      set font(_value: string) {},
      set lineWidth(_value: number) {},
      set strokeStyle(_value: string) {}
    };

    const originalCreateElement = document.createElement.bind(document);
    vi.spyOn(document, "createElement").mockImplementation((tagName: string) => {
      if (tagName === "canvas") {
        return {
          width: 0,
          height: 0,
          getContext: () => context,
          toDataURL: () => "data:image/png;base64,exported"
        } as unknown as HTMLCanvasElement;
      }

      return originalCreateElement(tagName);
    });

    class MockImage {
      onerror: (() => void) | null = null;
      onload: (() => void) | null = null;

      set src(_value: string) {
        queueMicrotask(() => this.onload?.());
      }
    }

    vi.stubGlobal("Image", MockImage);

    const image: LoadedImage = {
      src: "data:image/png;base64,input",
      name: "part.png",
      width: 640,
      height: 480
    };
    const annotations: Annotation[] = [
      {
        id: "text-1",
        type: "text",
        x: 20,
        y: 40,
        text: "馬達",
        fontSize: 18,
        color: "#111827",
        background: "#ffffff",
        stroke: "#dc2626",
        fill: "transparent",
        strokeWidth: 3
      },
      {
        id: "box-1",
        type: "rectangle",
        x: 100,
        y: 120,
        width: 80,
        height: 40,
        stroke: "#dc2626",
        fill: "transparent",
        strokeWidth: 3
      }
    ];

    const result = await exportAnnotatedPng({ image, annotations });

    expect(result).toBe("data:image/png;base64,exported");
    expect(drawImage).toHaveBeenCalled();
    expect(fillText).toHaveBeenCalledWith("馬達", 20, 40);
    expect(strokeRect).toHaveBeenCalledWith(100, 120, 80, 40);
  });
});
