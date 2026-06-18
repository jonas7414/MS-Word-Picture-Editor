import {
  createEvent,
  fireEvent,
  render,
  screen,
  waitFor
} from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { App } from "./App";

vi.mock("./image/imageValidation", () => ({
  loadImageFile: vi.fn(async (file: File) => ({
    src: "data:image/png;base64,input",
    name: file.name,
    width: 640,
    height: 480
  }))
}));

vi.mock("./render/exportPng", () => ({
  exportAnnotatedPng: vi.fn(async () => "data:image/png;base64,exported")
}));

vi.mock("./word/wordClient", () => ({
  insertImageAtSelection: vi.fn(async () => undefined)
}));

describe("App workflow", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("uploads an image, adds a text label, edits it, and inserts into Word", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.upload(
      screen.getByLabelText("上傳圖片"),
      new File(["image"], "assembly.png", { type: "image/png" })
    );

    expect(await screen.findByText("assembly.png")).toBeInTheDocument();

    await user.click(screen.getByRole("button", { name: "文字" }));
    fireEvent.pointerDown(screen.getByLabelText("圖片標註畫布"), {
      clientX: 80,
      clientY: 100
    });

    expect(screen.getByDisplayValue("零件名稱")).toBeInTheDocument();
    await user.clear(screen.getByLabelText("標註文字"));
    await user.type(screen.getByLabelText("標註文字"), "馬達");

    expect(screen.getByText("馬達")).toBeInTheDocument();

    await user.click(screen.getByRole("button", { name: "插入 Word" }));

    await waitFor(() => {
      expect(screen.getByText("已插入 Word。")).toBeInTheDocument();
    });
  });

  it("resizes a rectangle by dragging its canvas handle", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.upload(
      screen.getByLabelText("上傳圖片"),
      new File(["image"], "assembly.png", { type: "image/png" })
    );

    await user.click(screen.getByRole("button", { name: "方框" }));
    const canvas = screen.getByLabelText("圖片標註畫布");
    expect(canvas).toHaveClass("tool-cursor-rectangle");
    canvas.getBoundingClientRect = () =>
      ({
        bottom: 480,
        height: 480,
        left: 0,
        right: 640,
        top: 0,
        width: 640,
        x: 0,
        y: 0,
        toJSON: () => ({})
      }) as DOMRect;
    const downEvent = createEvent.pointerDown(canvas);
    Object.defineProperty(downEvent, "clientX", { value: 20 });
    Object.defineProperty(downEvent, "clientY", { value: 30 });
    fireEvent(canvas, downEvent);

    const handle = document.querySelector('[data-testid^="resize-handle-"]');
    expect(handle).toBeInTheDocument();

    fireEvent.pointerDown(handle as Element);
    const moveEvent = createEvent.pointerMove(canvas);
    Object.defineProperty(moveEvent, "clientX", { value: 260 });
    Object.defineProperty(moveEvent, "clientY", { value: 190 });
    fireEvent(canvas, moveEvent);
    fireEvent.pointerUp(canvas);

    expect(screen.getByLabelText("寬")).toHaveValue(240);
    expect(screen.getByLabelText("高")).toHaveValue(160);
  });

  it("moves a rectangle by dragging the annotation body", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.upload(
      screen.getByLabelText("上傳圖片"),
      new File(["image"], "assembly.png", { type: "image/png" })
    );

    await user.click(screen.getByRole("button", { name: "方框" }));
    const canvas = screen.getByLabelText("圖片標註畫布");
    canvas.getBoundingClientRect = () =>
      ({
        bottom: 480,
        height: 480,
        left: 0,
        right: 640,
        top: 0,
        width: 640,
        x: 0,
        y: 0,
        toJSON: () => ({})
      }) as DOMRect;

    const addEvent = createEvent.pointerDown(canvas);
    Object.defineProperty(addEvent, "clientX", { value: 20 });
    Object.defineProperty(addEvent, "clientY", { value: 30 });
    fireEvent(canvas, addEvent);

    const annotation = screen.getByLabelText("選取標註 rectangle");
    const dragStart = createEvent.pointerDown(annotation);
    Object.defineProperty(dragStart, "clientX", { value: 20 });
    Object.defineProperty(dragStart, "clientY", { value: 30 });
    fireEvent(annotation, dragStart);

    const dragMove = createEvent.pointerMove(canvas);
    Object.defineProperty(dragMove, "clientX", { value: 70 });
    Object.defineProperty(dragMove, "clientY", { value: 90 });
    fireEvent(canvas, dragMove);
    fireEvent.pointerUp(canvas);

    const outline = document.querySelector('[data-testid^="selected-outline-"]');
    expect(outline).toHaveAttribute("x", "66");
    expect(outline).toHaveAttribute("y", "86");
    expect(screen.getByLabelText("寬")).toHaveValue(120);
    expect(screen.getByLabelText("高")).toHaveValue(72);
  });

  it("places annotations at the clicked image position when the canvas has letterboxing", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.upload(
      screen.getByLabelText("上傳圖片"),
      new File(["image"], "assembly.png", { type: "image/png" })
    );

    await user.click(screen.getByRole("button", { name: "方框" }));
    const canvas = screen.getByLabelText("圖片標註畫布");
    canvas.getBoundingClientRect = () =>
      ({
        bottom: 800,
        height: 800,
        left: 0,
        right: 640,
        top: 0,
        width: 640,
        x: 0,
        y: 0,
        toJSON: () => ({})
      }) as DOMRect;

    const addEvent = createEvent.pointerDown(canvas);
    Object.defineProperty(addEvent, "clientX", { value: 320 });
    Object.defineProperty(addEvent, "clientY", { value: 400 });
    fireEvent(canvas, addEvent);

    const outline = document.querySelector('[data-testid^="selected-outline-"]');
    expect(outline).toHaveAttribute("x", "316");
    expect(outline).toHaveAttribute("y", "236");
  });

  it("shows a same-size preview before placing a shape", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.upload(
      screen.getByLabelText("上傳圖片"),
      new File(["image"], "assembly.png", { type: "image/png" })
    );

    await user.click(screen.getByRole("button", { name: "方框" }));
    const canvas = screen.getByLabelText("圖片標註畫布");
    canvas.getBoundingClientRect = () =>
      ({
        bottom: 480,
        height: 480,
        left: 0,
        right: 640,
        top: 0,
        width: 640,
        x: 0,
        y: 0,
        toJSON: () => ({})
      }) as DOMRect;

    const moveEvent = createEvent.pointerMove(canvas);
    Object.defineProperty(moveEvent, "clientX", { value: 40 });
    Object.defineProperty(moveEvent, "clientY", { value: 50 });
    fireEvent(canvas, moveEvent);

    const preview = document.querySelector(".tool-preview");
    expect(preview).toHaveAttribute("width", "120");
    expect(preview).toHaveAttribute("height", "72");

    const addEvent = createEvent.pointerDown(canvas);
    Object.defineProperty(addEvent, "clientX", { value: 40 });
    Object.defineProperty(addEvent, "clientY", { value: 50 });
    fireEvent(canvas, addEvent);

    expect(screen.getByLabelText("寬")).toHaveValue(120);
    expect(screen.getByLabelText("高")).toHaveValue(72);
  });

  it("allows thick annotation strokes up to 96", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.upload(
      screen.getByLabelText("上傳圖片"),
      new File(["image"], "assembly.png", { type: "image/png" })
    );

    await user.click(screen.getByRole("button", { name: "方框" }));
    const canvas = screen.getByLabelText("圖片標註畫布");
    canvas.getBoundingClientRect = () =>
      ({
        bottom: 480,
        height: 480,
        left: 0,
        right: 640,
        top: 0,
        width: 640,
        x: 0,
        y: 0,
        toJSON: () => ({})
      }) as DOMRect;

    const addEvent = createEvent.pointerDown(canvas);
    Object.defineProperty(addEvent, "clientX", { value: 20 });
    Object.defineProperty(addEvent, "clientY", { value: 30 });
    fireEvent(canvas, addEvent);

    const strokeWidth = screen.getByLabelText("線寬");
    expect(strokeWidth).toHaveAttribute("max", "96");
    await user.clear(strokeWidth);
    await user.type(strokeWidth, "96");
    expect(strokeWidth).toHaveValue(96);
  });
});
