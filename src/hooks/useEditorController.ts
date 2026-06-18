import { useCallback, useEffect, useMemo, useReducer, useState } from "react";
import { createAnnotation } from "../editor/defaults";
import { editorReducer, initialEditorState } from "../editor/reducer";
import type { Annotation, LoadedImage, Point, Tool } from "../editor/types";
import { loadImageFile } from "../image/imageValidation";

export type EditorController = {
  state: typeof initialEditorState;
  status: string;
  error: string;
  selectedAnnotation: Annotation | null;
  setTool: (tool: Tool) => void;
  loadImage: (image: LoadedImage) => void;
  loadFile: (file: File) => Promise<void>;
  addAnnotationAt: (x: number, y: number) => void;
  selectAnnotation: (id: string | null) => void;
  updateAnnotation: (id: string, patch: Partial<Annotation>) => void;
  updateSelected: (patch: Partial<Annotation>) => void;
  deleteSelected: () => void;
  duplicateSelected: () => void;
  moveSelected: (delta: Point) => void;
  moveAnnotation: (id: string, delta: Point) => void;
  resizeSelected: (width: number, height: number) => void;
  setZoom: (zoom: number, anchor?: Point) => void;
  panViewport: (delta: Point) => void;
  resetViewport: () => void;
  undo: () => void;
  redo: () => void;
  setStatus: (message: string) => void;
  setError: (message: string) => void;
};

export function useEditorController(): EditorController {
  const [state, dispatch] = useReducer(editorReducer, initialEditorState);
  const [status, setStatus] = useState("");
  const [error, setError] = useState("");

  const selectedAnnotation = useMemo(
    () =>
      state.annotations.find((annotation) => annotation.id === state.selectedId) ??
      null,
    [state.annotations, state.selectedId]
  );

  const loadFile = useCallback(async (file: File) => {
    setError("");
    setStatus("正在載入圖片...");

    try {
      const image = await loadImageFile(file);
      dispatch({ type: "image/set", image });
      setStatus("圖片已載入。");
    } catch (caught) {
      const message = caught instanceof Error ? caught.message : "無法載入圖片。";
      setError(message);
      setStatus("");
    }
  }, []);

  const loadImage = useCallback((image: LoadedImage) => {
    setError("");
    dispatch({ type: "image/set", image });
    setStatus("圖片已載入。");
  }, []);

  const addAnnotationAt = useCallback(
    (x: number, y: number) => {
      if (state.activeTool === "select") return;

      dispatch({
        type: "annotation/add",
        annotation: scaleAnnotation(
          createAnnotationForTool(state.activeTool, { x, y }),
          1 / state.zoom,
          { x, y }
        )
      });
      dispatch({ type: "tool/set", tool: "select" });
    },
    [state.activeTool, state.zoom]
  );

  const updateSelected = useCallback(
    (patch: Partial<Annotation>) => {
      if (!state.selectedId) return;
      dispatch({ type: "annotation/update", id: state.selectedId, patch });
    },
    [state.selectedId]
  );

  const updateAnnotation = useCallback((id: string, patch: Partial<Annotation>) => {
    dispatch({ type: "annotation/update", id, patch });
  }, []);

  useEffect(() => {
    function onKeyDown(event: KeyboardEvent) {
      const target = event.target;
      const isEditableTarget =
        target instanceof HTMLInputElement ||
        target instanceof HTMLTextAreaElement ||
        (target instanceof HTMLElement && target.isContentEditable);
      const isModifier = event.metaKey || event.ctrlKey;
      const key = event.key.toLowerCase();

      if (!isEditableTarget && (event.key === "Delete" || event.key === "Backspace")) {
        event.preventDefault();
        dispatch({ type: "annotation/delete-selected" });
        return;
      }

      if (isEditableTarget) return;

      if (isModifier && key === "z" && event.shiftKey) {
        event.preventDefault();
        dispatch({ type: "history/redo" });
        return;
      }

      if (isModifier && key === "z") {
        event.preventDefault();
        dispatch({ type: "history/undo" });
        return;
      }

      if (isModifier && key === "y") {
        event.preventDefault();
        dispatch({ type: "history/redo" });
      }
    }

    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, []);

  return {
    state,
    status,
    error,
    selectedAnnotation,
    setTool: (tool) => dispatch({ type: "tool/set", tool }),
    loadImage,
    loadFile,
    addAnnotationAt,
    selectAnnotation: (id) => dispatch({ type: "annotation/select", id }),
    updateAnnotation,
    updateSelected,
    deleteSelected: () => dispatch({ type: "annotation/delete-selected" }),
    duplicateSelected: () => dispatch({ type: "annotation/duplicate-selected" }),
    moveAnnotation: (id, delta) =>
      dispatch({ type: "annotation/move", id, delta }),
    moveSelected: (delta) =>
      dispatch({ type: "annotation/move-selected", delta }),
    resizeSelected: (width, height) =>
      dispatch({ type: "annotation/resize-selected", width, height }),
    setZoom: (zoom, anchor) => {
      const nextZoom = Math.min(4, Math.max(0.25, zoom));
      if (state.image) {
        const currentView = {
          width: state.image.width / state.zoom,
          height: state.image.height / state.zoom
        };
        const center = anchor ?? {
          x: state.pan.x + currentView.width / 2,
          y: state.pan.y + currentView.height / 2
        };
        const anchorRatio = anchor
          ? {
              x: (anchor.x - state.pan.x) / currentView.width,
              y: (anchor.y - state.pan.y) / currentView.height
            }
          : { x: 0.5, y: 0.5 };
        const nextView = {
          width: state.image.width / nextZoom,
          height: state.image.height / nextZoom
        };
        dispatch({
          type: "viewport/set-pan",
          pan: {
            x: Math.min(
              Math.max(0, center.x - nextView.width * anchorRatio.x),
              Math.max(0, state.image.width - nextView.width)
            ),
            y: Math.min(
              Math.max(0, center.y - nextView.height * anchorRatio.y),
              Math.max(0, state.image.height - nextView.height)
            )
          }
        });
      }
      dispatch({ type: "viewport/set-zoom", zoom: nextZoom });
    },
    panViewport: (delta) =>
      dispatch({
        type: "viewport/pan-by",
        delta
      }),
    resetViewport: () => {
      dispatch({ type: "viewport/set-zoom", zoom: 1 });
      dispatch({ type: "viewport/set-pan", pan: { x: 0, y: 0 } });
    },
    undo: () => dispatch({ type: "history/undo" }),
    redo: () => dispatch({ type: "history/redo" }),
    setStatus,
    setError
  };
}

function scaleAnnotation(
  annotation: Annotation,
  scale: number,
  origin: Point
): Annotation {
  if (scale === 1) return annotation;

  switch (annotation.type) {
    case "line":
    case "arrow":
      return {
        ...annotation,
        strokeWidth: Math.max(1, annotation.strokeWidth * scale),
        x2: origin.x + (annotation.x2 - origin.x) * scale,
        y2: origin.y + (annotation.y2 - origin.y) * scale
      };
    case "rectangle":
    case "ellipse":
      return {
        ...annotation,
        strokeWidth: Math.max(1, annotation.strokeWidth * scale),
        width: annotation.width * scale,
        height: annotation.height * scale
      };
    case "text":
      return {
        ...annotation,
        strokeWidth: Math.max(1, annotation.strokeWidth * scale),
        fontSize: Math.max(6, annotation.fontSize * scale)
      };
    case "callout":
      return {
        ...annotation,
        strokeWidth: Math.max(1, annotation.strokeWidth * scale),
        fontSize: Math.max(6, annotation.fontSize * scale),
        targetX: origin.x + (annotation.targetX - origin.x) * scale,
        targetY: origin.y + (annotation.targetY - origin.y) * scale
      };
  }
}

function createAnnotationForTool(
  tool: Exclude<Tool, "select">,
  point: { x: number; y: number }
): Annotation {
  switch (tool) {
    case "text":
      return createAnnotation("text", point);
    case "line":
      return createAnnotation("line", point);
    case "arrow":
      return createAnnotation("arrow", point);
    case "rectangle":
      return createAnnotation("rectangle", point);
    case "dashedRectangle":
      return createAnnotation("dashedRectangle", point);
    case "ellipse":
      return createAnnotation("ellipse", point);
    case "callout":
      return createAnnotation("callout", point);
  }
}
