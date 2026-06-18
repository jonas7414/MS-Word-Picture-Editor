import { useCallback, useEffect, useMemo, useReducer, useState } from "react";
import { createAnnotation } from "../editor/defaults";
import { editorReducer, initialEditorState } from "../editor/reducer";
import type { Annotation, Point, Tool } from "../editor/types";
import { loadImageFile } from "../image/imageValidation";

export type EditorController = {
  state: typeof initialEditorState;
  status: string;
  error: string;
  selectedAnnotation: Annotation | null;
  setTool: (tool: Tool) => void;
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

  const addAnnotationAt = useCallback(
    (x: number, y: number) => {
      if (state.activeTool === "select") return;

      dispatch({
        type: "annotation/add",
        annotation: createAnnotationForTool(state.activeTool, { x, y })
      });
      dispatch({ type: "tool/set", tool: "select" });
    },
    [state.activeTool]
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
    undo: () => dispatch({ type: "history/undo" }),
    redo: () => dispatch({ type: "history/redo" }),
    setStatus,
    setError
  };
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
