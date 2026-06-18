import type { Annotation, EditorAction, EditorState } from "./types";
import { moveAnnotation, resizeShape } from "./geometry";

export const initialEditorState: EditorState = {
  image: null,
  activeTool: "select",
  annotations: [],
  selectedId: null,
  zoom: 1,
  pan: { x: 0, y: 0 },
  past: [],
  future: []
};

let nextDuplicateId = 1;

function withHistory(
  state: EditorState,
  annotations: Annotation[],
  selectedId: string | null
): EditorState {
  return {
    ...state,
    annotations,
    selectedId,
    past: [...state.past, state.annotations],
    future: []
  };
}

function cloneWithOffset(annotation: Annotation): Annotation {
  const id = `${annotation.id}-copy-${nextDuplicateId}`;
  nextDuplicateId += 1;

  switch (annotation.type) {
    case "line":
    case "arrow":
      return {
        ...annotation,
        id,
        x1: annotation.x1 + 16,
        y1: annotation.y1 + 16,
        x2: annotation.x2 + 16,
        y2: annotation.y2 + 16
      };
    case "callout":
      return {
        ...annotation,
        id,
        x: annotation.x + 16,
        y: annotation.y + 16,
        targetX: annotation.targetX + 16,
        targetY: annotation.targetY + 16
      };
    case "text":
    case "rectangle":
    case "ellipse":
      return { ...annotation, id, x: annotation.x + 16, y: annotation.y + 16 };
  }
}

export function editorReducer(
  state: EditorState,
  action: EditorAction
): EditorState {
  switch (action.type) {
    case "image/set":
      return { ...state, image: action.image, selectedId: null };
    case "image/clear":
      return { ...initialEditorState };
    case "tool/set":
      return { ...state, activeTool: action.tool };
    case "annotation/add":
      return withHistory(
        state,
        [...state.annotations, action.annotation],
        action.annotation.id
      );
    case "annotation/select":
      return { ...state, selectedId: action.id, activeTool: "select" };
    case "annotation/update":
      return withHistory(
        state,
        state.annotations.map((annotation) =>
          annotation.id === action.id
            ? ({ ...annotation, ...action.patch } as Annotation)
            : annotation
        ),
        state.selectedId
      );
    case "annotation/delete-selected":
      if (!state.selectedId) return state;
      return withHistory(
        state,
        state.annotations.filter(
          (annotation) => annotation.id !== state.selectedId
        ),
        null
      );
    case "annotation/duplicate-selected": {
      const selected = state.annotations.find(
        (annotation) => annotation.id === state.selectedId
      );
      if (!selected) return state;
      const copy = cloneWithOffset(selected);
      return withHistory(state, [...state.annotations, copy], copy.id);
    }
    case "annotation/move":
      return withHistory(
        state,
        state.annotations.map((annotation) =>
          annotation.id === action.id
            ? moveAnnotation(annotation, action.delta)
            : annotation
        ),
        action.id
      );
    case "annotation/move-selected":
      if (!state.selectedId) return state;
      return withHistory(
        state,
        state.annotations.map((annotation) =>
          annotation.id === state.selectedId
            ? moveAnnotation(annotation, action.delta)
            : annotation
        ),
        state.selectedId
      );
    case "annotation/resize-selected":
      if (!state.selectedId) return state;
      return withHistory(
        state,
        state.annotations.map((annotation) =>
          annotation.id === state.selectedId
            ? resizeShape(annotation, action.width, action.height)
            : annotation
        ),
        state.selectedId
      );
    case "annotations/clear":
      return withHistory(state, [], null);
    case "viewport/set-zoom":
      return { ...state, zoom: Math.min(4, Math.max(0.25, action.zoom)) };
    case "viewport/set-pan":
      return { ...state, pan: action.pan };
    case "history/undo": {
      const previous = state.past.at(-1);
      if (!previous) return state;
      return {
        ...state,
        annotations: previous,
        selectedId: null,
        past: state.past.slice(0, -1),
        future: [state.annotations, ...state.future]
      };
    }
    case "history/redo": {
      const next = state.future[0];
      if (!next) return state;
      return {
        ...state,
        annotations: next,
        selectedId: null,
        past: [...state.past, state.annotations],
        future: state.future.slice(1)
      };
    }
  }
}
