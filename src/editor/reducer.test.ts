import { describe, expect, it } from "vitest";
import { createAnnotation } from "./defaults";
import { editorReducer, initialEditorState } from "./reducer";

describe("editorReducer", () => {
  it("adds and selects a text annotation", () => {
    const annotation = createAnnotation("text", { x: 40, y: 60 });
    const state = editorReducer(initialEditorState, {
      type: "annotation/add",
      annotation
    });

    expect(state.annotations).toHaveLength(1);
    expect(state.selectedId).toBe(annotation.id);
    expect(state.annotations[0]).toMatchObject({
      type: "text",
      text: "零件名稱",
      x: 40,
      y: 60
    });
  });

  it("updates a selected annotation style", () => {
    const annotation = createAnnotation("rectangle", { x: 10, y: 20 });
    const added = editorReducer(initialEditorState, {
      type: "annotation/add",
      annotation
    });

    const state = editorReducer(added, {
      type: "annotation/update",
      id: annotation.id,
      patch: { stroke: "#dc2626", strokeWidth: 5 }
    });

    expect(state.annotations[0]).toMatchObject({
      stroke: "#dc2626",
      strokeWidth: 5
    });
  });

  it("supports delete, undo, and redo", () => {
    const annotation = createAnnotation("arrow", { x: 20, y: 30 });
    const added = editorReducer(initialEditorState, {
      type: "annotation/add",
      annotation
    });
    const deleted = editorReducer(added, { type: "annotation/delete-selected" });

    expect(deleted.annotations).toHaveLength(0);

    const undone = editorReducer(deleted, { type: "history/undo" });
    expect(undone.annotations).toHaveLength(1);
    expect(undone.annotations[0].id).toBe(annotation.id);

    const redone = editorReducer(undone, { type: "history/redo" });
    expect(redone.annotations).toHaveLength(0);
  });

  it("duplicates the selected annotation with an offset", () => {
    const annotation = createAnnotation("ellipse", { x: 20, y: 30 });
    const added = editorReducer(initialEditorState, {
      type: "annotation/add",
      annotation
    });
    const duplicated = editorReducer(added, {
      type: "annotation/duplicate-selected"
    });

    expect(duplicated.annotations).toHaveLength(2);
    expect(duplicated.annotations[1]).toMatchObject({
      type: "ellipse",
      x: annotation.x + 16,
      y: annotation.y + 16
    });
    expect(duplicated.selectedId).toBe(duplicated.annotations[1].id);
  });

  it("creates a dashed rectangle annotation for dashed frame markup", () => {
    const annotation = createAnnotation("dashedRectangle", { x: 12, y: 24 });

    expect(annotation).toMatchObject({
      type: "rectangle",
      strokeStyle: "dashed",
      width: 120,
      height: 72
    });
  });

  it("creates a circle with equal default width and height", () => {
    const annotation = createAnnotation("ellipse", { x: 12, y: 24 });

    expect(annotation).toMatchObject({
      type: "ellipse",
      width: 96,
      height: 96
    });
  });
});
