export type Tool =
  | "select"
  | "text"
  | "line"
  | "arrow"
  | "rectangle"
  | "dashedRectangle"
  | "ellipse"
  | "callout";

export type Point = {
  x: number;
  y: number;
};

export type AnnotationBase = {
  id: string;
  stroke: string;
  fill: string;
  strokeWidth: number;
  strokeStyle?: "solid" | "dashed";
};

export type TextAnnotation = AnnotationBase & {
  type: "text";
  x: number;
  y: number;
  text: string;
  fontSize: number;
  color: string;
  background: string;
};

export type LineAnnotation = AnnotationBase & {
  type: "line" | "arrow";
  x1: number;
  y1: number;
  x2: number;
  y2: number;
};

export type ShapeAnnotation = AnnotationBase & {
  type: "rectangle" | "ellipse";
  x: number;
  y: number;
  width: number;
  height: number;
};

export type CalloutAnnotation = AnnotationBase & {
  type: "callout";
  x: number;
  y: number;
  text: string;
  fontSize: number;
  color: string;
  background: string;
  targetX: number;
  targetY: number;
};

export type Annotation =
  | TextAnnotation
  | LineAnnotation
  | ShapeAnnotation
  | CalloutAnnotation;

export type LoadedImage = {
  src: string;
  name: string;
  width: number;
  height: number;
};

export type EditorState = {
  image: LoadedImage | null;
  activeTool: Tool;
  annotations: Annotation[];
  selectedId: string | null;
  zoom: number;
  pan: Point;
  past: Annotation[][];
  future: Annotation[][];
};

export type EditorAction =
  | { type: "image/set"; image: LoadedImage }
  | { type: "image/clear" }
  | { type: "tool/set"; tool: Tool }
  | { type: "annotation/add"; annotation: Annotation }
  | { type: "annotation/select"; id: string | null }
  | { type: "annotation/update"; id: string; patch: Partial<Annotation> }
  | { type: "annotation/delete-selected" }
  | { type: "annotation/duplicate-selected" }
  | { type: "annotation/move"; id: string; delta: Point }
  | { type: "annotation/move-selected"; delta: Point }
  | { type: "annotation/resize-selected"; width: number; height: number }
  | { type: "annotations/clear" }
  | { type: "viewport/set-zoom"; zoom: number }
  | { type: "viewport/set-pan"; pan: Point }
  | { type: "viewport/pan-by"; delta: Point }
  | { type: "history/undo" }
  | { type: "history/redo" };
