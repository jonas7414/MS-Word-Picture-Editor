# Word Image Annotation Add-in Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a first usable Word Office Add-in task pane that lets engineers upload an image, add text and basic shape annotations, export the result as PNG, and insert it into the current Word document.

**Architecture:** Use a client-only React app inside a Word task pane. Keep annotation editing as structured state, render an SVG overlay for interactive editing, flatten the image plus annotations through a canvas exporter, and isolate Office.js behind a small insertion adapter that can be mocked in tests.

**Tech Stack:** Vite, React, TypeScript, Vitest, Testing Library, SVG for interactive overlays, Canvas for PNG export, Office.js Word JavaScript API, npm scripts `dev`, `build`, `preview`, `test`, and `test:run`.

## Global Constraints

- Product form: Word Office Add-in with a right-side task pane.
- Primary user: engineers writing technical Word documents.
- Core tools: text, line, arrow, rectangle, ellipse, and callout.
- First version should be client-side only; no backend, cloud storage, or user accounts.
- Export should flatten base image and annotations into one raster image before insertion into Word.
- Initial export format: PNG.
- Word insertion should use Office.js Word APIs at the current selection.
- UI copy should start in Traditional Chinese and stay concise.
- Do not build advanced photo editing features: filters, brightness, contrast, masking, blur, background removal, OCR, automatic part detection, AI labels, CAD dimensioning, mobile Word support, or PowerPoint/Excel support.
- Repository starts empty except for docs and skills, so the first task creates the application scaffold and test setup.

---

## File Structure

- Create `package.json`: npm scripts and runtime/test dependencies.
- Create `tsconfig.json`, `tsconfig.node.json`, `vite.config.ts`: TypeScript, Vite, and Vitest configuration.
- Create `index.html`: task pane host HTML.
- Create `public/manifest.xml`: Word add-in manifest for local sideloading.
- Create `src/main.tsx`: React entrypoint.
- Create `src/App.tsx`: top-level task pane composition and Word insertion flow.
- Create `src/styles.css`: compact engineering-tool UI styling.
- Create `src/test/setup.ts`: Testing Library and DOM test setup.
- Create `src/test/officeMock.ts`: Office.js mock helpers for adapter tests.
- Create `src/editor/types.ts`: annotation, tool, image, and editor-state types.
- Create `src/editor/defaults.ts`: annotation defaults and factory helpers.
- Create `src/editor/reducer.ts`: editor reducer and pure state transitions.
- Create `src/editor/geometry.ts`: hit testing, resize, and coordinate helpers.
- Create `src/image/imageValidation.ts`: upload validation and image loading contract.
- Create `src/render/annotationSvg.tsx`: SVG rendering of annotations for the editor.
- Create `src/render/exportPng.ts`: canvas-based PNG export.
- Create `src/word/wordClient.ts`: Office.js insertion adapter.
- Create `src/components/Toolbar.tsx`: compact tool and action buttons.
- Create `src/components/CanvasStage.tsx`: image preview, SVG overlay, pointer interactions.
- Create `src/components/PropertyPanel.tsx`: selected annotation properties.
- Create `src/components/ImagePicker.tsx`: image upload UI.
- Create `src/components/StatusMessage.tsx`: user-visible status/errors.
- Create `src/hooks/useEditorController.ts`: coordinates reducer, keyboard shortcuts, and UI commands.
- Create `docs/dev/word-sideloading.md`: local run and sideload instructions.
- Create tests beside their modules with `.test.ts` or `.test.tsx` suffix.

---

### Task 1: Project Scaffold and Smoke Test

**Files:**
- Create: `package.json`
- Create: `tsconfig.json`
- Create: `tsconfig.node.json`
- Create: `vite.config.ts`
- Create: `index.html`
- Create: `src/main.tsx`
- Create: `src/App.tsx`
- Create: `src/styles.css`
- Create: `src/test/setup.ts`
- Test: `src/App.test.tsx`

**Interfaces:**
- Consumes: no project code.
- Produces:
  - `App(): JSX.Element`
  - npm scripts: `dev`, `build`, `preview`, `test`, `test:run`, `typecheck`

- [ ] **Step 1: Initialize git if needed**

Run:

```bash
git status --short
```

Expected if repo is already initialized: command exits `0`.

Expected if repo is not initialized: command exits `128` with `fatal: not a git repository`.

If it exits `128`, run:

```bash
git init
```

Expected: output includes `Initialized empty Git repository`.

- [ ] **Step 2: Create the failing smoke test**

Create `src/App.test.tsx`:

```tsx
import { render, screen } from "@testing-library/react";
import { describe, expect, it } from "vitest";
import { App } from "./App";

describe("App", () => {
  it("renders the Word image annotation task pane", () => {
    render(<App />);

    expect(screen.getByRole("heading", { name: "圖片標註" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "插入 Word" })).toBeDisabled();
  });
});
```

- [ ] **Step 3: Run test to verify scaffold is missing**

Run:

```bash
npm test:run -- src/App.test.tsx
```

Expected: FAIL because `package.json` and the test runner do not exist yet.

- [ ] **Step 4: Create npm and TypeScript configuration**

Create `package.json`:

```json
{
  "name": "word-pic-editor",
  "version": "0.1.0",
  "private": true,
  "type": "module",
  "scripts": {
    "dev": "vite --host 127.0.0.1",
    "build": "tsc --noEmit && vite build",
    "preview": "vite preview --host 127.0.0.1",
    "test": "vitest",
    "test:run": "vitest run",
    "typecheck": "tsc --noEmit"
  },
  "dependencies": {
    "@vitejs/plugin-react": "^5.0.0",
    "lucide-react": "^0.468.0",
    "office-js": "^1.1.97",
    "react": "^18.3.1",
    "react-dom": "^18.3.1"
  },
  "devDependencies": {
    "@testing-library/jest-dom": "^6.6.3",
    "@testing-library/react": "^16.1.0",
    "@testing-library/user-event": "^14.5.2",
    "@types/node": "^22.10.2",
    "@types/react": "^18.3.18",
    "@types/react-dom": "^18.3.5",
    "jsdom": "^25.0.1",
    "typescript": "^5.7.2",
    "vite": "^6.0.5",
    "vitest": "^2.1.8"
  }
}
```

Create `tsconfig.json`:

```json
{
  "compilerOptions": {
    "target": "ES2020",
    "useDefineForClassFields": true,
    "lib": ["DOM", "DOM.Iterable", "ES2020"],
    "allowJs": false,
    "skipLibCheck": true,
    "esModuleInterop": true,
    "allowSyntheticDefaultImports": true,
    "strict": true,
    "forceConsistentCasingInFileNames": true,
    "module": "ESNext",
    "moduleResolution": "Node",
    "resolveJsonModule": true,
    "isolatedModules": true,
    "noEmit": true,
    "jsx": "react-jsx",
    "types": ["vitest/globals", "vite/client"]
  },
  "include": ["src", "vite.config.ts"]
}
```

Create `tsconfig.node.json`:

```json
{
  "compilerOptions": {
    "skipLibCheck": true,
    "module": "ESNext",
    "moduleResolution": "Node",
    "allowSyntheticDefaultImports": true,
    "strict": true
  },
  "include": ["vite.config.ts"]
}
```

Create `vite.config.ts`:

```ts
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

export default defineConfig({
  plugins: [react()],
  test: {
    environment: "jsdom",
    setupFiles: ["./src/test/setup.ts"],
    globals: true
  }
});
```

- [ ] **Step 5: Create the minimal app shell**

Create `index.html`:

```html
<!doctype html>
<html lang="zh-Hant">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Word 圖片標註</title>
  </head>
  <body>
    <div id="root"></div>
    <script type="module" src="/src/main.tsx"></script>
  </body>
</html>
```

Create `src/test/setup.ts`:

```ts
import "@testing-library/jest-dom/vitest";
```

Create `src/main.tsx`:

```tsx
import React from "react";
import ReactDOM from "react-dom/client";
import { App } from "./App";
import "./styles.css";

ReactDOM.createRoot(document.getElementById("root") as HTMLElement).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
```

Create `src/App.tsx`:

```tsx
export function App() {
  return (
    <main className="task-pane" aria-label="Word 圖片標註工具">
      <header className="app-header">
        <h1>圖片標註</h1>
        <button type="button" disabled>
          插入 Word
        </button>
      </header>
      <section className="empty-state" aria-label="尚未載入圖片">
        <p>上傳圖片後即可加入文字、箭頭與基礎圖案。</p>
      </section>
    </main>
  );
}
```

Create `src/styles.css`:

```css
:root {
  color: #1f2933;
  background: #f6f7f9;
  font-family:
    Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont,
    "Segoe UI", sans-serif;
}

* {
  box-sizing: border-box;
}

body {
  margin: 0;
  min-width: 320px;
  background: #f6f7f9;
}

button,
input,
select,
textarea {
  font: inherit;
}

.task-pane {
  min-height: 100vh;
  padding: 12px;
}

.app-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
}

.app-header h1 {
  margin: 0;
  font-size: 20px;
  font-weight: 700;
}

.app-header button {
  border: 1px solid #2563eb;
  border-radius: 6px;
  background: #2563eb;
  color: white;
  padding: 8px 10px;
}

.app-header button:disabled {
  border-color: #b9c2cf;
  background: #d8dee8;
  color: #64748b;
}

.empty-state {
  margin-top: 16px;
  border: 1px dashed #9aa7b7;
  border-radius: 8px;
  padding: 16px;
  background: white;
}

.empty-state p {
  margin: 0;
  color: #52606d;
  line-height: 1.5;
}
```

- [ ] **Step 6: Install dependencies**

Run:

```bash
npm install
```

Expected: dependencies install and `package-lock.json` is created.

- [ ] **Step 7: Run test to verify it passes**

Run:

```bash
npm test:run -- src/App.test.tsx
```

Expected: PASS for `renders the Word image annotation task pane`.

- [ ] **Step 8: Run typecheck**

Run:

```bash
npm run typecheck
```

Expected: PASS with no TypeScript errors.

- [ ] **Step 9: Commit**

```bash
git add package.json package-lock.json tsconfig.json tsconfig.node.json vite.config.ts index.html src/main.tsx src/App.tsx src/styles.css src/test/setup.ts src/App.test.tsx
git commit -m "chore: scaffold word annotation add-in"
```

---

### Task 2: Annotation Domain State

**Files:**
- Create: `src/editor/types.ts`
- Create: `src/editor/defaults.ts`
- Create: `src/editor/reducer.ts`
- Test: `src/editor/reducer.test.ts`

**Interfaces:**
- Consumes: no app modules from Task 1.
- Produces:
  - `type Tool = "select" | "text" | "line" | "arrow" | "rectangle" | "ellipse" | "callout"`
  - `type Annotation`
  - `type EditorState`
  - `const initialEditorState: EditorState`
  - `createAnnotation(tool: Exclude<Tool, "select">, point: Point): Annotation`
  - `editorReducer(state: EditorState, action: EditorAction): EditorState`

- [ ] **Step 1: Write failing reducer tests**

Create `src/editor/reducer.test.ts`:

```ts
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
    const duplicated = editorReducer(added, { type: "annotation/duplicate-selected" });

    expect(duplicated.annotations).toHaveLength(2);
    expect(duplicated.annotations[1]).toMatchObject({
      type: "ellipse",
      x: annotation.x + 16,
      y: annotation.y + 16
    });
    expect(duplicated.selectedId).toBe(duplicated.annotations[1].id);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test:run -- src/editor/reducer.test.ts
```

Expected: FAIL with import errors for `./defaults` and `./reducer`.

- [ ] **Step 3: Implement domain types and defaults**

Create `src/editor/types.ts`:

```ts
export type Tool =
  | "select"
  | "text"
  | "line"
  | "arrow"
  | "rectangle"
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
  | { type: "annotations/clear" }
  | { type: "viewport/set-zoom"; zoom: number }
  | { type: "viewport/set-pan"; pan: Point }
  | { type: "history/undo" }
  | { type: "history/redo" };
```

Create `src/editor/defaults.ts`:

```ts
import type { Annotation, Point, Tool } from "./types";

export const defaultStroke = "#dc2626";
export const defaultFill = "transparent";
export const defaultTextColor = "#111827";
export const defaultTextBackground = "#ffffff";
export const defaultStrokeWidth = 3;
export const defaultFontSize = 18;

let nextId = 1;

export function resetAnnotationIdsForTests() {
  nextId = 1;
}

function createId() {
  const id = `annotation-${nextId}`;
  nextId += 1;
  return id;
}

export function createAnnotation(
  tool: Exclude<Tool, "select">,
  point: Point
): Annotation {
  const base = {
    id: createId(),
    stroke: defaultStroke,
    fill: defaultFill,
    strokeWidth: defaultStrokeWidth
  };

  if (tool === "text") {
    return {
      ...base,
      type: "text",
      x: point.x,
      y: point.y,
      text: "零件名稱",
      fontSize: defaultFontSize,
      color: defaultTextColor,
      background: defaultTextBackground
    };
  }

  if (tool === "line" || tool === "arrow") {
    return {
      ...base,
      type: tool,
      x1: point.x,
      y1: point.y,
      x2: point.x + 96,
      y2: point.y
    };
  }

  if (tool === "callout") {
    return {
      ...base,
      type: "callout",
      x: point.x,
      y: point.y,
      text: "零件名稱",
      fontSize: defaultFontSize,
      color: defaultTextColor,
      background: defaultTextBackground,
      targetX: point.x + 96,
      targetY: point.y + 40
    };
  }

  return {
    ...base,
    type: tool,
    x: point.x,
    y: point.y,
    width: 120,
    height: 72
  };
}
```

- [ ] **Step 4: Implement reducer**

Create `src/editor/reducer.ts`:

```ts
import type { Annotation, EditorAction, EditorState } from "./types";

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
  const id = `${annotation.id}-copy-${Date.now()}`;
  if (annotation.type === "line" || annotation.type === "arrow") {
    return {
      ...annotation,
      id,
      x1: annotation.x1 + 16,
      y1: annotation.y1 + 16,
      x2: annotation.x2 + 16,
      y2: annotation.y2 + 16
    };
  }
  if (annotation.type === "callout") {
    return {
      ...annotation,
      id,
      x: annotation.x + 16,
      y: annotation.y + 16,
      targetX: annotation.targetX + 16,
      targetY: annotation.targetY + 16
    };
  }
  return { ...annotation, id, x: annotation.x + 16, y: annotation.y + 16 };
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
        state.annotations.filter((annotation) => annotation.id !== state.selectedId),
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
```

- [ ] **Step 5: Run reducer tests**

Run:

```bash
npm test:run -- src/editor/reducer.test.ts
```

Expected: PASS for all reducer tests.

- [ ] **Step 6: Commit**

```bash
git add src/editor/types.ts src/editor/defaults.ts src/editor/reducer.ts src/editor/reducer.test.ts
git commit -m "feat: add annotation editor state"
```

---

### Task 3: Image Upload Validation and Loading Contract

**Files:**
- Create: `src/image/imageValidation.ts`
- Test: `src/image/imageValidation.test.ts`

**Interfaces:**
- Consumes:
  - `LoadedImage` from `src/editor/types.ts`
- Produces:
  - `const maxImageBytes: number`
  - `validateImageFile(file: File): { ok: true } | { ok: false; message: string }`
  - `loadImageFile(file: File): Promise<LoadedImage>`

- [ ] **Step 1: Write failing validation tests**

Create `src/image/imageValidation.test.ts`:

```ts
import { describe, expect, it } from "vitest";
import { maxImageBytes, validateImageFile } from "./imageValidation";

function makeFile(name: string, type: string, size: number) {
  return new File([new Uint8Array(size)], name, { type });
}

describe("validateImageFile", () => {
  it("accepts png, jpeg, and webp images", () => {
    expect(validateImageFile(makeFile("part.png", "image/png", 1000))).toEqual({
      ok: true
    });
    expect(validateImageFile(makeFile("part.jpg", "image/jpeg", 1000))).toEqual({
      ok: true
    });
    expect(validateImageFile(makeFile("part.webp", "image/webp", 1000))).toEqual({
      ok: true
    });
  });

  it("rejects unsupported file types with Traditional Chinese copy", () => {
    expect(validateImageFile(makeFile("part.gif", "image/gif", 1000))).toEqual({
      ok: false,
      message: "僅支援 PNG、JPG、JPEG、WEBP 圖片。"
    });
  });

  it("rejects files larger than the configured limit", () => {
    expect(
      validateImageFile(makeFile("huge.png", "image/png", maxImageBytes + 1))
    ).toEqual({
      ok: false,
      message: "圖片不可超過 10 MB。"
    });
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test:run -- src/image/imageValidation.test.ts
```

Expected: FAIL with import error for `./imageValidation`.

- [ ] **Step 3: Implement validation and loading**

Create `src/image/imageValidation.ts`:

```ts
import type { LoadedImage } from "../editor/types";

export const maxImageBytes = 10 * 1024 * 1024;

const supportedTypes = new Set(["image/png", "image/jpeg", "image/webp"]);

export type ImageValidationResult =
  | { ok: true }
  | { ok: false; message: string };

export function validateImageFile(file: File): ImageValidationResult {
  if (!supportedTypes.has(file.type)) {
    return { ok: false, message: "僅支援 PNG、JPG、JPEG、WEBP 圖片。" };
  }

  if (file.size > maxImageBytes) {
    return { ok: false, message: "圖片不可超過 10 MB。" };
  }

  return { ok: true };
}

export async function loadImageFile(file: File): Promise<LoadedImage> {
  const validation = validateImageFile(file);
  if (!validation.ok) {
    throw new Error(validation.message);
  }

  const src = await readFileAsDataUrl(file);
  const dimensions = await readImageDimensions(src);

  return {
    src,
    name: file.name,
    width: dimensions.width,
    height: dimensions.height
  };
}

function readFileAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(new Error("無法讀取圖片。"));
    reader.readAsDataURL(file);
  });
}

function readImageDimensions(src: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () =>
      resolve({
        width: image.naturalWidth,
        height: image.naturalHeight
      });
    image.onerror = () => reject(new Error("無法載入圖片。"));
    image.src = src;
  });
}
```

- [ ] **Step 4: Run validation tests**

Run:

```bash
npm test:run -- src/image/imageValidation.test.ts
```

Expected: PASS for all validation tests.

- [ ] **Step 5: Commit**

```bash
git add src/image/imageValidation.ts src/image/imageValidation.test.ts
git commit -m "feat: validate image uploads"
```

---

### Task 4: Annotation SVG Rendering

**Files:**
- Create: `src/render/annotationSvg.tsx`
- Test: `src/render/annotationSvg.test.tsx`

**Interfaces:**
- Consumes:
  - `Annotation` from `src/editor/types.ts`
- Produces:
  - `AnnotationSvg(props: { annotations: Annotation[]; selectedId: string | null; onSelect(id: string): void }): JSX.Element`

- [ ] **Step 1: Write failing SVG rendering test**

Create `src/render/annotationSvg.test.tsx`:

```tsx
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

    render(
      <svg>
        <AnnotationSvg
          annotations={annotations}
          selectedId="text-1"
          onSelect={onSelect}
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
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test:run -- src/render/annotationSvg.test.tsx
```

Expected: FAIL with import error for `./annotationSvg`.

- [ ] **Step 3: Implement SVG rendering**

Create `src/render/annotationSvg.tsx`:

```tsx
import type { PointerEvent } from "react";
import type { Annotation } from "../editor/types";

type Props = {
  annotations: Annotation[];
  selectedId: string | null;
  onSelect: (id: string) => void;
};

export function AnnotationSvg({ annotations, selectedId, onSelect }: Props) {
  return (
    <g>
      <defs>
        <marker
          id="arrow-head"
          markerWidth="10"
          markerHeight="10"
          refX="8"
          refY="3"
          orient="auto"
          markerUnits="strokeWidth"
        >
          <path d="M0,0 L0,6 L9,3 z" fill="#dc2626" />
        </marker>
      </defs>
      {annotations.map((annotation) => (
        <AnnotationNode
          key={annotation.id}
          annotation={annotation}
          selected={annotation.id === selectedId}
          onSelect={onSelect}
        />
      ))}
    </g>
  );
}

function AnnotationNode({
  annotation,
  selected,
  onSelect
}: {
  annotation: Annotation;
  selected: boolean;
  onSelect: (id: string) => void;
}) {
  const common = {
    onPointerDown: (event: PointerEvent) => {
      event.stopPropagation();
      onSelect(annotation.id);
    }
  };

  if (annotation.type === "text") {
    return (
      <g {...common} aria-label={`選取標註 ${annotation.text}`} role="button">
        <rect
          x={annotation.x - 4}
          y={annotation.y - annotation.fontSize}
          width={annotation.text.length * annotation.fontSize}
          height={annotation.fontSize + 10}
          rx="3"
          fill={annotation.background}
          stroke={selected ? "#2563eb" : "transparent"}
          data-testid={selected ? `selected-outline-${annotation.id}` : undefined}
        />
        <text
          x={annotation.x}
          y={annotation.y}
          fontSize={annotation.fontSize}
          fill={annotation.color}
        >
          {annotation.text}
        </text>
      </g>
    );
  }

  if (annotation.type === "line" || annotation.type === "arrow") {
    return (
      <g {...common} aria-label={`選取標註 ${annotation.type}`} role="button">
        <line
          x1={annotation.x1}
          y1={annotation.y1}
          x2={annotation.x2}
          y2={annotation.y2}
          stroke={annotation.stroke}
          strokeWidth={annotation.strokeWidth}
          markerEnd={annotation.type === "arrow" ? "url(#arrow-head)" : undefined}
          data-testid={`${annotation.type}-${annotation.id}`}
        />
        {selected ? (
          <circle
            cx={annotation.x2}
            cy={annotation.y2}
            r="5"
            fill="#2563eb"
            data-testid={`selected-outline-${annotation.id}`}
          />
        ) : null}
      </g>
    );
  }

  if (annotation.type === "callout") {
    return (
      <g {...common} aria-label={`選取標註 ${annotation.text}`} role="button">
        <line
          x1={annotation.x}
          y1={annotation.y}
          x2={annotation.targetX}
          y2={annotation.targetY}
          stroke={annotation.stroke}
          strokeWidth={annotation.strokeWidth}
          markerEnd="url(#arrow-head)"
        />
        <rect
          x={annotation.x - 4}
          y={annotation.y - annotation.fontSize}
          width={annotation.text.length * annotation.fontSize}
          height={annotation.fontSize + 10}
          rx="3"
          fill={annotation.background}
          stroke={selected ? "#2563eb" : annotation.stroke}
        />
        <text
          x={annotation.x}
          y={annotation.y}
          fontSize={annotation.fontSize}
          fill={annotation.color}
        >
          {annotation.text}
        </text>
      </g>
    );
  }

  const shape =
    annotation.type === "rectangle" ? (
      <rect
        x={annotation.x}
        y={annotation.y}
        width={annotation.width}
        height={annotation.height}
        fill={annotation.fill}
        stroke={annotation.stroke}
        strokeWidth={annotation.strokeWidth}
      />
    ) : (
      <ellipse
        cx={annotation.x + annotation.width / 2}
        cy={annotation.y + annotation.height / 2}
        rx={annotation.width / 2}
        ry={annotation.height / 2}
        fill={annotation.fill}
        stroke={annotation.stroke}
        strokeWidth={annotation.strokeWidth}
      />
    );

  return (
    <g {...common} aria-label={`選取標註 ${annotation.type}`} role="button">
      {shape}
      {selected ? (
        <rect
          x={annotation.x - 4}
          y={annotation.y - 4}
          width={annotation.width + 8}
          height={annotation.height + 8}
          fill="transparent"
          stroke="#2563eb"
          strokeDasharray="4 4"
          data-testid={`selected-outline-${annotation.id}`}
        />
      ) : null}
    </g>
  );
}
```

- [ ] **Step 4: Run SVG tests**

Run:

```bash
npm test:run -- src/render/annotationSvg.test.tsx
```

Expected: PASS for `renders text and arrow annotations and reports selection`.

- [ ] **Step 5: Commit**

```bash
git add src/render/annotationSvg.tsx src/render/annotationSvg.test.tsx
git commit -m "feat: render annotation overlay"
```

---

### Task 5: PNG Exporter

**Files:**
- Create: `src/render/exportPng.ts`
- Test: `src/render/exportPng.test.ts`

**Interfaces:**
- Consumes:
  - `LoadedImage`, `Annotation` from `src/editor/types.ts`
- Produces:
  - `exportAnnotatedPng(input: { image: LoadedImage; annotations: Annotation[] }): Promise<string>`

- [ ] **Step 1: Write failing exporter test**

Create `src/render/exportPng.test.ts`:

```ts
import { describe, expect, it, vi } from "vitest";
import type { Annotation, LoadedImage } from "../editor/types";
import { exportAnnotatedPng } from "./exportPng";

describe("exportAnnotatedPng", () => {
  it("returns a PNG data URL with the source image dimensions", async () => {
    const drawImage = vi.fn();
    const fillText = vi.fn();
    const strokeRect = vi.fn();
    const context = {
      drawImage,
      fillText,
      strokeRect,
      beginPath: vi.fn(),
      moveTo: vi.fn(),
      lineTo: vi.fn(),
      stroke: vi.fn(),
      ellipse: vi.fn(),
      fillRect: vi.fn(),
      save: vi.fn(),
      restore: vi.fn(),
      fill: vi.fn(),
      closePath: vi.fn(),
      set fillStyle(value: string) {},
      set strokeStyle(value: string) {},
      set lineWidth(value: number) {},
      set font(value: string) {}
    };

    vi.spyOn(document, "createElement").mockImplementation((tagName: string) => {
      if (tagName === "canvas") {
        return {
          width: 0,
          height: 0,
          getContext: () => context,
          toDataURL: () => "data:image/png;base64,exported"
        } as unknown as HTMLCanvasElement;
      }
      return document.createElement(tagName);
    });

    class MockImage {
      naturalWidth = 640;
      naturalHeight = 480;
      onload: (() => void) | null = null;
      onerror: (() => void) | null = null;
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
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test:run -- src/render/exportPng.test.ts
```

Expected: FAIL with import error for `./exportPng`.

- [ ] **Step 3: Implement exporter**

Create `src/render/exportPng.ts`:

```ts
import type { Annotation, LoadedImage } from "../editor/types";

export async function exportAnnotatedPng({
  image,
  annotations
}: {
  image: LoadedImage;
  annotations: Annotation[];
}): Promise<string> {
  const bitmap = await loadBitmap(image.src);
  const canvas = document.createElement("canvas");
  canvas.width = image.width;
  canvas.height = image.height;

  const context = canvas.getContext("2d");
  if (!context) {
    throw new Error("無法建立匯出畫布。");
  }

  context.drawImage(bitmap, 0, 0, image.width, image.height);
  annotations.forEach((annotation) => drawAnnotation(context, annotation));

  return canvas.toDataURL("image/png");
}

function loadBitmap(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("無法載入匯出圖片。"));
    image.src = src;
  });
}

function drawAnnotation(context: CanvasRenderingContext2D, annotation: Annotation) {
  context.save();
  context.strokeStyle = annotation.stroke;
  context.lineWidth = annotation.strokeWidth;

  if (annotation.type === "text") {
    drawText(context, annotation.x, annotation.y, annotation.text, annotation.fontSize, annotation.color, annotation.background);
    context.restore();
    return;
  }

  if (annotation.type === "line" || annotation.type === "arrow") {
    drawLine(context, annotation.x1, annotation.y1, annotation.x2, annotation.y2);
    if (annotation.type === "arrow") {
      drawArrowHead(context, annotation.x1, annotation.y1, annotation.x2, annotation.y2);
    }
    context.restore();
    return;
  }

  if (annotation.type === "rectangle") {
    context.strokeRect(annotation.x, annotation.y, annotation.width, annotation.height);
    context.restore();
    return;
  }

  if (annotation.type === "ellipse") {
    context.beginPath();
    context.ellipse(
      annotation.x + annotation.width / 2,
      annotation.y + annotation.height / 2,
      annotation.width / 2,
      annotation.height / 2,
      0,
      0,
      Math.PI * 2
    );
    context.stroke();
    context.restore();
    return;
  }

  drawLine(context, annotation.x, annotation.y, annotation.targetX, annotation.targetY);
  drawArrowHead(context, annotation.x, annotation.y, annotation.targetX, annotation.targetY);
  drawText(context, annotation.x, annotation.y, annotation.text, annotation.fontSize, annotation.color, annotation.background);
  context.restore();
}

function drawText(
  context: CanvasRenderingContext2D,
  x: number,
  y: number,
  text: string,
  fontSize: number,
  color: string,
  background: string
) {
  context.font = `${fontSize}px sans-serif`;
  context.fillStyle = background;
  context.fillRect(x - 4, y - fontSize, text.length * fontSize, fontSize + 10);
  context.fillStyle = color;
  context.fillText(text, x, y);
}

function drawLine(
  context: CanvasRenderingContext2D,
  x1: number,
  y1: number,
  x2: number,
  y2: number
) {
  context.beginPath();
  context.moveTo(x1, y1);
  context.lineTo(x2, y2);
  context.stroke();
}

function drawArrowHead(
  context: CanvasRenderingContext2D,
  x1: number,
  y1: number,
  x2: number,
  y2: number
) {
  const angle = Math.atan2(y2 - y1, x2 - x1);
  const length = 14;
  context.beginPath();
  context.moveTo(x2, y2);
  context.lineTo(
    x2 - length * Math.cos(angle - Math.PI / 6),
    y2 - length * Math.sin(angle - Math.PI / 6)
  );
  context.lineTo(
    x2 - length * Math.cos(angle + Math.PI / 6),
    y2 - length * Math.sin(angle + Math.PI / 6)
  );
  context.closePath();
  context.fillStyle = String(context.strokeStyle);
  context.fill();
}
```

- [ ] **Step 4: Run exporter tests**

Run:

```bash
npm test:run -- src/render/exportPng.test.ts
```

Expected: PASS for `returns a PNG data URL with the source image dimensions`.

- [ ] **Step 5: Commit**

```bash
git add src/render/exportPng.ts src/render/exportPng.test.ts
git commit -m "feat: export annotated image as png"
```

---

### Task 6: Word Insertion Adapter

**Files:**
- Create: `src/word/wordClient.ts`
- Create: `src/test/officeMock.ts`
- Test: `src/word/wordClient.test.ts`

**Interfaces:**
- Consumes: no previous app code.
- Produces:
  - `insertImageAtSelection(dataUrl: string): Promise<void>`
  - `stripDataUrlPrefix(dataUrl: string): string`

- [ ] **Step 1: Write failing Word adapter tests**

Create `src/word/wordClient.test.ts`:

```ts
import { afterEach, describe, expect, it, vi } from "vitest";
import { installWordMock } from "../test/officeMock";
import { insertImageAtSelection, stripDataUrlPrefix } from "./wordClient";

describe("wordClient", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("strips the PNG data URL prefix for Office.js insertion", () => {
    expect(stripDataUrlPrefix("data:image/png;base64,abc123")).toBe("abc123");
  });

  it("inserts the image at the current Word selection", async () => {
    const mock = installWordMock();

    await insertImageAtSelection("data:image/png;base64,abc123");

    expect(mock.insertInlinePictureFromBase64).toHaveBeenCalledWith(
      "abc123",
      "Replace"
    );
    expect(mock.sync).toHaveBeenCalled();
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test:run -- src/word/wordClient.test.ts
```

Expected: FAIL with import errors for `../test/officeMock` and `./wordClient`.

- [ ] **Step 3: Create Office.js mock helper**

Create `src/test/officeMock.ts`:

```ts
import { vi } from "vitest";

export function installWordMock() {
  const insertInlinePictureFromBase64 = vi.fn();
  const sync = vi.fn().mockResolvedValue(undefined);
  const getSelection = vi.fn(() => ({ insertInlinePictureFromBase64 }));
  const run = vi.fn(async (callback: (context: unknown) => Promise<void>) => {
    await callback({
      document: { getSelection },
      sync
    });
  });

  vi.stubGlobal("Word", {
    run,
    InsertLocation: {
      replace: "Replace"
    }
  });

  return {
    run,
    sync,
    getSelection,
    insertInlinePictureFromBase64
  };
}
```

- [ ] **Step 4: Implement Word insertion adapter**

Create `src/word/wordClient.ts`:

```ts
type WordContext = {
  document: {
    getSelection: () => {
      insertInlinePictureFromBase64: (
        base64: string,
        location: "Replace"
      ) => void;
    };
  };
  sync: () => Promise<void>;
};

type WordGlobal = {
  run: (callback: (context: WordContext) => Promise<void>) => Promise<void>;
  InsertLocation?: {
    replace?: "Replace";
  };
};

declare global {
  interface Window {
    Word?: WordGlobal;
  }

  const Word: WordGlobal | undefined;
}

export function stripDataUrlPrefix(dataUrl: string): string {
  return dataUrl.replace(/^data:image\/png;base64,/, "");
}

export async function insertImageAtSelection(dataUrl: string): Promise<void> {
  const word = globalThis.Word;
  if (!word?.run) {
    throw new Error("目前環境無法使用 Word 插入功能。");
  }

  await word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertInlinePictureFromBase64(
      stripDataUrlPrefix(dataUrl),
      word.InsertLocation?.replace ?? "Replace"
    );
    await context.sync();
  });
}
```

- [ ] **Step 5: Run Word adapter tests**

Run:

```bash
npm test:run -- src/word/wordClient.test.ts
```

Expected: PASS for all Word adapter tests.

- [ ] **Step 6: Commit**

```bash
git add src/word/wordClient.ts src/test/officeMock.ts src/word/wordClient.test.ts
git commit -m "feat: add word insertion adapter"
```

---

### Task 7: Task Pane UI and Editing Workflow

**Files:**
- Create: `src/components/ImagePicker.tsx`
- Create: `src/components/Toolbar.tsx`
- Create: `src/components/CanvasStage.tsx`
- Create: `src/components/PropertyPanel.tsx`
- Create: `src/components/StatusMessage.tsx`
- Create: `src/hooks/useEditorController.ts`
- Modify: `src/App.tsx`
- Modify: `src/styles.css`
- Test: `src/App.workflow.test.tsx`

**Interfaces:**
- Consumes:
  - `initialEditorState`, `editorReducer`
  - `createAnnotation`
  - `loadImageFile`
  - `AnnotationSvg`
  - `exportAnnotatedPng`
  - `insertImageAtSelection`
- Produces:
  - `useEditorController(): EditorController`
  - task pane UI that supports upload, tool selection, annotation creation, property edits, delete, duplicate, undo, redo, export, and Word insertion.

- [ ] **Step 1: Write failing workflow test**

Create `src/App.workflow.test.tsx`:

```tsx
import { fireEvent, render, screen, waitFor } from "@testing-library/react";
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
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test:run -- src/App.workflow.test.tsx
```

Expected: FAIL with missing component and hook imports after `App.tsx` is changed, or current smoke-only app lacks upload/editing controls.

- [ ] **Step 3: Implement editor controller hook**

Create `src/hooks/useEditorController.ts`:

```tsx
import { useCallback, useMemo, useReducer, useState } from "react";
import { createAnnotation } from "../editor/defaults";
import { editorReducer, initialEditorState } from "../editor/reducer";
import type { Annotation, Tool } from "../editor/types";
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
  updateSelected: (patch: Partial<Annotation>) => void;
  deleteSelected: () => void;
  duplicateSelected: () => void;
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
        annotation: createAnnotation(state.activeTool, { x, y })
      });
      dispatch({ type: "tool/set", tool: "select" });
    },
    [state.activeTool]
  );

  return {
    state,
    status,
    error,
    selectedAnnotation,
    setTool: (tool) => dispatch({ type: "tool/set", tool }),
    loadFile,
    addAnnotationAt,
    selectAnnotation: (id) => dispatch({ type: "annotation/select", id }),
    updateSelected: (patch) => {
      if (state.selectedId) {
        dispatch({ type: "annotation/update", id: state.selectedId, patch });
      }
    },
    deleteSelected: () => dispatch({ type: "annotation/delete-selected" }),
    duplicateSelected: () => dispatch({ type: "annotation/duplicate-selected" }),
    undo: () => dispatch({ type: "history/undo" }),
    redo: () => dispatch({ type: "history/redo" }),
    setStatus,
    setError
  };
}
```

- [ ] **Step 4: Implement UI components**

Create `src/components/ImagePicker.tsx`:

```tsx
export function ImagePicker({
  imageName,
  onFile
}: {
  imageName: string | null;
  onFile: (file: File) => void;
}) {
  return (
    <section className="panel-section">
      <label className="file-input">
        <span>上傳圖片</span>
        <input
          aria-label="上傳圖片"
          type="file"
          accept="image/png,image/jpeg,image/webp"
          onChange={(event) => {
            const file = event.currentTarget.files?.[0];
            if (file) onFile(file);
          }}
        />
      </label>
      {imageName ? <p className="image-name">{imageName}</p> : null}
    </section>
  );
}
```

Create `src/components/Toolbar.tsx`:

```tsx
import { Circle, MousePointer2, Square, Type, Undo2, Redo2, Trash2, Copy, ArrowRight, Minus, MessageSquare } from "lucide-react";
import type { Tool } from "../editor/types";

const tools: Array<{ tool: Tool; label: string; icon: JSX.Element }> = [
  { tool: "select", label: "選取", icon: <MousePointer2 size={16} /> },
  { tool: "text", label: "文字", icon: <Type size={16} /> },
  { tool: "arrow", label: "箭頭", icon: <ArrowRight size={16} /> },
  { tool: "line", label: "線條", icon: <Minus size={16} /> },
  { tool: "rectangle", label: "方框", icon: <Square size={16} /> },
  { tool: "ellipse", label: "圓形", icon: <Circle size={16} /> },
  { tool: "callout", label: "標籤箭頭", icon: <MessageSquare size={16} /> }
];

export function Toolbar({
  activeTool,
  canEdit,
  onTool,
  onUndo,
  onRedo,
  onDuplicate,
  onDelete
}: {
  activeTool: Tool;
  canEdit: boolean;
  onTool: (tool: Tool) => void;
  onUndo: () => void;
  onRedo: () => void;
  onDuplicate: () => void;
  onDelete: () => void;
}) {
  return (
    <section className="toolbar" aria-label="標註工具">
      {tools.map((item) => (
        <button
          key={item.tool}
          type="button"
          aria-pressed={activeTool === item.tool}
          aria-label={item.label}
          title={item.label}
          onClick={() => onTool(item.tool)}
        >
          {item.icon}
          <span>{item.label}</span>
        </button>
      ))}
      <button type="button" aria-label="復原" title="復原" onClick={onUndo}>
        <Undo2 size={16} />
      </button>
      <button type="button" aria-label="重做" title="重做" onClick={onRedo}>
        <Redo2 size={16} />
      </button>
      <button type="button" aria-label="複製" title="複製" disabled={!canEdit} onClick={onDuplicate}>
        <Copy size={16} />
      </button>
      <button type="button" aria-label="刪除" title="刪除" disabled={!canEdit} onClick={onDelete}>
        <Trash2 size={16} />
      </button>
    </section>
  );
}
```

Create `src/components/StatusMessage.tsx`:

```tsx
export function StatusMessage({ status, error }: { status: string; error: string }) {
  if (error) return <p className="status error">{error}</p>;
  if (status) return <p className="status">{status}</p>;
  return null;
}
```

Create `src/components/CanvasStage.tsx`:

```tsx
import type { LoadedImage, Annotation } from "../editor/types";
import { AnnotationSvg } from "../render/annotationSvg";

export function CanvasStage({
  image,
  annotations,
  selectedId,
  onAddAt,
  onSelect
}: {
  image: LoadedImage | null;
  annotations: Annotation[];
  selectedId: string | null;
  onAddAt: (x: number, y: number) => void;
  onSelect: (id: string | null) => void;
}) {
  if (!image) {
    return (
      <section className="empty-state" aria-label="尚未載入圖片">
        <p>上傳圖片後即可加入文字、箭頭與基礎圖案。</p>
      </section>
    );
  }

  return (
    <section className="canvas-wrap">
      <svg
        aria-label="圖片標註畫布"
        className="canvas-stage"
        viewBox={`0 0 ${image.width} ${image.height}`}
        onPointerDown={(event) => {
          const rect = event.currentTarget.getBoundingClientRect();
          const x = ((event.clientX - rect.left) / rect.width) * image.width;
          const y = ((event.clientY - rect.top) / rect.height) * image.height;
          onSelect(null);
          onAddAt(Math.round(x), Math.round(y));
        }}
      >
        <image href={image.src} width={image.width} height={image.height} />
        <AnnotationSvg
          annotations={annotations}
          selectedId={selectedId}
          onSelect={onSelect}
        />
      </svg>
    </section>
  );
}
```

Create `src/components/PropertyPanel.tsx`:

```tsx
import type { Annotation } from "../editor/types";

export function PropertyPanel({
  annotation,
  onChange
}: {
  annotation: Annotation | null;
  onChange: (patch: Partial<Annotation>) => void;
}) {
  if (!annotation) {
    return (
      <section className="panel-section" aria-label="標註屬性">
        <p className="muted">選取標註後可調整文字、顏色與線寬。</p>
      </section>
    );
  }

  return (
    <section className="panel-section properties" aria-label="標註屬性">
      {"text" in annotation ? (
        <label>
          標註文字
          <input
            aria-label="標註文字"
            value={annotation.text}
            onChange={(event) => onChange({ text: event.currentTarget.value } as Partial<Annotation>)}
          />
        </label>
      ) : null}
      <label>
        線條顏色
        <input
          aria-label="線條顏色"
          type="color"
          value={annotation.stroke}
          onChange={(event) => onChange({ stroke: event.currentTarget.value } as Partial<Annotation>)}
        />
      </label>
      <label>
        線寬
        <input
          aria-label="線寬"
          type="number"
          min="1"
          max="12"
          value={annotation.strokeWidth}
          onChange={(event) => onChange({ strokeWidth: Number(event.currentTarget.value) } as Partial<Annotation>)}
        />
      </label>
    </section>
  );
}
```

- [ ] **Step 5: Wire App to components and insertion flow**

Replace `src/App.tsx` with:

```tsx
import { CanvasStage } from "./components/CanvasStage";
import { ImagePicker } from "./components/ImagePicker";
import { PropertyPanel } from "./components/PropertyPanel";
import { StatusMessage } from "./components/StatusMessage";
import { Toolbar } from "./components/Toolbar";
import { useEditorController } from "./hooks/useEditorController";
import { exportAnnotatedPng } from "./render/exportPng";
import { insertImageAtSelection } from "./word/wordClient";

export function App() {
  const editor = useEditorController();
  const canInsert = Boolean(editor.state.image);

  async function handleInsert() {
    if (!editor.state.image) {
      editor.setError("請先上傳圖片。");
      return;
    }
    editor.setError("");
    editor.setStatus("正在插入 Word...");
    try {
      const png = await exportAnnotatedPng({
        image: editor.state.image,
        annotations: editor.state.annotations
      });
      await insertImageAtSelection(png);
      editor.setStatus("已插入 Word。");
    } catch (caught) {
      const message = caught instanceof Error ? caught.message : "無法插入 Word。";
      editor.setError(message);
      editor.setStatus("");
    }
  }

  return (
    <main className="task-pane" aria-label="Word 圖片標註工具">
      <header className="app-header">
        <h1>圖片標註</h1>
        <button type="button" disabled={!canInsert} onClick={handleInsert}>
          插入 Word
        </button>
      </header>

      <ImagePicker
        imageName={editor.state.image?.name ?? null}
        onFile={(file) => void editor.loadFile(file)}
      />

      <Toolbar
        activeTool={editor.state.activeTool}
        canEdit={Boolean(editor.selectedAnnotation)}
        onTool={editor.setTool}
        onUndo={editor.undo}
        onRedo={editor.redo}
        onDuplicate={editor.duplicateSelected}
        onDelete={editor.deleteSelected}
      />

      <CanvasStage
        image={editor.state.image}
        annotations={editor.state.annotations}
        selectedId={editor.state.selectedId}
        onAddAt={editor.addAnnotationAt}
        onSelect={editor.selectAnnotation}
      />

      <PropertyPanel
        annotation={editor.selectedAnnotation}
        onChange={editor.updateSelected}
      />

      <StatusMessage status={editor.status} error={editor.error} />
    </main>
  );
}
```

- [ ] **Step 6: Replace styles with full task pane styling**

Replace `src/styles.css` with:

```css
:root {
  color: #1f2933;
  background: #f6f7f9;
  font-family:
    Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont,
    "Segoe UI", sans-serif;
}

* {
  box-sizing: border-box;
}

body {
  margin: 0;
  min-width: 320px;
  background: #f6f7f9;
}

button,
input,
select,
textarea {
  font: inherit;
}

.task-pane {
  min-height: 100vh;
  padding: 12px;
  display: flex;
  flex-direction: column;
  gap: 10px;
}

.app-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
}

.app-header h1 {
  margin: 0;
  font-size: 20px;
  font-weight: 700;
}

.app-header button,
.toolbar button {
  border: 1px solid #2563eb;
  border-radius: 6px;
  background: #2563eb;
  color: white;
  padding: 8px 10px;
}

.app-header button:disabled,
.toolbar button:disabled {
  border-color: #b9c2cf;
  background: #d8dee8;
  color: #64748b;
}

.panel-section {
  border: 1px solid #d5dce6;
  border-radius: 8px;
  background: white;
  padding: 10px;
}

.file-input {
  display: grid;
  gap: 6px;
  font-weight: 600;
}

.image-name,
.muted,
.status {
  margin: 8px 0 0;
  color: #52606d;
  line-height: 1.5;
}

.status {
  border: 1px solid #b7d7c2;
  border-radius: 6px;
  background: #eefaf1;
  color: #166534;
  padding: 8px;
}

.status.error {
  border-color: #fecaca;
  background: #fef2f2;
  color: #b91c1c;
}

.toolbar {
  display: grid;
  grid-template-columns: repeat(4, minmax(0, 1fr));
  gap: 6px;
}

.toolbar button {
  min-height: 38px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 4px;
  padding: 7px;
}

.toolbar button[aria-pressed="true"] {
  outline: 2px solid #111827;
  outline-offset: 1px;
}

.toolbar span {
  font-size: 12px;
}

.empty-state {
  border: 1px dashed #9aa7b7;
  border-radius: 8px;
  padding: 16px;
  background: white;
}

.empty-state p {
  margin: 0;
  color: #52606d;
  line-height: 1.5;
}

.canvas-wrap {
  border: 1px solid #d5dce6;
  border-radius: 8px;
  background: #ffffff;
  overflow: auto;
}

.canvas-stage {
  display: block;
  width: 100%;
  height: auto;
  max-height: 58vh;
}

.properties {
  display: grid;
  gap: 8px;
}

.properties label {
  display: grid;
  gap: 4px;
  color: #334155;
  font-size: 13px;
}

.properties input {
  min-height: 34px;
  border: 1px solid #cbd5e1;
  border-radius: 6px;
  padding: 6px 8px;
}
```

- [ ] **Step 7: Run workflow test**

Run:

```bash
npm test:run -- src/App.workflow.test.tsx
```

Expected: PASS for `uploads an image, adds a text label, edits it, and inserts into Word`.

- [ ] **Step 8: Run smoke test**

Run:

```bash
npm test:run -- src/App.test.tsx
```

Expected: PASS with the heading visible and the `插入 Word` button disabled before upload.

- [ ] **Step 9: Commit**

```bash
git add src/components/ImagePicker.tsx src/components/Toolbar.tsx src/components/CanvasStage.tsx src/components/PropertyPanel.tsx src/components/StatusMessage.tsx src/hooks/useEditorController.ts src/App.tsx src/styles.css src/App.workflow.test.tsx
git commit -m "feat: add task pane annotation workflow"
```

---

### Task 8: Keyboard Shortcuts, Movement, Resize, and Zoom

**Files:**
- Create: `src/editor/geometry.ts`
- Modify: `src/editor/types.ts`
- Modify: `src/editor/reducer.ts`
- Modify: `src/components/CanvasStage.tsx`
- Modify: `src/hooks/useEditorController.ts`
- Test: `src/editor/geometry.test.ts`
- Test: `src/App.keyboard.test.tsx`

**Interfaces:**
- Consumes:
  - `Annotation`, `EditorAction`, `EditorState`
- Produces:
  - `moveAnnotation(annotation: Annotation, delta: Point): Annotation`
  - `resizeShape(annotation: Annotation, width: number, height: number): Annotation`
  - keyboard behavior for Delete, Ctrl/Cmd+Z, Ctrl/Cmd+Shift+Z, and Ctrl/Cmd+Y
  - zoom actions clamped from `0.25` to `4`

- [ ] **Step 1: Write geometry tests**

Create `src/editor/geometry.test.ts`:

```ts
import { describe, expect, it } from "vitest";
import type { Annotation } from "./types";
import { moveAnnotation, resizeShape } from "./geometry";

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
```

- [ ] **Step 2: Write keyboard workflow test**

Create `src/App.keyboard.test.tsx`:

```tsx
import { fireEvent, render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, expect, it, vi } from "vitest";
import { App } from "./App";

vi.mock("./image/imageValidation", () => ({
  loadImageFile: vi.fn(async (file: File) => ({
    src: "data:image/png;base64,input",
    name: file.name,
    width: 640,
    height: 480
  }))
}));

describe("keyboard editing", () => {
  it("deletes selected annotations and supports undo", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.upload(
      screen.getByLabelText("上傳圖片"),
      new File(["image"], "assembly.png", { type: "image/png" })
    );
    await user.click(screen.getByRole("button", { name: "文字" }));
    fireEvent.pointerDown(screen.getByLabelText("圖片標註畫布"), {
      clientX: 80,
      clientY: 100
    });

    expect(screen.getByText("零件名稱")).toBeInTheDocument();

    await user.keyboard("{Delete}");
    expect(screen.queryByText("零件名稱")).not.toBeInTheDocument();

    await user.keyboard("{Meta>}z{/Meta}");
    expect(screen.getByText("零件名稱")).toBeInTheDocument();
  });
});
```

- [ ] **Step 3: Run tests to verify failures**

Run:

```bash
npm test:run -- src/editor/geometry.test.ts src/App.keyboard.test.tsx
```

Expected: FAIL with import error for `./geometry` and missing keyboard behavior.

- [ ] **Step 4: Implement geometry helpers**

Create `src/editor/geometry.ts`:

```ts
import type { Annotation, Point } from "./types";

export function moveAnnotation(annotation: Annotation, delta: Point): Annotation {
  if (annotation.type === "line" || annotation.type === "arrow") {
    return {
      ...annotation,
      x1: annotation.x1 + delta.x,
      y1: annotation.y1 + delta.y,
      x2: annotation.x2 + delta.x,
      y2: annotation.y2 + delta.y
    };
  }

  if (annotation.type === "callout") {
    return {
      ...annotation,
      x: annotation.x + delta.x,
      y: annotation.y + delta.y,
      targetX: annotation.targetX + delta.x,
      targetY: annotation.targetY + delta.y
    };
  }

  return {
    ...annotation,
    x: annotation.x + delta.x,
    y: annotation.y + delta.y
  };
}

export function resizeShape(annotation: Annotation, width: number, height: number): Annotation {
  if (annotation.type !== "rectangle" && annotation.type !== "ellipse") {
    return annotation;
  }

  return {
    ...annotation,
    width: Math.max(12, width),
    height: Math.max(12, height)
  };
}
```

- [ ] **Step 5: Extend reducer actions**

Modify `src/editor/types.ts` by adding these actions to `EditorAction`:

```ts
  | { type: "annotation/move-selected"; delta: Point }
  | { type: "annotation/resize-selected"; width: number; height: number }
```

Modify `src/editor/reducer.ts` imports:

```ts
import { moveAnnotation, resizeShape } from "./geometry";
```

Add these cases before `annotations/clear`:

```ts
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
```

- [ ] **Step 6: Add keyboard handling to controller**

Modify `src/hooks/useEditorController.ts` imports:

```tsx
import { useCallback, useEffect, useMemo, useReducer, useState } from "react";
```

Inside `useEditorController`, before `return`, add:

```tsx
  useEffect(() => {
    function onKeyDown(event: KeyboardEvent) {
      const isModifier = event.metaKey || event.ctrlKey;
      if (event.key === "Delete" || event.key === "Backspace") {
        dispatch({ type: "annotation/delete-selected" });
        return;
      }
      if (isModifier && event.key.toLowerCase() === "z" && event.shiftKey) {
        event.preventDefault();
        dispatch({ type: "history/redo" });
        return;
      }
      if (isModifier && event.key.toLowerCase() === "z") {
        event.preventDefault();
        dispatch({ type: "history/undo" });
        return;
      }
      if (isModifier && event.key.toLowerCase() === "y") {
        event.preventDefault();
        dispatch({ type: "history/redo" });
      }
    }

    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, []);
```

Add these two properties to the returned controller object:

```tsx
    moveSelected: (delta) =>
      dispatch({ type: "annotation/move-selected", delta }),
    resizeSelected: (width, height) =>
      dispatch({ type: "annotation/resize-selected", width, height }),
```

Add these two properties to `EditorController`:

```tsx
  moveSelected: (delta: { x: number; y: number }) => void;
  resizeSelected: (width: number, height: number) => void;
```

- [ ] **Step 7: Run keyboard and geometry tests**

Run:

```bash
npm test:run -- src/editor/geometry.test.ts src/App.keyboard.test.tsx
```

Expected: PASS for all tests.

- [ ] **Step 8: Run full test suite**

Run:

```bash
npm test:run
```

Expected: PASS for all tests created so far.

- [ ] **Step 9: Commit**

```bash
git add src/editor/geometry.ts src/editor/geometry.test.ts src/editor/types.ts src/editor/reducer.ts src/hooks/useEditorController.ts src/App.keyboard.test.tsx
git commit -m "feat: add keyboard editing and geometry helpers"
```

---

### Task 9: Office Manifest and Sideload Documentation

**Files:**
- Create: `public/manifest.xml`
- Create: `docs/dev/word-sideloading.md`
- Modify: `package.json`
- Test: `public/manifest.xml` via XML parser command

**Interfaces:**
- Consumes: Vite dev server from Task 1.
- Produces:
  - Word task pane manifest pointing to `https://localhost:3000/index.html`
  - npm script `dev:https` for add-in hosting
  - developer instructions for local sideloading

- [ ] **Step 1: Add manifest file**

Create `public/manifest.xml`:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">
  <Id>7f7e9f49-65df-4d52-9f4e-b5ad8f3f9d7a</Id>
  <Version>0.1.0.0</Version>
  <ProviderName>word-pic-editor</ProviderName>
  <DefaultLocale>zh-TW</DefaultLocale>
  <DisplayName DefaultValue="圖片標註" />
  <Description DefaultValue="在 Word 中上傳圖片、加入工程標註並插入文件。" />
  <IconUrl DefaultValue="https://localhost:3000/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://localhost:3000/icon-64.png" />
  <SupportUrl DefaultValue="https://localhost:3000" />
  <AppDomains>
    <AppDomain>https://localhost:3000</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Document" />
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:3000/index.html" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Document">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title" />
            <Description resid="GetStarted.Description" />
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl" />
          </GetStarted>
          <FunctionFile resid="Commands.Url" />
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="AnnotationGroup">
                <Label resid="AnnotationGroup.Label" />
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16" />
                  <bt:Image size="32" resid="Icon.32x32" />
                  <bt:Image size="80" resid="Icon.80x80" />
                </Icon>
                <Control xsi:type="Button" id="OpenPaneButton">
                  <Label resid="OpenPaneButton.Label" />
                  <Supertip>
                    <Title resid="OpenPaneButton.Label" />
                    <Description resid="OpenPaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16" />
                    <bt:Image size="32" resid="Icon.32x32" />
                    <bt:Image size="80" resid="Icon.80x80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonId1</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/icon-16.png" />
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/icon-32.png" />
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/icon-80.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/index.html" />
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/index.html" />
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://localhost:3000" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="圖片標註已就緒" />
        <bt:String id="AnnotationGroup.Label" DefaultValue="圖片標註" />
        <bt:String id="OpenPaneButton.Label" DefaultValue="圖片標註" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="開啟右側面板，上傳圖片並加入工程標註。" />
        <bt:String id="OpenPaneButton.Tooltip" DefaultValue="開啟圖片標註工具面板。" />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

- [ ] **Step 2: Add HTTPS dev script**

Modify `package.json` scripts to include:

```json
{
  "scripts": {
    "dev": "vite --host 127.0.0.1",
    "dev:https": "vite --host 127.0.0.1 --port 3000 --https",
    "build": "tsc --noEmit && vite build",
    "preview": "vite preview --host 127.0.0.1",
    "test": "vitest",
    "test:run": "vitest run",
    "typecheck": "tsc --noEmit"
  }
}
```

- [ ] **Step 3: Add sideloading guide**

Create `docs/dev/word-sideloading.md`:

```md
# Word Add-in Local Sideloading

## Run the task pane locally

1. Install dependencies:

   ```bash
   npm install
   ```

2. Start the HTTPS dev server:

   ```bash
   npm run dev:https
   ```

3. Open `https://localhost:3000` once in the browser and accept the local certificate warning for development.

## Sideload into Word

1. Use `public/manifest.xml` as the Office Add-in manifest.
2. In Word desktop, open the add-ins management flow for sideloaded add-ins.
3. Load the manifest.
4. Open the `圖片標註` task pane from the Home ribbon.
5. Upload an image, add a text label and arrow, then click `插入 Word`.

## Manual acceptance checklist

- The task pane opens on the right side of Word.
- PNG, JPG, JPEG, and WEBP uploads are accepted.
- Unsupported image types show `僅支援 PNG、JPG、JPEG、WEBP 圖片。`
- A text label can be edited to a part name such as `馬達`.
- Arrow, line, rectangle, ellipse, and callout tools create visible annotations.
- Delete removes the selected annotation.
- Undo restores the deleted annotation.
- `插入 Word` inserts one flattened PNG at the current Word selection.
- Inserted labels remain readable after resizing the image in Word.
```

- [ ] **Step 4: Validate XML syntax**

Run:

```bash
python3 -m xml.dom.minidom public/manifest.xml >/tmp/word-pic-editor-manifest.xml
```

Expected: command exits `0` with no XML parsing error.

- [ ] **Step 5: Run build and tests**

Run:

```bash
npm run build
npm test:run
```

Expected: build succeeds and all tests pass.

- [ ] **Step 6: Commit**

```bash
git add public/manifest.xml docs/dev/word-sideloading.md package.json package-lock.json
git commit -m "chore: add word add-in manifest and sideload docs"
```

---

### Task 10: Final Verification and Product Fit Pass

**Files:**
- Modify: `src/styles.css`
- Modify: `docs/dev/word-sideloading.md`
- Test: all automated tests and manual browser smoke

**Interfaces:**
- Consumes: completed app, manifest, tests, and docs from prior tasks.
- Produces: verified first-version add-in ready for implementation review.

- [ ] **Step 1: Run the complete automated suite**

Run:

```bash
npm run typecheck
npm test:run
npm run build
```

Expected: all commands pass.

- [ ] **Step 2: Start local app for browser smoke**

Run:

```bash
npm run dev -- --port 5173
```

Expected: Vite reports a local URL at `http://127.0.0.1:5173/`.

- [ ] **Step 3: Browser smoke test**

Open `http://127.0.0.1:5173/` and verify:

```md
- Heading `圖片標註` is visible.
- `插入 Word` is disabled before image upload.
- Upload control is visible.
- Tool buttons are visible: 選取, 文字, 箭頭, 線條, 方框, 圓形, 標籤箭頭.
- Empty state text says `上傳圖片後即可加入文字、箭頭與基礎圖案。`
- No toolbar text wraps outside its button at 320px width.
```

- [ ] **Step 4: Verify PRD coverage**

Read `docs/prd/word-image-annotation-addin-prd.md` and confirm each item maps to implementation:

```md
- Word Office Add-in: Task 9 manifest.
- Right-side task pane: Task 9 manifest and Task 7 app shell.
- Upload image: Task 3 and Task 7.
- Text annotations: Task 2, Task 4, Task 7.
- Line, arrow, rectangle, ellipse, callout: Task 2, Task 4, Task 7.
- Edit selected annotation: Task 7 property panel.
- Delete, duplicate, undo, redo: Task 2, Task 7, Task 8.
- PNG export: Task 5.
- Word insertion: Task 6 and Task 7.
- Validation errors: Task 3 and Task 7.
- Manual Word acceptance: Task 9 docs.
```

- [ ] **Step 5: Commit verification updates if files changed**

If Step 3 or Step 4 required style or docs edits, run:

```bash
git add src/styles.css docs/dev/word-sideloading.md
git commit -m "polish: verify first annotation workflow"
```

If no files changed, run:

```bash
git status --short
```

Expected: no uncommitted changes from this task.

---

## Self-Review Notes

- Spec coverage: covered right task pane, image upload, text, basic shapes, callout, edit/delete/duplicate, undo/redo, validation, PNG export, Word insertion, Traditional Chinese copy, and manual Word verification.
- Scope control: excluded advanced photo editing, OCR, AI labeling, cloud storage, collaboration, CAD measurement, mobile Word, and PowerPoint/Excel.
- Test seams: state reducer, validation, SVG rendering, PNG exporter, Word adapter, task pane workflow, keyboard workflow, build, XML manifest validation, and manual Word sideload acceptance.
- Type consistency: shared names are `Tool`, `Annotation`, `LoadedImage`, `EditorState`, `createAnnotation`, `editorReducer`, `AnnotationSvg`, `exportAnnotatedPng`, and `insertImageAtSelection`.
