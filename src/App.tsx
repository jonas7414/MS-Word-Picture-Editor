import { useEffect, useRef, useState } from "react";
import { CanvasStage } from "./components/CanvasStage";
import { ImagePicker } from "./components/ImagePicker";
import { PropertyPanel } from "./components/PropertyPanel";
import { StatusMessage } from "./components/StatusMessage";
import { Toolbar } from "./components/Toolbar";
import { useEditorController } from "./hooks/useEditorController";
import { exportAnnotatedPng } from "./render/exportPng";
import {
  cleanupEditableWordImage,
  type EditableWordImage,
  insertImageAtSelection,
  loadSelectedImageForEditing,
  replaceEditableWordImage
} from "./word/wordClient";

export function App() {
  const editor = useEditorController();
  const [editTarget, setEditTarget] = useState<Pick<
    EditableWordImage,
    "targetDisplaySize" | "targetTag"
  > | null>(null);
  const [isLoadingSelectedImage, setIsLoadingSelectedImage] = useState(false);
  const [autoDismissStatus, setAutoDismissStatus] = useState(false);
  const editTargetRef = useRef(editTarget);
  const canInsert = Boolean(editor.state.image);
  const submitLabel = editTarget ? "取代圖片" : "插入 Word";

  useEffect(() => {
    editTargetRef.current = editTarget;
  }, [editTarget]);

  useEffect(() => {
    return () => {
      const current = editTargetRef.current;
      if (current) {
        void cleanupEditableWordImage(current.targetTag);
      }
    };
  }, []);

  function setWorkingStatus(message: string) {
    setAutoDismissStatus(false);
    editor.setStatus(message);
  }

  function setDoneStatus(message: string) {
    setAutoDismissStatus(true);
    editor.setStatus(message);
  }

  async function cleanupCurrentEditTarget() {
    if (!editTarget) return;

    try {
      await cleanupEditableWordImage(editTarget.targetTag);
    } catch {
      // Cleanup should not block switching workflows.
    }
    setEditTarget(null);
  }

  async function handleInsert() {
    if (!editor.state.image) {
      editor.setError("請先上傳圖片。");
      return;
    }

    editor.setError("");
    setWorkingStatus(editTarget ? "正在取代 Word 圖片..." : "正在插入 Word...");

    try {
      const png = await exportAnnotatedPng({
        image: editor.state.image,
        annotations: editor.state.annotations
      });
      if (editTarget) {
        await replaceEditableWordImage({
          dataUrl: png,
          targetDisplaySize: editTarget.targetDisplaySize,
          targetTag: editTarget.targetTag
        });
        setEditTarget(null);
        setDoneStatus("已取代 Word 圖片。");
      } else {
        await insertImageAtSelection(png);
        setDoneStatus("已插入 Word。");
      }
    } catch (caught) {
      const message =
        caught instanceof Error
          ? caught.message
          : editTarget
            ? "無法取代 Word 圖片。"
            : "無法插入 Word。";
      editor.setError(message);
      editor.setStatus("");
    }
  }

  async function handleLoadSelectedWordImage() {
    editor.setError("");
    setWorkingStatus("正在讀取 Word 選取圖片...");
    setIsLoadingSelectedImage(true);

    try {
      await cleanupCurrentEditTarget();
      const editable = await loadSelectedImageForEditing();
      editor.loadImage(editable.image);
      setEditTarget({
        targetDisplaySize: editable.targetDisplaySize,
        targetTag: editable.targetTag
      });
      setDoneStatus("已載入 Word 選取圖片，完成標註後可取代原圖。");
    } catch (caught) {
      const message =
        caught instanceof Error ? caught.message : "無法讀取 Word 選取圖片。";
      editor.setError(message);
      editor.setStatus("");
    } finally {
      setIsLoadingSelectedImage(false);
    }
  }

  async function handleUploadFile(file: File) {
    await cleanupCurrentEditTarget();
    await editor.loadFile(file);
  }

  return (
    <main aria-label="Word 圖片標註工具" className="task-pane">
      <header className="app-header">
        <div>
          <h1>圖片標註</h1>
          <p>Word 圖片註解工具</p>
        </div>
        <ImagePicker
          imageName={editor.state.image?.name ?? null}
          onFile={(file) => void handleUploadFile(file)}
        />
        <button
          className="secondary-action"
          disabled={isLoadingSelectedImage}
          type="button"
          onClick={() => void handleLoadSelectedWordImage()}
        >
          {isLoadingSelectedImage ? "讀取中..." : "編輯選取圖片"}
        </button>
        <button disabled={!canInsert} type="button" onClick={handleInsert}>
          {submitLabel}
        </button>
      </header>

      <section className="workspace" aria-label="圖片編輯工作區">
        <CanvasStage
          annotations={editor.state.annotations}
          activeTool={editor.state.activeTool}
          image={editor.state.image}
          pan={editor.state.pan}
          selectedId={editor.state.selectedId}
          zoom={editor.state.zoom}
          onAddAt={editor.addAnnotationAt}
          onMove={editor.moveAnnotation}
          onPan={editor.panViewport}
          onSelect={editor.selectAnnotation}
          onUpdate={editor.updateAnnotation}
          onZoom={editor.setZoom}
        />

        <div className="floating-toolbar">
          <Toolbar
            activeTool={editor.state.activeTool}
            canEdit={Boolean(editor.selectedAnnotation)}
            zoom={editor.state.zoom}
            onDelete={editor.deleteSelected}
            onDuplicate={editor.duplicateSelected}
            onRedo={editor.redo}
            onTool={editor.setTool}
            onUndo={editor.undo}
            onZoomIn={() => editor.setZoom(editor.state.zoom * 1.25)}
            onZoomOut={() => editor.setZoom(editor.state.zoom / 1.25)}
            onZoomReset={editor.resetViewport}
          />
        </div>

        <div className="floating-properties">
          <PropertyPanel
            annotation={editor.selectedAnnotation}
            onChange={editor.updateSelected}
          />
        </div>

        <div className="floating-status">
          <StatusMessage
            error={editor.error}
            status={editor.status}
            autoDismiss={autoDismissStatus}
            onDismiss={() => {
              editor.setError("");
              editor.setStatus("");
              setAutoDismissStatus(false);
            }}
          />
        </div>
      </section>
    </main>
  );
}
