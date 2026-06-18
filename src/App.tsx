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
    <main aria-label="Word 圖片標註工具" className="task-pane">
      <header className="app-header">
        <div>
          <h1>圖片標註</h1>
          <p>Word 圖片註解工具</p>
        </div>
        <ImagePicker
          imageName={editor.state.image?.name ?? null}
          onFile={(file) => void editor.loadFile(file)}
        />
        <button disabled={!canInsert} type="button" onClick={handleInsert}>
          插入 Word
        </button>
      </header>

      <section className="workspace" aria-label="圖片編輯工作區">
        <CanvasStage
          annotations={editor.state.annotations}
          image={editor.state.image}
          selectedId={editor.state.selectedId}
          onAddAt={editor.addAnnotationAt}
          onMove={editor.moveAnnotation}
          onSelect={editor.selectAnnotation}
          onUpdate={editor.updateAnnotation}
        />

        <div className="floating-toolbar">
          <Toolbar
            activeTool={editor.state.activeTool}
            canEdit={Boolean(editor.selectedAnnotation)}
            onDelete={editor.deleteSelected}
            onDuplicate={editor.duplicateSelected}
            onRedo={editor.redo}
            onTool={editor.setTool}
            onUndo={editor.undo}
          />
        </div>

        <div className="floating-properties">
          <PropertyPanel
            annotation={editor.selectedAnnotation}
            onChange={editor.updateSelected}
          />
        </div>

        <div className="floating-status">
          <StatusMessage error={editor.error} status={editor.status} />
        </div>
      </section>
    </main>
  );
}
