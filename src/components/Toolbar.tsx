import {
  ArrowRight,
  Circle,
  Copy,
  MessageSquare,
  Minus,
  MousePointer2,
  Redo2,
  Square,
  SquareDashed,
  Trash2,
  Type,
  Undo2,
  ZoomIn,
  ZoomOut
} from "lucide-react";
import type { ReactNode } from "react";
import type { Tool } from "../editor/types";

const tools: Array<{ tool: Tool; label: string; icon: ReactNode }> = [
  { tool: "select", label: "選取", icon: <MousePointer2 size={16} /> },
  { tool: "text", label: "文字", icon: <Type size={16} /> },
  { tool: "arrow", label: "箭頭", icon: <ArrowRight size={16} /> },
  { tool: "line", label: "線條", icon: <Minus size={16} /> },
  { tool: "rectangle", label: "方框", icon: <Square size={16} /> },
  { tool: "dashedRectangle", label: "虛線框", icon: <SquareDashed size={16} /> },
  { tool: "ellipse", label: "圓形", icon: <Circle size={16} /> },
  { tool: "callout", label: "標籤箭頭", icon: <MessageSquare size={16} /> }
];

export function Toolbar({
  activeTool,
  canEdit,
  zoom,
  onTool,
  onZoomIn,
  onZoomOut,
  onZoomReset,
  onUndo,
  onRedo,
  onDuplicate,
  onDelete
}: {
  activeTool: Tool;
  canEdit: boolean;
  zoom: number;
  onTool: (tool: Tool) => void;
  onZoomIn: () => void;
  onZoomOut: () => void;
  onZoomReset: () => void;
  onUndo: () => void;
  onRedo: () => void;
  onDuplicate: () => void;
  onDelete: () => void;
}) {
  return (
    <section aria-label="標註工具" className="toolbar">
      {tools.map((item) => (
        <button
          aria-label={item.label}
          aria-pressed={activeTool === item.tool}
          key={item.tool}
          title={item.label}
          type="button"
          onClick={() => onTool(item.tool)}
        >
          {item.icon}
          <span>{item.label}</span>
        </button>
      ))}
      <button aria-label="復原" title="復原" type="button" onClick={onUndo}>
        <Undo2 size={16} />
      </button>
      <button aria-label="重做" title="重做" type="button" onClick={onRedo}>
        <Redo2 size={16} />
      </button>
      <button aria-label="縮小" title="縮小" type="button" onClick={onZoomOut}>
        <ZoomOut size={16} />
      </button>
      <button
        aria-label="重設縮放"
        className="zoom-value"
        title="重設縮放"
        type="button"
        onClick={onZoomReset}
      >
        {Math.round(zoom * 100)}%
      </button>
      <button aria-label="放大" title="放大" type="button" onClick={onZoomIn}>
        <ZoomIn size={16} />
      </button>
      <button
        aria-label="複製"
        disabled={!canEdit}
        title="複製"
        type="button"
        onClick={onDuplicate}
      >
        <Copy size={16} />
      </button>
      <button
        aria-label="刪除"
        disabled={!canEdit}
        title="刪除"
        type="button"
        onClick={onDelete}
      >
        <Trash2 size={16} />
      </button>
    </section>
  );
}
