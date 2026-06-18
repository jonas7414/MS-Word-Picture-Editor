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
      <section aria-label="標註屬性" className="panel-section">
        <p className="muted">選取標註後可調整文字、顏色與線寬。</p>
      </section>
    );
  }

  return (
    <section aria-label="標註屬性" className="panel-section properties">
      {"text" in annotation ? (
        <label>
          文字
          <input
            aria-label="標註文字"
            value={annotation.text}
            onChange={(event) =>
              onChange({ text: event.currentTarget.value } as Partial<Annotation>)
            }
          />
        </label>
      ) : null}
      {"width" in annotation ? (
        <div className="property-grid">
          <label>
            寬
            <input
              aria-label="寬"
              min="12"
              type="number"
              value={Math.round(annotation.width)}
              onChange={(event) =>
                onChange({
                  width: Number(event.currentTarget.value)
                } as Partial<Annotation>)
              }
            />
          </label>
          <label>
            高
            <input
              aria-label="高"
              min="12"
              type="number"
              value={Math.round(annotation.height)}
              onChange={(event) =>
                onChange({
                  height: Number(event.currentTarget.value)
                } as Partial<Annotation>)
              }
            />
          </label>
        </div>
      ) : null}
      {"x2" in annotation ? (
        <div className="property-grid">
          <label>
            X
            <input
              aria-label="終點 X"
              type="number"
              value={Math.round(annotation.x2)}
              onChange={(event) =>
                onChange({
                  x2: Number(event.currentTarget.value)
                } as Partial<Annotation>)
              }
            />
          </label>
          <label>
            Y
            <input
              aria-label="終點 Y"
              type="number"
              value={Math.round(annotation.y2)}
              onChange={(event) =>
                onChange({
                  y2: Number(event.currentTarget.value)
                } as Partial<Annotation>)
              }
            />
          </label>
        </div>
      ) : null}
      {"targetX" in annotation ? (
        <div className="property-grid">
          <label>
            X
            <input
              aria-label="指向 X"
              type="number"
              value={Math.round(annotation.targetX)}
              onChange={(event) =>
                onChange({
                  targetX: Number(event.currentTarget.value)
                } as Partial<Annotation>)
              }
            />
          </label>
          <label>
            Y
            <input
              aria-label="指向 Y"
              type="number"
              value={Math.round(annotation.targetY)}
              onChange={(event) =>
                onChange({
                  targetY: Number(event.currentTarget.value)
                } as Partial<Annotation>)
              }
            />
          </label>
        </div>
      ) : null}
      <label>
        框線
        <select
          aria-label="框線"
          value={annotation.strokeStyle ?? "solid"}
          onChange={(event) =>
            onChange({
              strokeStyle: event.currentTarget.value as "solid" | "dashed"
            } as Partial<Annotation>)
          }
        >
          <option value="solid">實線</option>
          <option value="dashed">虛線</option>
        </select>
      </label>
      <label>
        顏色
        <input
          aria-label="線條顏色"
          type="color"
          value={annotation.stroke}
          onChange={(event) =>
            onChange({ stroke: event.currentTarget.value } as Partial<Annotation>)
          }
        />
      </label>
      <label>
        線寬
        <input
          aria-label="線寬"
          max="12"
          min="1"
          type="number"
          value={annotation.strokeWidth}
          onChange={(event) =>
            onChange({
              strokeWidth: Number(event.currentTarget.value)
            } as Partial<Annotation>)
          }
        />
      </label>
    </section>
  );
}
