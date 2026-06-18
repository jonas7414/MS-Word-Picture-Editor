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
          accept="image/png,image/jpeg,image/webp"
          type="file"
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
