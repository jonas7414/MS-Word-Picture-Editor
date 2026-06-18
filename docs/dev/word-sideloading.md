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

   This starts Vite with a local self-signed certificate at `https://localhost:3000`, which matches the manifest.

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
