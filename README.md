# MS Word Picture Editor

Microsoft Word 圖片標註增益集。可在 Word 右側面板中上傳圖片、加入文字、箭頭、方框、虛線框、圓形與標籤箭頭，最後插入 Word；也可以載入文件中已選取的圖片，標註後取代原圖。

## 線上版本

App：

```text
https://jonas7414.github.io/MS-Word-Picture-Editor/
```

Manifest：

```text
https://jonas7414.github.io/MS-Word-Picture-Editor/manifest.xml
```

GitHub Releases：

```text
https://github.com/jonas7414/MS-Word-Picture-Editor/releases
```

## Word Web 版安裝方式

1. 下載 release 裡的 `manifest.xml`，或開啟：

   ```text
   https://jonas7414.github.io/MS-Word-Picture-Editor/manifest.xml
   ```

   然後另存成 `manifest.xml`。

2. 開啟 Word Web 版：

   ```text
   https://www.office.com/
   ```

3. 開啟任一 Word 文件。

4. 到 `插入` -> `增益集` -> `我的增益集`。

5. 選擇 `上傳我的增益集`。

6. 選取剛剛下載的 `manifest.xml`。

7. 成功後，功能區會出現 `圖片標註`，點擊即可開啟右側面板。

## Word 桌面版安裝方式

### Windows

1. 建立資料夾，例如：

   ```text
   C:\OfficeAddins
   ```

2. 下載 `manifest.xml`，放進該資料夾。

3. 開啟 Word。

4. 到 `檔案` -> `選項` -> `信任中心` -> `信任中心設定`。

5. 找到 `受信任的增益集目錄`。

6. 新增共用資料夾路徑：

   ```text
   C:\OfficeAddins
   ```

7. 勾選允許顯示在功能表中，儲存後重新啟動 Word。

8. 到 `插入` -> `我的增益集` -> `共用資料夾`。

9. 選擇 `圖片標註`。

### Mac

1. 下載 `manifest.xml`。

2. 將檔案放到 Office add-ins sideload 目錄。常見位置：

   ```text
   ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
   ```

   如果 `wef` 資料夾不存在，請自行建立。

3. 重新啟動 Word。

4. 到 `插入` -> `我的增益集`，選擇 `圖片標註`。

## 使用方式

### 插入新圖片

1. 開啟 `圖片標註` 面板。
2. 點 `上傳圖片`。
3. 使用工具列加入文字、箭頭、線條、方框、虛線框、圓形或標籤箭頭。
4. 可用 `Ctrl + 滾輪` 放大縮小，或工具列的縮放按鈕。
5. 完成後點 `插入 Word`。

### 編輯文件中已有圖片

1. 在 Word 文件中選取一張嵌入式圖片。
2. 開啟 `圖片標註` 面板。
3. 點 `編輯選取圖片`。
4. 標註完成後點 `取代圖片`。

注意：目前支援 Word 的嵌入式圖片。如果圖片是浮動版面物件，請先將文繞圖改成 `與文字排列`。

## 開發

```bash
npm install
npm run dev:https
```

本機 sideload 使用：

```text
public/manifest.xml
```

## 發佈

GitHub Pages 會在 push 到 `master` 後自動部署。

要發新 GitHub Release，更新 `package.json` 的版本號，例如：

```json
"version": "0.1.1"
```

commit 並 push 後，Release workflow 會自動建立：

```text
v0.1.1
```

並附上：

```text
manifest.xml
word-pic-editor-v0.1.1.zip
```
