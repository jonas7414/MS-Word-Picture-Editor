# Word 圖片標註小工具 PRD

## Problem Statement

工程師在撰寫 Word 技術文件、報告、教學文件或截圖說明時，經常需要在圖片上標出零件位置、零件名稱、注意事項或操作步驟。目前常見做法是在 Word、PowerPoint、截圖工具或外部圖片編輯軟體之間切換，流程破碎、圖片容易失真，也不方便反覆修改。

使用者需要一個簡單、穩定、專注在工程文件標註情境的 Word Office Add-in。它不需要複雜修圖功能，而是要能在 Word 右側工具面板中匯入圖片、加上文字與基礎圖案標註，最後把標註後的圖片插入 Word 文件。

## Solution

建立一個 Word Office Add-in，在 Word 右側開啟圖片標註工具面板。使用者可以從本機上傳圖片，在面板內使用文字、線條、箭頭、矩形、圓形與簡單樣式設定標記圖片中的零件、位置與說明，完成後將標註結果匯出成單張圖片並插入目前的 Word 文件游標位置。

工具設計以工程師寫 document 的日常需求為主：操作要直覺、功能要少而完整、標註結果要清楚可讀、插入 Word 後要穩定呈現。核心成功標準不是做成完整圖片編輯器，而是讓使用者快速完成零件名稱與圖片區域標記。

## User Stories

1. As an engineer, I want to open the image annotation tool from Word, so that I can annotate images without leaving my document.
2. As an engineer, I want the tool to appear in the Word task pane, so that I can edit annotations beside the document I am writing.
3. As an engineer, I want to upload a local image, so that I can annotate photos, screenshots, or product images.
4. As an engineer, I want to replace the current image with another uploaded image, so that I can correct mistakes without reopening the add-in.
5. As an engineer, I want to see the uploaded image on an editing canvas, so that I can place annotations accurately.
6. As an engineer, I want the image to fit inside the task pane preview area, so that I can work with large images without horizontal scrolling.
7. As an engineer, I want to zoom in and out, so that I can precisely mark small parts.
8. As an engineer, I want to pan around a zoomed image, so that I can annotate different areas of a large image.
9. As an engineer, I want to reset zoom to fit, so that I can quickly return to the full picture.
10. As an engineer, I want to add text labels, so that I can write part names directly on the image.
11. As an engineer, I want to edit text label content after creating it, so that I can fix spelling or naming mistakes.
12. As an engineer, I want to move text labels, so that labels do not cover important image details.
13. As an engineer, I want to change text size, so that labels remain readable after insertion into Word.
14. As an engineer, I want to change text color, so that labels are visible against different image backgrounds.
15. As an engineer, I want to add a text background or outline, so that labels remain readable on busy photos.
16. As an engineer, I want to add straight lines, so that I can point to boundaries, axes, or measurement-related areas.
17. As an engineer, I want to add arrows, so that I can point from a label to a specific part.
18. As an engineer, I want to add rectangles, so that I can highlight a component or region.
19. As an engineer, I want to add circles or ellipses, so that I can mark round parts or localized details.
20. As an engineer, I want to add callouts that combine text and an arrow, so that I can label parts quickly.
21. As an engineer, I want to select an annotation, so that I can modify it after placing it.
22. As an engineer, I want to drag annotations, so that I can align them with the image.
23. As an engineer, I want to resize rectangles and ellipses, so that highlighted regions match the real part size.
24. As an engineer, I want to adjust line and arrow endpoints, so that pointers land on the correct location.
25. As an engineer, I want to delete selected annotations, so that I can remove mistakes.
26. As an engineer, I want undo and redo, so that I can recover from accidental edits.
27. As an engineer, I want to duplicate an annotation, so that repeated part labels can be created quickly.
28. As an engineer, I want to set stroke color, so that annotations stand out from the image.
29. As an engineer, I want to set stroke thickness, so that lines remain visible in Word.
30. As an engineer, I want simple preset colors, so that I can choose readable engineering markup colors quickly.
31. As an engineer, I want default annotation styles, so that I do not need to configure each new mark.
32. As an engineer, I want annotations to use consistent defaults across a session, so that one document has a consistent visual style.
33. As an engineer, I want selected annotations to show handles, so that I understand what can be changed.
34. As an engineer, I want keyboard delete to remove selected annotations, so that editing feels efficient.
35. As an engineer, I want keyboard shortcuts for undo and redo, so that I can work faster.
36. As an engineer, I want annotations to snap lightly or align visually when practical, so that documents look cleaner.
37. As an engineer, I want to preview the final annotated image before inserting it, so that I can catch readability issues.
38. As an engineer, I want to insert the annotated image at the current Word cursor position, so that it appears where I am writing.
39. As an engineer, I want the inserted image to include the base image and all annotations as one flattened image, so that Word layout remains stable.
40. As an engineer, I want the inserted image to preserve reasonable resolution, so that part names remain legible in printed or exported documents.
41. As an engineer, I want clear error messages if image upload fails, so that I know what to correct.
42. As an engineer, I want clear error messages if insertion into Word fails, so that I can retry without losing work.
43. As an engineer, I want the add-in to support common image formats, so that I can use photos and screenshots without conversion.
44. As an engineer, I want the tool to reject unsupported or very large files gracefully, so that Word does not become unstable.
45. As an engineer, I want my current editing work to remain in the pane while I adjust the Word document, so that I do not lose annotations accidentally.
46. As an engineer, I want a simple clear-all action, so that I can restart annotation on the current image.
47. As an engineer, I want a confirmation before discarding existing annotation work, so that I do not lose edits by mistake.
48. As an engineer, I want the UI labels to be concise, so that the tool feels like an engineering utility rather than a design app.
49. As an engineer, I want the tool to work on Word desktop where Office Add-ins are supported, so that it fits my normal document workflow.
50. As an engineer, I want the tool to avoid unnecessary special effects, so that the final document looks professional and technical.

## Implementation Decisions

- Build the product as a Word Office Add-in with a task pane surface.
- The first version should prioritize Word desktop compatibility where Office Add-ins and the Word JavaScript APIs are available.
- The add-in should use a client-side image editor in the task pane. No backend is required for the initial scope.
- The editor should model annotations as structured objects rather than directly mutating pixels during editing.
- Supported initial annotation types are text, line, arrow, rectangle, ellipse, and callout.
- The editor should keep the base image separate from annotation state until export.
- Export should flatten the base image and annotations into one raster image before insertion into Word.
- Word insertion should use the Office.js Word API to insert the exported image at the current selection.
- The initial export format should be PNG for sharp text and predictable Word insertion behavior.
- The UI should provide a compact toolbar for tools, a property area for the selected annotation, and a canvas preview.
- Annotation defaults should be optimized for technical documents: high contrast colors, readable text size, moderate stroke width, and minimal decoration.
- The add-in should avoid advanced photo editing features in the first version.
- The add-in should validate uploaded image type and size before loading the image into the editor.
- The add-in should preserve editor state in memory while the task pane remains open.
- Undo and redo should be implemented at the annotation state level.
- Error handling should distinguish upload errors, render/export errors, and Word insertion errors.
- Accessibility should include keyboard operation for common edit actions and visible selected states.
- Localization can start with Traditional Chinese UI copy, with implementation structured so English can be added later if needed.

## Testing Decisions

- Good tests should verify user-visible behavior and exported results, not private implementation details.
- Because the repository currently has no app source or existing test suite, the first implementation should establish tests at the highest practical seams.
- Test the task pane workflow as an end-to-end user path: upload image, add annotations, edit annotations, export, and trigger Word insertion.
- Test the annotation state model with unit tests for create, select, update, move, resize, delete, duplicate, undo, and redo behavior.
- Test the rendering/export seam by verifying that a base image plus representative annotations produces a non-empty PNG with expected dimensions.
- Test the Word insertion seam with a thin adapter boundary so Office.js can be mocked in automated tests.
- Test validation behavior for unsupported file type, oversized file, no image loaded, and insertion failure.
- Test keyboard behavior for delete, undo, and redo from the task pane.
- Visual regression tests should cover a representative annotated engineering image if the frontend stack supports screenshot testing.
- Manual acceptance testing should include inserting the final image into an actual Word document and verifying readability after resizing the inserted image.
- Prior art is not present in this repository, so test structure should be introduced with the first application scaffold.

## Out of Scope

- Full photo editing features such as filters, brightness, contrast, cropping presets, layers panel, masking, blur, or background removal.
- OCR or automatic part detection.
- AI-generated labels or automatic annotation placement.
- Collaborative editing.
- Cloud storage, user accounts, or server-side image processing.
- Editing annotations after the image has already been inserted into Word unless the user reopens or recreates the image in the add-in.
- Complex vector export formats.
- Mobile Word support.
- PowerPoint or Excel add-in support.
- Measurement calibration or CAD-like dimensioning tools.
- Rich diagramming features beyond the basic engineering annotation set.

## Further Notes

- The key product constraint is simplicity. This should feel like a focused engineering documentation tool, not a general-purpose design application.
- The primary user is an engineer preparing documents that explain components, assemblies, screenshots, procedures, or inspection points.
- The minimum useful version is: upload image, add text, add arrow, add rectangle or ellipse, edit/delete annotations, export, and insert into Word.
- The most important quality bar is that inserted images remain readable and visually stable in Word.
- Suggested test seams for implementation are task pane workflow, annotation state model, canvas renderer/exporter, and Word insertion adapter.
- The repository currently has no Git remote, no application source, and no issue tracker configuration, so publishing this PRD to an issue tracker requires the target repository name.
