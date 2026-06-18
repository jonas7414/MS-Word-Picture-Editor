import { afterEach, describe, expect, it, vi } from "vitest";
import { installWordMock } from "../test/officeMock";
import {
  cleanupEditableWordImage,
  insertImageAtSelection,
  loadSelectedImageForEditing,
  replaceEditableWordImage,
  stripDataUrlPrefix
} from "./wordClient";

describe("wordClient", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  function installImageDimensionMock(width = 1280, height = 960) {
    vi.stubGlobal(
      "Image",
      class {
        naturalHeight = height;
        naturalWidth = width;
        onerror: (() => void) | null = null;
        onload: (() => void) | null = null;

        set src(_src: string) {
          this.onload?.();
        }
      }
    );
  }

  it("strips the PNG data URL prefix for Office.js insertion", () => {
    expect(stripDataUrlPrefix("data:image/png;base64,abc123")).toBe("abc123");
  });

  it("inserts the image at the current Word selection", async () => {
    const mock = installWordMock();

    await insertImageAtSelection("data:image/png;base64,abc123");

    expect(mock.onReady).toHaveBeenCalled();
    expect(mock.insertInlinePictureFromBase64).toHaveBeenCalledWith(
      "abc123",
      "Replace"
    );
    expect(mock.sync).toHaveBeenCalled();
  });

  it("loads the selected inline picture for editing", async () => {
    const mock = installWordMock();
    installImageDimensionMock(1280, 960);

    const editable = await loadSelectedImageForEditing();

    expect(mock.selectedPicture.load).toHaveBeenCalledWith("width,height");
    expect(mock.selectedPicture.getBase64ImageSrc).toHaveBeenCalled();
    expect(mock.selectedPicture.insertContentControl).toHaveBeenCalled();
    expect(mock.selectedContentControl.tag).toMatch(/^word-pic-editor-/);
    expect(mock.selectedContentControl.title).toBe("");
    expect(editable.image).toMatchObject({
      src: "data:image/png;base64,selected-image",
      name: "Word 選取圖片",
      width: 1280,
      height: 960
    });
    expect(editable.targetTag).toBe(mock.selectedContentControl.tag);
    expect(editable.targetDisplaySize).toEqual({
      width: 640,
      height: 480
    });
  });

  it("replaces the tagged editable Word image without changing its display size", async () => {
    const mock = installWordMock();

    await replaceEditableWordImage({
      dataUrl: "data:image/png;base64,edited",
      targetDisplaySize: {
        width: 320,
        height: 240
      },
      targetTag: "word-pic-editor-test"
    });

    expect(mock.getByTag).toHaveBeenCalledWith("word-pic-editor-test");
    expect(mock.replaceTaggedPicture).toHaveBeenCalledWith("edited", "Replace");
    expect(mock.replacementPicture).toMatchObject({
      width: 320,
      height: 240
    });
    expect(mock.taggedContentControl.delete).toHaveBeenCalledWith(true);
    expect(mock.sync).toHaveBeenCalled();
  });

  it("cleans up a tagged editable Word image without replacing it", async () => {
    const mock = installWordMock();

    await cleanupEditableWordImage("word-pic-editor-test");

    expect(mock.getByTag).toHaveBeenCalledWith("word-pic-editor-test");
    expect(mock.taggedContentControl.delete).toHaveBeenCalledWith(true);
    expect(mock.replaceTaggedPicture).not.toHaveBeenCalled();
  });
});
