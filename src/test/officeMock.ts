import { vi } from "vitest";

export function installWordMock() {
  const onReady = vi.fn().mockResolvedValue({ host: "Word" });
  const insertedPicture = { height: 0, width: 0 };
  const replacementPicture = { height: 0, width: 0 };
  const insertInlinePictureFromBase64 = vi.fn(() => insertedPicture);
  const replaceTaggedPicture = vi.fn(() => replacementPicture);
  const sync = vi.fn().mockResolvedValue(undefined);
  const selectedContentControl = {
    tag: "",
    title: "",
    delete: vi.fn(),
    insertInlinePictureFromBase64: replaceTaggedPicture,
    load: vi.fn()
  };
  const selectedPicture = {
    isNullObject: false,
    width: 640,
    height: 480,
    getBase64ImageSrc: vi.fn(() => ({ value: "selected-image" })),
    insertContentControl: vi.fn(() => selectedContentControl),
    load: vi.fn()
  };
  const taggedContentControl = {
    isNullObject: false,
    tag: "word-pic-editor-test",
    delete: vi.fn(),
    insertInlinePictureFromBase64: replaceTaggedPicture,
    load: vi.fn()
  };
  const getFirstOrNullObject = vi.fn(() => selectedPicture);
  const getSelection = vi.fn(() => ({
    inlinePictures: { getFirstOrNullObject },
    insertInlinePictureFromBase64
  }));
  const getByTag = vi.fn(() => ({
    getFirstOrNullObject: () => taggedContentControl
  }));
  const run = vi.fn(async (callback: (context: unknown) => Promise<unknown>) => {
    return callback({
      document: {
        getSelection,
        contentControls: {
          getByTag
        }
      },
      sync
    });
  });

  vi.stubGlobal("Office", {
    onReady,
    HostType: {
      Word: "Word"
    }
  });

  vi.stubGlobal("Word", {
    run,
    InsertLocation: {
      replace: "Replace"
    }
  });

  return {
    onReady,
    run,
    sync,
    getSelection,
    getByTag,
    selectedContentControl,
    selectedPicture,
    taggedContentControl,
    insertedPicture,
    replacementPicture,
    insertInlinePictureFromBase64,
    replaceTaggedPicture
  };
}
