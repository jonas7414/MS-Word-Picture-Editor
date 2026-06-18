import { readImageDimensions } from "../image/imageValidation";

type WordContext = {
  document: {
    getSelection: () => {
      inlinePictures: {
        getFirstOrNullObject: () => WordInlinePicture;
      };
      insertInlinePictureFromBase64: (
        base64: string,
        location: "Replace"
      ) => WordInlinePicture;
    };
    contentControls: {
      getByTag: (tag: string) => {
        getFirstOrNullObject: () => WordContentControl;
      };
    };
  };
  sync: () => Promise<void>;
};

type WordClientResult<T> = {
  value: T;
};

type WordInlinePicture = {
  isNullObject?: boolean;
  width: number;
  height: number;
  getBase64ImageSrc: () => WordClientResult<string>;
  insertContentControl: () => WordContentControl;
  load: (properties: string) => void;
};

type WordContentControl = {
  isNullObject?: boolean;
  tag: string;
  title: string;
  delete: (keepContent: boolean) => void;
  insertInlinePictureFromBase64: (
    base64: string,
    location: "Replace"
  ) => WordInlinePicture;
  load: (properties: string) => void;
};

type WordGlobal = {
  run: <T>(callback: (context: WordContext) => Promise<T>) => Promise<T>;
  InsertLocation?: {
    replace?: "Replace";
  };
};

type OfficeGlobal = {
  onReady?: () => Promise<{ host?: string }>;
  HostType?: {
    Word?: string;
  };
};

declare global {
  interface Window {
    Office?: OfficeGlobal;
    Word?: WordGlobal;
  }
}

export function stripDataUrlPrefix(dataUrl: string): string {
  return dataUrl.replace(/^data:image\/png;base64,/, "");
}

export type EditableWordImage = {
  image: {
    src: string;
    name: string;
    width: number;
    height: number;
  };
  targetTag: string;
  targetDisplaySize: {
    width: number;
    height: number;
  };
};

let officeReadyPromise: Promise<void> | null = null;
const officeReadyTimeoutMs = 8000;
const wordOperationTimeoutMs = 12000;

function withTimeout<T>(
  promise: Promise<T>,
  timeoutMs: number,
  timeoutMessage: string
): Promise<T> {
  return new Promise((resolve, reject) => {
    const timeoutId = globalThis.setTimeout(() => {
      reject(new Error(timeoutMessage));
    }, timeoutMs);

    promise.then(
      (value) => {
        globalThis.clearTimeout(timeoutId);
        resolve(value);
      },
      (error: unknown) => {
        globalThis.clearTimeout(timeoutId);
        reject(error);
      }
    );
  });
}

export async function waitForOfficeReady(): Promise<void> {
  if (officeReadyPromise) {
    return officeReadyPromise;
  }

  officeReadyPromise = (async () => {
    const office = globalThis.window?.Office;
    if (!office?.onReady) {
      return;
    }

    const word = globalThis.window?.Word;
    if (word?.run) {
      office
        .onReady()
        .then((info) => {
          const wordHost = office.HostType?.Word;
          if (wordHost && info.host && info.host !== wordHost) {
            throw new Error("請在 Microsoft Word 中使用此工具。");
          }
        })
        .catch(() => undefined);
      return;
    }

    const info = await withTimeout(
      office.onReady(),
      officeReadyTimeoutMs,
      "Office 初始化逾時，請確認此工具是在 Microsoft Word 增益集中開啟，或重新開啟右側工具面板。"
    );
    const wordHost = office.HostType?.Word;
    if (wordHost && info.host && info.host !== wordHost) {
      throw new Error("請在 Microsoft Word 中使用此工具。");
    }
  })();

  try {
    return await officeReadyPromise;
  } catch (error) {
    officeReadyPromise = null;
    throw error;
  }
}

export async function insertImageAtSelection(dataUrl: string): Promise<void> {
  await waitForOfficeReady();

  const word = globalThis.window?.Word;
  if (!word?.run) {
    throw new Error("目前環境無法使用 Word 插入功能。");
  }

  await withTimeout(
    word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertInlinePictureFromBase64(
        stripDataUrlPrefix(dataUrl),
        word.InsertLocation?.replace ?? "Replace"
      );
      await context.sync();
    }),
    wordOperationTimeoutMs,
    "Word 插入圖片逾時，請稍後再試。"
  );
}

export async function loadSelectedImageForEditing(): Promise<EditableWordImage> {
  await waitForOfficeReady();

  const word = globalThis.window?.Word;
  if (!word?.run) {
    throw new Error("目前環境無法讀取 Word 圖片。");
  }

  return withTimeout(
    (async () => {
      const editable = await word.run(async (context) => {
        const selection = context.document.getSelection();
        const picture = selection.inlinePictures.getFirstOrNullObject();
        picture.load("width,height");
        await context.sync();

        if (picture.isNullObject) {
          throw new Error("請先在 Word 中選取一張圖片。");
        }

        const contentControl = picture.insertContentControl();
        const targetTag = `word-pic-editor-${Date.now()}`;
        contentControl.tag = targetTag;
        const base64 = picture.getBase64ImageSrc();
        await context.sync();

        const src = `data:image/png;base64,${base64.value}`;
        return {
          image: {
            src,
            name: "Word 選取圖片",
            width: 0,
            height: 0
          },
          targetTag,
          targetDisplaySize: {
            width: picture.width,
            height: picture.height
          }
        };
      });
      const dimensions = await readImageDimensions(editable.image.src);

      return {
        ...editable,
        image: {
          ...editable.image,
          width: dimensions.width,
          height: dimensions.height
        }
      };
    })(),
    wordOperationTimeoutMs,
    "讀取 Word 選取圖片逾時。請確認選到的是嵌入式圖片，若圖片是浮動版面物件，請先將文繞圖改成「與文字排列」。"
  );
}

export async function replaceEditableWordImage({
  dataUrl,
  targetDisplaySize,
  targetTag
}: {
  dataUrl: string;
  targetDisplaySize?: {
    width: number;
    height: number;
  };
  targetTag: string;
}): Promise<void> {
  await waitForOfficeReady();

  const word = globalThis.window?.Word;
  if (!word?.run) {
    throw new Error("目前環境無法取代 Word 圖片。");
  }

  await withTimeout(
    word.run(async (context) => {
      const target = context.document.contentControls
        .getByTag(targetTag)
        .getFirstOrNullObject();
      target.load("tag");
      await context.sync();

      if (target.isNullObject) {
        throw new Error("找不到原本要取代的圖片，請重新選取圖片後再編輯。");
      }

      const replacement = target.insertInlinePictureFromBase64(
        stripDataUrlPrefix(dataUrl),
        word.InsertLocation?.replace ?? "Replace"
      );
      if (targetDisplaySize) {
        replacement.width = targetDisplaySize.width;
        replacement.height = targetDisplaySize.height;
      }
      target.delete(true);
      await context.sync();
    }),
    wordOperationTimeoutMs,
    "取代 Word 圖片逾時，請重新選取圖片後再試。"
  );
}

export async function cleanupEditableWordImage(targetTag: string): Promise<void> {
  await waitForOfficeReady();

  const word = globalThis.window?.Word;
  if (!word?.run) {
    return;
  }

  await withTimeout(
    word.run(async (context) => {
      const target = context.document.contentControls
        .getByTag(targetTag)
        .getFirstOrNullObject();
      target.load("tag");
      await context.sync();

      if (!target.isNullObject) {
        target.delete(true);
        await context.sync();
      }
    }),
    wordOperationTimeoutMs,
    "清理 Word 編輯目標逾時，請稍後再試。"
  );
}
