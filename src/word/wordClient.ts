type WordContext = {
  document: {
    getSelection: () => {
      insertInlinePictureFromBase64: (
        base64: string,
        location: "Replace"
      ) => void;
    };
  };
  sync: () => Promise<void>;
};

type WordGlobal = {
  run: (callback: (context: WordContext) => Promise<void>) => Promise<void>;
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

let officeReadyPromise: Promise<void> | null = null;

export async function waitForOfficeReady(): Promise<void> {
  if (officeReadyPromise) {
    return officeReadyPromise;
  }

  officeReadyPromise = (async () => {
    const office = globalThis.window?.Office;
    if (!office?.onReady) {
      return;
    }

    const info = await office.onReady();
    const wordHost = office.HostType?.Word;
    if (wordHost && info.host && info.host !== wordHost) {
      throw new Error("請在 Microsoft Word 中使用此工具。");
    }
  })();

  return officeReadyPromise;
}

export async function insertImageAtSelection(dataUrl: string): Promise<void> {
  await waitForOfficeReady();

  const word = globalThis.window?.Word;
  if (!word?.run) {
    throw new Error("目前環境無法使用 Word 插入功能。");
  }

  await word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertInlinePictureFromBase64(
      stripDataUrlPrefix(dataUrl),
      word.InsertLocation?.replace ?? "Replace"
    );
    await context.sync();
  });
}
