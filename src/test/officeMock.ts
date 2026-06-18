import { vi } from "vitest";

export function installWordMock() {
  const onReady = vi.fn().mockResolvedValue({ host: "Word" });
  const insertInlinePictureFromBase64 = vi.fn();
  const sync = vi.fn().mockResolvedValue(undefined);
  const getSelection = vi.fn(() => ({ insertInlinePictureFromBase64 }));
  const run = vi.fn(async (callback: (context: unknown) => Promise<void>) => {
    await callback({
      document: { getSelection },
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
    insertInlinePictureFromBase64
  };
}
