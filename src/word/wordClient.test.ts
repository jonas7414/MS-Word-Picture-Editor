import { afterEach, describe, expect, it, vi } from "vitest";
import { installWordMock } from "../test/officeMock";
import { insertImageAtSelection, stripDataUrlPrefix } from "./wordClient";

describe("wordClient", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

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
});
