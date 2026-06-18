import { describe, expect, it } from "vitest";
import { maxImageBytes, validateImageFile } from "./imageValidation";

function makeFile(name: string, type: string, size: number) {
  return new File([new Uint8Array(size)], name, { type });
}

describe("validateImageFile", () => {
  it("accepts png, jpeg, and webp images", () => {
    expect(validateImageFile(makeFile("part.png", "image/png", 1000))).toEqual({
      ok: true
    });
    expect(validateImageFile(makeFile("part.jpg", "image/jpeg", 1000))).toEqual({
      ok: true
    });
    expect(validateImageFile(makeFile("part.webp", "image/webp", 1000))).toEqual({
      ok: true
    });
  });

  it("rejects unsupported file types with Traditional Chinese copy", () => {
    expect(validateImageFile(makeFile("part.gif", "image/gif", 1000))).toEqual({
      ok: false,
      message: "僅支援 PNG、JPG、JPEG、WEBP 圖片。"
    });
  });

  it("rejects files larger than the configured limit", () => {
    expect(
      validateImageFile(makeFile("huge.png", "image/png", maxImageBytes + 1))
    ).toEqual({
      ok: false,
      message: "圖片不可超過 10 MB。"
    });
  });
});
