import { fireEvent, render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, expect, it, vi } from "vitest";
import { App } from "./App";

vi.mock("./image/imageValidation", () => ({
  loadImageFile: vi.fn(async (file: File) => ({
    src: "data:image/png;base64,input",
    name: file.name,
    width: 640,
    height: 480
  }))
}));

describe("keyboard editing", () => {
  it("deletes selected annotations and supports undo", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.upload(
      screen.getByLabelText("上傳圖片"),
      new File(["image"], "assembly.png", { type: "image/png" })
    );
    await user.click(screen.getByRole("button", { name: "文字" }));
    fireEvent.pointerDown(screen.getByLabelText("圖片標註畫布"), {
      clientX: 80,
      clientY: 100
    });

    expect(screen.getByText("零件名稱")).toBeInTheDocument();

    await user.keyboard("{Delete}");
    expect(screen.queryByText("零件名稱")).not.toBeInTheDocument();

    await user.keyboard("{Meta>}z{/Meta}");
    expect(screen.getByText("零件名稱")).toBeInTheDocument();
  });
});
