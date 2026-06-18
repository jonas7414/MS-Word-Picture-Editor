import { render, screen } from "@testing-library/react";
import { describe, expect, it } from "vitest";
import { App } from "./App";

describe("App", () => {
  it("renders the Word image annotation task pane", () => {
    render(<App />);

    expect(screen.getByRole("heading", { name: "圖片標註" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "插入 Word" })).toBeDisabled();
  });
});
