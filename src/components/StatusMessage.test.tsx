import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, expect, it, vi } from "vitest";
import { StatusMessage } from "./StatusMessage";

describe("StatusMessage", () => {
  it("dismisses status messages after a short delay", () => {
    vi.useFakeTimers();
    const onDismiss = vi.fn();

    render(
      <StatusMessage
        autoDismiss
        error=""
        status="已插入 Word。"
        onDismiss={onDismiss}
      />
    );

    vi.advanceTimersByTime(3000);

    expect(onDismiss).toHaveBeenCalledOnce();
    vi.useRealTimers();
  });

  it("keeps working status messages visible by default", () => {
    vi.useFakeTimers();
    const onDismiss = vi.fn();

    render(
      <StatusMessage
        error=""
        status="正在讀取 Word 選取圖片..."
        onDismiss={onDismiss}
      />
    );

    vi.advanceTimersByTime(3000);

    expect(onDismiss).not.toHaveBeenCalled();
    vi.useRealTimers();
  });

  it("dismisses messages when clicking the close button", async () => {
    const user = userEvent.setup();
    const onDismiss = vi.fn();

    render(<StatusMessage error="無法插入 Word。" status="" onDismiss={onDismiss} />);

    await user.click(screen.getByRole("button", { name: "關閉通知" }));

    expect(onDismiss).toHaveBeenCalledOnce();
  });
});
