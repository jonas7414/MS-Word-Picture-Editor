import { X } from "lucide-react";
import { useEffect } from "react";

export function StatusMessage({
  status,
  error,
  autoDismiss,
  onDismiss
}: {
  status: string;
  error: string;
  autoDismiss?: boolean;
  onDismiss: () => void;
}) {
  const message = error || status;

  useEffect(() => {
    if (!status || error || !autoDismiss) return undefined;

    const timeoutId = window.setTimeout(onDismiss, 3000);
    return () => window.clearTimeout(timeoutId);
  }, [autoDismiss, error, onDismiss, status]);

  if (!message) return null;

  return (
    <div className={`status${error ? " error" : ""}`} role="status">
      <span>{message}</span>
      <button aria-label="關閉通知" type="button" onClick={onDismiss}>
        <X size={14} />
      </button>
    </div>
  );
}
