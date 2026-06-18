export function StatusMessage({
  status,
  error
}: {
  status: string;
  error: string;
}) {
  if (error) return <p className="status error">{error}</p>;
  if (status) return <p className="status">{status}</p>;
  return null;
}
