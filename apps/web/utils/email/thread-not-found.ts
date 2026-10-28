import { isOutlookItemNotFoundError } from "@/utils/error";
import { OutlookThreadNotFoundError } from "@/utils/outlook/thread-helpers";

/**
 * Whether the provider rejected an operation because the thread is gone, so
 * callers can tell "already handled" apart from a real failure.
 */
export function isThreadNotFoundError(error: unknown): boolean {
  if (isOutlookItemNotFoundError(error)) return true;
  if (error instanceof OutlookThreadNotFoundError) return true;

  return (
    error instanceof Error &&
    error.message.includes("Requested entity was not found")
  );
}
