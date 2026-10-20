const MAX_LENGTH = 200;

/**
 * The text the chat stream sends to the browser when a step fails. The AI SDK
 * defaults to "An error occurred", which hides the cause; this keeps the first
 * line of the real message and strips anything that looks like a credential.
 */
export function getStreamErrorText(error: unknown): string {
  if (!(error instanceof Error)) return "An error occurred";

  const firstLine = error.message.split("\n")[0]?.trim() ?? "";
  const message = firstLine
    .replace(/\/\/[^/\s:]+:[^@\s]+@/g, "//[redacted]@")
    .replace(/bearer\s+\S+/gi, "bearer [redacted]")
    .replace(/(key|secret|auth)=\S+/gi, "$1=[redacted]")
    .slice(0, MAX_LENGTH);

  return message || "An error occurred";
}
