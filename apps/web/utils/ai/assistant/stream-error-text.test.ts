import { describe, expect, it } from "vitest";
import { getStreamErrorText } from "./stream-error-text";

describe("getStreamErrorText", () => {
  it("keeps the first line of the real error", () => {
    expect(getStreamErrorText(new Error("Search failed: 429\nstack"))).toBe(
      "Search failed: 429",
    );
  });

  it("redacts credentials", () => {
    expect(
      getStreamErrorText(new Error("bad call https://u:pw@host key=abc123")),
    ).toBe("bad call https://[redacted]@host key=[redacted]");
  });

  it("falls back for non-errors and empty messages", () => {
    expect(getStreamErrorText("boom")).toBe("An error occurred");
    expect(getStreamErrorText(new Error("\n"))).toBe("An error occurred");
  });
});
