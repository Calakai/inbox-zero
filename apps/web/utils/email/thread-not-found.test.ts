import { describe, expect, it } from "vitest";
import { OutlookThreadNotFoundError } from "@/utils/outlook/thread-helpers";
import { isThreadNotFoundError } from "./thread-not-found";

describe("isThreadNotFoundError", () => {
  it("recognises an Outlook thread with no messages left as not found", () => {
    expect(isThreadNotFoundError(new OutlookThreadNotFoundError())).toBe(true);
  });

  it("recognises a provider's not-found message", () => {
    expect(
      isThreadNotFoundError(new Error("Requested entity was not found.")),
    ).toBe(true);
  });

  it("treats any other error as a real failure", () => {
    expect(isThreadNotFoundError(new Error("Mailbox is read-only"))).toBe(
      false,
    );
  });
});
