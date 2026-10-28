import { describe, expect, it } from "vitest";
import {
  archiveFolderSummary,
  createDraftSummary,
  createRuleSummary,
  deleteFolderSummary,
  deleteRuleSummary,
  emails,
  labelThreadSummary,
  markReadSummary,
  moveFolderToInboxSummary,
  runRulesSummary,
  updateRuleSummary,
  updateThreadSummary,
} from "./summaries";

describe("emails", () => {
  it.each([
    [0, "0 emails"],
    [1, "1 email"],
    [12, "12 emails"],
  ])("counts %i", (n, expected) => {
    expect(emails(n)).toBe(expected);
  });
});

describe("archiveFolderSummary", () => {
  it.each([
    [
      "one",
      { archived: 1, failed: 0, moreLeft: false },
      "Archived 1 email from Marketing; it is now empty.",
    ],
    [
      "many",
      { archived: 12, failed: 0, moreLeft: false },
      "Archived 12 emails from Marketing; it is now empty.",
    ],
    [
      "partial",
      { archived: 10, failed: 2, moreLeft: false },
      "Archived 10 of 12 emails from Marketing; 2 could not be moved and are still there.",
    ],
    [
      "partial, one failed",
      { archived: 10, failed: 1, moreLeft: false },
      "Archived 10 of 11 emails from Marketing; 1 could not be moved and is still there.",
    ],
    [
      "more left",
      { archived: 1000, failed: 0, moreLeft: true },
      "Archived 1000 emails from Marketing; more are left, so call archive_folder again.",
    ],
    [
      "partial and more left",
      { archived: 998, failed: 2, moreLeft: true },
      "Archived 998 of 1000 emails from Marketing; 2 could not be moved and are still there; more are left, so call archive_folder again.",
    ],
    [
      "zero",
      { archived: 0, failed: 0, moreLeft: false },
      "Archived nothing: Marketing was already empty.",
    ],
    [
      "all failed",
      { archived: 0, failed: 3, moreLeft: false },
      "Archived nothing: 3 emails in Marketing could not be moved and are still there.",
    ],
    [
      "one failed, none moved",
      { archived: 0, failed: 1, moreLeft: false },
      "Archived nothing: 1 email in Marketing could not be moved and is still there.",
    ],
  ])("%s", (_case, counts, expected) => {
    expect(archiveFolderSummary({ path: "Marketing", ...counts })).toBe(
      expected,
    );
  });
});

describe("moveFolderToInboxSummary", () => {
  it.each([
    [
      "one",
      { moved: 1, failed: 0, moreLeft: false },
      "Moved 1 email from Marketing to the Inbox; it is now empty.",
    ],
    [
      "many",
      { moved: 12, failed: 0, moreLeft: false },
      "Moved 12 emails from Marketing to the Inbox; it is now empty.",
    ],
    [
      "partial",
      { moved: 10, failed: 2, moreLeft: false },
      "Moved 10 of 12 emails from Marketing to the Inbox; 2 could not be moved and are still there.",
    ],
    [
      "more left",
      { moved: 1000, failed: 0, moreLeft: true },
      "Moved 1000 emails from Marketing to the Inbox; more are left, so call move_folder_to_inbox again.",
    ],
    [
      "zero",
      { moved: 0, failed: 0, moreLeft: false },
      "Moved nothing: Marketing was already empty.",
    ],
    [
      "all failed",
      { moved: 0, failed: 3, moreLeft: false },
      "Moved nothing: 3 emails in Marketing could not be moved and are still there.",
    ],
  ])("%s", (_case, counts, expected) => {
    expect(moveFolderToInboxSummary({ path: "Marketing", ...counts })).toBe(
      expected,
    );
  });
});

describe("markReadSummary", () => {
  const folder = { path: "Newsletters" };
  const sender = { sender: "news@shop.example" };
  const both = { ...folder, ...sender };

  it.each([
    [
      "one",
      { ...folder, read: true, marked: 1, failed: 0, moreLeft: false },
      "Marked 1 email in Newsletters read.",
    ],
    [
      "many",
      { ...folder, read: true, marked: 5, failed: 0, moreLeft: false },
      "Marked 5 emails in Newsletters read.",
    ],
    [
      "sender",
      { ...sender, read: true, marked: 5, failed: 0, moreLeft: false },
      "Marked 5 emails from news@shop.example read.",
    ],
    [
      "sender in a folder",
      { ...both, read: true, marked: 5, failed: 0, moreLeft: false },
      "Marked 5 emails from news@shop.example in Newsletters read.",
    ],
    [
      "read: false",
      { ...folder, read: false, marked: 5, failed: 0, moreLeft: false },
      "Marked 5 emails in Newsletters unread.",
    ],
    [
      "partial",
      { ...folder, read: true, marked: 3, failed: 2, moreLeft: false },
      "Marked 3 of 5 emails in Newsletters read; 2 could not be changed and are still unread.",
    ],
    [
      "partial, one failed",
      { ...folder, read: true, marked: 3, failed: 1, moreLeft: false },
      "Marked 3 of 4 emails in Newsletters read; 1 could not be changed and is still unread.",
    ],
    [
      "partial, read: false",
      { ...sender, read: false, marked: 3, failed: 2, moreLeft: false },
      "Marked 3 of 5 emails from news@shop.example unread; 2 could not be changed and are still read.",
    ],
    [
      "more left",
      { ...folder, read: true, marked: 1000, failed: 0, moreLeft: true },
      "Marked 1000 emails in Newsletters read; more are left, so call mark_read again.",
    ],
    [
      "zero",
      { ...folder, read: true, marked: 0, failed: 0, moreLeft: false },
      "Marked nothing: no unread emails in Newsletters.",
    ],
    [
      "zero, read: false",
      { ...folder, read: false, marked: 0, failed: 0, moreLeft: false },
      "Marked nothing: no read emails in Newsletters.",
    ],
    [
      "zero, sender",
      { ...sender, read: true, marked: 0, failed: 0, moreLeft: false },
      "Marked nothing: no unread emails from news@shop.example.",
    ],
    [
      "zero, sender in a folder",
      { ...both, read: true, marked: 0, failed: 0, moreLeft: false },
      "Marked nothing: no unread emails from news@shop.example in Newsletters.",
    ],
    [
      "all failed",
      { ...folder, read: true, marked: 0, failed: 3, moreLeft: false },
      "Marked nothing: 3 emails in Newsletters could not be changed and are still unread.",
    ],
  ])("%s", (_case, input, expected) => {
    expect(markReadSummary(input)).toBe(expected);
  });
});

describe("every number in a count sentence is a count it was given", () => {
  const grid = [0, 1, 2, 7, 12, 1000];
  const cases = grid.flatMap((done) =>
    grid.flatMap((failed) =>
      [false, true].map((moreLeft) => ({ done, failed, moreLeft })),
    ),
  );

  function numbersIn(sentence: string) {
    return (sentence.match(/\d+/g) ?? []).map(Number);
  }

  function expectHonestCounts(
    sentence: string,
    { done, failed }: { done: number; failed: number },
  ) {
    const allowed = new Set([done, failed, done + failed]);
    for (const n of numbersIn(sentence)) expect(allowed).toContain(n);
    if (done > 0) expect(numbersIn(sentence)).toContain(done);
    if (failed > 0) expect(numbersIn(sentence)).toContain(failed);
    if (done === 0) expect(sentence).toMatch(/ nothing/);
  }

  it.each(cases)("archive_folder %o", ({ done, failed, moreLeft }) => {
    expectHonestCounts(
      archiveFolderSummary({
        path: "Marketing",
        archived: done,
        failed,
        moreLeft,
      }),
      { done, failed },
    );
  });

  it.each(cases)("move_folder_to_inbox %o", ({ done, failed, moreLeft }) => {
    expectHonestCounts(
      moveFolderToInboxSummary({
        path: "Marketing",
        moved: done,
        failed,
        moreLeft,
      }),
      { done, failed },
    );
  });

  it.each(cases)("mark_read %o", ({ done, failed, moreLeft }) => {
    for (const read of [true, false])
      expectHonestCounts(
        markReadSummary({
          path: "Newsletters",
          sender: "news@shop.example",
          read,
          marked: done,
          failed,
          moreLeft,
        }),
        { done, failed },
      );
  });

  it("claims the folder is empty only when nothing failed and nothing is left", () => {
    for (const { done, failed, moreLeft } of cases) {
      const sentence = archiveFolderSummary({
        path: "Marketing",
        archived: done,
        failed,
        moreLeft,
      });
      const claimsEmpty = /empty/.test(sentence);
      expect(claimsEmpty).toBe(failed === 0 && !moreLeft);
    }
  });
});

describe("deleteFolderSummary", () => {
  it("names the folder it deleted", () => {
    expect(deleteFolderSummary("Old Projects")).toBe(
      "Deleted the empty folder Old Projects.",
    );
  });
});

describe("updateThreadSummary", () => {
  it.each([
    [
      "archive",
      "Deleted Items",
      "Archived 1 conversation; unarchive moves it back.",
    ],
    ["unarchive", "Deleted Items", "Moved 1 conversation back to the Inbox."],
    [
      "trash",
      "Deleted Items",
      "Moved 1 conversation to Deleted Items; it can be restored.",
    ],
    ["trash", "Trash", "Moved 1 conversation to Trash; it can be restored."],
    ["mark_read", "Trash", "Marked 1 conversation read."],
    ["mark_unread", "Trash", "Marked 1 conversation unread."],
  ] as const)("%s (%s)", (action, trashFolder, expected) => {
    expect(updateThreadSummary({ action, trashFolder })).toBe(expected);
  });
});

describe("labelThreadSummary", () => {
  it.each([
    [
      "one message",
      { labelName: "FYI", removed: false, messages: 1 },
      'Added the label "FYI" to 1 message in 1 conversation.',
    ],
    [
      "many messages",
      { labelName: "FYI", removed: false, messages: 2 },
      'Added the label "FYI" to 2 messages in 1 conversation.',
    ],
    [
      "no messages",
      { labelName: "FYI", removed: false, messages: 0 },
      'Added nothing: the conversation has no messages, so the label "FYI" was not added.',
    ],
    [
      "removed",
      { labelName: "FYI", removed: true },
      'Removed the label "FYI" from 1 conversation.',
    ],
  ])("%s", (_case, input, expected) => {
    expect(labelThreadSummary(input)).toBe(expected);
  });
});

describe("createDraftSummary", () => {
  it("says a new draft was not sent", () => {
    expect(createDraftSummary({ to: "a@example.com", reply: false })).toBe(
      "Saved a draft to a@example.com; it was not sent.",
    );
  });

  it("says a reply draft sits in the same conversation and was not sent", () => {
    expect(createDraftSummary({ to: "a@example.com", reply: true })).toBe(
      "Saved a reply draft to a@example.com in the same conversation; it was not sent.",
    );
  });
});

describe("runRulesSummary", () => {
  const receipts = {
    ruleName: "Receipts",
    alreadyProcessed: false,
    actions: [
      { type: "LABEL", label: "Receipts", folderName: null },
      { type: "ARCHIVE", label: null, folderName: null },
    ],
  };
  const noMatch = {
    ruleName: null,
    alreadyProcessed: false,
    status: "SKIPPED",
    actions: [],
  };

  it.each([
    [
      "preview, match",
      { preview: true, results: [{ ...receipts, status: "SKIPPED" }] },
      'Preview only, nothing changed: the rule "Receipts" would match (label "Receipts", archive).',
    ],
    [
      "preview, two matches",
      {
        preview: true,
        results: [
          { ...receipts, status: "APPLIED" },
          {
            ruleName: "Projects",
            alreadyProcessed: false,
            status: "APPLIED",
            actions: [{ type: "MOVE_FOLDER", label: null, folderName: "Work" }],
          },
        ],
      },
      'Preview only, nothing changed: the rule "Receipts" would match (label "Receipts", archive); the rule "Projects" would match (move to the folder Work).',
    ],
    [
      "preview, no match",
      { preview: true, results: [noMatch] },
      "Preview only, nothing changed: no rule matches this email.",
    ],
    [
      "applied, no match",
      { preview: false, results: [noMatch] },
      "No rule matched this email; nothing changed.",
    ],
    [
      "applied, ran",
      { preview: false, results: [{ ...receipts, status: "APPLIED" }] },
      'Ran the rule "Receipts" (label "Receipts", archive).',
    ],
    [
      "applied, scheduled",
      { preview: false, results: [{ ...receipts, status: "APPLYING" }] },
      'Started the rule "Receipts" (label "Receipts", archive); some of its actions are scheduled for later.',
    ],
    [
      "applied, failed",
      { preview: false, results: [{ ...receipts, status: "ERROR" }] },
      'The rule "Receipts" matched but failed while running (label "Receipts", archive); some of it may have gone through.',
    ],
    [
      "applied, matched but skipped",
      { preview: false, results: [{ ...receipts, status: "SKIPPED" }] },
      'The rule "Receipts" matched but was skipped; nothing changed.',
    ],
    [
      "already handled",
      {
        preview: false,
        results: [{ ...receipts, status: "APPLIED", alreadyProcessed: true }],
      },
      'Already handled earlier by the rule "Receipts" (label "Receipts", archive); nothing ran again.',
    ],
    [
      "already handled, during a preview",
      {
        preview: true,
        results: [{ ...receipts, status: "APPLIED", alreadyProcessed: true }],
      },
      'Already handled earlier by the rule "Receipts" (label "Receipts", archive); nothing ran again.',
    ],
    [
      "already handled with no match",
      { preview: false, results: [{ ...noMatch, alreadyProcessed: true }] },
      "Already handled earlier: no rule matched this email; nothing ran again.",
    ],
    [
      "a rule with no actions",
      {
        preview: false,
        results: [{ ...receipts, actions: [], status: "APPLIED" }],
      },
      'Ran the rule "Receipts" (no actions).',
    ],
  ])("%s", (_case, input, expected) => {
    expect(runRulesSummary(input)).toBe(expected);
  });
});

describe("rule summaries", () => {
  it("created", () => {
    expect(createRuleSummary("Digest")).toBe('Created the rule "Digest".');
  });

  it("replaced", () => {
    expect(updateRuleSummary("Digest")).toBe(
      'Replaced the rule "Digest" with the conditions and actions given.',
    );
  });

  it("deleted", () => {
    expect(deleteRuleSummary("Digest")).toBe('Deleted the rule "Digest".');
  });
});
