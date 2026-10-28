import { createHash } from "node:crypto";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { createEmailProvider } from "@/utils/email/provider";
import { resolveMcpEmailAccount } from "@/utils/mcp/account-selection";
import { createTestLogger } from "@/__tests__/helpers";
import type { OutlookFolder } from "@/utils/outlook/folders";
import type { ParsedMessage } from "@/utils/types";
import { runRulesOnMessage } from "@/utils/ai/choose-rule/run-rules-on-message";
import { checkRateLimit, createRateLimitKey } from "@/utils/rate-limit";
import { redis } from "@/utils/redis";
import { SafeError } from "@/utils/error";
import { OutlookThreadNotFoundError } from "@/utils/outlook/thread-helpers";
import {
  archiveFolderForMcp,
  createDraftForMcp,
  deleteFolderForMcp,
  labelThreadForMcp,
  listFoldersForMcp,
  markReadForMcp,
  moveFolderToInboxForMcp,
  readThreadForMcp,
  runRulesForMcp,
  searchInboxForMcp,
  updateThreadForMcp,
} from "./email-tools";

vi.mock("server-only", () => ({}));
vi.mock("@/utils/email/provider");
vi.mock("@/utils/ai/choose-rule/run-rules-on-message", () => ({
  runRulesOnMessage: vi.fn(),
}));
vi.mock("@/utils/rate-limit", async (importOriginal) => ({
  ...(await importOriginal<typeof import("@/utils/rate-limit")>()),
  checkRateLimit: vi.fn(),
}));
vi.mock("@/utils/redis", () => ({
  redis: { get: vi.fn(), set: vi.fn() },
}));
vi.mock("@/utils/mcp/account-selection", () => ({
  resolveMcpEmailAccount: vi.fn(),
}));

const logger = createTestLogger();
const emailAccount = {
  id: "account_1",
  email: "owner@example.com",
  name: null,
  provider: "google",
};

function message(overrides: Partial<ParsedMessage> = {}): ParsedMessage {
  return {
    id: "message-1",
    threadId: "thread-1",
    snippet: "Can we meet Thursday?",
    historyId: "",
    inline: [],
    headers: {
      from: "contact@example.com",
      to: "owner@example.com",
      subject: "Question",
      date: "2026-02-18T00:00:00.000Z",
    },
    subject: "Question",
    date: "2026-02-18T00:00:00.000Z",
    textPlain: "Can we meet Thursday about the contract?",
    ...overrides,
  };
}

describe("MCP email tools", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(resolveMcpEmailAccount).mockResolvedValue(emailAccount);
    vi.mocked(checkRateLimit).mockResolvedValue({
      limited: false,
      limit: 2,
      remaining: 1,
    });
  });

  it("returns search metadata without message bodies", async () => {
    const searchMessages = vi.fn().mockResolvedValue({
      messages: [message({ textPlain: "SECRET BODY" })],
      nextPageToken: "page-2",
    });
    vi.mocked(createEmailProvider).mockResolvedValue({
      searchMessages,
    } as never);

    const result = await searchInboxForMcp({
      userId: "user_1",
      query: "from:contact@example.com",
      logger,
    });

    expect(searchMessages).toHaveBeenCalledWith({
      query: "from:contact@example.com",
      maxResults: 10,
      pageToken: undefined,
    });
    expect(result.hasMore).toBe(true);
    expect(result.messages[0]).toMatchObject({
      messageId: "message-1",
      threadId: "thread-1",
      snippet: "Can we meet Thursday?",
      from: "contact@example.com",
    });
    expect(JSON.stringify(result)).not.toContain("SECRET BODY");
  });

  it("adds amounts and the due date from the body without returning the body", async () => {
    vi.mocked(createEmailProvider).mockResolvedValue({
      searchMessages: vi.fn().mockResolvedValue({
        messages: [
          message({
            subject: "Your October statement",
            snippet: "Your statement is ready.",
            textPlain:
              "SECRET BODY. Statement balance: $1,284.17. Due date: November 18.",
          }),
        ],
      }),
    } as never);

    const result = await searchInboxForMcp({
      userId: "user_1",
      query: "from:bank@example.com",
      logger,
    });

    expect(result.messages[0]).toMatchObject({
      amounts: ["$1,284.17"],
      dueDate: "November 18",
    });
    expect(JSON.stringify(result)).not.toContain("SECRET BODY");
  });

  it("leaves amounts and dueDate out when the email states none", async () => {
    vi.mocked(createEmailProvider).mockResolvedValue({
      searchMessages: vi.fn().mockResolvedValue({ messages: [message()] }),
    } as never);

    const result = await searchInboxForMcp({
      userId: "user_1",
      query: "from:contact@example.com",
      logger,
    });

    expect(result.messages[0]).not.toHaveProperty("amounts");
    expect(result.messages[0]).not.toHaveProperty("dueDate");
  });

  it("finds an amount inside an HTML table", async () => {
    vi.mocked(createEmailProvider).mockResolvedValue({
      searchMessages: vi.fn().mockResolvedValue({
        messages: [
          message({
            snippet: "Your receipt",
            textPlain: undefined,
            textHtml: "<table><tr><td>Total</td><td>$84.12</td></tr></table>",
          }),
        ],
      }),
    } as never);

    const result = await searchInboxForMcp({
      userId: "user_1",
      query: "receipt",
      logger,
    });

    expect(result.messages[0]).toMatchObject({ amounts: ["$84.12"] });
  });

  it("lists one folder's mail when given a folderId and no query", async () => {
    const searchMessages = vi.fn().mockResolvedValue({ messages: [message()] });
    vi.mocked(resolveMcpEmailAccount).mockResolvedValue({
      ...emailAccount,
      provider: "microsoft",
    });
    vi.mocked(createEmailProvider).mockResolvedValue({
      searchMessages,
      getFolders: vi
        .fn()
        .mockResolvedValue([
          folder({ id: "marketing", displayName: "Marketing" }),
        ]),
    } as never);

    const result = await searchInboxForMcp({
      userId: "user_1",
      folderId: "marketing",
      logger,
    });

    expect(searchMessages).toHaveBeenCalledWith(
      expect.objectContaining({ query: "", folderId: "marketing" }),
    );
    expect(result).toMatchObject({
      folderId: "marketing",
      messages: [{ messageId: "message-1" }],
    });
  });

  it("searches a folder given by name using the folder's real ID", async () => {
    const searchMessages = vi.fn().mockResolvedValue({ messages: [message()] });
    vi.mocked(resolveMcpEmailAccount).mockResolvedValue({
      ...emailAccount,
      provider: "microsoft",
    });
    vi.mocked(createEmailProvider).mockResolvedValue({
      searchMessages,
      getFolders: vi.fn().mockResolvedValue([
        folder({
          id: "AQMkADAwATMwMAItYzE4-long-real-id",
          displayName: "Marketing",
        }),
      ]),
    } as never);

    const result = await searchInboxForMcp({
      userId: "user_1",
      query: "sale",
      folderId: "marketing",
      logger,
    });

    expect(searchMessages).toHaveBeenCalledWith(
      expect.objectContaining({
        query: "sale",
        folderId: "AQMkADAwATMwMAItYzE4-long-real-id",
      }),
    );
    expect(result.folderId).toBe("AQMkADAwATMwMAItYzE4-long-real-id");
  });

  it("does not search when the folder name matches nothing", async () => {
    const searchMessages = vi.fn();
    vi.mocked(resolveMcpEmailAccount).mockResolvedValue({
      ...emailAccount,
      provider: "microsoft",
    });
    vi.mocked(createEmailProvider).mockResolvedValue({
      searchMessages,
      getFolders: vi.fn().mockResolvedValue([folder({ id: "a" })]),
    } as never);

    await expect(
      searchInboxForMcp({ userId: "user_1", folderId: "Nope", logger }),
    ).rejects.toThrow("list_folders");
    expect(searchMessages).not.toHaveBeenCalled();
  });

  it("refuses a search with neither a query nor a folderId", async () => {
    const searchMessages = vi.fn();
    vi.mocked(createEmailProvider).mockResolvedValue({
      searchMessages,
    } as never);

    await expect(
      searchInboxForMcp({ userId: "user_1", logger }),
    ).rejects.toThrow("query or folderId");
    expect(searchMessages).not.toHaveBeenCalled();
  });

  it("refuses a folderId on a Gmail account", async () => {
    const searchMessages = vi.fn();
    vi.mocked(createEmailProvider).mockResolvedValue({
      searchMessages,
    } as never);

    await expect(
      searchInboxForMcp({ userId: "user_1", folderId: "Label_1", logger }),
    ).rejects.toThrow("only available for Outlook");
    expect(searchMessages).not.toHaveBeenCalled();
  });

  it("reads the latest messages in a thread", async () => {
    const getThreadMessages = vi.fn().mockResolvedValue([
      message({ id: "old", textPlain: "older" }),
      message({
        id: "new",
        textPlain: "latest question",
        snippet: "latest question",
      }),
    ]);
    vi.mocked(createEmailProvider).mockResolvedValue({
      getThreadMessages,
    } as never);

    const result = await readThreadForMcp({
      userId: "user_1",
      threadId: "thread-1",
      maxMessages: 1,
      logger,
    });

    expect(result.truncated).toBe(true);
    expect(result.messages).toHaveLength(1);
    expect(result.messages[0]?.messageId).toBe("new");
    expect(result.messages[0]?.content).toContain("latest question");
  });

  it("creates a mailbox draft and does not send", async () => {
    const createDraft = vi.fn().mockResolvedValue({ id: "draft-1" });
    const sendEmail = vi.fn();
    vi.mocked(createEmailProvider).mockResolvedValue({
      createDraft,
      sendEmail,
    } as never);

    const result = await createDraftForMcp({
      userId: "user_1",
      to: "contact@example.com",
      subject: "Thursday",
      body: "See you then.",
      logger,
    });

    expect(createDraft).toHaveBeenCalledWith({
      to: "contact@example.com",
      subject: "Thursday",
      messageHtml: "<p>See you then.</p>",
    });
    expect(sendEmail).not.toHaveBeenCalled();
    expect(result).toMatchObject({
      summary: "Saved a draft to contact@example.com; it was not sent.",
      draftId: "draft-1",
      sent: false,
    });
  });

  it("drafts a reply inside the original thread when given a message id", async () => {
    const createDraft = vi.fn().mockResolvedValue({ id: "draft-2" });
    vi.mocked(createEmailProvider).mockResolvedValue({ createDraft } as never);

    const result = await createDraftForMcp({
      userId: "user_1",
      to: "contact@example.com",
      subject: "Re: Thursday",
      body: "Works for me.",
      replyToMessageId: "message-1",
      logger,
    });

    expect(createDraft).toHaveBeenCalledWith(
      expect.objectContaining({ replyToMessageId: "message-1" }),
    );
    expect(result).toMatchObject({
      summary:
        "Saved a reply draft to contact@example.com in the same conversation; it was not sent.",
      draftId: "draft-2",
      replyToMessageId: "message-1",
      sent: false,
    });
  });

  it("archives a thread for the mailbox owner without deleting it", async () => {
    const archiveThread = vi.fn();
    const trashThread = vi.fn();
    vi.mocked(createEmailProvider).mockResolvedValue({
      archiveThread,
      trashThread,
    } as never);

    const result = await updateThreadForMcp({
      userId: "user_1",
      threadId: "thread-1",
      action: "archive",
      logger,
    });

    expect(archiveThread).toHaveBeenCalledWith(
      "thread-1",
      "owner@example.com",
      { requireChange: true },
    );
    expect(trashThread).not.toHaveBeenCalled();
    expect(result).toMatchObject({
      summary: "Archived 1 conversation; unarchive moves it back.",
      action: "archive",
      applied: true,
    });
  });

  it("marks a thread unread", async () => {
    const markReadThread = vi.fn();
    vi.mocked(createEmailProvider).mockResolvedValue({
      markReadThread,
    } as never);

    const result = await updateThreadForMcp({
      userId: "user_1",
      threadId: "thread-1",
      action: "mark_unread",
      logger,
    });

    expect(markReadThread).toHaveBeenCalledWith("thread-1", false, {
      requireChange: true,
    });
    expect(result.summary).toBe("Marked 1 conversation unread.");
  });

  it("moves a thread to the trash, where it can still be restored", async () => {
    const trashThread = vi.fn();
    vi.mocked(createEmailProvider).mockResolvedValue({
      trashThread,
    } as never);

    const result = await updateThreadForMcp({
      userId: "user_1",
      threadId: "thread-1",
      action: "trash",
      logger,
    });

    expect(trashThread).toHaveBeenCalledWith(
      "thread-1",
      emailAccount.email,
      "automation",
    );
    expect(result).toMatchObject({
      summary: "Moved 1 conversation to Trash; it can be restored.",
      action: "trash",
      applied: true,
    });
  });

  it("names Deleted Items when an Outlook thread goes to the trash", async () => {
    vi.mocked(resolveMcpEmailAccount).mockResolvedValue({
      ...emailAccount,
      provider: "microsoft",
    });
    vi.mocked(createEmailProvider).mockResolvedValue({
      getThreadFolderIds: vi.fn().mockResolvedValue(["inbox"]),
      trashThread: vi.fn(),
    } as never);

    const result = await updateThreadForMcp({
      userId: "user_1",
      threadId: "thread-1",
      action: "trash",
      logger,
    });

    expect(result.summary).toBe(
      "Moved 1 conversation to Deleted Items; it can be restored.",
    );
  });

  it.each([
    ["archive", "archiveThread"],
    ["unarchive", "unarchiveThread"],
    ["trash", "trashThread"],
  ] as const)("does not claim to %s an Outlook thread that was not found", async (action, method) => {
    vi.mocked(resolveMcpEmailAccount).mockResolvedValue({
      ...emailAccount,
      provider: "microsoft",
    });
    const move = vi.fn();
    vi.mocked(createEmailProvider).mockResolvedValue({
      getThreadFolderIds: vi.fn().mockResolvedValue([]),
      [method]: move,
    } as never);

    await expect(
      updateThreadForMcp({
        userId: "user_1",
        threadId: "gone",
        action,
        logger,
      }),
    ).rejects.toThrow(
      new SafeError(
        "Not moved: the conversation was not found or has no messages; nothing was moved.",
      ),
    );
    expect(move).not.toHaveBeenCalled();
  });

  it("passes a provider's thread-not-found on as a refusal when marking read", async () => {
    vi.mocked(resolveMcpEmailAccount).mockResolvedValue({
      ...emailAccount,
      provider: "microsoft",
    });
    vi.mocked(createEmailProvider).mockResolvedValue({
      markReadThread: vi
        .fn()
        .mockRejectedValue(new OutlookThreadNotFoundError()),
    } as never);

    const error = await updateThreadForMcp({
      userId: "user_1",
      threadId: "gone",
      action: "mark_read",
      logger,
    }).catch((e: unknown) => e);

    expect(error).toBeInstanceOf(SafeError);
    expect(error).toHaveProperty(
      "message",
      "Conversation not found or it has no messages; nothing was changed.",
    );
  });

  it("says a removed label is no longer on the conversation", async () => {
    vi.mocked(createEmailProvider).mockResolvedValue({
      getLabelByName: vi.fn().mockResolvedValue({ id: "label-1", name: "FYI" }),
      removeThreadLabel: vi.fn(),
    } as never);

    const result = await labelThreadForMcp({
      userId: "user_1",
      threadId: "thread-1",
      labelName: "FYI",
      remove: true,
      logger,
    });

    expect(result.summary).toBe(
      'The conversation no longer has the label "FYI".',
    );
  });

  it("says which rule actions ran and which are scheduled for later", async () => {
    vi.mocked(runRulesOnMessage).mockResolvedValue([
      {
        rule: { name: "Newsletter" },
        status: "APPLIED",
        actionItems: [
          { type: "LABEL", label: "News", delayInMinutes: null },
          { type: "ARCHIVE", label: null, delayInMinutes: 1440 },
        ],
      },
    ] as never);

    const result = await runRulesForMcp({
      userId: "user_1",
      messageId: "message-1",
      apply: true,
      logger,
    });

    expect(result.summary).toBe(
      'Ran the rule "Newsletter" (label "News"); scheduled for later: archive.',
    );
  });

  it("labels every message in the thread with an existing label", async () => {
    const labelMessage = vi.fn().mockResolvedValue({});
    vi.mocked(createEmailProvider).mockResolvedValue({
      getLabelByName: vi.fn().mockResolvedValue({ id: "label-1", name: "FYI" }),
      getThreadMessages: vi
        .fn()
        .mockResolvedValue([message(), message({ id: "message-2" })]),
      labelMessage,
    } as never);

    const result = await labelThreadForMcp({
      userId: "user_1",
      threadId: "thread-1",
      labelName: "FYI",
      logger,
    });

    expect(labelMessage).toHaveBeenCalledTimes(2);
    expect(result.summary).toBe(
      'Added the label "FYI" to 2 messages in 1 conversation.',
    );
    expect(labelMessage).toHaveBeenCalledWith({
      messageId: "message-2",
      labelId: "label-1",
      labelName: "FYI",
    });
  });

  it("does not create a label that does not exist and lists the real ones", async () => {
    const createLabel = vi.fn();
    const labelMessage = vi.fn();
    vi.mocked(createEmailProvider).mockResolvedValue({
      getLabelByName: vi.fn().mockResolvedValue(null),
      getLabels: vi.fn().mockResolvedValue([{ id: "l1", name: "Newsletter" }]),
      createLabel,
      labelMessage,
    } as never);

    await expect(
      labelThreadForMcp({
        userId: "user_1",
        threadId: "thread-1",
        labelName: "Invented",
        logger,
      }),
    ).rejects.toThrow("Newsletter");
    expect(createLabel).not.toHaveBeenCalled();
    expect(labelMessage).not.toHaveBeenCalled();
  });

  it("previews rules without applying them unless asked", async () => {
    vi.mocked(runRulesOnMessage).mockResolvedValue([
      {
        rule: { name: "Newsletter" },
        reason: "Bulk sender",
        status: "SKIPPED",
        actionItems: [{ type: "ARCHIVE", label: null, folderName: null }],
      },
    ] as never);

    const result = await runRulesForMcp({
      userId: "user_1",
      messageId: "message-1",
      apply: false,
      logger,
    });

    expect(runRulesOnMessage).toHaveBeenCalledWith(
      expect.objectContaining({
        emailAccountId: "account_1",
        messageId: "message-1",
        isTest: true,
      }),
    );
    expect(result).toMatchObject({
      summary:
        'Preview only, nothing changed: the rule "Newsletter" would match (archive).',
      preview: true,
      results: [{ ruleName: "Newsletter", actions: [{ type: "ARCHIVE" }] }],
    });
  });

  describe("folders", () => {
    const outlookAccount = { ...emailAccount, provider: "microsoft" };

    beforeEach(() => {
      const store = new Map<string, unknown>();
      vi.mocked(redis.get).mockImplementation(
        async (key: string) => store.get(key) ?? null,
      );
      vi.mocked(redis.set).mockImplementation(async (key: string, value) => {
        store.set(key, value);
        return "OK";
      });
    });
    const folderTree = [
      folder({ id: "inbox", displayName: "Inbox", systemType: "INBOX" }),
      folder({
        id: "projects",
        displayName: "Projects",
        childFolderCount: 1,
        childFolders: [folder({ id: "old", displayName: "Old" })],
      }),
      folder({ id: "receipts", displayName: "Receipts", totalItemCount: 3 }),
    ];

    function mockOutlook(tree: OutlookFolder[] = folderTree) {
      const deleteFolder = vi.fn().mockResolvedValue(undefined);
      const searchMessages = vi
        .fn()
        .mockResolvedValueOnce({
          messages: [
            message({ id: "r1", threadId: "ta" }),
            message({ id: "r2", threadId: "ta" }),
          ],
          nextPageToken: "page-2",
        })
        .mockResolvedValueOnce({
          messages: [message({ id: "r3", threadId: "tb" })],
        });
      const bulkArchiveThreads = vi.fn().mockResolvedValue({
        succeededThreadIds: ["ta"],
        failedThreadIds: ["tb"],
      });
      const unarchiveMessages = vi.fn().mockResolvedValue(undefined);
      const markMessagesReadState = vi.fn().mockResolvedValue(undefined);
      const getFolders = vi.fn().mockResolvedValue(tree);
      vi.mocked(resolveMcpEmailAccount).mockResolvedValue(outlookAccount);
      vi.mocked(createEmailProvider).mockResolvedValue({
        getFolders,
        deleteFolder,
        searchMessages,
        bulkArchiveThreads,
        unarchiveMessages,
        markMessagesReadState,
      } as never);
      return {
        getFolders,
        deleteFolder,
        searchMessages,
        bulkArchiveThreads,
        unarchiveMessages,
        markMessagesReadState,
      };
    }

    it("lists nested folders and marks only empty custom folders deletable", async () => {
      mockOutlook();

      const result = await listFoldersForMcp({ userId: "user_1", logger });

      expect(result.folders).toEqual([
        expect.objectContaining({
          folderId: "inbox",
          canDelete: false,
          deleteBlockedReason: "is a system folder",
        }),
        expect.objectContaining({
          folderId: "projects",
          subfolders: 1,
          canDelete: false,
          deleteBlockedReason: "has 1 subfolder; delete those first",
        }),
        expect.objectContaining({
          folderId: "old",
          path: "Projects ⦙ Old",
          canDelete: true,
        }),
        expect.objectContaining({
          folderId: "receipts",
          total: 3,
          canDelete: false,
          deleteBlockedReason:
            "still holds 3 emails; archive or move them first",
        }),
      ]);
      expect(result.folders[2]).not.toHaveProperty("deleteBlockedReason");
    });

    it.each([
      ["a system folder", "inbox", "system folder"],
      ["a folder that still holds mail", "receipts", "3 emails"],
      ["a folder with subfolders", "projects", "1 subfolder"],
      ["an unknown folder", "missing", "list_folders"],
    ])("refuses to delete %s", async (_case, folderId, message) => {
      const { deleteFolder } = mockOutlook();

      await expect(
        deleteFolderForMcp({ userId: "user_1", folderId, logger }),
      ).rejects.toThrow(message);
      expect(deleteFolder).not.toHaveBeenCalled();
      expect(checkRateLimit).not.toHaveBeenCalled();
    });

    it("stops deleting folders once the limit is reached", async () => {
      const { deleteFolder } = mockOutlook();
      vi.mocked(checkRateLimit).mockResolvedValue({
        limited: true,
        limit: 2,
        retryAfterSeconds: 600,
      });

      await expect(
        deleteFolderForMcp({ userId: "user_1", folderId: "old", logger }),
      ).rejects.toThrow("At most 2 folders");
      expect(deleteFolder).not.toHaveBeenCalled();
    });

    it("deletes an empty custom folder", async () => {
      const { deleteFolder } = mockOutlook();

      const result = await deleteFolderForMcp({
        userId: "user_1",
        folderId: "old",
        logger,
      });

      expect(deleteFolder).toHaveBeenCalledWith("old");
      expect(result).toMatchObject({
        summary: "Deleted the empty folder Projects ⦙ Old.",
        path: "Projects ⦙ Old",
        deleted: true,
      });

      const { rule } = vi.mocked(checkRateLimit).mock.calls[0][0];
      expect(rule.limit).toBe(2);
      expect(rule.windowSeconds).toBe(600);
      expect(rule.key).toBe(
        createRateLimitKey([
          "rate-limit",
          "mcp",
          "folder-delete",
          outlookAccount.id,
        ]),
      );
    });

    describe("deleting a folder that recently held mail", () => {
      const receiptsEmptied = [
        folder({ id: "receipts", displayName: "Receipts" }),
      ];
      const heldMessage = (path: string) =>
        `Not deleted: "${path}" held email or subfolders until recently (now moved). Ask the user first; delete it only after they say yes, by calling delete_folder again with userConfirmed: true.`;

      it("refuses to delete a folder after archive_folder emptied it", async () => {
        await archiveFolder();
        const { deleteFolder } = mockOutlook(receiptsEmptied);

        await expect(
          deleteFolderForMcp({
            userId: "user_1",
            folderId: "receipts",
            logger,
          }),
        ).rejects.toThrow(heldMessage("Receipts"));
        expect(deleteFolder).not.toHaveBeenCalled();
        expect(checkRateLimit).not.toHaveBeenCalled();
      });

      it("deletes it once the user has confirmed", async () => {
        await archiveFolder();
        const { deleteFolder } = mockOutlook(receiptsEmptied);

        const result = await deleteFolderForMcp({
          userId: "user_1",
          folderId: "receipts",
          userConfirmed: true,
          logger,
        });

        expect(deleteFolder).toHaveBeenCalledWith("receipts");
        expect(result).toMatchObject({ path: "Receipts", deleted: true });
        expect(checkRateLimit).toHaveBeenCalledTimes(1);
      });

      it("treats userConfirmed false like a missing flag", async () => {
        await archiveFolder();
        const { deleteFolder } = mockOutlook(receiptsEmptied);

        await expect(
          deleteFolderForMcp({
            userId: "user_1",
            folderId: "receipts",
            userConfirmed: false,
            logger,
          }),
        ).rejects.toThrow("Ask the user first");
        expect(deleteFolder).not.toHaveBeenCalled();
      });

      it("refuses to delete a folder after move_folder_to_inbox emptied it", async () => {
        const { unarchiveMessages } = mockOutlook();
        await moveFolderToInboxForMcp({
          userId: "user_1",
          folderId: "receipts",
          logger,
        });
        expect(unarchiveMessages).toHaveBeenCalled();
        const { deleteFolder } = mockOutlook(receiptsEmptied);

        await expect(
          deleteFolderForMcp({
            userId: "user_1",
            folderId: "receipts",
            logger,
          }),
        ).rejects.toThrow(heldMessage("Receipts"));
        expect(deleteFolder).not.toHaveBeenCalled();
      });

      it("remembers a folder for a day, per email account", async () => {
        await archiveFolder();

        expect(redis.set).toHaveBeenCalledTimes(1);
        const [key, , options] = vi.mocked(redis.set).mock.calls[0];
        expect(key).toBe(
          createRateLimitKey([
            "mcp",
            "folder-held-mail",
            outlookAccount.id,
            createHash("sha256").update("receipts").digest("hex"),
          ]),
        );
        expect(options).toEqual({ ex: 24 * 60 * 60 });

        const { deleteFolder } = mockOutlook(receiptsEmptied);
        vi.mocked(resolveMcpEmailAccount).mockResolvedValue({
          ...outlookAccount,
          id: "account_2",
        });
        await deleteFolderForMcp({
          userId: "user_1",
          folderId: "receipts",
          logger,
        });
        expect(deleteFolder).toHaveBeenCalledWith("receipts");
      });

      it("does not remember a folder when no email moved out of it", async () => {
        const { searchMessages } = mockOutlook();
        searchMessages.mockReset().mockResolvedValue({ messages: [] });
        await archiveFolderForMcp({
          userId: "user_1",
          folderId: "receipts",
          logger,
        });
        expect(redis.set).not.toHaveBeenCalled();
      });

      it("marks the parent when a subfolder is deleted", async () => {
        const { deleteFolder } = mockOutlook();
        await deleteFolderForMcp({ userId: "user_1", folderId: "old", logger });
        expect(deleteFolder).toHaveBeenCalledWith("old");

        const projectsNowEmpty = mockOutlook([
          folder({ id: "projects", displayName: "Projects" }),
        ]);
        await expect(
          deleteFolderForMcp({
            userId: "user_1",
            folderId: "projects",
            logger,
          }),
        ).rejects.toThrow(heldMessage("Projects"));
        expect(projectsNowEmpty.deleteFolder).not.toHaveBeenCalled();
      });

      it("does not mark anything when a top-level folder is deleted", async () => {
        mockOutlook([folder({ id: "empty", displayName: "Empty" })]);
        await deleteFolderForMcp({
          userId: "user_1",
          folderId: "empty",
          logger,
        });
        expect(redis.set).not.toHaveBeenCalled();
      });

      it("deletes a folder that was never emptied without confirmation", async () => {
        const { deleteFolder } = mockOutlook([
          folder({ id: "empty", displayName: "Empty" }),
        ]);

        const result = await deleteFolderForMcp({
          userId: "user_1",
          folderId: "empty",
          logger,
        });

        expect(deleteFolder).toHaveBeenCalledWith("empty");
        expect(result.deleted).toBe(true);
      });

      it("requires confirmation when Redis cannot be read", async () => {
        const { deleteFolder } = mockOutlook([
          folder({ id: "empty", displayName: "Empty" }),
        ]);
        vi.mocked(redis.get).mockRejectedValue(new Error("redis down"));

        await expect(
          deleteFolderForMcp({ userId: "user_1", folderId: "empty", logger }),
        ).rejects.toThrow(heldMessage("Empty"));
        expect(deleteFolder).not.toHaveBeenCalled();

        await deleteFolderForMcp({
          userId: "user_1",
          folderId: "empty",
          userConfirmed: true,
          logger,
        });
        expect(deleteFolder).toHaveBeenCalledWith("empty");
      });

      it("moves nothing when Redis cannot record the folder", async () => {
        const archiving = mockOutlook();
        vi.mocked(redis.set).mockRejectedValue(new Error("redis down"));

        await expect(
          archiveFolderForMcp({
            userId: "user_1",
            folderId: "receipts",
            logger,
          }),
        ).rejects.toThrow(
          'Not archived: could not record "Receipts" for the delete check; nothing was moved.',
        );
        const moving = mockOutlook();
        await expect(
          moveFolderToInboxForMcp({
            userId: "user_1",
            folderId: "receipts",
            logger,
          }),
        ).rejects.toThrow(
          'Not moved: could not record "Receipts" for the delete check; nothing was moved.',
        );
        expect(archiving.bulkArchiveThreads).not.toHaveBeenCalled();
        expect(moving.unarchiveMessages).not.toHaveBeenCalled();
      });

      it("deletes nothing when Redis cannot record the parent", async () => {
        const { deleteFolder } = mockOutlook();
        vi.mocked(redis.set).mockRejectedValue(new Error("redis down"));

        await expect(
          deleteFolderForMcp({ userId: "user_1", folderId: "old", logger }),
        ).rejects.toThrow(
          'Not deleted: could not record "Projects" for the delete check; nothing was deleted.',
        );
        expect(deleteFolder).not.toHaveBeenCalled();
      });

      it("reports the system and not-empty blockers before asking for confirmation", async () => {
        await archiveFolder();
        mockOutlook();

        await expect(
          deleteFolderForMcp({
            userId: "user_1",
            folderId: "receipts",
            logger,
          }),
        ).rejects.toThrow("still holds 3 emails");
        await expect(
          deleteFolderForMcp({ userId: "user_1", folderId: "inbox", logger }),
        ).rejects.toThrow("system folder");
      });

      describe("folders the server has seen holding mail", () => {
        const heldKey = (folderId: string) =>
          createRateLimitKey([
            "mcp",
            "folder-held-mail",
            outlookAccount.id,
            createHash("sha256").update(folderId).digest("hex"),
          ]);
        const recordedKeys = () =>
          vi.mocked(redis.set).mock.calls.map(([key]) => key);

        it("records every non-empty custom folder when list_folders shows it", async () => {
          mockOutlook();

          await listFoldersForMcp({ userId: "user_1", logger });

          expect(recordedKeys().sort()).toEqual(
            [heldKey("projects"), heldKey("receipts")].sort(),
          );
          for (const [, , options] of vi.mocked(redis.set).mock.calls)
            expect(options).toEqual({ ex: 24 * 60 * 60 });
        });

        it("skips empty and system folders when listing", async () => {
          mockOutlook([
            folder({ id: "inbox", displayName: "Inbox", systemType: "INBOX" }),
            folder({
              id: "sent",
              displayName: "Sent",
              systemType: "SENT",
              totalItemCount: 40,
            }),
            folder({ id: "empty", displayName: "Empty" }),
          ]);

          await listFoldersForMcp({ userId: "user_1", logger });

          expect(redis.set).not.toHaveBeenCalled();
        });

        it("records a folder when search_inbox finds mail in it", async () => {
          mockOutlook();

          await searchInboxForMcp({
            userId: "user_1",
            folderId: "Receipts",
            logger,
          });

          expect(recordedKeys()).toEqual([heldKey("receipts")]);
        });

        it("does not record a folder when the search finds nothing or it is a system folder", async () => {
          const empty = mockOutlook();
          empty.searchMessages.mockReset().mockResolvedValue({ messages: [] });
          await searchInboxForMcp({
            userId: "user_1",
            folderId: "Receipts",
            logger,
          });

          mockOutlook();
          await searchInboxForMcp({
            userId: "user_1",
            folderId: "Inbox",
            logger,
          });

          expect(redis.set).not.toHaveBeenCalled();
        });

        it("records the folders a whole-mailbox search finds mail in", async () => {
          const { searchMessages } = mockOutlook();
          searchMessages.mockReset().mockResolvedValue({
            messages: [
              message({ id: "r1", parentFolderId: "receipts" }),
              message({ id: "r2", parentFolderId: "receipts" }),
              message({ id: "r3" }),
            ],
          });

          await searchInboxForMcp({ userId: "user_1", query: "sale", logger });

          expect(recordedKeys()).toEqual([heldKey("receipts")]);
        });

        it("refuses to delete a folder emptied one email at a time with update_thread", async () => {
          mockOutlook();
          await listFoldersForMcp({ userId: "user_1", logger });
          // The thread sits outside Receipts, so only the earlier list or
          // search records Receipts; update_thread's own record can't.
          vi.mocked(createEmailProvider).mockResolvedValue({
            getThreadFolderIds: vi.fn().mockResolvedValue(["inbox"]),
            trashThread: vi.fn().mockResolvedValue(undefined),
          } as never);
          await updateThreadForMcp({
            userId: "user_1",
            threadId: "ta",
            action: "trash",
            logger,
          });
          const { deleteFolder } = mockOutlook([
            folder({ id: "receipts", displayName: "Receipts" }),
          ]);

          await expect(
            deleteFolderForMcp({
              userId: "user_1",
              folderId: "receipts",
              logger,
            }),
          ).rejects.toThrow(heldMessage("Receipts"));
          expect(deleteFolder).not.toHaveBeenCalled();

          await deleteFolderForMcp({
            userId: "user_1",
            folderId: "receipts",
            userConfirmed: true,
            logger,
          });
          expect(deleteFolder).toHaveBeenCalledWith("receipts");
        });

        it("refuses to delete a folder emptied after a whole-mailbox search", async () => {
          const { searchMessages } = mockOutlook();
          searchMessages.mockReset().mockResolvedValue({
            messages: [
              message({ id: "r1", threadId: "ta", parentFolderId: "receipts" }),
            ],
          });
          await searchInboxForMcp({
            userId: "user_1",
            query: "receipt",
            logger,
          });
          // The thread sits outside Receipts, so only the earlier list or
          // search records Receipts; update_thread's own record can't.
          vi.mocked(createEmailProvider).mockResolvedValue({
            getThreadFolderIds: vi.fn().mockResolvedValue(["inbox"]),
            trashThread: vi.fn().mockResolvedValue(undefined),
          } as never);
          await updateThreadForMcp({
            userId: "user_1",
            threadId: "ta",
            action: "trash",
            logger,
          });
          const { deleteFolder } = mockOutlook([
            folder({ id: "receipts", displayName: "Receipts" }),
          ]);

          await expect(
            deleteFolderForMcp({
              userId: "user_1",
              folderId: "receipts",
              logger,
            }),
          ).rejects.toThrow(heldMessage("Receipts"));
          expect(deleteFolder).not.toHaveBeenCalled();

          await deleteFolderForMcp({
            userId: "user_1",
            folderId: "receipts",
            userConfirmed: true,
            logger,
          });
          expect(deleteFolder).toHaveBeenCalledWith("receipts");
        });

        it.each([
          ["archive", "archiveThread"],
          ["unarchive", "unarchiveThread"],
          ["trash", "trashThread"],
        ] as const)("refuses to delete a folder that update_thread %s emptied by moving the whole conversation", async (action, method) => {
          vi.mocked(resolveMcpEmailAccount).mockResolvedValue(outlookAccount);
          const calls: string[] = [];
          const getThreadFolderIds = vi.fn(async () => {
            calls.push("read");
            return ["inbox", "receipts"];
          });
          vi.mocked(createEmailProvider).mockResolvedValue({
            getThreadFolderIds,
            [method]: vi.fn(async () => {
              calls.push("move");
            }),
          } as never);
          await updateThreadForMcp({
            userId: "user_1",
            threadId: "ta",
            action,
            logger,
          });
          expect(getThreadFolderIds).toHaveBeenCalledWith("ta");
          expect(calls).toEqual(["read", "move"]);
          const { deleteFolder } = mockOutlook([
            folder({ id: "receipts", displayName: "Receipts" }),
          ]);

          await expect(
            deleteFolderForMcp({
              userId: "user_1",
              folderId: "receipts",
              logger,
            }),
          ).rejects.toThrow(heldMessage("Receipts"));
          expect(deleteFolder).not.toHaveBeenCalled();

          await deleteFolderForMcp({
            userId: "user_1",
            folderId: "receipts",
            userConfirmed: true,
            logger,
          });
          expect(deleteFolder).toHaveBeenCalledWith("receipts");
        });

        it("moves nothing when the thread's folders can't be read", async () => {
          vi.mocked(resolveMcpEmailAccount).mockResolvedValue(outlookAccount);
          const archiveThread = vi.fn().mockResolvedValue(undefined);
          vi.mocked(createEmailProvider).mockResolvedValue({
            getThreadFolderIds: vi.fn().mockRejectedValue(new Error("429")),
            archiveThread,
          } as never);

          await expect(
            updateThreadForMcp({
              userId: "user_1",
              threadId: "ta",
              action: "archive",
              logger,
            }),
          ).rejects.toThrow("nothing was moved");
          expect(archiveThread).not.toHaveBeenCalled();
        });

        it("moves nothing when the thread's folders can't be recorded", async () => {
          vi.mocked(resolveMcpEmailAccount).mockResolvedValue(outlookAccount);
          vi.mocked(redis.set).mockRejectedValue(new Error("redis down"));
          const trashThread = vi.fn().mockResolvedValue(undefined);
          vi.mocked(createEmailProvider).mockResolvedValue({
            getThreadFolderIds: vi.fn().mockResolvedValue(["receipts"]),
            trashThread,
          } as never);

          await expect(
            updateThreadForMcp({
              userId: "user_1",
              threadId: "ta",
              action: "trash",
              logger,
            }),
          ).rejects.toThrow("nothing was moved");
          expect(trashThread).not.toHaveBeenCalled();
        });

        it("does not read the thread's folders to mark it read, or for Gmail", async () => {
          const getThreadFolderIds = vi.fn().mockResolvedValue(["receipts"]);
          vi.mocked(resolveMcpEmailAccount).mockResolvedValue(outlookAccount);
          vi.mocked(createEmailProvider).mockResolvedValue({
            getThreadFolderIds,
            markReadThread: vi.fn(),
          } as never);
          await updateThreadForMcp({
            userId: "user_1",
            threadId: "ta",
            action: "mark_read",
            logger,
          });

          vi.mocked(resolveMcpEmailAccount).mockResolvedValue(emailAccount);
          vi.mocked(createEmailProvider).mockResolvedValue({
            getThreadFolderIds,
            archiveThread: vi.fn(),
          } as never);
          await updateThreadForMcp({
            userId: "user_1",
            threadId: "ta",
            action: "archive",
            logger,
          });

          expect(getThreadFolderIds).not.toHaveBeenCalled();
          expect(redis.set).not.toHaveBeenCalled();
        });

        it("records a folder when delete_folder refuses because it still holds mail", async () => {
          mockOutlook();

          await expect(
            deleteFolderForMcp({
              userId: "user_1",
              folderId: "receipts",
              logger,
            }),
          ).rejects.toThrow("still holds 3 emails");
          expect(recordedKeys()).toEqual([heldKey("receipts")]);

          const { deleteFolder } = mockOutlook([
            folder({ id: "receipts", displayName: "Receipts" }),
          ]);
          await expect(
            deleteFolderForMcp({
              userId: "user_1",
              folderId: "receipts",
              logger,
            }),
          ).rejects.toThrow(heldMessage("Receipts"));
          expect(deleteFolder).not.toHaveBeenCalled();
        });

        it("does not record a system folder when delete_folder refuses it", async () => {
          mockOutlook();

          await expect(
            deleteFolderForMcp({ userId: "user_1", folderId: "inbox", logger }),
          ).rejects.toThrow("system folder");

          expect(redis.set).not.toHaveBeenCalled();
        });

        it("still lists folders, searches and refuses when Redis cannot record", async () => {
          mockOutlook();
          vi.mocked(redis.set).mockRejectedValue(new Error("redis down"));

          const listed = await listFoldersForMcp({ userId: "user_1", logger });
          expect(listed.folders).toHaveLength(4);

          const searched = await searchInboxForMcp({
            userId: "user_1",
            folderId: "Receipts",
            logger,
          });
          expect(searched.messages).toHaveLength(2);

          await expect(
            deleteFolderForMcp({
              userId: "user_1",
              folderId: "receipts",
              logger,
            }),
          ).rejects.toThrow("still holds 3 emails");
        });
      });

      async function archiveFolder() {
        mockOutlook();
        await archiveFolderForMcp({
          userId: "user_1",
          folderId: "receipts",
          logger,
        });
      }
    });

    it("refuses Gmail accounts before touching the mailbox", async () => {
      await expect(
        deleteFolderForMcp({ userId: "user_1", folderId: "old", logger }),
      ).rejects.toThrow("Outlook");
      expect(createEmailProvider).not.toHaveBeenCalled();
    });

    it("archives every email in a folder in one call and reports failures", async () => {
      const { searchMessages, bulkArchiveThreads } = mockOutlook();

      const result = await archiveFolderForMcp({
        userId: "user_1",
        folderId: "receipts",
        logger,
      });

      expect(searchMessages).toHaveBeenCalledTimes(2);
      expect(searchMessages).toHaveBeenLastCalledWith(
        expect.objectContaining({ folderId: "receipts", pageToken: "page-2" }),
      );
      expect(bulkArchiveThreads).toHaveBeenCalledWith(
        [
          { threadId: "ta", messageIds: ["r1", "r2"] },
          { threadId: "tb", messageIds: ["r3"] },
        ],
        outlookAccount.email,
      );
      expect(result).toMatchObject({
        summary:
          "Archived 2 of 3 emails from Receipts; 1 could not be moved and is still there.",
        path: "Receipts",
        archived: 2,
        failed: 1,
      });
    });

    it("says an empty folder was already empty and archives nothing", async () => {
      const { searchMessages } = mockOutlook();
      searchMessages.mockReset().mockResolvedValue({ messages: [] });

      const result = await archiveFolderForMcp({
        userId: "user_1",
        folderId: "receipts",
        logger,
      });

      expect(result).toMatchObject({
        summary: "Archived nothing: Receipts was already empty.",
        archived: 0,
        failed: 0,
      });
    });

    it.each([
      ["a system folder", "inbox", "system folder"],
      ["an unknown folder", "missing", "list_folders"],
    ])("refuses to archive %s", async (_case, folderId, message) => {
      const { bulkArchiveThreads } = mockOutlook();

      await expect(
        archiveFolderForMcp({ userId: "user_1", folderId, logger }),
      ).rejects.toThrow(message);
      expect(bulkArchiveThreads).not.toHaveBeenCalled();
    });

    describe("folder by name", () => {
      it("resolves a path case-insensitively", async () => {
        const { deleteFolder } = mockOutlook();

        const result = await deleteFolderForMcp({
          userId: "user_1",
          folderId: "projects ⦙ OLD",
          logger,
        });

        expect(deleteFolder).toHaveBeenCalledWith("old");
        expect(result.folderId).toBe("old");
      });

      it("resolves a display name case-insensitively", async () => {
        const { searchMessages, bulkArchiveThreads } = mockOutlook();

        const result = await archiveFolderForMcp({
          userId: "user_1",
          folderId: "RECEIPTS",
          logger,
        });

        expect(searchMessages).toHaveBeenCalledWith(
          expect.objectContaining({ folderId: "receipts" }),
        );
        expect(bulkArchiveThreads).toHaveBeenCalled();
        expect(result.folderId).toBe("receipts");
      });

      it("prefers an exact ID over a folder that happens to have that name", async () => {
        const { deleteFolder } = mockOutlook([
          folder({ id: "old", displayName: "Old" }),
          folder({ id: "other", displayName: "old" }),
        ]);

        await deleteFolderForMcp({ userId: "user_1", folderId: "old", logger });

        expect(deleteFolder).toHaveBeenCalledWith("old");
      });

      it("prefers a path match over a display name shared with a subfolder", async () => {
        const { deleteFolder } = mockOutlook([
          folder({ id: "top-old", displayName: "Old" }),
          folder({
            id: "projects",
            displayName: "Projects",
            childFolderCount: 1,
            childFolders: [folder({ id: "nested-old", displayName: "Old" })],
          }),
        ]);

        await deleteFolderForMcp({ userId: "user_1", folderId: "Old", logger });

        expect(deleteFolder).toHaveBeenCalledWith("top-old");
      });

      it("lists the matching paths when a name is ambiguous", async () => {
        const { deleteFolder } = mockOutlook([
          folder({
            id: "projects",
            displayName: "Projects",
            childFolderCount: 1,
            childFolders: [folder({ id: "p-old", displayName: "Old" })],
          }),
          folder({
            id: "clients",
            displayName: "Clients",
            childFolderCount: 1,
            childFolders: [folder({ id: "c-old", displayName: "Old" })],
          }),
        ]);

        const attempt = deleteFolderForMcp({
          userId: "user_1",
          folderId: "old",
          logger,
        });

        await expect(attempt).rejects.toThrow("Projects ⦙ Old");
        await expect(attempt).rejects.toThrow("Clients ⦙ Old");
        expect(deleteFolder).not.toHaveBeenCalled();
      });

      it("says a name was not found and points to list_folders", async () => {
        mockOutlook();

        await expect(
          archiveFolderForMcp({
            userId: "user_1",
            folderId: "Recipts",
            logger,
          }),
        ).rejects.toThrow(/"Recipts" not found.*list_folders/);
      });
    });

    describe("move_folder_to_inbox", () => {
      it("moves every email in a folder back to the inbox and reports failures", async () => {
        const { searchMessages, unarchiveMessages, bulkArchiveThreads } =
          mockOutlook();
        unarchiveMessages
          .mockRejectedValueOnce(new Error("batch failed"))
          .mockResolvedValueOnce(undefined)
          .mockResolvedValueOnce(undefined)
          .mockRejectedValueOnce(new Error("gone"));

        const result = await moveFolderToInboxForMcp({
          userId: "user_1",
          folderId: "Receipts",
          logger,
        });

        expect(searchMessages).toHaveBeenCalledTimes(2);
        expect(searchMessages).toHaveBeenCalledWith(
          expect.objectContaining({ folderId: "receipts" }),
        );
        expect(unarchiveMessages).toHaveBeenNthCalledWith(1, [
          "r1",
          "r2",
          "r3",
        ]);
        expect(unarchiveMessages).toHaveBeenCalledWith(["r1"]);
        expect(bulkArchiveThreads).not.toHaveBeenCalled();
        expect(result).toMatchObject({
          summary:
            "Moved 2 of 3 emails from Receipts to the Inbox; 1 could not be moved and is still there.",
          folderId: "receipts",
          path: "Receipts",
          moved: 2,
          failed: 1,
        });
      });

      it("moves everything in one call when nothing fails", async () => {
        const { unarchiveMessages } = mockOutlook();

        const result = await moveFolderToInboxForMcp({
          userId: "user_1",
          folderId: "receipts",
          logger,
        });

        expect(unarchiveMessages).toHaveBeenCalledTimes(1);
        expect(result).toMatchObject({
          summary:
            "Moved 3 emails from Receipts to the Inbox; it is now empty.",
          moved: 3,
          failed: 0,
        });
      });

      it.each([
        ["a system folder", "inbox", "system folder"],
        ["an unknown folder", "missing", "list_folders"],
      ])("refuses to move %s", async (_case, folderId, message) => {
        const { unarchiveMessages } = mockOutlook();

        await expect(
          moveFolderToInboxForMcp({ userId: "user_1", folderId, logger }),
        ).rejects.toThrow(message);
        expect(unarchiveMessages).not.toHaveBeenCalled();
      });

      it("refuses Gmail accounts before touching the mailbox", async () => {
        await expect(
          moveFolderToInboxForMcp({
            userId: "user_1",
            folderId: "old",
            logger,
          }),
        ).rejects.toThrow("Outlook");
        expect(createEmailProvider).not.toHaveBeenCalled();
      });
    });

    describe("mark_read", () => {
      it("marks every unread email in a folder across pages", async () => {
        const { searchMessages, markMessagesReadState } = mockOutlook();

        const result = await markReadForMcp({
          userId: "user_1",
          folderId: "Receipts",
          logger,
        });

        expect(searchMessages).toHaveBeenCalledTimes(2);
        expect(searchMessages).toHaveBeenCalledWith(
          expect.objectContaining({
            folderId: "receipts",
            readState: "unread",
          }),
        );
        expect(searchMessages).toHaveBeenLastCalledWith(
          expect.objectContaining({ pageToken: "page-2" }),
        );
        expect(markMessagesReadState).toHaveBeenCalledWith(
          ["r1", "r2", "r3"],
          true,
        );
        expect(result).toMatchObject({
          summary: "Marked 3 emails in Receipts read.",
          folderId: "receipts",
          path: "Receipts",
          read: true,
          marked: 3,
          failed: 0,
          moreLeft: false,
        });
      });

      it("collects every page before marking anything", async () => {
        const { searchMessages, markMessagesReadState } = mockOutlook();

        await markReadForMcp({
          userId: "user_1",
          folderId: "receipts",
          logger,
        });

        const lastSearch = Math.max(...searchMessages.mock.invocationCallOrder);
        expect(
          markMessagesReadState.mock.invocationCallOrder[0],
        ).toBeGreaterThan(lastSearch);
      });

      it("marks one sender's mail across the mailbox without listing folders", async () => {
        const { getFolders, searchMessages, markMessagesReadState } =
          mockOutlook();

        const result = await markReadForMcp({
          userId: "user_1",
          sender: "Daily Deals <news@shop.example>",
          logger,
        });

        expect(getFolders).not.toHaveBeenCalled();
        expect(searchMessages).toHaveBeenCalledWith(
          expect.objectContaining({
            fromEmail: "news@shop.example",
            folderId: undefined,
            readState: "unread",
          }),
        );
        expect(markMessagesReadState).toHaveBeenCalledWith(
          ["r1", "r2", "r3"],
          true,
        );
        expect(result).toMatchObject({
          summary: "Marked 3 emails from news@shop.example read.",
          sender: "news@shop.example",
          marked: 3,
        });
      });

      it("marks read mail unread with read: false", async () => {
        const { searchMessages, markMessagesReadState } = mockOutlook();

        const result = await markReadForMcp({
          userId: "user_1",
          folderId: "receipts",
          read: false,
          logger,
        });

        expect(searchMessages).toHaveBeenCalledWith(
          expect.objectContaining({ readState: "read" }),
        );
        expect(markMessagesReadState).toHaveBeenCalledWith(
          ["r1", "r2", "r3"],
          false,
        );
        expect(result).toMatchObject({
          summary: "Marked 3 emails in Receipts unread.",
          read: false,
          marked: 3,
        });
      });

      it("counts the emails it could not mark", async () => {
        const { markMessagesReadState } = mockOutlook();
        markMessagesReadState
          .mockRejectedValueOnce(new Error("batch failed"))
          .mockResolvedValueOnce(undefined)
          .mockResolvedValueOnce(undefined)
          .mockRejectedValueOnce(new Error("throttled"));

        const result = await markReadForMcp({
          userId: "user_1",
          folderId: "receipts",
          logger,
        });

        expect(markMessagesReadState).toHaveBeenCalledWith(["r3"], true);
        expect(result).toMatchObject({
          summary:
            "Marked 2 of 3 emails in Receipts read; 1 could not be changed and is still unread.",
          marked: 2,
          failed: 1,
        });
      });

      it("allows the Inbox", async () => {
        const { searchMessages } = mockOutlook();

        const result = await markReadForMcp({
          userId: "user_1",
          folderId: "Inbox",
          logger,
        });

        expect(searchMessages).toHaveBeenCalledWith(
          expect.objectContaining({ folderId: "inbox" }),
        );
        expect(result).toMatchObject({
          summary: "Marked 3 emails in Inbox read.",
          path: "Inbox",
          marked: 3,
        });
      });

      it("never moves, trashes or deletes mail and records no held folder", async () => {
        const { bulkArchiveThreads, unarchiveMessages, deleteFolder } =
          mockOutlook();

        await markReadForMcp({
          userId: "user_1",
          folderId: "receipts",
          logger,
        });

        expect(bulkArchiveThreads).not.toHaveBeenCalled();
        expect(unarchiveMessages).not.toHaveBeenCalled();
        expect(deleteFolder).not.toHaveBeenCalled();
        expect(redis.set).not.toHaveBeenCalled();
      });

      it("stops at 1,000 emails and says more are left", async () => {
        const { searchMessages, markMessagesReadState } = mockOutlook();
        let page = 0;
        searchMessages.mockReset().mockImplementation(async () => {
          page++;
          return {
            messages: Array.from({ length: 20 }, (_, i) =>
              message({ id: `m${page}-${i}`, threadId: `t${page}-${i}` }),
            ),
            nextPageToken: `page-${page + 1}`,
          };
        });

        const result = await markReadForMcp({
          userId: "user_1",
          folderId: "receipts",
          logger,
        });

        expect(searchMessages).toHaveBeenCalledTimes(50);
        expect(markMessagesReadState).toHaveBeenCalledTimes(50);
        expect(result).toMatchObject({
          summary:
            "Marked 1000 emails in Receipts read; more are left, so call mark_read again.",
          marked: 1000,
          failed: 0,
          moreLeft: true,
        });
      });

      it("marks nothing when nothing is unread", async () => {
        const { searchMessages, markMessagesReadState } = mockOutlook();
        searchMessages.mockReset().mockResolvedValue({ messages: [] });

        const result = await markReadForMcp({
          userId: "user_1",
          folderId: "receipts",
          logger,
        });

        expect(markMessagesReadState).not.toHaveBeenCalled();
        expect(result).toMatchObject({
          summary: "Marked nothing: no unread emails in Receipts.",
          marked: 0,
          failed: 0,
          moreLeft: false,
        });
      });

      it("refuses a call with neither a folder nor a sender", async () => {
        mockOutlook();

        await expect(
          markReadForMcp({ userId: "user_1", logger }),
        ).rejects.toThrow("folderId or sender");
        expect(createEmailProvider).not.toHaveBeenCalled();
      });

      it("refuses a sender that is not an email address", async () => {
        mockOutlook();

        await expect(
          markReadForMcp({ userId: "user_1", sender: "Daily Deals", logger }),
        ).rejects.toThrow("not an email address");
        expect(createEmailProvider).not.toHaveBeenCalled();
      });

      it("refuses more than one sender address", async () => {
        mockOutlook();

        await expect(
          markReadForMcp({
            userId: "user_1",
            sender: "a@shop.example, b@shop.example",
            logger,
          }),
        ).rejects.toThrow("one sender address");
        expect(createEmailProvider).not.toHaveBeenCalled();
      });

      it("limits a sender to one folder when both are given", async () => {
        const { searchMessages } = mockOutlook();

        await markReadForMcp({
          userId: "user_1",
          folderId: "Receipts",
          sender: "news@shop.example",
          logger,
        });

        expect(searchMessages).toHaveBeenCalledWith(
          expect.objectContaining({
            folderId: "receipts",
            fromEmail: "news@shop.example",
          }),
        );
      });

      it("refuses Gmail accounts before touching the mailbox", async () => {
        await expect(
          markReadForMcp({
            userId: "user_1",
            sender: "news@shop.example",
            logger,
          }),
        ).rejects.toThrow("Outlook");
        expect(createEmailProvider).not.toHaveBeenCalled();
      });
    });
  });
});

function folder(overrides: Partial<OutlookFolder> = {}): OutlookFolder {
  return {
    id: "folder",
    displayName: "Folder",
    childFolders: [],
    childFolderCount: 0,
    totalItemCount: 0,
    unreadItemCount: 0,
    ...overrides,
  };
}
