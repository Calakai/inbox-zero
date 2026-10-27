import chunk from "lodash/chunk";
import { z } from "zod";
import { runRulesOnMessage } from "@/utils/ai/choose-rule/run-rules-on-message";
import { createEmailProvider } from "@/utils/email/provider";
import { isMicrosoftProvider } from "@/utils/email/provider-types";
import { getEmailForLLM } from "@/utils/get-email-from-message";
import { resolveMcpEmailAccount } from "@/utils/mcp/account-selection";
import {
  flattenOutlookFolders,
  type OutlookFolder,
} from "@/utils/outlook/folders";
import { checkRateLimit, createRateLimitKey } from "@/utils/rate-limit";
import {
  markFolderHeldMail,
  wasFolderRecentlyHeldMail,
} from "@/utils/redis/mcp-folder-held-mail";
import { textToHtmlParagraphs } from "@/utils/string";
import type { EmailProvider } from "@/utils/email/types";
import type { Logger } from "@/utils/logger";
import type { ParsedMessage } from "@/utils/types";

const SEARCH_INBOX_MAX_RESULTS = 20;
const READ_THREAD_MAX_MESSAGES = 20;
const READ_MESSAGE_MAX_LENGTH = 4000;
const LABEL_NAMES_IN_ERROR = 50;
const ARCHIVE_FOLDER_MAX_MESSAGES = 1000;
const MOVE_TO_INBOX_CHUNK_SIZE = 20;
// Each delete takes one call, so a model that misreads a vague follow-up can
// walk the whole folder list; the cap stops it after a couple.
const FOLDER_DELETE_LIMIT = { limit: 2, windowSeconds: 10 * 60 };

export const mcpAccountSelectorShape = {
  emailAccountId: z.string().optional(),
  emailAddress: z.string().email().optional(),
};

export const searchInboxInputShape = {
  ...mcpAccountSelectorShape,
  query: z
    .string()
    .trim()
    .min(1)
    .optional()
    .describe(
      "Search query using the mailbox provider's search syntax (Gmail or Outlook). Required unless folderId is given.",
    ),
  folderId: z
    .string()
    .trim()
    .min(1)
    .optional()
    .describe(
      "Folder name or folderId, as list_folders shows them; prefer the name, since IDs are long and easy to miscopy. Limits results to that Outlook folder; without a query it lists the folder's mail, newest first.",
    ),
  maxResults: z.number().int().min(1).max(SEARCH_INBOX_MAX_RESULTS).optional(),
  pageToken: z.string().optional(),
};

export const readThreadInputShape = {
  ...mcpAccountSelectorShape,
  threadId: z.string().trim().min(1).describe("Thread ID from search_inbox."),
  maxMessages: z.number().int().min(1).max(READ_THREAD_MAX_MESSAGES).optional(),
};

export const createDraftInputShape = {
  ...mcpAccountSelectorShape,
  to: z
    .string()
    .trim()
    .min(1)
    .describe("Recipient email address or comma-separated addresses."),
  subject: z.string().trim().min(1).max(300),
  body: z
    .string()
    .trim()
    .min(1)
    .describe(
      "Draft body. Plain text is converted to HTML. This creates a mailbox draft and does not send.",
    ),
  replyToMessageId: z
    .string()
    .trim()
    .min(1)
    .optional()
    .describe(
      "messageId from search_inbox or read_thread. Set it to save the draft as a reply in that message's thread; omit it for a new message.",
    ),
};

export const updateThreadInputShape = {
  ...mcpAccountSelectorShape,
  threadId: z
    .string()
    .trim()
    .min(1)
    .describe("threadId from search_inbox or read_thread."),
  action: z
    .enum(["archive", "unarchive", "trash", "mark_read", "mark_unread"])
    .describe(
      "archive moves the thread out of the inbox (nothing is deleted); unarchive moves it back; trash moves it to Deleted Items (Outlook) or Trash (Gmail), where it can still be restored; mark_read and mark_unread set its read state.",
    ),
};

export const labelThreadInputShape = {
  ...mcpAccountSelectorShape,
  threadId: z
    .string()
    .trim()
    .min(1)
    .describe("threadId from search_inbox or read_thread."),
  labelName: z
    .string()
    .trim()
    .min(1)
    .describe(
      "Name of an existing label (Outlook category or Gmail label), matched exactly.",
    ),
  remove: z
    .boolean()
    .optional()
    .describe("true removes the label from the thread; omit to add it."),
};

export const runRulesInputShape = {
  ...mcpAccountSelectorShape,
  messageId: z
    .string()
    .trim()
    .min(1)
    .describe("messageId from search_inbox or read_thread."),
  apply: z
    .boolean()
    .describe(
      "false previews which rule would match and its actions without changing anything. true runs the matched rule's actions.",
    ),
};

export const listFoldersInputShape = mcpAccountSelectorShape;

const folderInputShape = {
  ...mcpAccountSelectorShape,
  folderId: z
    .string()
    .trim()
    .min(1)
    .describe(
      "Folder name or folderId, as list_folders shows them; prefer the name, since IDs are long and easy to miscopy.",
    ),
};

export const archiveFolderInputShape = folderInputShape;

export const moveFolderToInboxInputShape = folderInputShape;

export const deleteFolderInputShape = {
  ...folderInputShape,
  userConfirmed: z
    .boolean()
    .optional()
    .describe(
      "Set true only after the user has said yes to deleting this folder; the client may ask the user to approve the call. Required when the folder held email or subfolders until recently.",
    ),
};

export async function searchInboxForMcp({
  userId,
  query,
  folderId,
  maxResults,
  pageToken,
  emailAccountId,
  emailAddress,
  logger,
}: {
  userId: string;
  query?: string;
  folderId?: string;
  maxResults?: number;
  pageToken?: string;
  emailAccountId?: string;
  emailAddress?: string;
  logger: Logger;
}) {
  if (!query && !folderId)
    throw new Error("Give a query or folderId to search.");

  const emailAccount = await resolveMcpEmailAccount({
    userId,
    emailAccountId,
    emailAddress,
  });
  if (folderId && !isMicrosoftProvider(emailAccount.provider))
    throw new Error(
      "folderId is only available for Outlook accounts. Gmail uses labels; search with a label: query instead.",
    );

  const emailProvider = await createEmailProvider({
    emailAccountId: emailAccount.id,
    provider: emailAccount.provider,
    logger,
  });
  const folder = folderId
    ? findFolder(
        flattenOutlookFolders(await emailProvider.getFolders()),
        folderId,
      )
    : undefined;
  const searchResult = await emailProvider.searchMessages({
    query: query ?? "",
    folderId: folder?.id,
    maxResults: maxResults ?? 10,
    pageToken,
  });

  return {
    emailAccount,
    query,
    folderId: folder?.id,
    nextPageToken: searchResult.nextPageToken,
    hasMore: Boolean(searchResult.nextPageToken),
    messages: searchResult.messages.map(toSearchResultMessage),
  };
}

export async function readThreadForMcp({
  userId,
  threadId,
  maxMessages,
  emailAccountId,
  emailAddress,
  logger,
}: {
  userId: string;
  threadId: string;
  maxMessages?: number;
  emailAccountId?: string;
  emailAddress?: string;
  logger: Logger;
}) {
  const emailAccount = await resolveMcpEmailAccount({
    userId,
    emailAccountId,
    emailAddress,
  });
  const emailProvider = await createEmailProvider({
    emailAccountId: emailAccount.id,
    provider: emailAccount.provider,
    logger,
  });
  const messages = await emailProvider.getThreadMessages(threadId);
  const limit = maxMessages ?? 10;
  const selected = messages.slice(-limit);

  return {
    emailAccount,
    threadId,
    truncated: messages.length > selected.length,
    messages: selected.map(toThreadMessage),
  };
}

export async function createDraftForMcp({
  userId,
  to,
  subject,
  body,
  replyToMessageId,
  emailAccountId,
  emailAddress,
  logger,
}: {
  userId: string;
  to: string;
  subject: string;
  body: string;
  replyToMessageId?: string;
  emailAccountId?: string;
  emailAddress?: string;
  logger: Logger;
}) {
  const emailAccount = await resolveMcpEmailAccount({
    userId,
    emailAccountId,
    emailAddress,
  });
  const emailProvider = await createEmailProvider({
    emailAccountId: emailAccount.id,
    provider: emailAccount.provider,
    logger,
  });
  const draft = await emailProvider.createDraft({
    to,
    subject,
    messageHtml: bodyToHtml(body),
    replyToMessageId,
  });
  if (!draft.id) {
    throw new Error("Mailbox draft could not be created.");
  }

  return {
    emailAccount,
    draftId: draft.id,
    to,
    subject,
    replyToMessageId: replyToMessageId ?? null,
    sent: false,
  };
}

export async function updateThreadForMcp({
  userId,
  threadId,
  action,
  emailAccountId,
  emailAddress,
  logger,
}: {
  userId: string;
  threadId: string;
  action: "archive" | "unarchive" | "trash" | "mark_read" | "mark_unread";
  emailAccountId?: string;
  emailAddress?: string;
  logger: Logger;
}) {
  const emailAccount = await resolveMcpEmailAccount({
    userId,
    emailAccountId,
    emailAddress,
  });
  const emailProvider = await createEmailProvider({
    emailAccountId: emailAccount.id,
    provider: emailAccount.provider,
    logger,
  });

  if (action === "archive")
    await emailProvider.archiveThread(threadId, emailAccount.email);
  else if (action === "unarchive")
    await emailProvider.unarchiveThread(threadId);
  else if (action === "trash")
    await emailProvider.trashThread(threadId, emailAccount.email, "automation");
  else await emailProvider.markReadThread(threadId, action === "mark_read");

  return { emailAccount, threadId, action, applied: true };
}

export async function labelThreadForMcp({
  userId,
  threadId,
  labelName,
  remove,
  emailAccountId,
  emailAddress,
  logger,
}: {
  userId: string;
  threadId: string;
  labelName: string;
  remove?: boolean;
  emailAccountId?: string;
  emailAddress?: string;
  logger: Logger;
}) {
  const emailAccount = await resolveMcpEmailAccount({
    userId,
    emailAccountId,
    emailAddress,
  });
  const emailProvider = await createEmailProvider({
    emailAccountId: emailAccount.id,
    provider: emailAccount.provider,
    logger,
  });

  const label = await emailProvider.getLabelByName(labelName);
  if (!label) {
    const labels = await emailProvider.getLabels();
    const names = labels.map((l) => l.name).slice(0, LABEL_NAMES_IN_ERROR);
    throw new Error(
      `Label "${labelName}" not found. Existing labels: ${names.join(", ") || "none"}.`,
    );
  }

  if (remove) {
    await emailProvider.removeThreadLabel(threadId, label.id);
  } else {
    const messages = await emailProvider.getThreadMessages(threadId);
    for (const message of messages) {
      await emailProvider.labelMessage({
        messageId: message.id,
        labelId: label.id,
        labelName: label.name,
      });
    }
  }

  return {
    emailAccount,
    threadId,
    labelName: label.name,
    removed: Boolean(remove),
    applied: true,
  };
}

export async function listFoldersForMcp({
  userId,
  emailAccountId,
  emailAddress,
  logger,
}: {
  userId: string;
  emailAccountId?: string;
  emailAddress?: string;
  logger: Logger;
}) {
  const { emailAccount, folders } = await getOutlookFoldersForMcp({
    userId,
    emailAccountId,
    emailAddress,
    logger,
  });

  return {
    emailAccount,
    folders: folders.map((folder) => {
      const blocker = getFolderDeleteBlocker(folder);
      return {
        folderId: folder.id,
        path: folder.path,
        total: folder.totalItemCount ?? 0,
        unread: folder.unreadItemCount ?? 0,
        subfolders: subfolderCount(folder),
        ...(folder.systemType ? { systemType: folder.systemType } : {}),
        canDelete: blocker === null,
        ...(blocker ? { deleteBlockedReason: blocker } : {}),
      };
    }),
  };
}

export async function deleteFolderForMcp({
  userId,
  folderId,
  userConfirmed,
  emailAccountId,
  emailAddress,
  logger,
}: {
  userId: string;
  folderId: string;
  userConfirmed?: boolean;
  emailAccountId?: string;
  emailAddress?: string;
  logger: Logger;
}) {
  const { emailAccount, emailProvider, folders } =
    await getOutlookFoldersForMcp({
      userId,
      emailAccountId,
      emailAddress,
      logger,
    });

  const folder = findFolder(folders, folderId);

  // Enforced here rather than in the tool description, so a client that skips
  // list_folders still cannot delete mail along with a folder.
  const blocker = getFolderDeleteBlocker(folder);
  if (blocker) throw new Error(`Not deleted: "${folder.path}" ${blocker}.`);

  // A model can empty a folder and delete it in one go, so a folder that held
  // mail until recently needs the user's yes, which the client collects.
  if (
    userConfirmed !== true &&
    (await wasFolderRecentlyHeldMail({
      emailAccountId: emailAccount.id,
      folderId: folder.id,
      logger,
    }))
  )
    throw new Error(
      `Not deleted: "${folder.path}" held email or subfolders until recently (now moved). Ask the user first; delete it only after they say yes, by calling delete_folder again with userConfirmed: true.`,
    );

  const budget = await checkRateLimit({
    rule: {
      key: createRateLimitKey([
        "rate-limit",
        "mcp",
        "folder-delete",
        emailAccount.id,
      ]),
      ...FOLDER_DELETE_LIMIT,
    },
    logger,
  });
  if (budget.limited)
    throw new Error(
      `Not deleted: "${folder.path}". At most ${FOLDER_DELETE_LIMIT.limit} folders can be deleted every ${FOLDER_DELETE_LIMIT.windowSeconds / 60} minutes. Tell the user which folders are left; they can delete them in Outlook or ask again later.`,
    );

  await emailProvider.deleteFolder(folder.id);
  logger.info("Deleted empty folder via MCP", { folderId: folder.id });

  const parent = folders.find((candidate) =>
    candidate.childFolders.some((child) => child.id === folder.id),
  );
  if (parent)
    await markFolderHeldMail({
      emailAccountId: emailAccount.id,
      folderId: parent.id,
      logger,
    });

  return {
    emailAccount,
    folderId: folder.id,
    path: folder.path,
    deleted: true,
  };
}

export async function archiveFolderForMcp({
  userId,
  folderId,
  emailAccountId,
  emailAddress,
  logger,
}: {
  userId: string;
  folderId: string;
  emailAccountId?: string;
  emailAddress?: string;
  logger: Logger;
}) {
  const { emailAccount, emailProvider, folders } =
    await getOutlookFoldersForMcp({
      userId,
      emailAccountId,
      emailAddress,
      logger,
    });

  const folder = findFolder(folders, folderId);
  if (folder.systemType)
    throw new Error(`Not archived: "${folder.path}" is a system folder.`);

  const { messages, pageToken } = await collectFolderMessages(
    emailProvider,
    folder.id,
  );
  // Only this folder's messages move, so a thread's replies elsewhere stay put.
  const threads = new Map<string, string[]>();
  for (const message of messages) {
    threads.set(message.threadId, [
      ...(threads.get(message.threadId) ?? []),
      message.id,
    ]);
  }
  const collected = messages.length;

  const result = await emailProvider.bulkArchiveThreads(
    [...threads].map(([threadId, messageIds]) => ({ threadId, messageIds })),
    emailAccount.email,
  );
  const failed = new Set(result.failedThreadIds);
  const archived = [...threads]
    .filter(([threadId]) => !failed.has(threadId))
    .reduce((sum, [, messageIds]) => sum + messageIds.length, 0);
  logger.info("Archived folder via MCP", {
    folderId: folder.id,
    archived,
    failedThreads: failed.size,
  });
  if (archived > 0)
    await markFolderHeldMail({
      emailAccountId: emailAccount.id,
      folderId: folder.id,
      logger,
    });

  return {
    emailAccount,
    folderId: folder.id,
    path: folder.path,
    archived,
    failed: collected - archived,
    moreLeft: Boolean(pageToken),
  };
}

export async function moveFolderToInboxForMcp({
  userId,
  folderId,
  emailAccountId,
  emailAddress,
  logger,
}: {
  userId: string;
  folderId: string;
  emailAccountId?: string;
  emailAddress?: string;
  logger: Logger;
}) {
  const { emailAccount, emailProvider, folders } =
    await getOutlookFoldersForMcp({
      userId,
      emailAccountId,
      emailAddress,
      logger,
    });

  const folder = findFolder(folders, folderId);
  if (folder.systemType)
    throw new Error(`Not moved: "${folder.path}" is a system folder.`);

  const { messages, pageToken } = await collectFolderMessages(
    emailProvider,
    folder.id,
  );

  // unarchiveMessages moves each message to the inbox from any folder and
  // throws on the first failure, so a failed chunk is retried one by one to
  // count exactly what stayed behind.
  let moved = 0;
  for (const ids of chunk(
    messages.map((message) => message.id),
    MOVE_TO_INBOX_CHUNK_SIZE,
  )) {
    try {
      await emailProvider.unarchiveMessages(ids);
      moved += ids.length;
    } catch (error) {
      logger.warn("Moving a chunk to the inbox failed; retrying one by one", {
        error,
      });
      for (const id of ids) {
        try {
          await emailProvider.unarchiveMessages([id]);
          moved++;
        } catch {
          // Counted as failed below.
        }
      }
    }
  }
  logger.info("Moved folder to inbox via MCP", {
    folderId: folder.id,
    moved,
    failed: messages.length - moved,
  });
  if (moved > 0)
    await markFolderHeldMail({
      emailAccountId: emailAccount.id,
      folderId: folder.id,
      logger,
    });

  return {
    emailAccount,
    folderId: folder.id,
    path: folder.path,
    moved,
    failed: messages.length - moved,
    moreLeft: Boolean(pageToken),
  };
}

export async function runRulesForMcp({
  userId,
  messageId,
  apply,
  emailAccountId,
  emailAddress,
  logger,
}: {
  userId: string;
  messageId: string;
  apply: boolean;
  emailAccountId?: string;
  emailAddress?: string;
  logger: Logger;
}) {
  const emailAccount = await resolveMcpEmailAccount({
    userId,
    emailAccountId,
    emailAddress,
  });
  const results = await runRulesOnMessage({
    emailAccountId: emailAccount.id,
    provider: emailAccount.provider,
    messageId,
    isTest: !apply,
    logger: logger.with({ messageId }),
  });

  return {
    emailAccount,
    messageId,
    preview: !apply,
    results: results.map((result) => ({
      ruleName: result.rule?.name ?? null,
      reason: result.reason ?? null,
      status: result.status,
      alreadyProcessed: result.existing ?? false,
      actions: (result.actionItems ?? []).map((item) => ({
        type: item.type,
        label: item.label ?? null,
        folderName: item.folderName ?? null,
      })),
    })),
  };
}

function toSearchResultMessage(message: ParsedMessage) {
  return {
    messageId: message.id,
    threadId: message.threadId,
    subject: message.subject,
    from: message.headers.from,
    to: message.headers.to,
    date: message.date,
    snippet: message.snippet,
  };
}

function toThreadMessage(message: ParsedMessage) {
  const email = getEmailForLLM(message, { maxLength: READ_MESSAGE_MAX_LENGTH });
  return {
    messageId: message.id,
    threadId: message.threadId,
    from: email.from,
    to: email.to,
    cc: email.cc,
    subject: email.subject,
    date: email.date?.toISOString() ?? message.date,
    content: email.content,
    attachments: email.attachments,
  };
}

function bodyToHtml(body: string) {
  if (/<[a-z][\s\S]*>/i.test(body)) return body;
  return textToHtmlParagraphs(body);
}

async function getOutlookFoldersForMcp({
  userId,
  emailAccountId,
  emailAddress,
  logger,
}: {
  userId: string;
  emailAccountId?: string;
  emailAddress?: string;
  logger: Logger;
}) {
  const emailAccount = await resolveMcpEmailAccount({
    userId,
    emailAccountId,
    emailAddress,
  });
  if (!isMicrosoftProvider(emailAccount.provider))
    throw new Error(
      "Folders are only available for Outlook accounts. Gmail uses labels.",
    );

  const emailProvider = await createEmailProvider({
    emailAccountId: emailAccount.id,
    provider: emailAccount.provider,
    logger,
  });
  const folders = flattenOutlookFolders(await emailProvider.getFolders());

  return { emailAccount, emailProvider, folders };
}

type OutlookFlatFolder = ReturnType<typeof flattenOutlookFolders>[number];

function getFolderDeleteBlocker(folder: OutlookFolder) {
  if (folder.systemType) return "is a system folder";
  const total = folder.totalItemCount ?? 0;
  if (total > 0)
    return `still holds ${total} email${total === 1 ? "" : "s"}; archive or move them first`;
  const subfolders = subfolderCount(folder);
  if (subfolders > 0)
    return `has ${subfolders} subfolder${subfolders === 1 ? "" : "s"}; delete those first`;
  return null;
}

function subfolderCount(folder: OutlookFolder) {
  return Math.max(folder.childFolderCount ?? 0, folder.childFolders.length);
}

function findFolder(folders: OutlookFlatFolder[], nameOrId: string) {
  const byId = folders.find((folder) => folder.id === nameOrId);
  if (byId) return byId;

  const wanted = nameOrId.toLowerCase();
  for (const field of ["path", "displayName"] as const) {
    const matches = folders.filter(
      (folder) => folder[field].toLowerCase() === wanted,
    );
    if (matches.length === 1) return matches[0];
    if (matches.length > 1)
      throw new Error(
        `"${nameOrId}" matches more than one folder: ${matches.map((folder) => `"${folder.path}"`).join(", ")}. Use the full path.`,
      );
  }

  throw new Error(
    `Folder "${nameOrId}" not found. Call list_folders for current folder names.`,
  );
}

// Pages are collected before anything moves: Graph pages by offset, so moving
// mail out of the folder mid-listing would skip a page for every page moved.
async function collectFolderMessages(
  emailProvider: EmailProvider,
  folderId: string,
) {
  const messages: Array<{ id: string; threadId: string }> = [];
  let pageToken: string | undefined;
  do {
    const page = await emailProvider.searchMessages({
      query: "",
      folderId,
      maxResults: SEARCH_INBOX_MAX_RESULTS,
      pageToken,
    });
    for (const message of page.messages)
      messages.push({ id: message.id, threadId: message.threadId });
    pageToken = page.nextPageToken;
  } while (pageToken && messages.length < ARCHIVE_FOLDER_MAX_MESSAGES);

  return { messages, pageToken };
}
