import { assertCanUseDigestsIfNeeded } from "@/utils/premium/server";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { WebStandardStreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/webStandardStreamableHttp.js";
import { z } from "zod";
import { getStatsByPeriod } from "@/app/api/user/stats/by-period/controller";
import { getResponseTimeStats } from "@/utils/stats/response-time/controller";
import { toRuleWriteInput } from "@/app/api/v1/rules/request";
import { apiRuleSelect, serializeRule } from "@/app/api/v1/rules/serializers";
import { ruleRequestBodySchema } from "@/app/api/v1/rules/validation";
import { BRAND_NAME } from "@/utils/branding";
import { createEmailProvider } from "@/utils/email/provider";
import { SafeError } from "@/utils/error";
import { createScopedLogger } from "@/utils/logger";
import prisma from "@/utils/prisma";
import { createRule, deleteRule, updateRule } from "@/utils/rule/rule";
import {
  listMcpEmailAccounts,
  resolveMcpEmailAccount,
} from "@/utils/mcp/account-selection";
import type { MCP_SCOPES } from "@/utils/mcp/config";
import { isMcpServerEnabledForUser } from "@/utils/mcp/access";
import {
  registerMailClient,
  getMailPermissions,
} from "@/utils/mcp/mail-client-ui";
import {
  archiveFolderForMcp,
  archiveFolderInputShape,
  createDraftForMcp,
  createDraftInputShape,
  deleteFolderForMcp,
  deleteFolderInputShape,
  labelThreadForMcp,
  labelThreadInputShape,
  listFoldersForMcp,
  listFoldersInputShape,
  markReadForMcp,
  markReadInputShape,
  moveFolderToInboxForMcp,
  moveFolderToInboxInputShape,
  mcpAccountSelectorShape,
  readThreadForMcp,
  readThreadInputShape,
  runRulesForMcp,
  runRulesInputShape,
  searchInboxForMcp,
  searchInboxInputShape,
  updateThreadForMcp,
  updateThreadInputShape,
} from "@/utils/mcp/email-tools";
import {
  createRuleSummary,
  deleteRuleSummary,
  updateRuleSummary,
} from "@/utils/mcp/summaries";

const logger = createScopedLogger("mcp-server");
type ToolResultData = Record<string, unknown>;
type WriteToolResultData = ToolResultData & { summary: string };
const readOnlyAnnotations = {
  readOnlyHint: true,
  destructiveHint: false,
  openWorldHint: false,
};
const ruleWriteAnnotations = {
  readOnlyHint: false,
  destructiveHint: true,
  openWorldHint: true,
};
const destructiveAnnotations = {
  readOnlyHint: false,
  destructiveHint: true,
  openWorldHint: false,
};
const mailboxReadAnnotations = {
  readOnlyHint: true,
  destructiveHint: false,
  openWorldHint: true,
};
const mailboxWriteAnnotations = {
  readOnlyHint: false,
  destructiveHint: false,
  openWorldHint: true,
};

export async function handleMcpServerRequest(
  request: Request,
  session: { userId: string; scopes: string[]; clientId?: string },
) {
  const userId = session.userId;
  if (!userId) {
    return new Response(null, { status: 401 });
  }

  if (!(await isMcpServerEnabledForUser(userId))) {
    return new Response(null, { status: 403 });
  }

  const toolLogger = logger.with({
    userId,
    ...(session.clientId ? { clientId: session.clientId } : {}),
  });

  const server = new McpServer(
    { name: `${BRAND_NAME} MCP`, version: "1.0.0" },
    {
      instructions:
        "Use list_email_accounts when targeting a specific inbox. Search and read tools return mail; create_draft saves a mailbox draft and never sends; update_thread archives or marks threads; label_thread adds or removes a label; list_folders, archive_folder, move_folder_to_inbox and delete_folder manage Outlook folders; mark_read marks a whole Outlook folder or one sender's mail read; run_rules runs the user's rules on one email. Rule and stats tools accept emailAccountId or emailAddress and default to the first linked account.",
    },
  );

  const runTool =
    (
      name: string,
      required: (typeof MCP_SCOPES)[number],
      handler: (args: Record<string, unknown>) => Promise<ToolResultData>,
    ) =>
    async (args: Record<string, unknown>) => {
      try {
        assertMcpScope(session.scopes, required);
        const data = await handler(args);
        toolLogger.info("MCP tool call", {
          tool: name,
          ...emailAccountLogFields(data),
        });
        return createToolResult(data);
      } catch (error) {
        toolLogger.warn("MCP tool failed", {
          tool: name,
          error: error instanceof Error ? error.message : error,
        });
        throw error;
      }
    };

  // A write tool must say what it changed in one sentence, first, so a model
  // that repeats it (or a client that truncates the JSON) keeps the real counts.
  const runWriteTool =
    (
      name: string,
      title: string,
      handler: (args: Record<string, unknown>) => Promise<WriteToolResultData>,
    ) =>
    async (args: Record<string, unknown>) => {
      try {
        assertMcpScope(session.scopes, "mcp:write");
        const { summary, ...data } = await handler(args);
        toolLogger.info("MCP tool call", {
          tool: name,
          ...emailAccountLogFields(data),
        });
        return createToolResult({ summary, ...data });
      } catch (error) {
        toolLogger.warn("MCP tool failed", {
          tool: name,
          error: error instanceof Error ? error.message : error,
        });
        throw toWriteToolError(title, error);
      }
    };

  server.registerTool(
    "list_email_accounts",
    {
      title: "List email accounts",
      description: "List the inbox accounts linked to the authenticated user.",
      annotations: { ...readOnlyAnnotations, title: "List email accounts" },
    },
    runTool("list_email_accounts", "mcp:read", async () => ({
      accounts: await listMcpEmailAccounts(userId),
      permissions: getMailPermissions(session.scopes),
    })),
  );

  server.registerTool(
    "search_inbox",
    {
      title: "Search inbox",
      description:
        "Search one inbox, or one Outlook folder by folder name or folderId, and return message metadata and snippets. Does not return full message bodies, send email, or change mail. Results may carry amounts (currency amounts as written in the email) and dueDate (a stated due or payment date, as written); each is left out when the email states none.",
      inputSchema: searchInboxInputShape,
      annotations: { ...mailboxReadAnnotations, title: "Search inbox" },
    },
    runTool("search_inbox", "mcp:read", async (args) =>
      searchInboxForMcp({
        userId,
        query: typeof args.query === "string" ? args.query : undefined,
        folderId: typeof args.folderId === "string" ? args.folderId : undefined,
        maxResults:
          typeof args.maxResults === "number" ? args.maxResults : undefined,
        pageToken:
          typeof args.pageToken === "string" ? args.pageToken : undefined,
        emailAccountId:
          typeof args.emailAccountId === "string"
            ? args.emailAccountId
            : undefined,
        emailAddress:
          typeof args.emailAddress === "string" ? args.emailAddress : undefined,
        logger: toolLogger,
      }),
    ),
  );

  server.registerTool(
    "read_thread",
    {
      title: "Read thread",
      description:
        "Read messages from one inbox thread by thread ID. Returns plain-text bodies truncated per message.",
      inputSchema: readThreadInputShape,
      annotations: { ...mailboxReadAnnotations, title: "Read thread" },
    },
    runTool("read_thread", "mcp:read", async (args) =>
      readThreadForMcp({
        userId,
        threadId: String(args.threadId),
        maxMessages:
          typeof args.maxMessages === "number" ? args.maxMessages : undefined,
        emailAccountId:
          typeof args.emailAccountId === "string"
            ? args.emailAccountId
            : undefined,
        emailAddress:
          typeof args.emailAddress === "string" ? args.emailAddress : undefined,
        logger: toolLogger,
      }),
    ),
  );

  server.registerTool(
    "create_draft",
    {
      title: "Create draft",
      description:
        "Create a mailbox draft with a recipient, subject, and body. Pass replyToMessageId to draft a reply inside an existing thread. Saves the draft for review without sending email.",
      inputSchema: createDraftInputShape,
      annotations: { ...mailboxWriteAnnotations, title: "Create draft" },
    },
    runWriteTool("create_draft", "Create draft", async (args) =>
      createDraftForMcp({
        userId,
        to: String(args.to),
        subject: String(args.subject),
        body: String(args.body),
        replyToMessageId:
          typeof args.replyToMessageId === "string"
            ? args.replyToMessageId
            : undefined,
        emailAccountId:
          typeof args.emailAccountId === "string"
            ? args.emailAccountId
            : undefined,
        emailAddress:
          typeof args.emailAddress === "string" ? args.emailAddress : undefined,
        logger: toolLogger,
      }),
    ),
  );

  server.registerTool(
    "update_thread",
    {
      title: "Update thread",
      description:
        "Archive, unarchive, trash, mark read or mark unread one inbox thread by thread ID. Archiving only moves the thread out of the inbox and can be undone with unarchive. Trash moves it to Deleted Items or Trash, where it can still be restored; no tool permanently deletes mail or empties the trash. Nothing is sent.",
      inputSchema: updateThreadInputShape,
      annotations: { ...mailboxWriteAnnotations, title: "Update thread" },
    },
    runWriteTool("update_thread", "Update thread", async (args) =>
      updateThreadForMcp({
        userId,
        threadId: String(args.threadId),
        action: args.action as
          | "archive"
          | "unarchive"
          | "trash"
          | "mark_read"
          | "mark_unread",
        ...accountSelector(args),
        logger: toolLogger,
      }),
    ),
  );

  server.registerTool(
    "label_thread",
    {
      title: "Label thread",
      description:
        "Add or remove an existing label (an Outlook category or Gmail label) on every message in one thread. It never creates or deletes labels; an unknown name fails and lists the existing labels.",
      inputSchema: labelThreadInputShape,
      annotations: { ...mailboxWriteAnnotations, title: "Label thread" },
    },
    runWriteTool("label_thread", "Label thread", async (args) =>
      labelThreadForMcp({
        userId,
        threadId: String(args.threadId),
        labelName: String(args.labelName),
        remove: args.remove === true,
        ...accountSelector(args),
        logger: toolLogger,
      }),
    ),
  );

  server.registerTool(
    "list_folders",
    {
      title: "List folders",
      description:
        "List the Outlook account's mail folders with their path, email count, unread count and subfolder count. Other folder tools accept a folder's name or path from here as well as its folderId. canDelete marks the folders delete_folder will accept; when false, deleteBlockedReason says why. Outlook only; Gmail has labels instead of folders.",
      inputSchema: listFoldersInputShape,
      annotations: { ...mailboxReadAnnotations, title: "List folders" },
    },
    runTool("list_folders", "mcp:read", async (args) =>
      listFoldersForMcp({
        userId,
        ...accountSelector(args),
        logger: toolLogger,
      }),
    ),
  );

  server.registerTool(
    "archive_folder",
    {
      title: "Archive folder",
      description:
        "Move every email in one Outlook folder to Archive in a single call, by folder name or folderId from list_folders. Nothing is deleted and the folder stays; use this to empty a folder before delete_folder. Refuses system folders. Returns how many emails moved and how many failed.",
      inputSchema: archiveFolderInputShape,
      annotations: { ...mailboxWriteAnnotations, title: "Archive folder" },
    },
    runWriteTool("archive_folder", "Archive folder", async (args) =>
      archiveFolderForMcp({
        userId,
        folderId: String(args.folderId),
        ...accountSelector(args),
        logger: toolLogger,
      }),
    ),
  );

  server.registerTool(
    "mark_read",
    {
      title: "Mark read",
      description:
        "Mark every email in one Outlook folder (by name; Inbox allowed) or from one sender read in a single call; read: false marks them unread. Never moves or deletes mail. Use this instead of update_thread per thread when the user asks to mark a folder or a sender's mail read. Returns how many changed, how many failed and whether more are left (call again if moreLeft is true).",
      inputSchema: markReadInputShape,
      annotations: { ...mailboxWriteAnnotations, title: "Mark read" },
    },
    runWriteTool("mark_read", "Mark read", async (args) =>
      markReadForMcp({
        userId,
        folderId: typeof args.folderId === "string" ? args.folderId : undefined,
        sender: typeof args.sender === "string" ? args.sender : undefined,
        read: args.read !== false,
        ...accountSelector(args),
        logger: toolLogger,
      }),
    ),
  );

  server.registerTool(
    "move_folder_to_inbox",
    {
      title: "Move folder to inbox",
      description:
        "Move every email in one Outlook folder back to the Inbox in a single call, by folder name or folderId from list_folders. Nothing is deleted and the folder stays; once it is empty, delete_folder can remove it. Refuses system folders. Returns how many emails moved and how many failed.",
      inputSchema: moveFolderToInboxInputShape,
      annotations: {
        ...mailboxWriteAnnotations,
        title: "Move folder to inbox",
      },
    },
    runWriteTool("move_folder_to_inbox", "Move folder to inbox", async (args) =>
      moveFolderToInboxForMcp({
        userId,
        folderId: String(args.folderId),
        ...accountSelector(args),
        logger: toolLogger,
      }),
    ),
  );

  server.registerTool(
    "delete_folder",
    {
      title: "Delete folder",
      description:
        "Delete one empty Outlook mail folder by folder name or folderId from list_folders. Refuses system folders (Inbox, Sent, Drafts, Archive, Deleted Items, Junk) and any folder that still holds email or subfolders; empty it first with archive_folder or move_folder_to_inbox. At most 2 folders can be deleted every 10 minutes, so delete only the folders the user named. A folder that held email or subfolders in the last day (listed, searched, emptied, or its subfolder deleted) is refused unless the user has said yes and you pass userConfirmed: true.",
      inputSchema: deleteFolderInputShape,
      annotations: {
        ...mailboxWriteAnnotations,
        destructiveHint: true,
        title: "Delete folder",
      },
    },
    runWriteTool("delete_folder", "Delete folder", async (args) =>
      deleteFolderForMcp({
        userId,
        folderId: String(args.folderId),
        userConfirmed: args.userConfirmed === true,
        ...accountSelector(args),
        logger: toolLogger,
      }),
    ),
  );

  server.registerTool(
    "run_rules",
    {
      title: "Run rules",
      description:
        "Run the user's enabled rules on one email by message ID and report which rule matches, why, and its actions. apply=false is a preview and changes nothing. apply=true runs the matched rule's actions (for example labelling, archiving or drafting), the same as when new mail arrives; an email that was already processed returns its earlier result instead of running again. Prefer apply=false unless the user asked to apply the rules.",
      inputSchema: runRulesInputShape,
      annotations: { ...ruleWriteAnnotations, title: "Run rules" },
    },
    runWriteTool("run_rules", "Run rules", async (args) =>
      runRulesForMcp({
        userId,
        messageId: String(args.messageId),
        apply: args.apply === true,
        ...accountSelector(args),
        logger: toolLogger,
      }),
    ),
  );

  server.registerTool(
    "list_rules",
    {
      title: "List rules",
      description: "List automation rules for one inbox account.",
      inputSchema: mcpAccountSelectorShape,
      annotations: { ...readOnlyAnnotations, title: "List rules" },
    },
    runTool("list_rules", "mcp:read", async (args) => {
      const emailAccount = await resolveMcpEmailAccount({
        userId,
        ...accountSelector(args),
      });
      const rules = await prisma.rule.findMany({
        where: { emailAccountId: emailAccount.id },
        select: apiRuleSelect,
        orderBy: { createdAt: "asc" },
      });

      return {
        emailAccount,
        rules: rules.map(serializeRule),
      };
    }),
  );

  server.registerTool(
    "get_rule",
    {
      title: "Get rule",
      description: "Get one automation rule by ID for one inbox account.",
      inputSchema: {
        ...mcpAccountSelectorShape,
        id: z.string(),
      },
      annotations: { ...readOnlyAnnotations, title: "Get rule" },
    },
    runTool("get_rule", "mcp:read", async (args) => {
      const emailAccount = await resolveMcpEmailAccount({
        userId,
        ...accountSelector(args),
      });
      const rule = await prisma.rule.findFirst({
        where: { id: String(args.id), emailAccountId: emailAccount.id },
        select: apiRuleSelect,
      });

      if (!rule) {
        throw new SafeError("Rule not found for the selected email account.");
      }

      return {
        emailAccount,
        rule: serializeRule(rule),
      };
    }),
  );

  server.registerTool(
    "create_rule",
    {
      title: "Create rule",
      description: "Create an automation rule for one inbox account.",
      inputSchema: {
        ...mcpAccountSelectorShape,
        rule: ruleRequestBodySchema,
      },
      annotations: { ...ruleWriteAnnotations, title: "Create rule" },
    },
    runWriteTool("create_rule", "Create rule", async (args) => {
      const emailAccount = await resolveMcpEmailAccount({
        userId,
        ...accountSelector(args),
      });
      const ruleInput = toRuleWriteInput(
        ruleRequestBodySchema.parse(args.rule),
      );
      const scopedLogger = toolLogger.with({
        emailAccountId: emailAccount.id,
      });

      await assertCanUseDigestsIfNeeded(userId, ruleInput.actions);

      const createdRule = await createRule({
        result: {
          name: ruleInput.name,
          condition: ruleInput.condition,
          actions: ruleInput.actions,
        },
        emailAccountId: emailAccount.id,
        provider: emailAccount.provider,
        runOnThreads: ruleInput.runOnThreads,
        logger: scopedLogger,
      });

      const storedRule = await prisma.rule.findUnique({
        where: { id: createdRule.id },
        select: apiRuleSelect,
      });

      if (!storedRule) {
        throw new Error("Created rule could not be loaded.");
      }

      return {
        summary: createRuleSummary(storedRule.name),
        emailAccount,
        rule: serializeRule(storedRule),
      };
    }),
  );

  server.registerTool(
    "update_rule",
    {
      title: "Update rule",
      description: "Replace an automation rule for one inbox account.",
      inputSchema: {
        ...mcpAccountSelectorShape,
        id: z.string(),
        rule: ruleRequestBodySchema,
      },
      annotations: { ...ruleWriteAnnotations, title: "Update rule" },
    },
    runWriteTool("update_rule", "Update rule", async (args) => {
      const emailAccount = await resolveMcpEmailAccount({
        userId,
        ...accountSelector(args),
      });
      const existingRule = await prisma.rule.findFirst({
        where: { id: String(args.id), emailAccountId: emailAccount.id },
        select: { id: true, actions: { select: { type: true } } },
      });

      if (!existingRule) {
        throw new SafeError("Rule not found for the selected email account.");
      }

      const ruleInput = toRuleWriteInput(
        ruleRequestBodySchema.parse(args.rule),
      );
      const scopedLogger = toolLogger.with({
        emailAccountId: emailAccount.id,
        ruleId: String(args.id),
      });

      await assertCanUseDigestsIfNeeded(
        userId,
        ruleInput.actions,
        existingRule.actions,
      );

      await updateRule({
        ruleId: String(args.id),
        result: {
          name: ruleInput.name,
          condition: ruleInput.condition,
          actions: ruleInput.actions,
        },
        emailAccountId: emailAccount.id,
        provider: emailAccount.provider,
        logger: scopedLogger,
        runOnThreads: ruleInput.runOnThreads,
      });

      const updatedRule = await prisma.rule.findFirst({
        where: { id: String(args.id), emailAccountId: emailAccount.id },
        select: apiRuleSelect,
      });

      if (!updatedRule) {
        throw new Error("Updated rule could not be loaded.");
      }

      return {
        summary: updateRuleSummary(updatedRule.name),
        emailAccount,
        rule: serializeRule(updatedRule),
      };
    }),
  );

  server.registerTool(
    "delete_rule",
    {
      title: "Delete rule",
      description: "Delete an automation rule for one inbox account.",
      inputSchema: {
        ...mcpAccountSelectorShape,
        id: z.string(),
      },
      annotations: { ...destructiveAnnotations, title: "Delete rule" },
    },
    runWriteTool("delete_rule", "Delete rule", async (args) => {
      const emailAccount = await resolveMcpEmailAccount({
        userId,
        ...accountSelector(args),
      });
      const existingRule = await prisma.rule.findFirst({
        where: { id: String(args.id), emailAccountId: emailAccount.id },
        select: { groupId: true, name: true },
      });

      if (!existingRule) {
        throw new SafeError("Rule not found for the selected email account.");
      }

      await deleteRule({
        emailAccountId: emailAccount.id,
        ruleId: String(args.id),
        groupId: existingRule.groupId,
      });

      return {
        summary: deleteRuleSummary(existingRule.name),
        deleted: true,
        emailAccount,
        id: String(args.id),
      };
    }),
  );

  server.registerTool(
    "get_stats_by_period",
    {
      title: "Get stats by period",
      description: "Get email statistics grouped by day, week, month, or year.",
      inputSchema: {
        ...mcpAccountSelectorShape,
        period: z.enum(["day", "week", "month", "year"]).optional(),
        fromDate: z.number().int().optional(),
        toDate: z.number().int().optional(),
      },
      annotations: { ...readOnlyAnnotations, title: "Get stats by period" },
    },
    runTool("get_stats_by_period", "mcp:read", async (args) => {
      const emailAccount = await resolveMcpEmailAccount({
        userId,
        ...accountSelector(args),
      });
      const result = await getStatsByPeriod({
        period:
          args.period === "day" ||
          args.period === "week" ||
          args.period === "month" ||
          args.period === "year"
            ? args.period
            : "week",
        fromDate: typeof args.fromDate === "number" ? args.fromDate : undefined,
        toDate: typeof args.toDate === "number" ? args.toDate : undefined,
        emailAccountId: emailAccount.id,
      });

      return {
        emailAccount,
        ...result,
      };
    }),
  );

  server.registerTool(
    "get_response_time_stats",
    {
      title: "Get response time stats",
      description: "Get response time analytics for one inbox account.",
      inputSchema: {
        ...mcpAccountSelectorShape,
        fromDate: z.number().int().optional(),
        toDate: z.number().int().optional(),
      },
      annotations: { ...readOnlyAnnotations, title: "Get response time stats" },
    },
    runTool("get_response_time_stats", "mcp:read", async (args) => {
      const emailAccount = await resolveMcpEmailAccount({
        userId,
        ...accountSelector(args),
      });
      const scopedLogger = toolLogger.with({
        emailAccountId: emailAccount.id,
      });
      const emailProvider = await createEmailProvider({
        emailAccountId: emailAccount.id,
        provider: emailAccount.provider,
        logger: scopedLogger,
      });
      const result = await getResponseTimeStats({
        fromDate: typeof args.fromDate === "number" ? args.fromDate : undefined,
        toDate: typeof args.toDate === "number" ? args.toDate : undefined,
        emailAccountId: emailAccount.id,
        emailProvider,
        logger: scopedLogger,
      });

      return {
        emailAccount,
        ...serializeResponseTimeStats(result),
      };
    }),
  );

  registerMailClient(server, session, toolLogger);

  const transport = new WebStandardStreamableHTTPServerTransport({
    sessionIdGenerator: undefined,
    enableJsonResponse: true,
  });

  await server.connect(transport);

  return transport.handleRequest(request);
}

function createToolResult(data: ToolResultData) {
  return {
    content: [{ type: "text" as const, text: JSON.stringify(data, null, 2) }],
    structuredContent: data,
  };
}

function serializeResponseTimeStats(
  result: Awaited<ReturnType<typeof getResponseTimeStats>>,
) {
  return {
    ...result,
    trend: result.trend.map((entry) => ({
      ...entry,
      periodDate: entry.periodDate.toISOString(),
    })),
  };
}

function assertMcpScope(
  scopes: string[],
  required: (typeof MCP_SCOPES)[number],
) {
  if (!scopes.includes(required))
    throw new SafeError(`Missing required permission: ${required}`);
}

// Refusals already say nothing changed, so they reach the client as written.
// Anything else may have failed partway, so the model must not report success.
function toWriteToolError(title: string, error: unknown) {
  if (error instanceof SafeError || error instanceof z.ZodError) return error;
  const message = (error instanceof Error ? error.message : String(error))
    .trim()
    .replace(/\.+$/, "");
  return new Error(
    `${title} failed: ${message}. Some changes may have gone through; check with search_inbox or list_folders before saying what changed.`,
  );
}

function accountSelector(args: Record<string, unknown>) {
  return {
    emailAccountId:
      typeof args.emailAccountId === "string" ? args.emailAccountId : undefined,
    emailAddress:
      typeof args.emailAddress === "string" ? args.emailAddress : undefined,
  };
}

function emailAccountLogFields(data: ToolResultData) {
  const emailAccount = data.emailAccount;
  if (
    !emailAccount ||
    typeof emailAccount !== "object" ||
    !("id" in emailAccount) ||
    typeof emailAccount.id !== "string"
  ) {
    return {};
  }
  return { emailAccountId: emailAccount.id };
}
