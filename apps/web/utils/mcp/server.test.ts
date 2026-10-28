import { beforeEach, describe, expect, it, vi } from "vitest";

const {
  handleRequest,
  mcpServerConstructor,
  registerTool,
  transportConstructor,
} = vi.hoisted(() => {
  const registerTool = vi.fn();
  const connect = vi.fn();
  const handleRequest = vi.fn();
  const mcpServerConstructor = vi.fn(
    class MockMcpServer {
      registerTool = registerTool;
      registerResource = vi.fn();
      connect = connect;
    },
  );
  const transportConstructor = vi.fn(
    class MockTransport {
      handleRequest = handleRequest;
    },
  );

  return {
    handleRequest,
    mcpServerConstructor,
    registerTool,
    transportConstructor,
  };
});

vi.mock("@modelcontextprotocol/sdk/server/mcp.js", () => ({
  McpServer: mcpServerConstructor,
}));
vi.mock(
  "@modelcontextprotocol/sdk/server/webStandardStreamableHttp.js",
  () => ({
    WebStandardStreamableHTTPServerTransport: transportConstructor,
  }),
);
vi.mock("@/app/api/user/stats/by-period/controller", () => ({
  getStatsByPeriod: vi.fn(),
}));
vi.mock("@/utils/stats/response-time/controller", () => ({
  getResponseTimeStats: vi.fn(),
}));
vi.mock("@/app/api/v1/rules/request", () => ({
  toRuleWriteInput: vi.fn(),
}));
vi.mock("@/app/api/v1/rules/serializers", () => ({
  apiRuleSelect: {},
  serializeRule: vi.fn(),
}));
vi.mock("@/app/api/v1/rules/validation", async (importActual) => {
  const actual =
    await importActual<typeof import("@/app/api/v1/rules/validation")>();
  return actual;
});
vi.mock("@/utils/branding", () => ({
  BRAND_NAME: "Inbox Zero",
}));
vi.mock("@/utils/email/provider", () => ({
  createEmailProvider: vi.fn(),
}));
vi.mock("@/utils/mcp/access", () => ({
  isMcpServerEnabledForUser: vi.fn(),
}));
vi.mock("@/utils/mcp/account-selection", () => ({
  listMcpEmailAccounts: vi.fn(),
  resolveMcpEmailAccount: vi.fn(),
}));
vi.mock("@/utils/mcp/mail-client-bundle", () => ({
  getMailClientResourceUri: vi.fn(),
  readMailClientHtml: vi.fn(),
}));
vi.mock("@/utils/prisma");
vi.mock("@/utils/premium/server", () => ({
  assertCanUseDigestsIfNeeded: vi.fn(),
}));
vi.mock("@/utils/rule/rule", () => ({
  createRule: vi.fn(),
  deleteRule: vi.fn(),
  updateRule: vi.fn(),
}));

import { assertCanUseDigestsIfNeeded } from "@/utils/premium/server";
import { toRuleWriteInput } from "@/app/api/v1/rules/request";
import { createRule, updateRule, deleteRule } from "@/utils/rule/rule";
import { resolveMcpEmailAccount } from "@/utils/mcp/account-selection";
import prisma from "@/utils/__mocks__/prisma";
import { handleMcpServerRequest } from "@/utils/mcp/server";
import { isMcpServerEnabledForUser } from "@/utils/mcp/access";
import { getMailClientResourceUri } from "@/utils/mcp/mail-client-bundle";
import { createEmailProvider } from "@/utils/email/provider";
import { SafeError } from "@/utils/error";

describe("mcp-server", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    handleRequest.mockResolvedValue(new Response("ok"));
    vi.mocked(isMcpServerEnabledForUser).mockResolvedValue(true);
  });

  it("returns 401 when the MCP access token has no user id", async () => {
    const response = await handleMcpServerRequest(
      new Request("http://localhost/mcp", { method: "POST" }),
      {} as never,
    );

    expect(response.status).toBe(401);
    expect(mcpServerConstructor).not.toHaveBeenCalled();
  });

  it("annotates mailbox tools and delegates the request to the MCP transport", async () => {
    const request = new Request("http://localhost/mcp", {
      method: "POST",
    });

    const response = await handleMcpServerRequest(request, {
      userId: "user_1",
      scopes: ["mcp:read", "mcp:write"],
    } as never);

    expect(
      registerTool.mock.calls.find(([name]) => name === "search_inbox")?.[1],
    ).toMatchObject({
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: true,
      },
    });
    expect(
      registerTool.mock.calls.find(([name]) => name === "create_draft")?.[1],
    ).toMatchObject({
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: true,
      },
    });
    expect(
      registerTool.mock.calls.find(([name]) => name === "delete_rule")?.[1],
    ).toMatchObject({
      annotations: { readOnlyHint: false, destructiveHint: true },
    });
    expect(handleRequest).toHaveBeenCalledWith(request);
    expect(response).toBeInstanceOf(Response);
  });

  it("registers mail client tools only when the mail client UI is installed", async () => {
    const request = () =>
      handleMcpServerRequest(
        new Request("http://localhost/mcp", { method: "POST" }),
        { userId: "user_1", scopes: ["mcp:read", "mcp:send"] } as never,
      );
    const registered = () => registerTool.mock.calls.map(([name]) => name);

    vi.mocked(getMailClientResourceUri).mockReturnValue(null);
    await request();
    expect(registered()).toContain("search_inbox");
    expect(registered()).not.toContain("open_mail");
    expect(registered()).not.toContain("send_mail");

    registerTool.mockClear();
    vi.mocked(getMailClientResourceUri).mockReturnValue("ui://mail/test.html");
    await request();
    expect(registered()).toContain("open_mail");
    expect(registered()).toContain("send_mail");
  });

  it("returns 403 when MCP access is disabled for the user", async () => {
    vi.mocked(isMcpServerEnabledForUser).mockResolvedValue(false);

    const response = await handleMcpServerRequest(
      new Request("http://localhost/mcp", { method: "POST" }),
      { userId: "user_1", scopes: [] } as never,
    );

    expect(response.status).toBe(403);
    expect(mcpServerConstructor).not.toHaveBeenCalled();
  });
});

describe("MCP tool permissions and rule writes", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(getMailClientResourceUri).mockReturnValue("ui://mail/test.html");
    vi.mocked(isMcpServerEnabledForUser).mockResolvedValue(true);
    vi.mocked(resolveMcpEmailAccount).mockResolvedValue({
      id: "account_1",
      email: "owner@example.com",
      name: null,
      provider: "google",
    });
    vi.mocked(toRuleWriteInput).mockReturnValue({
      name: "Digest",
      condition: {},
      actions: [{ type: "DIGEST", fields: null }],
      runOnThreads: false,
    } as never);
  });

  it("does not expand an existing draft-write grant into send access", async () => {
    const tool = await getTool("send_mail", ["mcp:read", "mcp:write"]);
    await expect(tool({})).rejects.toThrow("mcp:send");
    expect(resolveMcpEmailAccount).not.toHaveBeenCalled();
    const registration = registerTool.mock.calls.find(
      ([name]) => name === "send_mail",
    );
    expect(registration?.[1]._meta.ui.visibility).toEqual(["app"]);
  });

  it.each([
    "create_draft",
    "update_thread",
    "label_thread",
    "archive_folder",
    "move_folder_to_inbox",
    "mark_read",
    "delete_folder",
    "run_rules",
    "create_rule",
    "update_rule",
    "delete_rule",
    "save_mail_draft",
    "change_mail",
    "send_mail",
  ])("denies %s to a read-only client before accessing an inbox", async (name) => {
    const tool = await getTool(name, ["mcp:read"]);
    await expect(tool({ id: "rule_1", rule: {} })).rejects.toThrow(
      "Missing required permission",
    );
    expect(resolveMcpEmailAccount).not.toHaveBeenCalled();
    expect(createRule).not.toHaveBeenCalled();
    expect(updateRule).not.toHaveBeenCalled();
    expect(deleteRule).not.toHaveBeenCalled();
  });

  it.each([
    "list_email_accounts",
    "search_inbox",
    "read_thread",
    "list_folders",
    "list_rules",
    "get_rule",
    "get_stats_by_period",
    "get_response_time_stats",
  ])("denies %s without a read grant", async (name) => {
    const tool = await getTool(name, ["offline_access"]);
    await expect(tool({ id: "rule_1" })).rejects.toThrow(
      "Missing required permission",
    );
  });

  it("rejects creating a paid digest rule before persistence", async () => {
    vi.mocked(assertCanUseDigestsIfNeeded).mockRejectedValueOnce(
      new Error("Upgrade required"),
    );
    const tool = await getTool("create_rule", ["mcp:write"]);
    await expect(tool({ rule: digestRuleBody })).rejects.toThrow(
      "Upgrade required",
    );
    expect(createRule).not.toHaveBeenCalled();
  });

  it("passes existing actions to the digest check when replacing a rule", async () => {
    const actions = [{ type: "DIGEST" }];
    prisma.rule.findFirst.mockResolvedValueOnce({
      id: "rule_1",
      actions,
    } as never);
    vi.mocked(assertCanUseDigestsIfNeeded).mockRejectedValueOnce(
      new Error("Upgrade required"),
    );
    const tool = await getTool("update_rule", ["mcp:write"]);
    await expect(tool({ id: "rule_1", rule: digestRuleBody })).rejects.toThrow(
      "Upgrade required",
    );
    expect(assertCanUseDigestsIfNeeded).toHaveBeenCalledWith(
      "user_1",
      expect.any(Array),
      actions,
    );
    expect(updateRule).not.toHaveBeenCalled();
  });

  it.each([
    "update_rule",
    "delete_rule",
  ])("rejects %s for a rule outside the selected inbox", async (name) => {
    prisma.rule.findFirst.mockResolvedValueOnce(null);
    const tool = await getTool(name, ["mcp:write"]);
    await expect(tool({ id: "foreign_rule", rule: {} })).rejects.toThrow(
      "Rule not found",
    );
    expect(prisma.rule.findFirst).toHaveBeenCalledWith(
      expect.objectContaining({
        where: { id: "foreign_rule", emailAccountId: "account_1" },
      }),
    );
    expect(updateRule).not.toHaveBeenCalled();
    expect(deleteRule).not.toHaveBeenCalled();
  });
});

describe("MCP write tool summaries", () => {
  const unexpected = (title: string, message: string) =>
    `${title} failed: ${message}. Some changes may have gone through; check with search_inbox or list_folders before saying what changed.`;

  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(getMailClientResourceUri).mockReturnValue(null);
    vi.mocked(isMcpServerEnabledForUser).mockResolvedValue(true);
    vi.mocked(resolveMcpEmailAccount).mockResolvedValue({
      id: "account_1",
      email: "owner@example.com",
      name: null,
      provider: "google",
    });
    vi.mocked(toRuleWriteInput).mockReturnValue({
      name: "Digest",
      condition: {},
      actions: [{ type: "ARCHIVE", fields: null }],
      runOnThreads: false,
    } as never);
  });

  it("a write tool's JSON starts with summary", async () => {
    prisma.rule.findFirst.mockResolvedValueOnce({
      groupId: null,
      name: "Digest",
    } as never);
    const tool = await getTool("delete_rule", ["mcp:write"]);

    const result = (await tool({ id: "rule_1" })) as ToolResult;

    const json = JSON.parse(result.content[0].text);
    expect(Object.keys(json)[0]).toBe("summary");
    expect(json.summary).toBe('Deleted the rule "Digest".');
    expect(Object.keys(result.structuredContent)[0]).toBe("summary");
    expect(prisma.rule.findFirst).toHaveBeenCalledWith(
      expect.objectContaining({
        select: expect.objectContaining({ name: true }),
      }),
    );
  });

  it("summarizes a created and a replaced rule by its stored name", async () => {
    vi.mocked(createRule).mockResolvedValueOnce({ id: "rule_1" } as never);
    prisma.rule.findUnique.mockResolvedValueOnce({
      id: "rule_1",
      name: "Digest",
    } as never);
    const create = await getTool("create_rule", ["mcp:write"]);
    const created = (await create({ rule: digestRuleBody })) as ToolResult;
    expect(JSON.parse(created.content[0].text).summary).toBe(
      'Created the rule "Digest".',
    );

    prisma.rule.findFirst
      .mockResolvedValueOnce({ id: "rule_1", actions: [] } as never)
      .mockResolvedValueOnce({ id: "rule_1", name: "Digest" } as never);
    const update = await getTool("update_rule", ["mcp:write"]);
    const updated = (await update({
      id: "rule_1",
      rule: digestRuleBody,
    })) as ToolResult;
    expect(JSON.parse(updated.content[0].text).summary).toBe(
      'Replaced the rule "Digest" with the conditions and actions given.',
    );
  });

  it("an unexpected error becomes a sentence that claims no success", async () => {
    prisma.rule.findFirst.mockResolvedValueOnce({
      groupId: null,
      name: "Digest",
    } as never);
    vi.mocked(deleteRule).mockRejectedValueOnce(new Error("Connection reset."));
    const deleteTool = await getTool("delete_rule", ["mcp:write"]);
    await expect(deleteTool({ id: "rule_1" })).rejects.toHaveProperty(
      "message",
      unexpected("Delete rule", "Connection reset"),
    );

    vi.mocked(createEmailProvider).mockResolvedValueOnce({
      archiveThread: vi
        .fn()
        .mockRejectedValue(new Error("Service unavailable")),
    } as never);
    const updateThread = await getTool("update_thread", ["mcp:write"]);
    await expect(
      updateThread({ threadId: "thread-1", action: "archive" }),
    ).rejects.toHaveProperty(
      "message",
      unexpected("Update thread", "Service unavailable"),
    );
  });

  it.each([
    {
      name: "a missing permission",
      tool: "delete_rule",
      scopes: ["mcp:read"],
      args: { id: "rule_1" },
      message: "Missing required permission: mcp:write",
    },
    {
      name: "a rule outside the inbox",
      tool: "delete_rule",
      scopes: ["mcp:write"],
      args: { id: "foreign_rule" },
      message: "Rule not found for the selected email account.",
    },
    {
      name: "a folder tool on Gmail",
      tool: "archive_folder",
      scopes: ["mcp:write"],
      args: { folderId: "Receipts" },
      message:
        "Folders are only available for Outlook accounts. Gmail uses labels.",
    },
  ])("a refusal's text reaches the client unchanged: $name", async ({
    tool: name,
    scopes,
    args,
    message,
  }) => {
    prisma.rule.findFirst.mockResolvedValueOnce(null);
    const tool = await getTool(name, scopes);
    await expect(tool(args)).rejects.toHaveProperty("message", message);
  });

  it("a refusal from a paid-plan check reaches the client unchanged", async () => {
    vi.mocked(assertCanUseDigestsIfNeeded).mockRejectedValueOnce(
      new SafeError("Digests are available on the Plus plan.", 403),
    );
    const tool = await getTool("create_rule", ["mcp:write"]);
    await expect(tool({ rule: digestRuleBody })).rejects.toHaveProperty(
      "message",
      "Digests are available on the Plus plan.",
    );
  });

  it("invalid rule input is a refusal, not a partial failure", async () => {
    const tool = await getTool("create_rule", ["mcp:write"]);
    const error = await tool({ rule: {} }).catch((e: Error) => e);
    expect(error).toBeInstanceOf(Error);
    expect((error as Error).message).not.toContain("failed:");
    expect(createRule).not.toHaveBeenCalled();
  });
});

type ToolResult = {
  content: { text: string }[];
  structuredContent: Record<string, unknown>;
};

const digestRuleBody = {
  name: "Digest",
  condition: { aiInstructions: "Match digest emails" },
  actions: [{ type: "DIGEST" }],
};

async function getTool(name: string, scopes: string[]) {
  await handleMcpServerRequest(
    new Request("http://localhost/mcp", { method: "POST" }),
    { userId: "user_1", scopes },
  );
  const registration = registerTool.mock.calls.find(
    ([toolName]) => toolName === name,
  );
  if (!registration) throw new Error(`Tool not registered: ${name}`);
  return registration[2] as (args: Record<string, unknown>) => Promise<unknown>;
}
