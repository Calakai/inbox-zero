import { beforeEach, describe, expect, it, vi } from "vitest";
import prisma from "@/utils/__mocks__/prisma";
import { ActionType, ExecutedRuleStatus } from "@/generated/prisma/enums";

const { createEmailProviderMock, executeActMock, getMessageMock } = vi.hoisted(
  () => ({
    createEmailProviderMock: vi.fn(),
    executeActMock: vi.fn(),
    getMessageMock: vi.fn(),
  }),
);

vi.mock("@/utils/prisma");
vi.mock("@/utils/auth", () => ({
  auth: vi.fn(async () => ({
    user: { id: "user-1", email: "user@example.com" },
  })),
}));
vi.mock("@/utils/email/provider", () => ({
  createEmailProvider: createEmailProviderMock,
}));
vi.mock("@/utils/ai/choose-rule/execute", () => ({
  executeAct: executeActMock,
}));

import {
  approveExecutedRuleAction,
  rejectExecutedRuleAction,
} from "@/utils/actions/executed-rule";

describe("approving held rule actions", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    prisma.emailAccount.findUnique.mockResolvedValue({
      email: "user@example.com",
      account: { userId: "user-1", provider: "google" },
    } as any);
    getMessageMock.mockResolvedValue({ id: "message-1", threadId: "thread-1" });
    createEmailProviderMock.mockResolvedValue({ getMessage: getMessageMock });
    executeActMock.mockResolvedValue(ExecutedRuleStatus.APPLIED);
  });

  it("runs the held actions once, for the owner's pending rule only", async () => {
    prisma.executedRule.updateMany.mockResolvedValue({ count: 1 } as any);
    prisma.executedRule.findUniqueOrThrow.mockResolvedValue({
      id: "er-1",
      messageId: "message-1",
      actionItems: [{ id: "a-1", type: ActionType.DELETE }],
    } as any);

    const result = await approveExecutedRuleAction("account-1", {
      executedRuleId: "er-1",
    });

    expect(prisma.executedRule.updateMany).toHaveBeenCalledWith({
      where: {
        id: "er-1",
        emailAccountId: "account-1",
        status: ExecutedRuleStatus.PENDING,
      },
      data: { status: ExecutedRuleStatus.APPLYING },
    });
    expect(executeActMock).toHaveBeenCalledWith(
      expect.objectContaining({ approved: true }),
    );
    expect(result?.data?.status).toBe(ExecutedRuleStatus.APPLIED);
  });

  it("runs nothing when the rule is not pending for this account", async () => {
    prisma.executedRule.updateMany.mockResolvedValue({ count: 0 } as any);

    const result = await approveExecutedRuleAction("account-1", {
      executedRuleId: "er-other",
    });

    expect(executeActMock).not.toHaveBeenCalled();
    expect(result?.serverError).toBe(
      "This item is no longer waiting for approval",
    );
  });

  it("rejecting marks the rule rejected and runs nothing", async () => {
    prisma.executedRule.updateMany.mockResolvedValue({ count: 1 } as any);

    await rejectExecutedRuleAction("account-1", { executedRuleId: "er-1" });

    expect(prisma.executedRule.updateMany).toHaveBeenCalledWith({
      where: {
        id: "er-1",
        emailAccountId: "account-1",
        status: ExecutedRuleStatus.PENDING,
      },
      data: { status: ExecutedRuleStatus.REJECTED },
    });
    expect(executeActMock).not.toHaveBeenCalled();
  });
});
