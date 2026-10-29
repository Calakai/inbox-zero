import { beforeEach, describe, expect, it, vi } from "vitest";
import { createTestLogger } from "@/__tests__/helpers";
import { ActionType } from "@/generated/prisma/enums";

const { updateRuleActionsMock, partialUpdateRuleMock, setRuleEnabledMock } =
  vi.hoisted(() => ({
    updateRuleActionsMock: vi.fn(),
    partialUpdateRuleMock: vi.fn(),
    setRuleEnabledMock: vi.fn(),
  }));

vi.mock("@/utils/rule/rule", async (importOriginal) => {
  const actual = await importOriginal<typeof import("@/utils/rule/rule")>();
  return {
    ...actual,
    updateRuleActions: updateRuleActionsMock,
    partialUpdateRule: partialUpdateRuleMock,
    setRuleEnabled: setRuleEnabledMock,
  };
});

vi.mock("@/utils/prisma");

import prisma from "@/utils/__mocks__/prisma";
import { confirmAssistantUpdateRuleForAccount } from "@/utils/actions/assistant-chat-confirmation";

describe("confirmAssistantUpdateRuleForAccount", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    prisma.rule.findUnique.mockResolvedValue(vendorBillingRule() as never);
    prisma.chatMessage.findFirst.mockResolvedValue({
      id: "cm-1",
      chatId: "chat-1",
      updatedAt: new Date("2026-10-29T00:00:00.000Z"),
      parts: [buildPendingUpdateRulePart()],
    } as never);
    prisma.chatMessage.updateMany.mockResolvedValue({ count: 1 } as never);
  });

  it("applies the confirmed risky update and persists the confirmed part", async () => {
    const result = await confirmAssistantUpdateRuleForAccount({
      chatId: "chat-1",
      chatMessageId: "cm-1",
      toolCallId: "tool-ur-1",
      emailAccountId: "ea_1",
      provider: "google",
      logger: createTestLogger(),
    });

    expect(result).toMatchObject({
      success: true,
      confirmationState: "confirmed",
      ruleId: "rule-id",
    });
    expect(updateRuleActionsMock).toHaveBeenCalledWith(
      expect.objectContaining({
        ruleId: "rule-id",
        emailAccountId: "ea_1",
        actions: expect.arrayContaining([
          expect.objectContaining({
            type: ActionType.FORWARD,
            fields: expect.objectContaining({ to: "someone@outside.example" }),
          }),
        ]),
      }),
    );

    const persistedParts = (
      prisma.chatMessage.updateMany.mock.calls[1][0] as {
        data: { parts: { output: Record<string, unknown> }[] };
      }
    ).data.parts;
    expect(persistedParts[0].output).toMatchObject({
      actionType: "update_rule",
      confirmationState: "confirmed",
      ruleId: "rule-id",
      confirmationResult: { ruleId: "rule-id" },
    });
  });

  it("does not apply the update twice once confirmed", async () => {
    prisma.chatMessage.findFirst.mockResolvedValue({
      id: "cm-1",
      chatId: "chat-1",
      updatedAt: new Date("2026-10-29T00:00:00.000Z"),
      parts: [
        buildPendingUpdateRulePart({
          confirmationState: "confirmed",
          confirmationResult: {
            ruleId: "rule-id",
            confirmedAt: "2026-10-29T00:01:00.000Z",
          },
        }),
      ],
    } as never);

    const result = await confirmAssistantUpdateRuleForAccount({
      chatId: "chat-1",
      chatMessageId: "cm-1",
      toolCallId: "tool-ur-1",
      emailAccountId: "ea_1",
      provider: "google",
      logger: createTestLogger(),
    });

    expect(result).toMatchObject({
      confirmationState: "confirmed",
      ruleId: "rule-id",
    });
    expect(updateRuleActionsMock).not.toHaveBeenCalled();
    expect(prisma.chatMessage.updateMany).not.toHaveBeenCalled();
  });
});

function buildPendingUpdateRulePart(output?: Record<string, unknown>) {
  return {
    type: "tool-updateRule",
    toolCallId: "tool-ur-1",
    state: "output-available",
    input: {
      ruleName: "Vendor Billing",
      updates: {
        actions: [
          {
            type: ActionType.LABEL,
            fields: { label: "Vendor Billing" },
            delayInMinutes: null,
          },
          {
            type: ActionType.FORWARD,
            fields: { to: "someone@outside.example" },
            delayInMinutes: null,
          },
        ],
      },
    },
    output: {
      success: true,
      actionType: "update_rule",
      requiresConfirmation: true,
      confirmationState: "pending",
      riskMessages: ["Medium Risk: Forward actions send a copy"],
      ruleId: "rule-id",
      originalName: "Vendor Billing",
      ...output,
    },
  };
}

function vendorBillingRule() {
  return {
    id: "rule-id",
    name: "Vendor Billing",
    enabled: true,
    updatedAt: new Date("2026-10-28T00:00:00.000Z"),
    organizationRuleId: null,
    emailAccount: { rulesRevision: 3 },
    instructions: "Billing notices.",
    from: "billing@vendor.example",
    to: null,
    subject: null,
    conditionalOperator: null,
    actions: [
      {
        type: ActionType.LABEL,
        content: null,
        label: "Vendor Billing",
        to: null,
        cc: null,
        bcc: null,
        subject: null,
        url: null,
        folderName: null,
        delayInMinutes: null,
      },
    ],
  };
}
