"use server";

import { actionClient } from "@/utils/actions/safe-action";
import { reviewExecutedRuleBody } from "@/utils/actions/executed-rule.validation";
import prisma from "@/utils/prisma";
import { SafeError } from "@/utils/error";
import { createEmailProvider } from "@/utils/email/provider";
import { executeAct } from "@/utils/ai/choose-rule/execute";
import { ExecutedRuleStatus } from "@/generated/prisma/enums";

export const approveExecutedRuleAction = actionClient
  .metadata({ name: "approveExecutedRule" })
  .inputSchema(reviewExecutedRuleBody)
  .action(
    async ({
      ctx: { emailAccountId, emailAccount, userId, provider, logger },
      parsedInput: { executedRuleId },
    }) => {
      // Claiming the row first means a second click or a second tab cannot run the actions twice.
      const claimed = await prisma.executedRule.updateMany({
        where: {
          id: executedRuleId,
          emailAccountId,
          status: ExecutedRuleStatus.PENDING,
        },
        data: { status: ExecutedRuleStatus.APPLYING },
      });
      if (!claimed.count) {
        throw new SafeError("This item is no longer waiting for approval");
      }

      const executedRule = await prisma.executedRule.findUniqueOrThrow({
        where: { id: executedRuleId },
        include: { actionItems: true },
      });

      const emailProvider = await createEmailProvider({
        emailAccountId,
        provider,
        logger,
      });
      const message = await emailProvider.getMessage(executedRule.messageId);

      const status = await executeAct({
        client: emailProvider,
        executedRule,
        message,
        emailAccount: {
          email: emailAccount.email,
          id: emailAccountId,
          userId,
        },
        logger,
        approved: true,
      });

      return { status };
    },
  );

export const rejectExecutedRuleAction = actionClient
  .metadata({ name: "rejectExecutedRule" })
  .inputSchema(reviewExecutedRuleBody)
  .action(
    async ({ ctx: { emailAccountId }, parsedInput: { executedRuleId } }) => {
      const rejected = await prisma.executedRule.updateMany({
        where: {
          id: executedRuleId,
          emailAccountId,
          status: ExecutedRuleStatus.PENDING,
        },
        data: { status: ExecutedRuleStatus.REJECTED },
      });
      if (!rejected.count) {
        throw new SafeError("This item is no longer waiting for approval");
      }

      return { success: true };
    },
  );
