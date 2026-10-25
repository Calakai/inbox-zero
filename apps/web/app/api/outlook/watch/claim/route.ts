import { NextResponse } from "next/server";
import prisma from "@/utils/prisma";
import { hasCronSecret } from "@/utils/cron";
import { withError } from "@/utils/middleware";
import { captureException } from "@/utils/error";
import { checkMailboxClaim } from "@/utils/outlook/mailbox-claim";

// Read by `inbox status` to show whether this copy owns the mailbox.
export const GET = withError("outlook/watch/claim", async (request) => {
  if (!hasCronSecret(request)) {
    captureException(
      new Error("Unauthorized cron request: api/outlook/watch/claim"),
    );
    return new Response("Unauthorized", { status: 401 });
  }

  const emailAccounts = await prisma.emailAccount.findMany({
    where: { account: { provider: "microsoft" } },
    select: { id: true },
  });

  const results = [];
  for (const { id: emailAccountId } of emailAccounts) {
    const claim = await checkMailboxClaim({
      emailAccountId,
      logger: request.logger,
    });
    results.push({ emailAccountId, ...claim });
  }

  return NextResponse.json({ results });
});
