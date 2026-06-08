import { NextResponse } from "next/server";
import { hasCronSecret, hasPostCronSecret } from "@/utils/cron";
import { withError } from "@/utils/middleware";
import { captureException } from "@/utils/error";
import prisma from "@/utils/prisma";
import { createEmailProvider } from "@/utils/email/provider";
import type { Logger } from "@/utils/logger";

export const maxDuration = 300;

const MAX_MESSAGES_PER_ACCOUNT = 20;

// IMAP has no push notifications, so new mail is picked up by polling here.
export const GET = withError("cron/imap-sync", async (request) => {
  if (!hasCronSecret(request)) {
    captureException(
      new Error("Unauthorized cron request: api/cron/imap-sync"),
    );
    return new Response("Unauthorized", { status: 401 });
  }
  return syncImapAccounts(request.logger);
});

export const POST = withError("cron/imap-sync", async (request) => {
  if (!(await hasPostCronSecret(request))) {
    captureException(
      new Error("Unauthorized cron request: api/cron/imap-sync"),
    );
    return new Response("Unauthorized", { status: 401 });
  }
  return syncImapAccounts(request.logger);
});

async function syncImapAccounts(logger: Logger) {
  const emailAccounts = await prisma.emailAccount.findMany({
    where: { account: { provider: "imap" } },
    select: { id: true },
  });

  const results = await Promise.allSettled(
    emailAccounts.map(async ({ id: emailAccountId }) => {
      const provider = await createEmailProvider({
        emailAccountId,
        provider: "imap",
        logger,
      });
      const messages = await provider.getInboxMessages(
        MAX_MESSAGES_PER_ACCOUNT,
      );
      logger.info("Synced IMAP inbox", {
        emailAccountId,
        count: messages.length,
      });
      return messages.length;
    }),
  );

  const synced = results.filter((r) => r.status === "fulfilled").length;
  const failed = results.length - synced;
  return NextResponse.json({ accounts: emailAccounts.length, synced, failed });
}
