"use server";

import { revalidatePath } from "next/cache";
import { actionClientUser } from "@/utils/actions/safe-action";
import { connectImapBody } from "@/utils/actions/imap.validation";
import prisma from "@/utils/prisma";
import { encryptToken } from "@/utils/encryption";
import { buildConfig, testImapConnection } from "@/utils/imap/client";
import { SafeError } from "@/utils/error";

export const connectImapAccountAction = actionClientUser
  .metadata({ name: "connectImapAccount" })
  .schema(connectImapBody)
  .action(async ({ ctx: { userId }, parsedInput }) => {
    const email = parsedInput.email.trim().toLowerCase();

    // Verify the credentials actually work before persisting anything.
    await testImapConnection(
      buildConfig({
        email,
        username: parsedInput.username,
        password: parsedInput.password,
        imapHost: parsedInput.imapHost,
        imapPort: parsedInput.imapPort,
        imapSecure: parsedInput.imapSecure,
        smtpHost: parsedInput.smtpHost,
        smtpPort: parsedInput.smtpPort,
        smtpSecure: parsedInput.smtpSecure,
      }),
    );

    const existing = await prisma.emailAccount.findUnique({
      where: { email },
      select: { userId: true },
    });
    if (existing) {
      throw new SafeError(
        existing.userId === userId
          ? "This email account is already connected"
          : "This email is already connected to another account",
      );
    }

    await prisma.account.create({
      data: {
        userId,
        type: "imap",
        provider: "imap",
        providerAccountId: `imap:${email}`,
        imapHost: parsedInput.imapHost,
        imapPort: parsedInput.imapPort,
        imapSecure: parsedInput.imapSecure,
        imapUsername: parsedInput.username,
        imapPassword: encryptToken(parsedInput.password),
        smtpHost: parsedInput.smtpHost,
        smtpPort: parsedInput.smtpPort,
        smtpSecure: parsedInput.smtpSecure,
        emailAccount: {
          create: {
            email,
            userId,
          },
        },
      },
    });

    revalidatePath("/accounts");
  });
