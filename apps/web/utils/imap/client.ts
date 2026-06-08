import { ImapFlow } from "imapflow";
import nodemailer, { type Transporter } from "nodemailer";
import prisma from "@/utils/prisma";
import { decryptToken } from "@/utils/encryption";
import { SafeError } from "@/utils/error";

export interface ImapConnectionConfig {
  email: string;
  imap: { host: string; port: number; secure: boolean };
  smtp: { host: string; port: number; secure: boolean };
  auth: { user: string; pass: string };
}

export interface ImapCredentialsInput {
  email: string;
  imapHost: string;
  imapPort: number;
  imapSecure: boolean;
  smtpHost: string;
  smtpPort: number;
  smtpSecure: boolean;
  username: string;
  password: string;
}

/**
 * Loads and decrypts the stored IMAP/SMTP connection details for an email account.
 * Throws if the account is not an IMAP account or is missing credentials.
 */
export async function getImapConnectionForEmail({
  emailAccountId,
}: {
  emailAccountId: string;
}): Promise<ImapConnectionConfig> {
  const emailAccount = await prisma.emailAccount.findUnique({
    where: { id: emailAccountId },
    select: {
      email: true,
      account: {
        select: {
          provider: true,
          imapHost: true,
          imapPort: true,
          imapSecure: true,
          imapUsername: true,
          imapPassword: true,
          smtpHost: true,
          smtpPort: true,
          smtpSecure: true,
        },
      },
    },
  });

  const account = emailAccount?.account;
  if (account?.provider !== "imap") {
    throw new SafeError("This email account is not an IMAP account");
  }

  const password = decryptToken(account.imapPassword ?? null);
  if (
    !account.imapHost ||
    !account.imapPort ||
    !account.imapUsername ||
    !account.smtpHost ||
    !account.smtpPort ||
    !password
  ) {
    throw new SafeError("IMAP account is missing connection details");
  }

  return buildConfig({
    email: emailAccount.email,
    imapHost: account.imapHost,
    imapPort: account.imapPort,
    imapSecure: account.imapSecure ?? true,
    smtpHost: account.smtpHost,
    smtpPort: account.smtpPort,
    smtpSecure: account.smtpSecure ?? true,
    username: account.imapUsername,
    password,
  });
}

export function buildConfig(input: ImapCredentialsInput): ImapConnectionConfig {
  return {
    email: input.email,
    imap: {
      host: input.imapHost,
      port: input.imapPort,
      secure: input.imapSecure,
    },
    smtp: {
      host: input.smtpHost,
      port: input.smtpPort,
      secure: input.smtpSecure,
    },
    auth: { user: input.username, pass: input.password },
  };
}

export function createImapClient(config: ImapConnectionConfig): ImapFlow {
  return new ImapFlow({
    host: config.imap.host,
    port: config.imap.port,
    secure: config.imap.secure,
    auth: { user: config.auth.user, pass: config.auth.pass },
    logger: false,
  });
}

export function createSmtpTransport(config: ImapConnectionConfig): Transporter {
  return nodemailer.createTransport({
    host: config.smtp.host,
    port: config.smtp.port,
    secure: config.smtp.secure,
    auth: { user: config.auth.user, pass: config.auth.pass },
  });
}

/**
 * Opens an IMAP connection, runs the callback, and always logs out afterwards.
 */
export async function withImapClient<T>(
  config: ImapConnectionConfig,
  fn: (client: ImapFlow) => Promise<T>,
): Promise<T> {
  const client = createImapClient(config);
  await client.connect();
  try {
    return await fn(client);
  } finally {
    await client.logout().catch(() => {});
  }
}

/**
 * Validates IMAP login + SMTP login. Used before persisting credentials so we
 * never store details that don't actually work. Throws SafeError on failure.
 */
export async function testImapConnection(
  config: ImapConnectionConfig,
): Promise<void> {
  try {
    await withImapClient(config, async () => {
      // Connecting + authenticating is enough to validate IMAP.
    });
  } catch (error) {
    throw new SafeError(
      `Could not connect to IMAP server: ${getErrorMessage(error)}`,
    );
  }

  try {
    const transport = createSmtpTransport(config);
    await transport.verify();
    transport.close();
  } catch (error) {
    throw new SafeError(
      `Could not connect to SMTP server: ${getErrorMessage(error)}`,
    );
  }
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) return error.message;
  return "unknown error";
}
