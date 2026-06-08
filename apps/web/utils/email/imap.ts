import type { ImapFlow, ListResponse } from "imapflow";
import { simpleParser } from "mailparser";
import { createScopedLogger, type Logger } from "@/utils/logger";
import { SafeError } from "@/utils/error";
import type {
  EmailProvider,
  EmailThread,
  EmailLabel,
  EmailFilter,
  EmailSignature,
} from "@/utils/email/types";
import type { ParsedMessage } from "@/utils/types";
import type { InboxZeroLabel } from "@/utils/label";
import type { OutlookFolder } from "@/utils/outlook/folders";
import type { ThreadsQuery } from "@/app/api/threads/validation";
import {
  type ImapConnectionConfig,
  createSmtpTransport,
  withImapClient,
} from "@/utils/imap/client";
import { decodeImapId, parseImapMessage } from "@/utils/imap/message";

const INBOX = "INBOX";
const DEFAULT_MAX_RESULTS = 20;

/**
 * IMAP/SMTP email provider.
 *
 * Core read/send/flag/move operations are implemented over IMAP (via imapflow)
 * and SMTP (via nodemailer). Concepts that don't exist in IMAP — server-side
 * labels with colours, filters, push notifications, and native threading — are
 * adapted (folders act as labels; each message is its own single-message
 * "thread") or are intentionally not supported and clearly surfaced.
 */
export class ImapProvider implements EmailProvider {
  readonly name = "imap";
  private readonly config: ImapConnectionConfig;
  private readonly logger: Logger;

  constructor(
    config: ImapConnectionConfig,
    logger?: Logger,
    emailAccountId?: string,
  ) {
    this.config = config;
    this.logger = (logger || createScopedLogger("imap-provider")).with({
      provider: "imap",
      emailAccountId,
    });
  }

  toJSON() {
    return { name: this.name, type: "ImapProvider" };
  }

  getAccessToken(): string {
    // IMAP uses stored credentials, not bearer tokens.
    return "";
  }

  // ---------------------------------------------------------------------------
  // Reading messages
  // ---------------------------------------------------------------------------

  async getInboxMessages(
    maxResults = DEFAULT_MAX_RESULTS,
  ): Promise<ParsedMessage[]> {
    return this.fetchRecent(INBOX, maxResults);
  }

  async getSentMessages(
    maxResults = DEFAULT_MAX_RESULTS,
  ): Promise<ParsedMessage[]> {
    return withImapClient(this.config, async (client) => {
      const sent = (await this.findSpecialMailbox(client, "\\Sent")) ?? "Sent";
      return this.fetchRecentWithClient(client, sent, maxResults);
    });
  }

  async getMessage(messageId: string): Promise<ParsedMessage> {
    const { mailbox, uid } = decodeImapId(messageId);
    return withImapClient(this.config, async (client) => {
      await client.mailboxOpen(mailbox, { readOnly: true });
      const msg = await client.fetchOne(
        String(uid),
        { source: true, flags: true },
        { uid: true },
      );
      if (!msg?.source) {
        throw new SafeError(`Message not found: ${messageId}`);
      }
      return parseImapMessage({
        source: msg.source,
        uid,
        mailbox,
        flags: msg.flags ?? new Set<string>(),
      });
    });
  }

  async getMessagesBatch(messageIds: string[]): Promise<ParsedMessage[]> {
    const results: ParsedMessage[] = [];
    for (const id of messageIds) {
      try {
        results.push(await this.getMessage(id));
      } catch (error) {
        this.logger.warn("Failed to fetch message in batch", { id, error });
      }
    }
    return results;
  }

  async getMessagesWithPagination(options: {
    query?: string;
    maxResults?: number;
    pageToken?: string;
    before?: Date;
    after?: Date;
    inboxOnly?: boolean;
    unreadOnly?: boolean;
  }): Promise<{ messages: ParsedMessage[]; nextPageToken?: string }> {
    const messages = await this.fetchRecent(
      INBOX,
      options.maxResults ?? DEFAULT_MAX_RESULTS,
    );
    const filtered = options.unreadOnly
      ? messages.filter((m) => m.labelIds?.includes("UNREAD"))
      : messages;
    // Single-page for the MVP; UID-cursor pagination can be layered on later.
    return { messages: filtered, nextPageToken: undefined };
  }

  // ---------------------------------------------------------------------------
  // Threads (each message is treated as its own single-message thread)
  // ---------------------------------------------------------------------------

  async getThreads(folderId?: string): Promise<EmailThread[]> {
    const messages = await this.fetchRecent(
      folderId || INBOX,
      DEFAULT_MAX_RESULTS,
    );
    return messages.map((message) => this.toThread(message));
  }

  async getThread(threadId: string): Promise<EmailThread> {
    const message = await this.getMessage(threadId);
    return this.toThread(message);
  }

  async getThreadMessages(threadId: string): Promise<ParsedMessage[]> {
    return [await this.getMessage(threadId)];
  }

  async getThreadMessagesInInbox(threadId: string): Promise<ParsedMessage[]> {
    return this.getThreadMessages(threadId);
  }

  async getLatestMessageInThread(
    threadId: string,
  ): Promise<ParsedMessage | null> {
    try {
      return await this.getMessage(threadId);
    } catch {
      return null;
    }
  }

  async getLatestMessageFromThreadSnapshot(
    thread: Pick<EmailThread, "id" | "messages">,
  ): Promise<ParsedMessage | null> {
    return thread.messages.at(-1) ?? null;
  }

  // ---------------------------------------------------------------------------
  // Folders / labels (IMAP mailboxes)
  // ---------------------------------------------------------------------------

  async getFolders(): Promise<OutlookFolder[]> {
    return withImapClient(this.config, async (client) => {
      const list = await client.list();
      return list.map((folder) => ({
        id: folder.path,
        displayName: folder.name || folder.path,
        childFolders: [],
        childFolderCount: 0,
      }));
    });
  }

  async getLabels(): Promise<EmailLabel[]> {
    return withImapClient(this.config, async (client) => {
      const list = await client.list();
      return list.map((folder) => this.toLabel(folder));
    });
  }

  async getLabelById(labelId: string): Promise<EmailLabel | null> {
    const labels = await this.getLabels();
    return labels.find((label) => label.id === labelId) ?? null;
  }

  async getLabelByName(name: string): Promise<EmailLabel | null> {
    const labels = await this.getLabels();
    return (
      labels.find((label) => label.name.toLowerCase() === name.toLowerCase()) ??
      null
    );
  }

  // ---------------------------------------------------------------------------
  // Flags / moving messages
  // ---------------------------------------------------------------------------

  async markReadThread(threadId: string, read: boolean): Promise<void> {
    const { mailbox, uid } = decodeImapId(threadId);
    await withImapClient(this.config, async (client) => {
      await client.mailboxOpen(mailbox);
      if (read) {
        await client.messageFlagsAdd(String(uid), ["\\Seen"], { uid: true });
      } else {
        await client.messageFlagsRemove(String(uid), ["\\Seen"], { uid: true });
      }
    });
  }

  async markRead(threadId: string): Promise<void> {
    await this.markReadThread(threadId, true);
  }

  async archiveThread(threadId: string, _ownerEmail: string): Promise<void> {
    await this.moveMessage(threadId, "\\Archive", "Archive");
  }

  async archiveMessage(messageId: string): Promise<void> {
    await this.moveMessage(messageId, "\\Archive", "Archive");
  }

  async trashThread(
    threadId: string,
    _ownerEmail: string,
    _actionSource: "user" | "automation",
  ): Promise<void> {
    await this.moveMessage(threadId, "\\Trash", "Trash");
  }

  async markSpam(threadId: string): Promise<void> {
    await this.moveMessage(threadId, "\\Junk", "Junk");
  }

  async moveThreadToFolder(
    threadId: string,
    _ownerEmail: string,
    folderName: string,
  ): Promise<void> {
    const { mailbox, uid } = decodeImapId(threadId);
    await withImapClient(this.config, async (client) => {
      await client.mailboxOpen(mailbox);
      await client.messageMove(String(uid), folderName, { uid: true });
    });
  }

  async getOrCreateFolderIdByName(folderName: string): Promise<string> {
    return withImapClient(this.config, async (client) => {
      const existing = await client.list();
      const match = existing.find(
        (f) => f.name.toLowerCase() === folderName.toLowerCase(),
      );
      if (match) return match.path;
      await client.mailboxCreate(folderName);
      return folderName;
    });
  }

  // ---------------------------------------------------------------------------
  // Sending (SMTP)
  // ---------------------------------------------------------------------------

  async sendEmail(args: {
    to: string;
    cc?: string;
    bcc?: string;
    subject: string;
    messageText: string;
  }): Promise<void> {
    await this.smtpSend({
      to: args.to,
      cc: args.cc,
      bcc: args.bcc,
      subject: args.subject,
      text: args.messageText,
    });
  }

  async sendEmailWithHtml(body: {
    replyToEmail?: {
      threadId: string;
      headerMessageId: string;
      references?: string;
      messageId?: string;
    };
    to: string;
    from?: string;
    cc?: string;
    bcc?: string;
    replyTo?: string;
    subject: string;
    messageHtml: string;
    attachments?: Array<{
      filename: string;
      content: string;
      contentType: string;
    }>;
  }): Promise<{ messageId: string; threadId: string }> {
    const info = await this.smtpSend({
      to: body.to,
      from: body.from,
      cc: body.cc,
      bcc: body.bcc,
      replyTo: body.replyTo,
      subject: body.subject,
      html: body.messageHtml,
      inReplyTo: body.replyToEmail?.headerMessageId,
      references: body.replyToEmail?.references,
      attachments: body.attachments?.map((a) => ({
        filename: a.filename,
        content: Buffer.from(a.content, "base64"),
        contentType: a.contentType,
      })),
    });
    const messageId = info.messageId ?? "";
    return {
      messageId,
      threadId: body.replyToEmail?.threadId ?? messageId,
    };
  }

  async replyToEmail(
    email: ParsedMessage,
    content: string,
    options?: { replyTo?: string; from?: string },
  ): Promise<void> {
    const to = email.headers["reply-to"] || email.headers.from;
    await this.smtpSend({
      to,
      from: options?.from,
      replyTo: options?.replyTo,
      subject: prefixSubject(email.subject, "Re: "),
      html: content,
      inReplyTo: email.headers["message-id"],
      references: joinReferences(email),
    });
  }

  async forwardEmail(
    email: ParsedMessage,
    args: { to: string; cc?: string; bcc?: string; content?: string },
  ): Promise<void> {
    const body = `${args.content ?? ""}<br/><br/>---------- Forwarded message ----------<br/>${
      email.textHtml ?? email.textPlain ?? ""
    }`;
    await this.smtpSend({
      to: args.to,
      cc: args.cc,
      bcc: args.bcc,
      subject: prefixSubject(email.subject, "Fwd: "),
      html: body,
    });
  }

  // ---------------------------------------------------------------------------
  // Attachments
  // ---------------------------------------------------------------------------

  async getAttachment(
    messageId: string,
    attachmentId: string,
  ): Promise<{ data: string; size: number }> {
    const message = await this.getRawAttachment(messageId, attachmentId);
    return message;
  }

  // ---------------------------------------------------------------------------
  // Stats
  // ---------------------------------------------------------------------------

  async getInboxStats(): Promise<{ total: number; unread: number }> {
    return withImapClient(this.config, async (client) => {
      const status = await client.status(INBOX, {
        messages: true,
        unseen: true,
      });
      return {
        total: status.messages ?? 0,
        unread: status.unseen ?? 0,
      };
    });
  }

  async getSignatures(): Promise<EmailSignature[]> {
    // IMAP has no concept of server-side signatures.
    return [];
  }

  // ---------------------------------------------------------------------------
  // Message heuristics
  // ---------------------------------------------------------------------------

  isReplyInThread(message: ParsedMessage): boolean {
    return !!(message.headers["in-reply-to"] || message.headers.references);
  }

  isSentMessage(message: ParsedMessage): boolean {
    return message.labelIds?.includes("SENT") ?? false;
  }

  // ---------------------------------------------------------------------------
  // Not supported on IMAP (clearly surfaced)
  // ---------------------------------------------------------------------------

  async getMessageByRfc822MessageId(): Promise<ParsedMessage | null> {
    return null;
  }

  async getOriginalMessage(): Promise<ParsedMessage | null> {
    return null;
  }

  async getSentMessageIds(): Promise<{ id: string; threadId: string }[]> {
    return [];
  }

  async getSentThreadsExcluding(): Promise<EmailThread[]> {
    return [];
  }

  async getPreviousConversationMessages(): Promise<ParsedMessage[]> {
    return [];
  }

  async getMessagesWithAttachments(): Promise<{
    messages: ParsedMessage[];
    nextPageToken?: string;
  }> {
    return { messages: [] };
  }

  async getMessagesFromSender(): Promise<{
    messages: ParsedMessage[];
    nextPageToken?: string;
  }> {
    return { messages: [] };
  }

  async getThreadsWithParticipant(): Promise<EmailThread[]> {
    return [];
  }

  async getThreadsWithLabel(): Promise<EmailThread[]> {
    return [];
  }

  async getThreadsWithQuery(_options: {
    query?: ThreadsQuery;
    maxResults?: number;
    pageToken?: string;
  }): Promise<{ threads: EmailThread[]; nextPageToken?: string }> {
    const threads = await this.getThreads();
    return { threads };
  }

  async getThreadsFromSenderWithSubject(): Promise<
    Array<{ id: string; snippet: string; subject: string }>
  > {
    return [];
  }

  async hasPreviousCommunicationsWithSenderOrDomain(): Promise<boolean> {
    return false;
  }

  async checkIfReplySent(): Promise<boolean> {
    return false;
  }

  async countReceivedMessages(): Promise<number> {
    return 0;
  }

  async getDrafts(): Promise<ParsedMessage[]> {
    return [];
  }

  async getDraft(): Promise<ParsedMessage | null> {
    return null;
  }

  async draftEmail(): Promise<{ draftId: string }> {
    throw this.notSupported("Drafts");
  }

  async createDraft(): Promise<{ id: string }> {
    throw this.notSupported("Drafts");
  }

  async updateDraft(): Promise<void> {
    throw this.notSupported("Drafts");
  }

  async deleteDraft(): Promise<void> {
    throw this.notSupported("Drafts");
  }

  async sendDraft(): Promise<{ messageId: string; threadId: string }> {
    throw this.notSupported("Drafts");
  }

  async createLabel(name: string): Promise<EmailLabel> {
    // Creating a "label" maps to creating an IMAP folder.
    const path = await this.getOrCreateFolderIdByName(name);
    return { id: path, name, type: "user" };
  }

  async deleteLabel(labelId: string): Promise<void> {
    await withImapClient(this.config, async (client) => {
      await client.mailboxDelete(labelId);
    });
  }

  async getOrCreateInboxZeroLabel(key: InboxZeroLabel): Promise<EmailLabel> {
    const path = await this.getOrCreateFolderIdByName(key);
    return { id: path, name: key, type: "user" };
  }

  async labelMessage(options: {
    messageId: string;
    labelId: string;
    labelName: string | null;
  }): Promise<{ usedFallback?: boolean; actualLabelId?: string }> {
    // Applying a label on IMAP = copying the message into that folder.
    const { mailbox, uid } = decodeImapId(options.messageId);
    await withImapClient(this.config, async (client) => {
      await client.mailboxOpen(mailbox);
      await client.messageCopy(String(uid), options.labelId, { uid: true });
    });
    return { actualLabelId: options.labelId };
  }

  async removeThreadLabel(): Promise<void> {
    this.logger.warn("removeThreadLabel is not supported for IMAP");
  }

  async removeThreadLabels(): Promise<void> {
    this.logger.warn("removeThreadLabels is not supported for IMAP");
  }

  async archiveThreadWithLabel(
    threadId: string,
    ownerEmail: string,
  ): Promise<void> {
    await this.archiveThread(threadId, ownerEmail);
  }

  async bulkArchiveFromSenders(): Promise<void> {
    this.logger.warn("bulkArchiveFromSenders is not supported for IMAP");
  }

  async bulkTrashFromSenders(): Promise<void> {
    this.logger.warn("bulkTrashFromSenders is not supported for IMAP");
  }

  async blockUnsubscribedEmail(): Promise<void> {
    this.logger.warn("blockUnsubscribedEmail is not supported for IMAP");
  }

  async getFiltersList(): Promise<EmailFilter[]> {
    return [];
  }

  async createFilter(): Promise<{ status: number }> {
    this.logger.warn("createFilter is not supported for IMAP");
    return { status: 501 };
  }

  async deleteFilter(): Promise<{ status: number }> {
    this.logger.warn("deleteFilter is not supported for IMAP");
    return { status: 501 };
  }

  async createAutoArchiveFilter(): Promise<{ status: number }> {
    this.logger.warn("createAutoArchiveFilter is not supported for IMAP");
    return { status: 501 };
  }

  // IMAP has no webhook/push. New mail is picked up by the polling cron instead.
  async watchEmails(): Promise<{
    expirationDate: Date;
    subscriptionId?: string;
  } | null> {
    return null;
  }

  async unwatchEmails(): Promise<void> {
    // No-op: nothing to unwatch without push subscriptions.
  }

  async processHistory(): Promise<void> {
    // No-op: IMAP has no history API; the polling cron drives processing.
  }

  // ---------------------------------------------------------------------------
  // Private helpers
  // ---------------------------------------------------------------------------

  private async fetchRecent(
    mailbox: string,
    maxResults: number,
  ): Promise<ParsedMessage[]> {
    return withImapClient(this.config, (client) =>
      this.fetchRecentWithClient(client, mailbox, maxResults),
    );
  }

  private async fetchRecentWithClient(
    client: ImapFlow,
    mailbox: string,
    maxResults: number,
  ): Promise<ParsedMessage[]> {
    const box = await client.mailboxOpen(mailbox, { readOnly: true });
    const total = box.exists;
    if (!total) return [];

    const start = Math.max(1, total - maxResults + 1);
    const messages: ParsedMessage[] = [];
    for await (const msg of client.fetch(`${start}:*`, {
      uid: true,
      flags: true,
      source: true,
    })) {
      if (!msg.source) continue;
      messages.push(
        await parseImapMessage({
          source: msg.source,
          uid: msg.uid,
          mailbox,
          flags: msg.flags ?? new Set<string>(),
        }),
      );
    }
    // IMAP returns oldest-first for an ascending range; newest first is nicer.
    return messages.reverse();
  }

  private async moveMessage(
    encodedId: string,
    specialUse: string,
    fallbackName: string,
  ): Promise<void> {
    const { mailbox, uid } = decodeImapId(encodedId);
    await withImapClient(this.config, async (client) => {
      const destination =
        (await this.findSpecialMailbox(client, specialUse)) ?? fallbackName;
      await client.mailboxOpen(mailbox);
      await client.messageMove(String(uid), destination, { uid: true });
    });
  }

  private async findSpecialMailbox(
    client: ImapFlow,
    specialUse: string,
  ): Promise<string | null> {
    const list = await client.list();
    return list.find((f) => f.specialUse === specialUse)?.path ?? null;
  }

  private async getRawAttachment(
    messageId: string,
    attachmentId: string,
  ): Promise<{ data: string; size: number }> {
    const { mailbox, uid } = decodeImapId(messageId);
    return withImapClient(this.config, async (client) => {
      await client.mailboxOpen(mailbox, { readOnly: true });
      const msg = await client.fetchOne(
        String(uid),
        { source: true },
        { uid: true },
      );
      if (!msg?.source) {
        throw new SafeError(`Message not found: ${messageId}`);
      }
      const parsed = await simpleParser(msg.source);
      const index = Number(attachmentId);
      const attachment = parsed.attachments?.[index];
      if (!attachment) {
        throw new SafeError(`Attachment not found: ${attachmentId}`);
      }
      return {
        data: attachment.content.toString("base64"),
        size: attachment.size ?? attachment.content.length,
      };
    });
  }

  private async smtpSend(message: {
    to: string;
    from?: string;
    cc?: string;
    bcc?: string;
    replyTo?: string;
    subject: string;
    text?: string;
    html?: string;
    inReplyTo?: string;
    references?: string;
    attachments?: Array<{
      filename: string;
      content: Buffer;
      contentType: string;
    }>;
  }): Promise<{ messageId?: string }> {
    const transport = createSmtpTransport(this.config);
    try {
      const info = await transport.sendMail({
        from: message.from || this.config.email,
        to: message.to,
        cc: message.cc,
        bcc: message.bcc,
        replyTo: message.replyTo,
        subject: message.subject,
        text: message.text,
        html: message.html,
        inReplyTo: message.inReplyTo,
        references: message.references,
        attachments: message.attachments,
      });
      return { messageId: info.messageId };
    } finally {
      transport.close();
    }
  }

  private toThread(message: ParsedMessage): EmailThread {
    return {
      id: message.id,
      messages: [message],
      snippet: message.snippet,
    };
  }

  private toLabel(folder: ListResponse): EmailLabel {
    return {
      id: folder.path,
      name: folder.name || folder.path,
      type: "user",
    };
  }

  private notSupported(feature: string): SafeError {
    this.logger.warn(`${feature} are not supported for IMAP accounts`);
    return new SafeError(`${feature} are not supported for IMAP accounts`);
  }
}

function prefixSubject(subject: string, prefix: string): string {
  const value = subject ?? "";
  return value.toLowerCase().startsWith(prefix.toLowerCase())
    ? value
    : `${prefix}${value}`;
}

function joinReferences(email: ParsedMessage): string | undefined {
  const references = email.headers.references;
  const messageId = email.headers["message-id"];
  return [references, messageId].filter(Boolean).join(" ") || undefined;
}
