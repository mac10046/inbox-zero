import { simpleParser, type ParsedMail, type AddressObject } from "mailparser";
import type {
  Attachment,
  ParsedMessage,
  ParsedMessageHeaders,
} from "@/utils/types";

const ID_SEPARATOR = "::";

// IMAP messages are addressed by (mailbox, UID). We encode both into the single
// `id`/`messageId` string the rest of the app passes around.
export function encodeImapId(mailbox: string, uid: number): string {
  return `${mailbox}${ID_SEPARATOR}${uid}`;
}

export function decodeImapId(id: string): { mailbox: string; uid: number } {
  const idx = id.lastIndexOf(ID_SEPARATOR);
  if (idx === -1) throw new Error(`Invalid IMAP message id: ${id}`);
  return {
    mailbox: id.slice(0, idx),
    uid: Number(id.slice(idx + ID_SEPARATOR.length)),
  };
}

export async function parseImapMessage({
  source,
  uid,
  mailbox,
  flags,
}: {
  source: Buffer | string;
  uid: number;
  mailbox: string;
  flags: Iterable<string>;
}): Promise<ParsedMessage> {
  const parsed = await simpleParser(source);
  const id = encodeImapId(mailbox, uid);
  const flagList = Array.from(flags ?? []);
  const date = parsed.date?.toISOString() ?? new Date().toISOString();
  const textHtml = typeof parsed.html === "string" ? parsed.html : undefined;

  return {
    id,
    threadId: getThreadId(parsed, id),
    labelIds: buildLabelIds(mailbox, flagList),
    snippet: makeSnippet(parsed.text ?? stripHtml(textHtml)),
    historyId: "",
    attachments: mapAttachments(parsed),
    inline: [],
    headers: buildHeaders(parsed, date),
    textPlain: parsed.text || undefined,
    textHtml,
    subject: parsed.subject ?? "",
    date,
    internalDate: date,
  };
}

function buildHeaders(parsed: ParsedMail, date: string): ParsedMessageHeaders {
  const references = Array.isArray(parsed.references)
    ? parsed.references.join(" ")
    : parsed.references;

  return {
    subject: parsed.subject ?? "",
    from: addressText(parsed.from),
    to: addressText(parsed.to),
    cc: addressText(parsed.cc) || undefined,
    bcc: addressText(parsed.bcc) || undefined,
    date,
    "message-id": parsed.messageId || undefined,
    "in-reply-to": parsed.inReplyTo || undefined,
    references: references || undefined,
  };
}

// IMAP has no native threads. Group by the root References id when present,
// otherwise the message stands alone as its own thread.
function getThreadId(parsed: ParsedMail, fallbackId: string): string {
  const refs = Array.isArray(parsed.references)
    ? parsed.references
    : parsed.references
      ? [parsed.references]
      : [];
  if (refs.length > 0) return refs[0];
  if (parsed.inReplyTo) return parsed.inReplyTo;
  if (parsed.messageId) return parsed.messageId;
  return fallbackId;
}

// Map the mailbox + IMAP flags onto the label-id strings the app expects so
// heuristics like isSentMessage / unread detection keep working.
function buildLabelIds(mailbox: string, flags: string[]): string[] {
  const labels: string[] = [];
  const lower = mailbox.toLowerCase();
  if (lower === "inbox") labels.push("INBOX");
  if (lower.includes("sent")) labels.push("SENT");
  if (lower.includes("draft")) labels.push("DRAFT");
  if (lower.includes("trash") || lower.includes("deleted"))
    labels.push("TRASH");
  if (lower.includes("junk") || lower.includes("spam")) labels.push("SPAM");
  if (!flags.includes("\\Seen")) labels.push("UNREAD");
  return labels;
}

function mapAttachments(parsed: ParsedMail): Attachment[] {
  return (parsed.attachments ?? []).map((att, index) => ({
    filename: att.filename ?? `attachment-${index}`,
    mimeType: att.contentType ?? "application/octet-stream",
    size: att.size ?? att.content?.length ?? 0,
    // We address attachments by their position when downloading later.
    attachmentId: String(index),
    headers: {
      "content-type": att.contentType ?? "application/octet-stream",
      "content-description": "",
      "content-transfer-encoding": att.contentDisposition ?? "",
      "content-id": att.contentId ?? "",
    },
  }));
}

function addressText(
  address: AddressObject | AddressObject[] | undefined,
): string {
  if (!address) return "";
  if (Array.isArray(address)) return address.map((a) => a.text).join(", ");
  return address.text ?? "";
}

function makeSnippet(text: string | undefined): string {
  if (!text) return "";
  return text.replace(/\s+/g, " ").trim().slice(0, 200);
}

function stripHtml(html: string | undefined): string | undefined {
  if (!html) return undefined;
  return html.replace(/<[^>]+>/g, " ");
}
