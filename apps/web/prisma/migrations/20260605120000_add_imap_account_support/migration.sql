-- Add IMAP/SMTP connection fields to Account (only populated when provider = 'imap').
-- imapPassword is stored encrypted (see utils/encryption.ts) and reused for SMTP auth.
ALTER TABLE "Account" ADD COLUMN "imapHost" TEXT;
ALTER TABLE "Account" ADD COLUMN "imapPort" INTEGER;
ALTER TABLE "Account" ADD COLUMN "imapSecure" BOOLEAN;
ALTER TABLE "Account" ADD COLUMN "imapUsername" TEXT;
ALTER TABLE "Account" ADD COLUMN "imapPassword" TEXT;
ALTER TABLE "Account" ADD COLUMN "smtpHost" TEXT;
ALTER TABLE "Account" ADD COLUMN "smtpPort" INTEGER;
ALTER TABLE "Account" ADD COLUMN "smtpSecure" BOOLEAN;
