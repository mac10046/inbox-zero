import { z } from "zod";

export const connectImapBody = z.object({
  email: z.string().email(),
  username: z.string().min(1),
  password: z.string().min(1),
  imapHost: z.string().min(1),
  imapPort: z.coerce.number().int().positive().default(993),
  imapSecure: z.boolean().default(true),
  smtpHost: z.string().min(1),
  smtpPort: z.coerce.number().int().positive().default(465),
  smtpSecure: z.boolean().default(true),
});
export type ConnectImapBody = z.infer<typeof connectImapBody>;
