"use client";

import { useState } from "react";
import { useRouter } from "next/navigation";
import { useForm, type SubmitHandler } from "react-hook-form";
import { useAction } from "next-safe-action/hooks";
import { zodResolver } from "@hookform/resolvers/zod";
import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
  DialogTrigger,
} from "@/components/ui/dialog";
import { Input } from "@/components/Input";
import { toastSuccess, toastError } from "@/components/Toast";
import { connectImapAccountAction } from "@/utils/actions/imap";
import {
  connectImapBody,
  type ConnectImapBody,
} from "@/utils/actions/imap.validation";

export function ConnectImapDialog() {
  const [open, setOpen] = useState(false);
  const router = useRouter();

  const { execute, isExecuting } = useAction(connectImapAccountAction, {
    onSuccess: () => {
      toastSuccess({ description: "IMAP account connected" });
      setOpen(false);
      router.refresh();
    },
    onError: ({ error }) => {
      toastError({
        title: "Could not connect IMAP account",
        description:
          error.serverError ?? "Check your server settings and try again.",
      });
    },
  });

  const {
    register,
    handleSubmit,
    formState: { errors },
  } = useForm<ConnectImapBody>({
    resolver: zodResolver(connectImapBody),
    defaultValues: {
      imapPort: 993,
      imapSecure: true,
      smtpPort: 465,
      smtpSecure: true,
    },
  });

  const onSubmit: SubmitHandler<ConnectImapBody> = (data) => execute(data);

  return (
    <Dialog open={open} onOpenChange={setOpen}>
      <DialogTrigger asChild>
        <Button variant="outline" className="w-full">
          Add IMAP account
        </Button>
      </DialogTrigger>
      <DialogContent className="max-h-[90vh] overflow-y-auto">
        <DialogHeader>
          <DialogTitle>Connect an IMAP account</DialogTitle>
          <DialogDescription>
            Works with Yahoo, Fastmail, and most hosting/enterprise mailboxes.
            For accounts with 2FA (e.g. Yahoo), use an app password.
          </DialogDescription>
        </DialogHeader>

        <form className="space-y-3" onSubmit={handleSubmit(onSubmit)}>
          <Input
            type="email"
            name="email"
            label="Email address"
            placeholder="you@example.com"
            registerProps={register("email")}
            error={errors.email}
          />
          <Input
            type="text"
            name="username"
            label="Username"
            placeholder="Usually your full email address"
            registerProps={register("username")}
            error={errors.username}
          />
          <Input
            type="password"
            name="password"
            label="Password / app password"
            registerProps={register("password")}
            error={errors.password}
          />

          <div className="grid grid-cols-2 gap-3">
            <Input
              type="text"
              name="imapHost"
              label="IMAP host"
              placeholder="imap.mail.yahoo.com"
              registerProps={register("imapHost")}
              error={errors.imapHost}
            />
            <Input
              type="number"
              name="imapPort"
              label="IMAP port"
              registerProps={register("imapPort")}
              error={errors.imapPort}
            />
          </div>

          <div className="grid grid-cols-2 gap-3">
            <Input
              type="text"
              name="smtpHost"
              label="SMTP host"
              placeholder="smtp.mail.yahoo.com"
              registerProps={register("smtpHost")}
              error={errors.smtpHost}
            />
            <Input
              type="number"
              name="smtpPort"
              label="SMTP port"
              registerProps={register("smtpPort")}
              error={errors.smtpPort}
            />
          </div>

          <label className="flex items-center gap-2 text-sm">
            <input type="checkbox" {...register("imapSecure")} />
            Use TLS for IMAP (recommended)
          </label>
          <label className="flex items-center gap-2 text-sm">
            <input type="checkbox" {...register("smtpSecure")} />
            Use TLS for SMTP (recommended)
          </label>

          <Button type="submit" loading={isExecuting} className="w-full">
            Connect account
          </Button>
        </form>
      </DialogContent>
    </Dialog>
  );
}
