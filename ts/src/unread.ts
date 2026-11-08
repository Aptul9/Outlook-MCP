// Cross-mailbox unread listing, bounded parallelism and explicit per-account completeness.
import { selectAccount, type Account } from "./tokens.ts";
import { type Outlook } from "./api.ts";
import { listEmailsPage, type EmailList } from "./mail.ts";
export interface UnreadMailbox {
  account: string;
  result?: EmailList;
  error?: string;
}
export async function listUnread(
  o: Outlook,
  options: {
    accounts?: string[];
    limit?: number;
    max_pages?: number;
    concurrency?: number;
  } = {},
): Promise<UnreadMailbox[]> {
  const all = await o.tokens.accounts();
  const errors: UnreadMailbox[] = [];
  const selected: Account[] = [];
  if (options.accounts?.length) {
    for (const selector of options.accounts) {
      try {
        selected.push(selectAccount(all, selector));
      } catch {
        errors.push({
          account: selector,
          error: "Mailbox selector not found or ambiguous",
        });
      }
    }
  } else selected.push(...all);
  const unique = [...new Map(selected.map((a) => [a.upn, a])).values()];
  const out: UnreadMailbox[] = new Array(unique.length);
  let next = 0;
  const workers = Math.min(
    unique.length,
    Math.max(1, Math.min(options.concurrency ?? 3, 5)),
  );
  await Promise.all(
    Array.from({ length: workers }, async () => {
      for (;;) {
        const i = next++;
        if (i >= unique.length) return;
        const account = unique[i]!.upn;
        if (unique[i]!.status !== "ok") {
          out[i] = {
            account,
            error: unique[i]!.message || "Mailbox sign-in requires attention",
          };
          continue;
        }
        try {
          out[i] = {
            account,
            result: await listEmailsPage(o, account, {
              folder: "all",
              unread_only: true,
              limit: options.limit ?? 100,
              max_pages: options.max_pages,
            }),
          };
        } catch {
          out[i] = { account, error: "Mailbox unread request failed" };
        }
      }
    }),
  );
  return [...out, ...errors];
}
