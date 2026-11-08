// Outlook REST v2.0 transport and mail reads (port of outlook_api.py). One transport for mail and calendars.

import { OUTLOOK_REST, type Account, type Tokens } from "./tokens.ts";

export type Obj = Record<string, any>;
export type Fetch = typeof fetch;

export class ApiError extends Error {}

export class Outlook {
  readonly tokens: Tokens;
  private readonly fetchFn: Fetch;
  constructor(tokens: Tokens, fetchFn: Fetch = fetch) {
    this.tokens = tokens;
    this.fetchFn = fetchFn;
  }

  account(account?: string): Promise<Account> {
    return this.tokens.pick(account);
  }

  /** One request; a 401 reloads tokens from the broker once and retries. Error bodies are cut short: they may quote mail content. */
  async request(
    account: string | undefined,
    p: string,
    method = "GET",
    data?: unknown,
    retry = true,
    extra?: Record<string, string>,
  ): Promise<Obj> {
    if (!/^me(?:[/?]|$)/.test(p) || /[\r\n\\]/.test(p))
      throw new ApiError("Invalid Outlook request path");
    const a = await this.tokens.pick(account, !retry);
    const headers: Record<string, string> = {
      ...extra,
      Authorization: `Bearer ${a.token}`,
      Accept: "application/json",
      "User-Agent": "OutlookMCP/3.0.0",
    };
    if (data !== undefined)
      headers["Content-Type"] = "application/json; charset=utf-8";
    const r = await this.fetchFn(`${OUTLOOK_REST}/${p.replace(/^\/+/, "")}`, {
      method,
      headers,
      body: data === undefined ? undefined : JSON.stringify(data),
      redirect: "error",
      signal: AbortSignal.timeout(30_000),
    });
    if (r.status === 401 && retry) {
      await r.body?.cancel();
      return this.request(account, p, method, data, false, extra);
    }
    const text = await r.text();
    if (!r.ok)
      throw new ApiError(
        `Outlook API error ${r.status} on ${method} ${p.split("?")[0]}: ${text.slice(0, 300) || r.statusText}`,
      );
    if (!text) return {};
    try {
      return JSON.parse(text) as Obj;
    } catch {
      throw new ApiError(
        `Outlook API answered non-JSON on ${method} ${p.split("?")[0]}`,
      );
    }
  }
}

export const q = encodeURIComponent;

const WELL_KNOWN: Record<string, string> = {
  inbox: "Inbox",
  "posta in arrivo": "Inbox",
  sent: "SentItems",
  sentitems: "SentItems",
  "sent items": "SentItems",
  "posta inviata": "SentItems",
  drafts: "Drafts",
  bozze: "Drafts",
  archive: "Archive",
  archivio: "Archive",
  deleted: "DeletedItems",
  deleteditems: "DeletedItems",
  "deleted items": "DeletedItems",
  trash: "DeletedItems",
  "posta eliminata": "DeletedItems",
  junk: "JunkEmail",
  junkemail: "JunkEmail",
  spam: "JunkEmail",
  "posta indesiderata": "JunkEmail",
  outbox: "Outbox",
  "posta in uscita": "Outbox",
};

export interface Folder {
  id: string;
  name: string;
  display_name?: string;
  total_count: number;
  unread_count: number;
  child_count: number;
}

export async function listFolders(
  o: Outlook,
  account?: string,
  children = true,
): Promise<Folder[]> {
  account = (await o.account(account)).upn;
  const { nextLinkPath } = await import("./mail.ts");
  const folders: Folder[] = [];
  async function scan(
    base: string,
    prefix: string,
    depth: number,
  ): Promise<void> {
    let p = `${base}?$top=100`;
    const seen = new Set<string>();
    for (let pages = 0; pages < 50; pages++) {
      if (seen.has(p)) throw new ApiError("Repeated folder pagination link");
      seen.add(p);
      const res = await o.request(account, p);
      for (const f of (res.value ?? []) as Obj[]) {
        const name = prefix + String(f.DisplayName ?? "");
        folders.push({
          id: f.Id,
          name,
          display_name: f.DisplayName,
          total_count: f.TotalItemCount ?? 0,
          unread_count: f.UnreadItemCount ?? 0,
          child_count: Number(f.ChildFolderCount) || 0,
        });
        if (children && f.ChildFolderCount && f.Id) {
          if (depth >= 20)
            throw new ApiError("Folder nesting exceeds 20 levels");
          await scan(
            `me/MailFolders/${q(f.Id)}/childfolders`,
            name + "/",
            depth + 1,
          );
        }
      }
      if (!res["@odata.nextLink"]) return;
      const next = nextLinkPath(res["@odata.nextLink"], base);
      if (!next) throw new ApiError("Untrusted folder pagination link");
      p = next;
    }
    throw new ApiError("Folder listing exceeds 50 pages");
  }
  await scan("me/MailFolders", "", 0);
  return folders;
}

export async function resolveFolderId(
  o: Outlook,
  account: string | undefined,
  folder?: string,
): Promise<string | null> {
  const clean = (folder ?? "").trim();
  if (!clean) return "Inbox";
  const known = WELL_KNOWN[clean.toLowerCase()];
  if (known) return known;
  if (clean.startsWith("AAMk") || clean.startsWith("AQMk")) return clean;
  const folders = await listFolders(o, account, true);
  const want = clean.toLowerCase();
  const exact = folders.find(
    (f) =>
      f.name.toLowerCase() === want ||
      (f.display_name ?? "").toLowerCase() === want,
  );
  return (
    exact?.id ??
    folders.find((f) => f.name.toLowerCase().includes(want))?.id ??
    null
  );
}
