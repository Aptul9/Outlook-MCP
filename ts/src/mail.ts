// Mail reads: list, search, message, thread, attachments (port of outlook_api.py).

import { q, resolveFolderId, type Obj, type Outlook } from "./api.ts";
import { OUTLOOK_REST } from "./tokens.ts";

/** Most messages one list or search returns. */
export const MAX_LIST_LIMIT = 1000;
/** Pages read per call unless max_pages says otherwise, and the ceiling max_pages is clamped to. */
export const DEFAULT_MAX_PAGES = 20;
export const MAX_PAGES_CAP = 50;
const PAGE_SIZE = 100;

export interface ListOptions {
  folder?: string;
  query?: string;
  /** 1 to MAX_LIST_LIMIT, default 20. */
  limit?: number;
  unread_only?: boolean;
  from_address?: string;
  has_attachments?: boolean;
  start_date?: string;
  end_date?: string;
  flagged_only?: boolean;
  /** Pages of PAGE_SIZE read at most, 1 to MAX_PAGES_CAP, default DEFAULT_MAX_PAGES. */
  max_pages?: number;
}

/** Why paging stopped. Only limit and exhausted mean the scan finished. */
export type StopReason =
  | "limit"
  | "exhausted"
  | "max_pages"
  | "untrusted_next_link"
  | "repeated_next_link";

export interface EmailList {
  items: Obj[];
  /** Messages past items exist; under local filters (from_address, flagged_only, filters next to a query) they may not match. */
  has_more: boolean;
  /** Paging stopped before limit and before the end of the collection: older matches may be missing. */
  incomplete: boolean;
  stop_reason: StopReason;
  /** Effective limit after clamping. */
  limit: number;
  /** Message pages requested. */
  pages: number;
  /** Messages received across those pages, before local filters. */
  scanned: number;
}

export function normalizeIsoDate(
  d: string | undefined,
  endOfDay = false,
): string | undefined {
  const s = d?.trim();
  if (!s) return undefined;
  if (/^\d{4}-\d{2}-\d{2}$/.test(s))
    return `${s}T${endOfDay ? "23:59:59" : "00:00:00"}Z`;
  const tail = s.slice(10);
  if (!s.endsWith("Z") && !tail.includes("+") && !tail.includes("-"))
    return `${s}Z`;
  return s;
}

export function matchesLocalFilters(m: Obj, o: ListOptions): boolean {
  if (o.unread_only && m.IsRead !== false) return false;
  if (o.has_attachments === true && !m.HasAttachments) return false;
  if (o.has_attachments === false && m.HasAttachments) return false;
  if (
    o.flagged_only &&
    String(m.Flag?.FlagStatus ?? "").toLowerCase() !== "flagged"
  )
    return false;
  if (o.from_address) {
    const s = (m.From ?? m.Sender ?? {}).EmailAddress ?? {};
    if (
      !`${s.Name ?? ""} ${s.Address ?? ""}`
        .toLowerCase()
        .includes(o.from_address.toLowerCase())
    )
      return false;
  }
  const at = String(m.ReceivedDateTime ?? m.CreatedDateTime ?? "");
  const start = normalizeIsoDate(o.start_date);
  const end = normalizeIsoDate(o.end_date, true);
  if (start && at && compareInstants(at, start) < 0) return false;
  if (end && at && compareInstants(at, end) > 0) return false;
  return true;
}

/** Instants compared as time, so +01:00 and Z bounds agree; text order only when one side does not parse. */
function compareInstants(a: string, b: string): number {
  const x = Date.parse(a);
  const y = Date.parse(b);
  if (Number.isFinite(x) && Number.isFinite(y)) return x - y;
  return a < b ? -1 : a > b ? 1 : 0;
}

const SUMMARY_FIELDS =
  "Id,Subject,From,Sender,ToRecipients,CcRecipients,ReceivedDateTime,CreatedDateTime,BodyPreview,HasAttachments,IsRead,IsDraft,Importance,Flag,ConversationId,ParentFolderId";
const ATTACHMENT_FIELDS = "$select=Id,Name,ContentType,Size,IsInline";

const clamp = (v: number | undefined, dflt: number, max: number) =>
  Math.max(1, Math.min(Math.trunc(v ?? dflt) || dflt, max));

const REST = new URL(OUTLOOK_REST);

/** Segments the API matches without regard to case: path words and well-known folder names. Opaque ids are case-sensitive and keep theirs. */
const CASELESS = new Set([
  "me",
  "mailfolders",
  "messages",
  "childfolders",
  "inbox",
  "sentitems",
  "drafts",
  "archive",
  "deleteditems",
  "junkemail",
  "outbox",
]);

/** Collection identity of a REST path: segments decoded one by one (an encoded "/" stays inside its segment), MailFolders('x') read as MailFolders/x, trailing slash ignored. */
function collectionKey(p: string): string | null {
  try {
    const parts: string[] = [];
    for (const seg of p
      .replace(/\/+$/, "")
      .split("/")
      .map(decodeURIComponent)) {
      const k = /^MailFolders\('((?:[^']|'')*)'\)$/i.exec(seg);
      if (k) parts.push("MailFolders", k[1]!.replace(/''/g, "'"));
      else parts.push(seg);
    }
    return JSON.stringify(
      parts.map((s) => (CASELESS.has(s.toLowerCase()) ? s.toLowerCase() : s)),
    );
  } catch {
    return null;
  }
}

/**
 * The path to request for an @odata.nextLink, or null when the link is not absolute https on the Outlook REST origin
 * (no credentials, default port), under its API path, and on the same collection as `base`. The request is rebuilt on
 * OUTLOOK_REST from that path, so the token never leaves the origin whatever the link says.
 */
export function nextLinkPath(link: unknown, base: string): string | null {
  if (typeof link !== "string" || !link) return null;
  let u: URL;
  try {
    u = new URL(link);
  } catch {
    return null;
  }
  const prefix = `${REST.pathname}/`;
  if (
    u.origin !== REST.origin ||
    u.username ||
    u.password ||
    !u.pathname.startsWith(prefix)
  )
    return null;
  const path = u.pathname.slice(prefix.length);
  const key = collectionKey(path);
  return key !== null && key === collectionKey(base) ? path + u.search : null;
}

/** Reads `base?query` and its nextLinks until `limit` kept messages, the end, `maxPages`, or a link that is refused or repeats. Duplicate Ids across pages are dropped. */
async function collect(
  o: Outlook,
  account: string | undefined,
  base: string,
  query: string,
  limit: number,
  maxPages: number,
  keep?: (m: Obj) => boolean,
): Promise<EmailList> {
  const items: Obj[] = [];
  const ids = new Set<unknown>();
  const visited = new Set<string>();
  const fresh = (m: Obj) => m.Id == null || !ids.has(m.Id);
  let path = `${base}?${query}`;
  let pages = 0;
  let scanned = 0;
  for (;;) {
    visited.add(path);
    const res = await o.request(account, path);
    pages++;
    const value = (Array.isArray(res.value) ? res.value : []) as Obj[];
    scanned += value.length;
    let i = 0;
    for (; i < value.length && items.length < limit; i++) {
      const m = value[i]!;
      if (!fresh(m)) continue;
      if (m.Id != null) ids.add(m.Id);
      if (!keep || keep(m)) items.push(m);
    }
    const link = res["@odata.nextLink"];
    const done = (stop_reason: StopReason, has_more: boolean): EmailList => ({
      items,
      has_more,
      incomplete: stop_reason !== "limit" && stop_reason !== "exhausted",
      stop_reason,
      limit,
      pages,
      scanned,
    });
    if (items.length >= limit)
      return done(
        "limit",
        !!link || value.slice(i).some((m) => fresh(m) && (!keep || keep(m))),
      );
    if (!link) return done("exhausted", false);
    const next = nextLinkPath(link, base);
    if (!next) return done("untrusted_next_link", true);
    if (visited.has(next)) return done("repeated_next_link", true);
    if (pages >= maxPages) return done("max_pages", true);
    path = next;
  }
}

/**
 * Newest first, up to limit (max MAX_LIST_LIMIT), following nextLink pages so local filters reach older mail.
 * Check `incomplete` before treating a short result as all there is. An unknown folder name is an error.
 */
export async function listEmailsPage(
  o: Outlook,
  account: string | undefined,
  opt: ListOptions = {},
): Promise<EmailList> {
  if (account && !account.includes("@"))
    account = (await o.account(account)).upn;
  const limit = clamp(opt.limit, 20, MAX_LIST_LIMIT);
  const maxPages = clamp(opt.max_pages, DEFAULT_MAX_PAGES, MAX_PAGES_CAP);
  const folder = opt.folder ?? "Inbox";
  let base = "me/messages";
  if (!["all", "*", "any"].includes(folder.trim().toLowerCase())) {
    const fid = await resolveFolderId(o, account, folder);
    if (!fid) throw new Error(`Folder not found: ${folder}`);
    base = `me/MailFolders/${q(fid)}/messages`;
  }
  const postFilter = !!(
    opt.from_address ||
    opt.flagged_only ||
    (opt.query &&
      (opt.unread_only ||
        opt.has_attachments !== undefined ||
        opt.start_date ||
        opt.end_date))
  );
  const top = postFilter ? PAGE_SIZE : Math.min(limit, PAGE_SIZE);
  const params = [`$top=${top}`, `$select=${SUMMARY_FIELDS}`];
  if (opt.query) {
    params.push(`$search=${q(`"${opt.query.replace(/"/g, "")}"`)}`);
  } else {
    const filters: string[] = [];
    if (opt.unread_only) filters.push("IsRead eq false");
    if (opt.has_attachments === true) filters.push("HasAttachments eq true");
    else if (opt.has_attachments === false)
      filters.push("HasAttachments eq false");
    const start = normalizeIsoDate(opt.start_date);
    const end = normalizeIsoDate(opt.end_date, true);
    if (start) filters.push(`ReceivedDateTime ge ${start}`);
    if (end) filters.push(`ReceivedDateTime le ${end}`);
    if (filters.length) params.push(`$filter=${q(filters.join(" and "))}`);
    params.push(`$orderby=${q("ReceivedDateTime desc")}`);
  }
  return collect(
    o,
    account,
    base,
    params.join("&"),
    limit,
    maxPages,
    postFilter ? (m) => matchesLocalFilters(m, opt) : undefined,
  );
}

/** Items only, for callers that do not report paging; same paging as listEmailsPage. */
export async function listEmails(
  o: Outlook,
  account: string | undefined,
  opt: ListOptions = {},
): Promise<Obj[]> {
  return (await listEmailsPage(o, account, opt)).items;
}

/** listEmailsPage across all folders unless one is named; query or from_address required. */
export async function searchEmailsPage(
  o: Outlook,
  account: string | undefined,
  opt: ListOptions,
): Promise<EmailList> {
  if (!opt.query && !opt.from_address)
    throw new Error("query or from_address is required");
  return listEmailsPage(o, account, { ...opt, folder: opt.folder || "all" });
}

export async function searchEmails(
  o: Outlook,
  account: string | undefined,
  opt: ListOptions,
): Promise<Obj[]> {
  return (await searchEmailsPage(o, account, opt)).items;
}

export async function listAttachments(
  o: Outlook,
  account: string | undefined,
  messageId: string,
): Promise<Obj[]> {
  const res = await o.request(
    account,
    `me/messages/${q(messageId)}/attachments?${ATTACHMENT_FIELDS}`,
  );
  return (res.value ?? []) as Obj[];
}

export async function getMessage(
  o: Outlook,
  account: string | undefined,
  messageId: string,
): Promise<Obj> {
  if (!messageId) throw new Error("message_id is required");
  const msg = await o.request(account, `me/messages/${q(messageId)}`);
  msg.AttachmentsList = msg.HasAttachments
    ? await listAttachments(o, account, messageId).catch(() => [])
    : [];
  return msg;
}

export async function readThreadPage(
  o: Outlook,
  account: string | undefined,
  threadId: string,
  limit = 50,
): Promise<EmailList> {
  if (!threadId) throw new Error("thread_id or message_id is required");
  let cid = threadId;
  if (!threadId.startsWith("AAQk")) {
    try {
      const m = await o.request(
        account,
        `me/messages/${q(threadId)}?$select=ConversationId`,
      );
      if (m.ConversationId) cid = m.ConversationId;
    } catch {
      // not a message id: used as a conversation id
    }
  }
  const fields =
    "Id,Subject,From,Sender,ToRecipients,CcRecipients,BccRecipients,ReplyTo,ReceivedDateTime,SentDateTime,CreatedDateTime,Body,BodyPreview,HasAttachments,IsRead,IsDraft,Importance,Flag,ConversationId";
  account = (await o.account(account)).upn;
  const top = clamp(limit, 50, PAGE_SIZE);
  const filter = q(`ConversationId eq '${cid.replace(/'/g, "''")}'`);
  const page = await collect(
    o,
    account,
    "me/messages",
    `$filter=${filter}&$top=${top}&$select=${fields}`,
    top,
    DEFAULT_MAX_PAGES,
  );
  const { items } = page;
  const at = (m: Obj) =>
    String(m.ReceivedDateTime ?? m.SentDateTime ?? m.CreatedDateTime ?? "");
  const msgs = items.sort((a, b) =>
    at(a) < at(b) ? -1 : at(a) > at(b) ? 1 : 0,
  );
  let nextAttachment = 0;
  await Promise.all(
    Array.from({ length: Math.min(3, msgs.length) }, async () => {
      for (;;) {
        const i = nextAttachment++;
        if (i >= msgs.length) return;
        const m = msgs[i]!;
        m.AttachmentsList = m.HasAttachments
          ? await listAttachments(o, account, m.Id).catch(() => [])
          : [];
      }
    }),
  );
  return { ...page, items: msgs };
}

export async function readThread(
  o: Outlook,
  account: string | undefined,
  threadId: string,
  limit = 50,
): Promise<Obj[]> {
  return (await readThreadPage(o, account, threadId, limit)).items;
}

export async function getAttachment(
  o: Outlook,
  account: string | undefined,
  messageId: string,
  attachmentId: string,
  expandItem = false,
): Promise<Obj> {
  let p = `me/messages/${q(messageId)}/attachments/${q(attachmentId)}`;
  if (expandItem) p += "?$expand=Microsoft.OutlookServices.ItemAttachment/Item";
  return o.request(account, p);
}
