// Mail reads: list, search, message, thread, attachments (port of outlook_api.py).

import { q, resolveFolderId, type Obj, type Outlook } from "./api.ts";

export interface ListOptions {
  folder?: string;
  query?: string;
  limit?: number;
  unread_only?: boolean;
  from_address?: string;
  has_attachments?: boolean;
  start_date?: string;
  end_date?: string;
  flagged_only?: boolean;
}

export function normalizeIsoDate(d: string | undefined, endOfDay = false): string | undefined {
  const s = d?.trim();
  if (!s) return undefined;
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return `${s}T${endOfDay ? "23:59:59" : "00:00:00"}Z`;
  const tail = s.slice(10);
  if (!s.endsWith("Z") && !tail.includes("+") && !tail.includes("-")) return `${s}Z`;
  return s;
}

export function matchesLocalFilters(m: Obj, o: ListOptions): boolean {
  if (o.unread_only && m.IsRead !== false) return false;
  if (o.has_attachments === true && !m.HasAttachments) return false;
  if (o.has_attachments === false && m.HasAttachments) return false;
  if (o.flagged_only && String(m.Flag?.FlagStatus ?? "").toLowerCase() !== "flagged") return false;
  if (o.from_address) {
    const s = (m.From ?? m.Sender ?? {}).EmailAddress ?? {};
    if (!`${s.Name ?? ""} ${s.Address ?? ""}`.toLowerCase().includes(o.from_address.toLowerCase())) return false;
  }
  const at = String(m.ReceivedDateTime ?? m.CreatedDateTime ?? "");
  const start = normalizeIsoDate(o.start_date);
  const end = normalizeIsoDate(o.end_date, true);
  if (start && at && at < start) return false;
  if (end && at && at > end) return false;
  return true;
}

const SUMMARY_FIELDS =
  "Id,Subject,From,Sender,ToRecipients,CcRecipients,ReceivedDateTime,CreatedDateTime,BodyPreview,HasAttachments,IsRead,IsDraft,Importance,Flag,ConversationId,ParentFolderId";
const ATTACHMENT_FIELDS = "$select=Id,Name,ContentType,Size,IsInline";

export async function listEmails(o: Outlook, account: string | undefined, opt: ListOptions = {}): Promise<Obj[]> {
  const limit = Math.max(1, Math.min(Math.trunc(opt.limit ?? 20) || 20, 100));
  const folder = opt.folder ?? "Inbox";
  let base = "me/messages";
  if (!["all", "*", "any"].includes(folder.trim().toLowerCase())) {
    const fid = await resolveFolderId(o, account, folder);
    if (fid) base = `me/MailFolders/${q(fid)}/messages`;
  }
  const postFilter = !!(opt.from_address || opt.flagged_only || (opt.query && (opt.unread_only || opt.has_attachments !== undefined || opt.start_date || opt.end_date)));
  const top = postFilter ? Math.min(Math.max(limit * 4, 50), 100) : limit;
  const params = [`$top=${top}`, `$select=${SUMMARY_FIELDS}`];
  if (opt.query) {
    params.push(`$search=${q(`"${opt.query.replace(/"/g, "")}"`)}`);
  } else {
    const filters: string[] = [];
    if (opt.unread_only) filters.push("IsRead eq false");
    if (opt.has_attachments === true) filters.push("HasAttachments eq true");
    else if (opt.has_attachments === false) filters.push("HasAttachments eq false");
    const start = normalizeIsoDate(opt.start_date);
    const end = normalizeIsoDate(opt.end_date, true);
    if (start) filters.push(`ReceivedDateTime ge ${start}`);
    if (end) filters.push(`ReceivedDateTime le ${end}`);
    if (filters.length) params.push(`$filter=${q(filters.join(" and "))}`);
    params.push(`$orderby=${q("ReceivedDateTime desc")}`);
  }
  const res = await o.request(account, `${base}?${params.join("&")}`);
  let items = (res.value ?? []) as Obj[];
  if (postFilter) items = items.filter((m) => matchesLocalFilters(m, opt));
  return items.slice(0, limit);
}

export async function searchEmails(o: Outlook, account: string | undefined, opt: ListOptions): Promise<Obj[]> {
  if (!opt.query && !opt.from_address) throw new Error("query or from_address is required");
  return listEmails(o, account, { ...opt, folder: opt.folder || "all" });
}

export async function listAttachments(o: Outlook, account: string | undefined, messageId: string): Promise<Obj[]> {
  const res = await o.request(account, `me/messages/${q(messageId)}/attachments?${ATTACHMENT_FIELDS}`);
  return (res.value ?? []) as Obj[];
}

export async function getMessage(o: Outlook, account: string | undefined, messageId: string): Promise<Obj> {
  if (!messageId) throw new Error("message_id is required");
  const msg = await o.request(account, `me/messages/${q(messageId)}`);
  msg.AttachmentsList = msg.HasAttachments ? await listAttachments(o, account, messageId).catch(() => []) : [];
  return msg;
}

export async function readThread(o: Outlook, account: string | undefined, threadId: string, limit = 50): Promise<Obj[]> {
  if (!threadId) throw new Error("thread_id or message_id is required");
  let cid = threadId;
  if (!threadId.startsWith("AAQk")) {
    try {
      const m = await o.request(account, `me/messages/${q(threadId)}?$select=ConversationId`);
      if (m.ConversationId) cid = m.ConversationId;
    } catch {
      // not a message id: used as a conversation id
    }
  }
  const fields =
    "Id,Subject,From,Sender,ToRecipients,CcRecipients,BccRecipients,ReplyTo,ReceivedDateTime,SentDateTime,CreatedDateTime,Body,BodyPreview,HasAttachments,IsRead,IsDraft,Importance,Flag,ConversationId";
  const top = Math.max(1, Math.min(Math.trunc(limit) || 50, 100));
  const filter = q(`ConversationId eq '${cid.replace(/'/g, "''")}'`);
  const res = await o.request(account, `me/messages?$filter=${filter}&$top=${top}&$select=${fields}`);
  const at = (m: Obj) => String(m.ReceivedDateTime ?? m.SentDateTime ?? m.CreatedDateTime ?? "");
  const msgs = ((res.value ?? []) as Obj[]).sort((a, b) => (at(a) < at(b) ? -1 : at(a) > at(b) ? 1 : 0));
  for (const m of msgs) m.AttachmentsList = m.HasAttachments ? await listAttachments(o, account, m.Id).catch(() => []) : [];
  return msgs;
}

export async function getAttachment(o: Outlook, account: string | undefined, messageId: string, attachmentId: string, expandItem = false): Promise<Obj> {
  let p = `me/messages/${q(messageId)}/attachments/${q(attachmentId)}`;
  if (expandItem) p += "?$expand=Microsoft.OutlookServices.ItemAttachment/Item";
  return o.request(account, p);
}
