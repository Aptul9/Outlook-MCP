// Text shapes of messages for the tools (port of the formatting helpers in outlook_mcp.py).

import { htmlToMarkdown } from "./html.ts";
import type { Obj } from "./api.ts";

/** Local wall time "YYYY-MM-DD HH:MM" of an ISO instant; Outlook REST sends UTC without a zone, so a missing zone means UTC. */
export function fmtTime(iso: unknown): string {
  if (typeof iso !== "string" || !iso) return "-";
  const s = /(Z|[+-]\d{2}:?\d{2})$/i.test(iso) ? iso : `${iso}Z`;
  const d = new Date(s);
  if (Number.isNaN(d.getTime())) return iso.slice(0, 16).replace("T", " ");
  const p = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

export function recipientStr(r: unknown): string {
  if (!r || typeof r !== "object") return r ? String(r) : "?";
  const e = ((r as Obj).EmailAddress ?? r) as Obj;
  const name = e.Name as string | undefined;
  const address = e.Address as string | undefined;
  if (name && address && name !== address) return `${name} <${address}>`;
  return address || name || "?";
}

export const recipientsStr = (recs: unknown) => (Array.isArray(recs) && recs.length ? recs.map(recipientStr).join(", ") : "-");

export function messageSummary(m: Obj): string {
  const preview = String(m.BodyPreview ?? "").replace(/[\r\n]+/g, " ").trim();
  const tags = [
    m.IsRead === false ? " [UNREAD]" : "",
    m.IsDraft === true ? " [DRAFT]" : "",
    m.HasAttachments ? " [ATTACHMENTS]" : "",
    String(m.Flag?.FlagStatus ?? "").toLowerCase() === "flagged" ? " [FLAGGED]" : "",
    String(m.Importance ?? "").toLowerCase() === "high" ? " [HIGH]" : "",
  ].join("");
  const lines = [`[${fmtTime(m.ReceivedDateTime ?? m.CreatedDateTime)}]${tags} From: ${recipientStr(m.From ?? m.Sender)}`];
  if (m.ToRecipients?.length) lines.push(`To: ${recipientsStr(m.ToRecipients)}`);
  if (m.CcRecipients?.length) lines.push(`Cc: ${recipientsStr(m.CcRecipients)}`);
  lines.push(
    `Subject: ${m.Subject || "(No Subject)"}`,
    `Preview: ${preview.length > 140 ? `${preview.slice(0, 140)}...` : preview}`,
    `Id: ${m.Id}`,
    `ThreadId: ${m.ConversationId}`,
  );
  return lines.join("\n");
}

export function messageFull(m: Obj): string {
  const body = (m.Body ?? {}) as Obj;
  const raw = String(body.Content ?? m.BodyPreview ?? "");
  const md = String(body.ContentType ?? "HTML").toUpperCase() === "HTML" ? htmlToMarkdown(raw) : raw.trim();
  const status: string[] = [];
  if (m.IsDraft === true) status.push("DRAFT");
  if (m.IsRead === false) status.push("UNREAD");
  const imp = String(m.Importance ?? "Normal");
  if (imp.toLowerCase() !== "normal") status.push(`Importance: ${imp}`);
  const h = ["==================================================", `From: ${recipientStr(m.From ?? m.Sender)}`, `To: ${recipientsStr(m.ToRecipients)}`];
  if (m.CcRecipients?.length) h.push(`Cc: ${recipientsStr(m.CcRecipients)}`);
  if (m.BccRecipients?.length) h.push(`Bcc: ${recipientsStr(m.BccRecipients)}`);
  if (m.ReplyTo?.length) h.push(`Reply-To: ${recipientsStr(m.ReplyTo)}`);
  h.push(`Date: ${fmtTime(m.ReceivedDateTime ?? m.SentDateTime ?? m.CreatedDateTime)}`);
  if (status.length) h.push(`Status: ${status.join(", ")}`);
  h.push(`Subject: ${m.Subject || "(No Subject)"}`, `Message ID: ${m.Id}`, `Thread ID: ${m.ConversationId}`);
  const atts = (m.AttachmentsList ?? []) as Obj[];
  if (atts.length) {
    h.push("Attachments:");
    atts.forEach((a, i) => {
      const kb = a.Size ? Math.max(1, Math.floor(a.Size / 1024)) : 0;
      h.push(`  #${i + 1} ${a.Name} (${kb} KB, ${a.ContentType}${a.IsInline ? ", inline" : ""}) [attachment_id: ${a.Id ?? ""}]`);
    });
  }
  h.push("--------------------------------------------------");
  return `${h.join("\n")}\n\n${md}`;
}
