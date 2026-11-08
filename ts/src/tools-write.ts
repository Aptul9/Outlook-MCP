// MCP tools that change a mailbox or write files: send, forward, manage, attachments (port of outlook_mcp.py).

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import type { CallToolResult, McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import type { Obj, Outlook } from "./api.ts";
import { messageFull } from "./format.ts";
import { getAttachment, listAttachments } from "./mail.ts";
import { MANAGE_ACTIONS, manageEmail, sendEmail, type SendOptions, type SendResult } from "./send.ts";
import { account, run, THIRD_PARTY } from "./tools-read.ts";

const SENDS = { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true };
export const DOWNLOAD_DIR = path.join(process.env.LOCALAPPDATA || path.join(os.homedir(), ".local", "share"), "OutlookMCP", "downloads");

const recipients = z.union([z.string().max(5000), z.array(z.string().max(500)).max(200)]);
/** Remote server: files travel in the request (http.ts caps the request body at 4 MiB). */
const inlineFiles = z
  .array(z.object({ name: z.string().min(1).max(255), content_base64: z.string().min(1).max(4_000_000), content_type: z.string().max(200).optional() }))
  .max(20)
  .optional()
  .describe("Files to attach, base64 content: this server cannot read the caller's disk");
const localFields = {
  account,
  to: recipients.optional().describe("Address or list; on a reply it replaces the To of the thread"),
  cc: recipients.optional(),
  bcc: recipients.optional(),
  subject: z.string().max(1000).optional(),
  body: z.string().max(2_000_000).optional().describe("Plain text or HTML"),
  is_html: z.boolean().optional().describe("Body is HTML (detected from tags when omitted)"),
  importance: z.enum(["Low", "Normal", "High"]).optional(),
  attachments: z.array(z.string().max(1000)).max(50).optional().describe("Local file paths to attach"),
  save_as_draft: z.boolean().optional().describe("Keep it in Drafts instead of sending"),
};
const { attachments: _localPaths, ...common } = localFields;
const remoteFields = { ...common, files: inlineFiles };

function sendText(upn: string, r: SendResult, label: string): string {
  const parts = [`${label} ${r.status === "draft" ? "saved as draft" : "successfully sent"} from ${upn}.`];
  if (r.subject) parts.push(`Subject: ${r.subject}`);
  parts.push(`To: [${r.to.join(", ") || "-"}]`);
  if (r.cc.length) parts.push(`Cc: [${r.cc.join(", ")}]`);
  if (r.bcc.length) parts.push(`Bcc: [${r.bcc.join(", ")}]`);
  if (r.draft_id) parts.push(`Draft ID: ${r.draft_id}`);
  return parts.join(" | ");
}

async function doSend(o: Outlook, a: SendOptions & { account?: string }): Promise<string> {
  const who = await o.account(a.account);
  if (a.reply_to_id) {
    if (!a.body?.trim() && !a.save_as_draft) return "body comment is required when replying";
    const r = await sendEmail(o, who.upn, a);
    return sendText(who.upn, r, `Reply (${a.reply_all ? "reply-all" : "selective/single reply"} to ${a.reply_to_id.slice(0, 24)}...)`);
  }
  if (a.forward_to_id) {
    if (!a.to && !a.save_as_draft) return "'to' recipient is required when forwarding an email";
    return sendText(who.upn, await sendEmail(o, who.upn, a), `Forward (of ${a.forward_to_id.slice(0, 24)}...)`);
  }
  if (!a.to && !a.save_as_draft) return "'to' recipient is required for sending new email";
  if (!a.body?.trim() && !a.subject?.trim()) return "Cannot send empty email (both subject and body are empty)";
  return sendText(who.upn, await sendEmail(o, who.upn, a), "Email");
}

export const safeFilename = (name: string | undefined, fallback: string) =>
  ((name || fallback).trim().replace(/[\\/:*?"<>|\r\n\t]/g, "_").replace(/^[. ]+|[. ]+$/g, "") || fallback).slice(0, 150);

/** A free name in dir: "a.pdf", then "a (2).pdf"; an existing file is never overwritten. */
function freePath(dir: string, name: string): string {
  const ext = path.extname(name);
  const stem = name.slice(0, name.length - ext.length);
  for (let i = 1; ; i++) {
    const p = path.join(dir, i === 1 ? name : `${stem} (${i})${ext}`);
    if (!fs.existsSync(p)) return p;
  }
}

async function attachmentsTool(
  o: Outlook,
  a: { account?: string; message_id: string; attachment_id?: string; index?: number; save_dir?: string },
  remote: boolean,
): Promise<CallToolResult | string> {
  const who = await o.account(a.account);
  const atts = await listAttachments(o, who.upn, a.message_id);
  if (!atts.length) return `No attachments on message ${a.message_id}.`;
  let picked: [number, Obj][];
  if (a.attachment_id) {
    picked = atts.map((x, i): [number, Obj] => [i + 1, x]).filter(([, x]) => x.Id === a.attachment_id);
    if (!picked.length) return `attachment_id not found on message ${a.message_id}; it has ${atts.length} attachment(s).`;
  } else if (a.index !== undefined) {
    if (a.index < 1 || a.index > atts.length) return `index ${a.index} out of range: message ${a.message_id} has ${atts.length} attachment(s).`;
    picked = [[a.index, atts[a.index - 1]!]];
  } else picked = atts.map((x, i): [number, Obj] => [i + 1, x]);
  const dir = a.save_dir || DOWNLOAD_DIR;
  if (!remote) fs.mkdirSync(dir, { recursive: true });
  const lines = [`${atts.length} attachment(s) on message ${a.message_id}:`];
  const blocks: CallToolResult["content"] = [];
  for (const [n, meta] of picked) {
    const full = await getAttachment(o, who.upn, a.message_id, meta.Id);
    const type = String(full["@odata.type"] ?? meta["@odata.type"] ?? "").toLowerCase();
    const name = String(full.Name ?? meta.Name ?? `attachment-${n}`);
    const mime = String(full.ContentType ?? meta.ContentType ?? "application/octet-stream").split(";")[0]!.trim();
    if (type.includes("itemattachment")) {
      const item = (await getAttachment(o, who.upn, a.message_id, meta.Id, true).catch(() => ({}) as Obj)).Item as Obj | undefined;
      if (!item) {
        lines.push(`  #${n} "${name}": embedded Outlook item (email/event).`);
        continue;
      }
      const md = messageFull(item);
      if (remote) {
        lines.push(`  #${n} "${name}": embedded Outlook item as Markdown:`, md);
        continue;
      }
      const p = freePath(dir, `${safeFilename(name, `item-${n}`)}.md`);
      fs.writeFileSync(p, md, "utf-8");
      lines.push(`  #${n} "${name}": embedded Outlook item saved as Markdown to ${p}:`, md);
      continue;
    }
    if (type.includes("referenceattachment")) {
      lines.push(`  #${n} "${name}": link to a cloud file (OneDrive/SharePoint), no bytes on the message. ${full.SourceUrl ?? ""}`.trimEnd());
      continue;
    }
    if (!full.ContentBytes) {
      lines.push(`  #${n} "${name}": no content bytes returned (type ${type || "unknown"}).`);
      continue;
    }
    const data = Buffer.from(String(full.ContentBytes), "base64");
    const size = `${(data.length / 1024).toFixed(1)} KB, ${mime}`;
    if (remote) {
      // the bytes go back in the result: the server's disk is not the caller's
      const uri = `attachment:///${encodeURIComponent(safeFilename(name, `attachment-${n}`))}`;
      lines.push(`  #${n} "${name}" (${size}): content block ${blocks.length + 2}`);
      blocks.push(
        mime.startsWith("image/")
          ? { type: "image", data: String(full.ContentBytes), mimeType: mime }
          : { type: "resource", resource: { uri, mimeType: mime, blob: String(full.ContentBytes) } },
      );
      continue;
    }
    const p = freePath(dir, safeFilename(name, `attachment-${n}`));
    fs.writeFileSync(p, data);
    lines.push(`  #${n} "${name}" saved: ${p} (${size})`);
    if (mime.startsWith("image/")) blocks.push({ type: "image", data: String(full.ContentBytes), mimeType: mime });
  }
  return { content: [{ type: "text", text: lines.join("\n") }, ...blocks] };
}

/** remote: the server runs elsewhere than its caller (http.ts): attachments travel as base64, never as local paths. */
export function registerWriteTools(server: McpServer, o: Outlook, remote = false) {
  registerAttachmentTool(server, o, remote);
  const sendFields = remote ? remoteFields : localFields;
  server.registerTool(
    "send_email",
    {
      title: "Send, reply or forward",
      description:
        `Sends a new email, replies (reply_to_id, reply_all) or forwards (forward_to_id), or saves any of them as a draft. HTML or text body, cc/bcc, selective reply (to, replace_cc, exclude_recipients), ${remote ? "attachments as base64 files" : "local file attachments"}. Sends immediately unless save_as_draft.`,
      inputSchema: z.object({
        ...sendFields,
        reply_to_id: z.string().max(500).optional().describe("Message id to reply to; keeps the thread"),
        reply_all: z.boolean().optional(),
        replace_cc: z.boolean().optional().describe("On reply-all, cc replaces the thread Cc instead of adding to it"),
        exclude_recipients: recipients.optional().describe("Addresses to drop from To/Cc/Bcc of a reply"),
        forward_to_id: z.string().max(500).optional().describe("Message id to forward, with its attachments"),
      }),
      annotations: SENDS,
    },
    (a) => run(() => doSend(o, a)),
  );
  server.registerTool(
    "forward_email",
    {
      title: "Forward an email",
      description: `Forwards an email with its attachments and quoted text, with an optional comment, cc/bcc and extra ${remote ? "base64 files" : "local attachments"}. Sends immediately unless save_as_draft.`,
      inputSchema: z.object({ ...sendFields, message_id: z.string().min(1).max(500), to: recipients }),
      annotations: SENDS,
    },
    ({ message_id, ...a }) => run(() => doSend(o, { ...a, forward_to_id: message_id })),
  );
  server.registerTool(
    "manage_email",
    {
      title: "Manage emails",
      description: "Marks emails read or unread, flags or unflags them, archives them, moves them to Deleted Items (delete) or to a folder (move).",
      inputSchema: z.object({
        account,
        message_id: z.union([z.string().min(1).max(500), z.array(z.string().min(1).max(500)).min(1).max(100)]),
        action: z.enum(MANAGE_ACTIONS),
        destination_folder: z.string().max(200).optional().describe("Folder for move"),
      }),
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: true, openWorldHint: false },
    },
    (a) =>
      run(async () => {
        const who = await o.account(a.account);
        const res = await manageEmail(o, who.upn, a.message_id, a.action, a.destination_folder);
        return [`Executed '${a.action}' on ${res.length} message(s) for ${who.upn}:`, ...res.map((r) => `- ${String(r.id).slice(0, 32)}...: ${r.status}`)].join("\n");
      }),
  );
}

export function registerAttachmentTool(server: McpServer, o: Outlook, remote = false) {
  const pick = { account, message_id: z.string().min(1).max(500), attachment_id: z.string().max(1000).optional(), index: z.number().int().min(1).optional() };
  server.registerTool(
    "get_attachment",
    {
      title: "Attachments of an email",
      description: remote
        ? `Returns the attachments of an email in the result: images inline, other files as embedded resources (base64), embedded Outlook items as Markdown. Nothing is written to disk. index from 1, or attachment_id; all when both are missing. ${THIRD_PARTY}`
        : `Saves the attachments of an email to disk (never overwriting a file), shows images inline, and turns embedded Outlook items into Markdown. index from 1, or attachment_id; all when both are missing. ${THIRD_PARTY}`,
      inputSchema: remote ? z.object(pick) : z.object({ ...pick, save_dir: z.string().max(1000).optional().describe(`Folder to save into (default ${DOWNLOAD_DIR})`) }),
      annotations: remote ? { readOnlyHint: true, openWorldHint: false } : { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
    },
    (a) => run(() => attachmentsTool(o, a, remote)),
  );
}
